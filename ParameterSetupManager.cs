using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace peilin
{
    public enum ParameterCategory
    {
        Camera,     // 相機參數
        Position,   // 位置參數  
        Detection,  // 檢測參數
        Timing,     // 時間參數
        Testing     // 測試驗證
    }

    public enum ParameterStatus
    {
        NotStarted,      // 未開始
        InProgress,      // 進行中
        Completed,       // 已完成
        RequiresExternal // 需要外部工具
    }

    public class ParameterCategoryCompletedEventArgs : EventArgs
    {
        public ParameterCategory Category { get; set; }
        public ParameterStatus Status { get; set; }
    }

    public class ParameterSetupManager
    {
        private Dictionary<ParameterCategory, ParameterStatus> categoryStatus;
        private Dictionary<ParameterCategory, int> categoryProgress;
        private Dictionary<ParameterCategory, int> categoryTotal;

        public event EventHandler<ParameterCategoryCompletedEventArgs> CategoryStatusChanged;

        public ParameterSetupManager()
        {
            InitializeStatus();
        }

        private void InitializeStatus()
        {
            categoryStatus = new Dictionary<ParameterCategory, ParameterStatus>
            {
                { ParameterCategory.Camera, ParameterStatus.NotStarted },
                { ParameterCategory.Position, ParameterStatus.RequiresExternal },
                { ParameterCategory.Detection, ParameterStatus.NotStarted },
                { ParameterCategory.Timing, ParameterStatus.NotStarted },
                { ParameterCategory.Testing, ParameterStatus.NotStarted }
            };

            categoryProgress = new Dictionary<ParameterCategory, int>();
            categoryTotal = new Dictionary<ParameterCategory, int>();

            foreach (ParameterCategory category in Enum.GetValues(typeof(ParameterCategory)))
            {
                categoryProgress[category] = 0;
                categoryTotal[category] = 0;
            }
        }

        public ParameterStatus GetCategoryStatus(ParameterCategory category)
        {
            return categoryStatus.ContainsKey(category) ? categoryStatus[category] : ParameterStatus.NotStarted;
        }

        public void UpdateCategoryStatus(ParameterCategory category, ParameterStatus status)
        {
            if (categoryStatus.ContainsKey(category))
            {
                categoryStatus[category] = status;
                CategoryStatusChanged?.Invoke(this, new ParameterCategoryCompletedEventArgs
                {
                    Category = category,
                    Status = status
                });
            }
        }
        // 【新增方法】：基於資料庫檢查分類完成狀態
        // 【同時修正】：UpdateCategoryStatusFromDatabase 方法也要用相同邏輯
        public void UpdateCategoryStatusFromDatabase(ParameterCategory category, string targetType)
        {
            try
            {
                using (var db = new MydbDB())
                {
                    bool hasCompletedParameters = false;

                    switch (category)
                    {
                        case ParameterCategory.Camera:
                            hasCompletedParameters = db.Cameras.Any(c => c.Type == targetType &&
                                (c.Name.Equals("exposure", StringComparison.OrdinalIgnoreCase) ||
                                 c.Name.Equals("gain", StringComparison.OrdinalIgnoreCase) ||
                                 c.Name.Equals("delay", StringComparison.OrdinalIgnoreCase)));
                            break;

                        case ParameterCategory.Position:
                            hasCompletedParameters = db.@params.Any(p => p.Type == targetType &&
                                (p.Name.Contains("center") || p.Name.Contains("radius") ||
                                 p.Name.Contains("chamfer") || p.Name.Contains("position") ||
                                 p.Name.Contains("objBias")));
                            break;

                        case ParameterCategory.Timing:
                            hasCompletedParameters = db.@params.Any(p => p.Type == targetType &&
                                (p.Name.Contains("time") || p.Name.Contains("fourTo")) &&
                                !p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase));
                            break;

                        case ParameterCategory.Detection:
                            // 【新邏輯】檢測參數：排除法
                            hasCompletedParameters = db.@params.Any(p => p.Type == targetType &&
                                // 不是位置參數
                                !(p.Name.Contains("center") || p.Name.Contains("radius") ||
                                  p.Name.Contains("chamfer") || p.Name.Contains("position") ||
                                  p.Name.Contains("objBias")) &&
                                // 不是時間參數  
                                !(p.Name.Contains("time") || p.Name.Contains("fourTo")));
                            break;
                    }

                    if (hasCompletedParameters)
                    {
                        UpdateCategoryStatus(category, ParameterStatus.Completed);
                    }
                    else if (category == ParameterCategory.Position)
                    {
                        UpdateCategoryStatus(category, ParameterStatus.RequiresExternal);
                    }
                    else
                    {
                        UpdateCategoryStatus(category, ParameterStatus.NotStarted);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"檢查分類狀態時發生錯誤: {ex.Message}");
            }
        }

        public void UpdateCategoryProgress(ParameterCategory category, int completed, int total)
        {
            categoryProgress[category] = completed;
            categoryTotal[category] = total;

            // 自動更新狀態
            if (total > 0)
            {
                if (completed == 0)
                {
                    if (category == ParameterCategory.Position)
                        UpdateCategoryStatus(category, ParameterStatus.RequiresExternal);
                    else
                        UpdateCategoryStatus(category, ParameterStatus.NotStarted);
                }
                else if (completed == total)
                {
                    UpdateCategoryStatus(category, ParameterStatus.Completed);
                }
                else
                {
                    UpdateCategoryStatus(category, ParameterStatus.InProgress);
                }
            }
        }

        public bool CanEnableTab(ParameterCategory category)
        {
            switch (category)
            {
                case ParameterCategory.Camera:
                    return true; // 相機參數可以直接設定

                case ParameterCategory.Position:
                    return true; // 位置參數可以設定（但需要外部工具）

                case ParameterCategory.Detection:
                    // 檢測參數可直接設定
                    //return GetCategoryStatus(ParameterCategory.Position) == ParameterStatus.Completed;
                    return true;

                case ParameterCategory.Timing:
                    // 時間參數可直接設定
                    //return GetCategoryStatus(ParameterCategory.Position) == ParameterStatus.Completed;
                    return true;

                case ParameterCategory.Testing:
                    // 測試需要位置和檢測都完成
                    return GetCategoryStatus(ParameterCategory.Position) == ParameterStatus.Completed &&
                           (GetCategoryStatus(ParameterCategory.Detection) == ParameterStatus.Completed ||
                            GetCategoryStatus(ParameterCategory.Timing) == ParameterStatus.Completed);

                default:
                    return false;
            }
        }

        public int GetOverallProgress()
        {
            int totalCompleted = categoryProgress.Values.Sum();
            int totalCount = categoryTotal.Values.Sum();

            return totalCount > 0 ? (totalCompleted * 100) / totalCount : 0;
        }

        public string GetProgressSummary()
        {
            var summary = new List<string>();

            foreach (var category in categoryStatus.Keys)
            {
                string icon = GetStatusIcon(categoryStatus[category]);
                int completed = categoryProgress[category];
                int total = categoryTotal[category];
                string categoryName = GetCategoryDisplayName(category);

                summary.Add($"{icon}{categoryName}({completed}/{total})");
            }

            return string.Join(" ", summary);
        }

        private string GetStatusIcon(ParameterStatus status)
        {
            switch (status)
            {
                case ParameterStatus.Completed: return "✅";
                case ParameterStatus.InProgress: return "⚠️";
                case ParameterStatus.RequiresExternal: return "🔧";
                case ParameterStatus.NotStarted:
                default: return "❌";
            }
        }

        private string GetCategoryDisplayName(ParameterCategory category)
        {
            switch (category)
            {
                case ParameterCategory.Camera: return "相機";
                case ParameterCategory.Position: return "位置";
                case ParameterCategory.Detection: return "檢測";
                case ParameterCategory.Timing: return "時間";
                case ParameterCategory.Testing: return "測試";
                default: return category.ToString();
            }
        }

        public void UpdateTabStates(TabControl tabControl)
        {
            if (tabControl.TabPages.Count >= 5)
            {
                // Tab索引對應: 0=相機, 1=位置, 2=檢測, 3=時間, 4=測試
                tabControl.TabPages[0].Enabled = CanEnableTab(ParameterCategory.Camera);
                tabControl.TabPages[1].Enabled = CanEnableTab(ParameterCategory.Position);
                tabControl.TabPages[2].Enabled = CanEnableTab(ParameterCategory.Detection);
                tabControl.TabPages[3].Enabled = CanEnableTab(ParameterCategory.Timing);
                tabControl.TabPages[4].Enabled = CanEnableTab(ParameterCategory.Testing);
            }
        }

        public List<ParameterItem> GetParametersByCategory(List<ParameterItem> allParameters, ParameterCategory category)
        {
            switch (category)
            {
                case ParameterCategory.Camera:
                    // 相機參數：從Camera表來的特定參數
                    return allParameters.Where(p =>
                        p.Name.Equals("exposure", StringComparison.OrdinalIgnoreCase) ||
                        p.Name.Equals("gain", StringComparison.OrdinalIgnoreCase) ||
                        p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                case ParameterCategory.Position:
                    // 位置參數：特定字串匹配
                    return allParameters.Where(p =>
                        p.Name.Contains("center") ||
                        p.Name.Contains("radius") ||
                        p.Name.Contains("chamfer") ||
                        p.Name.Contains("position") ||
                        p.Name.Contains("objBias")
                    ).ToList();

                case ParameterCategory.Timing:
                    // 時間參數：特定字串匹配（但排除delay，因為delay是相機參數）
                    return allParameters.Where(p =>
                        (p.Name.Contains("time") || p.Name.Contains("fourTo")) &&
                        !p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                case ParameterCategory.Detection:
                    // 【新邏輯】檢測參數：排除法 = 全部Param參數 - 位置參數 - 時間參數
                    return allParameters.Where(p =>
                        // 不是相機參數
                        !(p.Name.Equals("exposure", StringComparison.OrdinalIgnoreCase) ||
                          p.Name.Equals("gain", StringComparison.OrdinalIgnoreCase) ||
                          p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase)) &&
                        // 不是位置參數
                        !(p.Name.Contains("center") ||
                          p.Name.Contains("radius") ||
                          p.Name.Contains("chamfer") ||
                          p.Name.Contains("position") ||
                          p.Name.Contains("objBias")) &&
                        // 不是時間參數
                        !(p.Name.Contains("time") || p.Name.Contains("fourTo"))
                    ).ToList();

                case ParameterCategory.Testing:
                    return new List<ParameterItem>(); // 測試頁面不顯示參數

                default:
                    return new List<ParameterItem>();
            }
        }
    }
}