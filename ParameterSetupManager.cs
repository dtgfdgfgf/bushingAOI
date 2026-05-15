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
        public Dictionary<ParameterCategory, int> categoryProgress;
        public Dictionary<ParameterCategory, int> categoryTotal;

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
                                 p.Name.Contains("chamfer") || p.Name.Contains("position")));
                            break;

                        case ParameterCategory.Timing:
                            hasCompletedParameters = db.@params.Any(p => p.Type == targetType &&
                                (p.Name.Contains("time") || p.Name.Contains("fourTo")) &&
                                !p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase));
                            break;

                        case ParameterCategory.Detection:
                            // 【新邏輯】檢測參數：排除法，加入objBias
                            hasCompletedParameters = db.@params.Any(p => p.Type == targetType &&
                                // 不是位置參數
                                !(p.Name.Contains("center") || p.Name.Contains("radius") ||
                                  p.Name.Contains("chamfer") || p.Name.Contains("position")) &&
                                // 不是時間參數  
                                !(p.Name.Contains("time") || p.Name.Contains("fourTo")) &&
                                // 包含objBias參數
                                (p.Name.Contains("objBias") || 
                                 p.Name.Contains("threshold") || p.Name.Contains("detect") ||
                                 p.Name.Contains("contour") || p.Name.Contains("white") ||
                                 p.Name.Contains("gap")));
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

        // 修正 ParameterSetupManager 中的進度計算
        // 修正 ParameterSetupManager 中的進度計算
        public void UpdateCategoryProgress(ParameterCategory category, int completed, int total)
        {
            categoryProgress[category] = completed;
            categoryTotal[category] = total;

            // 🔸 輸出除錯訊息
            System.Diagnostics.Debug.WriteLine($"更新進度 {category}: {completed}/{total}");

            // 🔸 修正：自動更新狀態邏輯
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
            else
            {
                // 🔸 如果沒有參考區參數，設為未開始
                if (category == ParameterCategory.Position)
                    UpdateCategoryStatus(category, ParameterStatus.RequiresExternal);
                else
                    UpdateCategoryStatus(category, ParameterStatus.NotStarted);
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

        // 修正整體進度計算
        public int GetOverallProgress()
        {
            int totalCompleted = 0;
            int totalCount = 0;

            // 🔸 修正：只計算有參數的類別
            foreach (var category in categoryProgress.Keys)
            {
                if (categoryTotal.ContainsKey(category) && categoryTotal[category] > 0)
                {
                    totalCompleted += categoryProgress[category];
                    totalCount += categoryTotal[category];
                }
            }

            return totalCount > 0 ? (int)Math.Round((double)totalCompleted / totalCount * 100) : 0;
        }

        // 修正進度摘要顯示
        public string GetProgressSummary()
        {
            var summary = new List<string>();

            foreach (var category in categoryStatus.Keys)
            {
                string icon = GetStatusIcon(categoryStatus[category]);
                int completed = categoryProgress.ContainsKey(category) ? categoryProgress[category] : 0;
                int total = categoryTotal.ContainsKey(category) ? categoryTotal[category] : 0;
                string categoryName = GetCategoryDisplayName(category);

                // 🔸 修正：只顯示有參數的類別，避免 (0/0) 的情況
                if (total > 0 || category == ParameterCategory.Position)
                {
                    summary.Add($"{icon}{categoryName}({completed}/{total})");
                }
            }

            return string.Join(" ", summary);
        }
        public int GetCategoryProgressPercentage(ParameterCategory category)
        {
            if (!categoryProgress.ContainsKey(category) || !categoryTotal.ContainsKey(category))
                return 0;

            int completed = categoryProgress[category];
            int total = categoryTotal[category];

            if (total == 0)
                return 0;

            return (int)Math.Round((double)completed / total * 100);
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
                    // 位置參數：特定字串匹配（移除objBias）
                    return allParameters.Where(p =>
                        p.Name.Contains("center") ||
                        p.Name.Contains("radius") ||
                        p.Name.Contains("chamfer") ||
                        p.Name.Contains("position")
                    ).ToList();

                case ParameterCategory.Timing:
                    // 時間參數：特定字串匹配（但排除delay，因為delay是相機參數）
                    return allParameters.Where(p =>
                        (p.Name.Contains("time") || p.Name.Contains("fourTo")) &&
                        !p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                case ParameterCategory.Detection:
                    // 【新邏輯】檢測參數：排除法，包含objBias
                    return allParameters.Where(p =>
                        // 不是相機參數
                        !(p.Name.Equals("exposure", StringComparison.OrdinalIgnoreCase) ||
                          p.Name.Equals("gain", StringComparison.OrdinalIgnoreCase) ||
                          p.Name.Equals("delay", StringComparison.OrdinalIgnoreCase)) &&
                        // 不是位置參數（但包含objBias）
                        !(p.Name.Contains("center") ||
                          p.Name.Contains("radius") ||
                          p.Name.Contains("chamfer") ||
                          p.Name.Contains("position")) &&
                        // 不是時間參數
                        !(p.Name.Contains("time") || p.Name.Contains("fourTo")) &&
                        // 包含objBias或其他檢測相關參數
                        (p.Name.Contains("objBias") ||
                         p.Name.Contains("threshold") || p.Name.Contains("detect") ||
                         p.Name.Contains("contour") || p.Name.Contains("white") ||
                         p.Name.Contains("gap") || p.Name.Contains("min") || p.Name.Contains("max") ||
                         // 通用檢測參數：如果不是以上任何類別，且不是空白，則歸類為檢測參數
                         (!string.IsNullOrEmpty(p.Name) && 
                          !p.Name.Contains("center") && !p.Name.Contains("radius") &&
                          !p.Name.Contains("position") && !p.Name.Contains("chamfer") &&
                          !p.Name.Contains("time") && !p.Name.Contains("fourTo")))
                    ).ToList();

                case ParameterCategory.Testing:
                    return new List<ParameterItem>(); // 測試頁面不顯示參數

                default:
                    return new List<ParameterItem>();
            }
        }
    }
}