using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using LinqToDB;

namespace peilin
{
    public partial class ParameterConfigForm : Form
    {
        private ParameterSession currentSession;
        private List<ParameterItem> allParameters;
        private ParameterSetupManager parameterManager;
        private string settingPath = Path.Combine(Application.StartupPath, "setting");
        private string sessionFilePath;
        private ParameterCategory currentCategory = ParameterCategory.Camera;
        private Dictionary<ParameterCategory, bool> tabSavedStatus;
        
        // 【新增】參數狀態視覺化
        private Dictionary<string, ParameterStatus> parameterStatusMap = new Dictionary<string, ParameterStatus>();

        // 由 GitHub Copilot 產生 - 參數區域快取（用於加速顏色查詢，避免重複 LINQ 查詢）
        private Dictionary<string, ParameterZone> _parameterZoneCache = new Dictionary<string, ParameterZone>();

        // 【新增】參數狀態列舉
        public enum ParameterStatus
        {
            NotAdded,        // 未新增 - 白色背景
            AddedUnmodified, // 已新增未修改 - 淺綠色背景
            AddedModified    // 已新增已修改 - 淺藍色背景
        }

        public ParameterConfigForm(string targetType)
        {
            TargetType = targetType;
            InitializeComponent();
            CreateTabControls();
            //TargetType = targetType;
            sessionFilePath = Path.Combine(settingPath, $"param_session_{targetType}.json");

            parameterManager = new ParameterSetupManager();
            parameterManager.CategoryStatusChanged += OnCategoryStatusChanged;

            tabSavedStatus = new Dictionary<ParameterCategory, bool>();
            foreach (ParameterCategory category in Enum.GetValues(typeof(ParameterCategory)))
            {
                tabSavedStatus[category] = false;
            }

            InitializeForm();
        }
        #region 建立控件
        // 【新增方法】：建立所有Tab的控件
        private void CreateTabControls()
        {
            CreateCameraTabControls();
            CreatePositionTabControls();
            CreateDetectionTabControls();
            CreateTimingTabControls();
        }
        // 【新增方法】：建立相機參數Tab的控件
        private void CreateCameraTabControls()
        {
            // 建立控件
            this.grpCameraUnmodified = new GroupBox();
            this.grpCameraModified = new GroupBox();
            this.grpCameraFixed = new GroupBox();
            this.dgvCameraUnmodified = new DataGridView();
            this.dgvCameraModified = new DataGridView();
            this.dgvCameraFixed = new DataGridView();
            this.btnCameraMoveToModified = new Button();
            this.btnCameraMoveToFixed = new Button();
            this.btnCameraMoveToUnmodified = new Button();
            this.btnCameraSelectAll = new Button();
            this.btnCameraSelectNone = new Button();
            this.btnCameraSelectInvert = new Button();

            SetupThreeZones(this.tabCamera, "Camera",
                this.grpCameraUnmodified, this.grpCameraModified, this.grpCameraFixed,
                this.dgvCameraUnmodified, this.dgvCameraModified, this.dgvCameraFixed,
                this.btnCameraMoveToModified, this.btnCameraMoveToFixed, this.btnCameraMoveToUnmodified,
                this.btnCameraSelectAll, this.btnCameraSelectNone, this.btnCameraSelectInvert);
        }
        // 【新增方法】：建立位置參數Tab的控件
        private void CreatePositionTabControls()
        {
            // 建立控件
            this.grpPositionUnmodified = new GroupBox();
            this.grpPositionModified = new GroupBox();
            this.grpPositionFixed = new GroupBox();
            this.dgvPositionUnmodified = new DataGridView();
            this.dgvPositionModified = new DataGridView();
            this.dgvPositionFixed = new DataGridView();
            this.btnPositionMoveToModified = new Button();
            this.btnPositionMoveToFixed = new Button();
            this.btnPositionMoveToUnmodified = new Button();
            this.btnPositionSelectAll = new Button();
            this.btnPositionSelectNone = new Button();
            this.btnPositionSelectInvert = new Button();
            this.btnOpenPositionCalibration = new Button();
            this.lblPositionHint = new Label();

            SetupThreeZones(this.tabPosition, "Position",
                this.grpPositionUnmodified, this.grpPositionModified, this.grpPositionFixed,
                this.dgvPositionUnmodified, this.dgvPositionModified, this.dgvPositionFixed,
                this.btnPositionMoveToModified, this.btnPositionMoveToFixed, this.btnPositionMoveToUnmodified,
                this.btnPositionSelectAll, this.btnPositionSelectNone, this.btnPositionSelectInvert);

            // 位置Tab的特殊按鈕
            this.btnOpenPositionCalibration.Text = "🔧 開啟位置校正工具";
            this.btnOpenPositionCalibration.Location = new Point(300, 20);
            this.btnOpenPositionCalibration.Size = new Size(200, 30);
            this.btnOpenPositionCalibration.BackColor = Color.LightBlue;
            this.btnOpenPositionCalibration.Click += btnOpenPositionCalibration_Click;

            this.lblPositionHint.Text = "💡 位置參數需要使用專用的校正工具進行設定";
            this.lblPositionHint.Location = new Point(520, 25);
            this.lblPositionHint.Size = new Size(400, 20);
            this.lblPositionHint.ForeColor = Color.Blue;

            this.tabPosition.Controls.Add(this.btnOpenPositionCalibration);
            this.tabPosition.Controls.Add(this.lblPositionHint);
        }
        // 【新增方法】：建立檢測參數Tab的控件
        private void CreateDetectionTabControls()
        {
            this.grpDetectionUnmodified = new GroupBox();
            this.grpDetectionModified = new GroupBox();
            this.grpDetectionFixed = new GroupBox();
            this.dgvDetectionUnmodified = new DataGridView();
            this.dgvDetectionModified = new DataGridView();
            this.dgvDetectionFixed = new DataGridView();
            this.btnDetectionMoveToModified = new Button();
            this.btnDetectionMoveToFixed = new Button();
            this.btnDetectionMoveToUnmodified = new Button();
            this.btnDetectionSelectAll = new Button();
            this.btnDetectionSelectNone = new Button();
            this.btnDetectionSelectInvert = new Button();

            // 【新增】建立智能校正助手面板
            CreateDetectionCalibrationPanel();

            SetupThreeZones(this.tabDetection, "Detection",
                this.grpDetectionUnmodified, this.grpDetectionModified, this.grpDetectionFixed,
                this.dgvDetectionUnmodified, this.dgvDetectionModified, this.dgvDetectionFixed,
                this.btnDetectionMoveToModified, this.btnDetectionMoveToFixed, this.btnDetectionMoveToUnmodified,
                this.btnDetectionSelectAll, this.btnDetectionSelectNone, this.btnDetectionSelectInvert);
        }
        // 【新增方法】：建立時間參數Tab的控件
        private void CreateTimingTabControls()
        {
            this.grpTimingUnmodified = new GroupBox();
            this.grpTimingModified = new GroupBox();
            this.grpTimingFixed = new GroupBox();
            this.dgvTimingUnmodified = new DataGridView();
            this.dgvTimingModified = new DataGridView();
            this.dgvTimingFixed = new DataGridView();
            this.btnTimingMoveToModified = new Button();
            this.btnTimingMoveToFixed = new Button();
            this.btnTimingMoveToUnmodified = new Button();
            this.btnTimingSelectAll = new Button();
            this.btnTimingSelectNone = new Button();
            this.btnTimingSelectInvert = new Button();

            // 【新增】建立智能計算助手面板
            CreateTimingCalculatorPanel();

            SetupThreeZones(this.tabTiming, "Timing",
                this.grpTimingUnmodified, this.grpTimingModified, this.grpTimingFixed,
                this.dgvTimingUnmodified, this.dgvTimingModified, this.dgvTimingFixed,
                this.btnTimingMoveToModified, this.btnTimingMoveToFixed, this.btnTimingMoveToUnmodified,
                this.btnTimingSelectAll, this.btnTimingSelectNone, this.btnTimingSelectInvert);
        }
        // 【重構方法】：設定三區域佈局（新邏輯：參考區→已新增未修改→已新增已修改）
        private void SetupThreeZones(TabPage tab, string prefix,
            GroupBox grpUnmodified, GroupBox grpModified, GroupBox grpFixed,
            DataGridView dgvUnmodified, DataGridView dgvModified, DataGridView dgvFixed,
            Button btnMoveToModified, Button btnMoveToFixed, Button btnMoveToUnmodified,
            Button btnSelectAll, Button btnSelectNone, Button btnSelectInvert)
        {
            // 視窗1600x1000，Tab可用空間更大，調整為更寬敞的佈局
            // 【重新命名】三個區域
            // 來源料號顯示（若 currentSession.SourceType 有值則顯示）
            string sourceTypeText = currentSession != null && !string.IsNullOrEmpty(currentSession.SourceType)
                ? $"來源料號: {currentSession.SourceType}"
                : "參考區（來源料號）";
            grpUnmodified.Text = sourceTypeText;
            grpUnmodified.Location = new Point(10, 55);
            grpUnmodified.Size = new Size(510, 720);  

            grpModified.Text = "已新增未修改區";
            grpModified.Location = new Point(530, 55);  
            grpModified.Size = new Size(510, 720);  

            grpFixed.Text = "已新增已修改區";
            grpFixed.Location = new Point(1050, 55);  
            grpFixed.Size = new Size(510, 720);  

            // DataGridView設定
            SetupDataGridView(dgvUnmodified);
            SetupDataGridView(dgvModified);
            SetupDataGridView(dgvFixed);

            // 【重新設計】移動按鈕 - 符合新邏輯
            // 參考區 → 已新增未修改區（複製參數）
            btnMoveToModified.Text = "複製到新增區 →";
            btnMoveToModified.Location = new Point(10, 785);
            btnMoveToModified.Size = new Size(130, 30);  
            btnMoveToModified.Tag = prefix;
            btnMoveToModified.Click += btnCopyToAddedUnmodified_Click;

            // 【隱藏】不再需要的按鈕 - 參考區不能直接到已修改區
            btnMoveToFixed.Visible = false; 

            // 【新增】建立缺少的按鈕
            var btnModifiedToFixed = new Button();
            var btnFixedToModified = new Button();

            // 已新增未修改區 → 已新增已修改區（編輯後移動）
            btnModifiedToFixed.Text = "標記為已修改 →";
            btnModifiedToFixed.Location = new Point(530, 785);
            btnModifiedToFixed.Size = new Size(140, 30);
            btnModifiedToFixed.Tag = prefix;
            btnModifiedToFixed.Click += btnMoveToAddedModified_Click;

            // 已新增已修改區 → 已新增未修改區（復原修改）
            btnFixedToModified.Text = "← 復原修改";
            btnFixedToModified.Location = new Point(1050, 785);
            btnFixedToModified.Size = new Size(100, 30);
            btnFixedToModified.Tag = prefix;
            btnFixedToModified.Click += btnRevertToAddedUnmodified_Click;

            // 刪除按鈕（從已新增區域移除，回到未複製狀態）
            btnMoveToUnmodified.Text = "← 移除選取項目";
            btnMoveToUnmodified.Location = new Point(730, 840);
            btnMoveToUnmodified.Size = new Size(140, 30);  
            btnMoveToUnmodified.Tag = prefix;
            btnMoveToUnmodified.Click += btnRemoveFromAdded_Click;

            // 選擇按鈕
            btnSelectAll.Text = "全選";
            btnSelectAll.Location = new Point(10, 20);
            btnSelectAll.Size = new Size(80, 30);
            btnSelectAll.Tag = prefix;
            btnSelectAll.Click += btnSelectAll_Click;
            btnSelectAll.Visible = false;

            btnSelectNone.Text = "清除";
            btnSelectNone.Location = new Point(100, 20);
            btnSelectNone.Size = new Size(80, 30);
            btnSelectNone.Tag = prefix;
            btnSelectNone.Click += btnSelectNone_Click;
            btnSelectNone.Visible = false;

            btnSelectInvert.Text = "反選";
            btnSelectInvert.Location = new Point(190, 20);
            btnSelectInvert.Size = new Size(80, 30);
            btnSelectInvert.Tag = prefix;
            btnSelectInvert.Click += btnSelectInvert_Click;
            btnSelectInvert.Visible = false;

            // 加入控件到GroupBox
            grpUnmodified.Controls.Add(dgvUnmodified);
            grpModified.Controls.Add(dgvModified);
            grpFixed.Controls.Add(dgvFixed);

            // 加入控件到TabPage
            tab.Controls.Add(grpUnmodified);
            tab.Controls.Add(grpModified);
            tab.Controls.Add(grpFixed);
            tab.Controls.Add(btnMoveToModified);
            // tab.Controls.Add(btnMoveToFixed); // 隱藏
            tab.Controls.Add(btnMoveToUnmodified);
            tab.Controls.Add(btnModifiedToFixed);
            tab.Controls.Add(btnFixedToModified);
            tab.Controls.Add(btnSelectAll);
            tab.Controls.Add(btnSelectNone);
            tab.Controls.Add(btnSelectInvert);
        


            // 選擇按鈕
            btnSelectAll.Text = "全選";
            btnSelectAll.Location = new Point(10, 20);
            btnSelectAll.Size = new Size(80, 30);
            btnSelectAll.Tag = prefix;
            btnSelectAll.Click += btnSelectAll_Click;

            btnSelectNone.Text = "清除";
            btnSelectNone.Location = new Point(100, 20);
            btnSelectNone.Size = new Size(80, 30);
            btnSelectNone.Tag = prefix;
            btnSelectNone.Click += btnSelectNone_Click;

            btnSelectInvert.Text = "反選";
            btnSelectInvert.Location = new Point(190, 20);
            btnSelectInvert.Size = new Size(80, 30);
            btnSelectInvert.Tag = prefix;
            btnSelectInvert.Click += btnSelectInvert_Click;

            // 加入控件到GroupBox
            grpUnmodified.Controls.Add(dgvUnmodified);
            grpModified.Controls.Add(dgvModified);
            grpFixed.Controls.Add(dgvFixed);

            // 加入控件到TabPage
            tab.Controls.Add(grpUnmodified);
            tab.Controls.Add(grpModified);
            tab.Controls.Add(grpFixed);
            tab.Controls.Add(btnMoveToModified);
            tab.Controls.Add(btnMoveToFixed);
            //tab.Controls.Add(btnMoveToUnmodified);
            tab.Controls.Add(btnModifiedToFixed);
            tab.Controls.Add(btnFixedToModified);
            tab.Controls.Add(btnSelectAll);
            tab.Controls.Add(btnSelectNone);
            tab.Controls.Add(btnSelectInvert);
        }

        // 由 GitHub Copilot 產生 - 在 SetupDataGridView 加入 commit handler 呼叫
        private void SetupDataGridView(DataGridView dgv)
        {
            dgv.Location = new System.Drawing.Point(8, 28);
            dgv.Size = new System.Drawing.Size(494, 680);
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = true;
            dgv.AutoGenerateColumns = true;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dgv.RowPrePaint += DataGridView_RowPrePaint;
            dgv.DataBindingComplete += DataGridView_DataBindingComplete;
            dgv.CellFormatting += DataGridView_CellFormatting;
            AddCheckboxCommitHandler(dgv); // ★ 新增
        }
        // 由 GitHub Copilot 產生 - CellFormatting 也同步（雙保險）
        private void DataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv?.DataSource is List<ParameterItem> list &&
                e.RowIndex >= 0 && e.RowIndex < list.Count &&
                IsReferenceZoneDataGridView(dgv))
            {
                var item = list[e.RowIndex];
                if (item.Zone == ParameterZone.Reference)
                {
                    UpdateReferenceRowVisual(dgv.Rows[e.RowIndex], item);
                }
            }
        }
        // 由 GitHub Copilot 產生 - 勾選選取框後立即套色（避免看到閃爍或變白）
        private void AddCheckboxCommitHandler(DataGridView dgv)
        {
            dgv.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgv.IsCurrentCellDirty)
                {
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    ApplyReferenceZoneColors();
                }
            };
        }
        // 由 GitHub Copilot 產生 - 進度與顏色修正：使用快取查詢（O(1) 複雜度）
        private System.Drawing.Color GetReferenceRowBackground(ParameterItem refItem)
        {
            string key = $"{refItem.Name}_{refItem.Stop ?? 0}";

            if (_parameterZoneCache.TryGetValue(key, out var zone))
            {
                if (zone == ParameterZone.AddedModified) return System.Drawing.Color.LightBlue;
                if (zone == ParameterZone.AddedUnmodified) return System.Drawing.Color.LightGreen;
            }

            return System.Drawing.Color.White; // 尚未新增
        }

        // 由 GitHub Copilot 產生 - 建立參數區域快取（一次性建立，避免 O(N²) 查詢）
        private void BuildParameterZoneCache()
        {
            _parameterZoneCache.Clear();

            if (allParameters == null) return;

            // 遍歷所有非參考區參數，建立快取
            // 優先記錄 AddedModified（若同時存在則以 Modified 為準）
            foreach (var param in allParameters)
            {
                if (param.Zone == ParameterZone.Reference) continue;

                string key = $"{param.Name}_{param.Stop ?? 0}";

                if (param.Zone == ParameterZone.AddedModified)
                {
                    // AddedModified 優先級最高，直接覆蓋
                    _parameterZoneCache[key] = ParameterZone.AddedModified;
                }
                else if (param.Zone == ParameterZone.AddedUnmodified)
                {
                    // 僅在尚未有記錄時才加入 AddedUnmodified
                    if (!_parameterZoneCache.ContainsKey(key))
                    {
                        _parameterZoneCache[key] = ParameterZone.AddedUnmodified;
                    }
                }
            }
        }
        private void UpdateReferenceRowVisual(DataGridViewRow row, ParameterItem item)
        {
            if (row == null || item == null) return;
            if (item.Zone != ParameterZone.Reference) return;

            System.Drawing.Color baseColor = GetReferenceRowBackground(item);

            // 基礎底色
            row.DefaultCellStyle.BackColor = baseColor;

            // 選取狀態：使用 IsSelected (checkbox 勾選) 來決定顯示
            if (item.IsSelected)
            {
                // 未新增(白底) → 改成淡黃色；其他保持原色但可加邊框提示
                if (baseColor == System.Drawing.Color.White)
                {
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.LightGoldenrodYellow;
                }
                else
                {
                    // 對已新增(綠/藍)加微弱強調色（可選）：
                    row.DefaultCellStyle.BackColor = baseColor; // 不變
                }

                // 讓使用者仍可看見是選取：使用 SelectionBackColor 與 BackColor 區隔
                row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;
            }
            else
            {
                // 未勾選：維持狀態色；選取時不要反白造成干擾
                row.DefaultCellStyle.SelectionBackColor = baseColor;
                row.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;
            }
        }

        #region 參數狀態視覺化
        // 由 GitHub Copilot 產生 - RowPrePaint 保底
        private void DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv?.DataSource is List<ParameterItem> parameters &&
                e.RowIndex >= 0 && e.RowIndex < parameters.Count &&
                IsReferenceZoneDataGridView(dgv))
            {
                var param = parameters[e.RowIndex];
                if (param.Zone == ParameterZone.Reference)
                {
                    UpdateReferenceRowVisual(dgv.Rows[e.RowIndex], param);
                }
            }
        }

        private void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv?.DataSource is List<ParameterItem> list &&
                e.RowIndex >= 0 && e.RowIndex < list.Count)
            {
                var item = list[e.RowIndex];
                if (item.Zone == ParameterZone.Reference)
                {
                    UpdateReferenceRowVisual(dgv.Rows[e.RowIndex], item);
                }
            }
        }
        // 【新增】：數據綁定完成事件 - 更新參數狀態對映
        // 由 GitHub Copilot 產生 - DataBindingComplete 中強制上色
        private void DataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv?.DataSource is List<ParameterItem>)
            {
                // 確保事件不重複
                dgv.CellValueChanged -= DataGridView_CellValueChanged;
                dgv.CellValueChanged += DataGridView_CellValueChanged;

                ApplyReferenceZoneColors();
            }
        }


        // 【新增】：判斷是否為參考區的DataGridView
        private bool IsReferenceZoneDataGridView(DataGridView dgv)
        {
            return dgv == dgvCameraUnmodified || dgv == dgvPositionUnmodified || 
                   dgv == dgvDetectionUnmodified || dgv == dgvTimingUnmodified;
        }

        // 【新增】：根據參數狀態取得背景色
        private Color GetParameterStatusBackgroundColor(ParameterItem param)
        {
            // 如果是參考區參數，檢查它的狀態
            if (param.Zone == ParameterZone.Reference)
            {
                string paramKey = $"{param.Name}_{param.Stop}";
                
                // 檢查是否已新增到目標料號
                bool hasAddedUnmodified = allParameters.Any(p => 
                    p.Name == param.Name && p.Stop == param.Stop && p.Zone == ParameterZone.AddedUnmodified);
                bool hasAddedModified = allParameters.Any(p => 
                    p.Name == param.Name && p.Stop == param.Stop && p.Zone == ParameterZone.AddedModified);

                if (hasAddedModified)
                {
                    return Color.LightBlue;      // 已新增已修改 - 淺藍色
                }
                else if (hasAddedUnmodified)
                {
                    return Color.LightGreen;     // 已新增未修改 - 淺綠色
                }
                else
                {
                    return Color.White;          // 尚未新增 - 白色
                }
            }
            
            return Color.White; // 預設白色
        }

        // 【新增】：更新參數狀態對映
        private void UpdateParameterStatusMap()
        {
            parameterStatusMap.Clear();
            
            foreach (var param in allParameters)
            {
                string paramKey = $"{param.Name}_{param.Stop}";
                
                if (param.Zone == ParameterZone.Reference)
                {
                    parameterStatusMap[paramKey] = ParameterStatus.NotAdded;
                }
                else if (param.Zone == ParameterZone.AddedUnmodified)
                {
                    // 檢查參考區是否有相同參數，如果有，更新其狀態
                    string refKey = paramKey;
                    if (parameterStatusMap.ContainsKey(refKey) || 
                        allParameters.Any(p => p.Name == param.Name && p.Stop == param.Stop && p.Zone == ParameterZone.Reference))
                    {
                        parameterStatusMap[refKey] = ParameterStatus.AddedUnmodified;
                    }
                }
                else if (param.Zone == ParameterZone.AddedModified)
                {
                    // 檢查參考區是否有相同參數，如果有，更新其狀態
                    string refKey = paramKey;
                    if (parameterStatusMap.ContainsKey(refKey) || 
                        allParameters.Any(p => p.Name == param.Name && p.Stop == param.Stop && p.Zone == ParameterZone.Reference))
                    {
                        parameterStatusMap[refKey] = ParameterStatus.AddedModified;
                    }
                }
            }
        }
        #endregion
        #endregion

        #region 時間參數計算助手
        // 【新增】：建立時間參數計算助手面板
        private void CreateTimingCalculatorPanel()
        {
            // 主面板
            var grpCalculator = new GroupBox();
            grpCalculator.Text = "⚡ 時間參數智能計算助手";
            grpCalculator.Location = new Point(10, 10);
            grpCalculator.Size = new Size(1550, 40);
            grpCalculator.BackColor = Color.LightYellow;

            // OD顯示標籤
            var lblOD = new Label();
            lblOD.Text = $"🔍 當前料號OD: {GetCurrentTypeOD()} mm";
            lblOD.Location = new Point(10, 15);
            lblOD.Size = new Size(200, 20);
            lblOD.Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold);

            // 計算按鈕
            var btnCalculateBasedOnOD = new Button();
            btnCalculateBasedOnOD.Text = "🧮 根據OD自動計算時間參數";
            btnCalculateBasedOnOD.Location = new Point(220, 12);
            btnCalculateBasedOnOD.Size = new Size(180, 25);
            btnCalculateBasedOnOD.BackColor = Color.LightGreen;
            btnCalculateBasedOnOD.Click += BtnCalculateBasedOnOD_Click;
            btnCalculateBasedOnOD.Visible = false;

            // 參考範例按鈕
            var btnShowReference = new Button();
            btnShowReference.Text = "📊 查看參考數值";
            btnShowReference.Location = new Point(410, 12);
            btnShowReference.Size = new Size(120, 25);
            btnShowReference.BackColor = Color.LightBlue;
            btnShowReference.Click += BtnShowReference_Click;

            // 手動調整說明
            var lblManualHint = new Label();
            lblManualHint.Text = "💡 提示：計算後的數值會自動填入「已修改區」，您可以在該區域中進一步調整";
            lblManualHint.Location = new Point(540, 15);
            lblManualHint.Size = new Size(500, 20);
            lblManualHint.ForeColor = Color.Blue;

            // 組裝面板
            grpCalculator.Controls.Add(lblOD);
            grpCalculator.Controls.Add(btnCalculateBasedOnOD);
            grpCalculator.Controls.Add(btnShowReference);
            grpCalculator.Controls.Add(lblManualHint);

            this.tabTiming.Controls.Add(grpCalculator);
        }

        // 【新增】：根據OD計算時間參數
        private void BtnCalculateBasedOnOD_Click(object sender, EventArgs e)
        {
            try
            {
                double od = GetCurrentTypeOD();

                if (od <= 0)
                {
                    MessageBox.Show("無法取得當前料號的OD數值，請檢查type.txt檔案", "錯誤",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 顯示計算預覽對話框
                ShowTimingCalculationDialog(od);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"計算時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 【新增】：顯示參考數值
        private void BtnShowReference_Click(object sender, EventArgs e)
        {
            try
            {
                string paramPath = Path.Combine(Application.StartupPath, "setting", "param.txt");
                if (!File.Exists(paramPath))
                {
                    MessageBox.Show("找不到參考參數檔案：param.txt", "錯誤",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ShowReferenceDialog(paramPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"顯示參考數值時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 【修正】：取得當前料號的OD - 從資料庫讀取
        private double GetCurrentTypeOD()
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 從 Types 表查詢當前料號的 OD 值
                    var typeInfo = db.Types
                        .Where(t => t.TypeColumn == TargetType)
                        .FirstOrDefault();

                    if (typeInfo != null)
                    {
                        // 假設 OD 欄位名稱是 OD，您需要根據實際的欄位名稱調整
                        return typeInfo.OD;
                    }
                    else
                    {
                        // 如果在資料庫中找不到該料號，記錄警告
                        MessageBox.Show($"在資料庫中找不到料號 {TargetType} 的資料", "警告",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取料號資料時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }

        // 由 GitHub Copilot 產生 - 顯示時間參數計算對話框（fourTo複製 + delay線性預測）
        private void ShowTimingCalculationDialog(double od)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "時間參數計算結果";
                dialog.Size = new Size(900, 650);
                dialog.StartPosition = FormStartPosition.CenterParent;

                // 計算時間參數（fourTo複製 + delay線性預測）
                var calculatedParams = CalculateTimingParameters(od);

                if (calculatedParams.Count == 0)
                {
                    MessageBox.Show("無法取得時間參數，請確認來源料號是否有設定相關參數", "警告",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 公式說明區域
                var txtFormula = new TextBox();
                txtFormula.Text = GetCalculationFormula(od);
                txtFormula.Location = new Point(10, 10);
                txtFormula.Size = new Size(860, 180);
                txtFormula.Multiline = true;
                txtFormula.ReadOnly = true;
                txtFormula.ScrollBars = ScrollBars.Vertical;
                txtFormula.Font = new Font("Microsoft JhengHei", 9);

                // 標題標籤
                var lblFormula = new Label();
                lblFormula.Text = $"📊 fourTo參數從來源料號複製 | delay參數根據OD線性預測 | 當前OD: {od:F2}mm";
                lblFormula.Location = new Point(10, 200);
                lblFormula.Size = new Size(700, 25);
                lblFormula.Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold);
                lblFormula.ForeColor = Color.DarkBlue;

                // 結果表格
                var dgvResults = new DataGridView();
                dgvResults.Location = new Point(10, 230);
                dgvResults.Size = new Size(860, 320);
                dgvResults.DataSource = calculatedParams.Select(p => new {
                    參數名稱 = p.Name,
                    計算結果 = p.Name == "delay" ? $"{p.Value} ms" : p.Value,
                    站點 = p.Stop,
                    說明 = p.ChineseName,
                    來源 = p.Name.StartsWith("fourTo") ? "來源料號複製" : "OD線性預測"
                }).ToList();
                dgvResults.ReadOnly = true;
                dgvResults.AllowUserToAddRows = false;
                dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                // 按鈕
                var btnOK = new Button();
                btnOK.Text = "✅ 使用這些數值";
                btnOK.Location = new Point(680, 570);
                btnOK.Size = new Size(120, 30);
                btnOK.BackColor = Color.LightGreen;
                btnOK.Click += (s, args) =>
                {
                    AddCalculatedParametersToModified(calculatedParams);
                    dialog.DialogResult = DialogResult.OK;
                };

                var btnCancel = new Button();
                btnCancel.Text = "❌ 取消";
                btnCancel.Location = new Point(810, 570);
                btnCancel.Size = new Size(80, 30);
                btnCancel.Click += (s, args) => dialog.DialogResult = DialogResult.Cancel;

                dialog.Controls.AddRange(new Control[] { txtFormula, lblFormula, dgvResults, btnOK, btnCancel });

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    RefreshCurrentTabDisplay();
                    SaveSession();

                    // 統計結果
                    int fourToCount = calculatedParams.Count(p => p.Name.StartsWith("fourTo"));
                    int delayCount = calculatedParams.Count(p => p.Name == "delay");

                    MessageBox.Show($"✅ 成功加入 {calculatedParams.Count} 個時間參數到已修改區\n\n" +
                        $"• fourTo 系列參數：{fourToCount} 個（從來源料號複製）\n" +
                        $"• delay 參數：{delayCount} 個（根據OD線性預測）",
                        "計算完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        #region 時間參數演算法
        // 由 GitHub Copilot 產生 - 計算時間參數（fourTo從來源複製 + delay線性預測）
        private List<ParameterItem> CalculateTimingParameters(double od)
        {
            var calculatedParams = new List<ParameterItem>();
            var warnings = new List<string>();

            // ========== 1. fourTo 系列參數：從來源料號的參考區複製 ==========
            string[] fourToParamNames = { "fourToOK1_time_ms", "fourToOK2_time_ms", "fourToNG_time_ms", "fourToNULL_time_ms" };
            var fourToDescriptions = new Dictionary<string, string>
            {
                { "fourToOK1_time_ms", "站4到OK1出料時間" },
                { "fourToOK2_time_ms", "站4到OK2出料時間" },
                { "fourToNG_time_ms", "站4到NG出料時間" },
                { "fourToNULL_time_ms", "站4到NULL出料時間" }
            };

            // 預設值（當來源料號無參數時使用）
            var fourToDefaults = new Dictionary<string, string>
            {
                { "fourToOK1_time_ms", "1250" },
                { "fourToOK2_time_ms", "1920" },
                { "fourToNG_time_ms", "4500" },
                { "fourToNULL_time_ms", "5000" }
            };

            foreach (var paramName in fourToParamNames)
            {
                // 從參考區尋找來源料號的參數
                var refParam = allParameters.FirstOrDefault(p =>
                    p.Zone == ParameterZone.Reference &&
                    p.Name == paramName &&
                    (p.Stop ?? 0) == 4);

                string value;
                string chineseName = fourToDescriptions[paramName];

                if (refParam != null)
                {
                    // 從來源料號複製
                    value = refParam.Value;
                }
                else
                {
                    // 使用預設值並記錄警告
                    value = fourToDefaults[paramName];
                    warnings.Add($"{paramName}: 來源料號無此參數，使用預設值 {value}");
                }

                calculatedParams.Add(new ParameterItem
                {
                    Name = paramName,
                    Value = value,
                    Stop = 4,
                    ChineseName = chineseName,
                    Type = TargetType,
                    Zone = ParameterZone.AddedModified
                });
            }

            // ========== 2. delay 參數：根據 OD 線性預測（儲存到 camera 資料表） ==========
            // 線性公式：基於 OD=35 的基準值，使用斜率進行線性預測
            // delay_1 = 434 + (OD - 35) × 5.2
            // delay_2 = 427 + (OD - 35) × 3.6
            // delay_3 = 418 + (OD - 35) × 5.8
            // delay_4 = 520 + (OD - 35) × 6.2
            var delayCalculations = new[]
            {
                new { Stop = 1, BaseValue = 434.0, Slope = 5.2 },
                new { Stop = 2, BaseValue = 427.0, Slope = 3.6 },
                new { Stop = 3, BaseValue = 418.0, Slope = 5.8 },
                new { Stop = 4, BaseValue = 520.0, Slope = 6.2 }
            };

            foreach (var calc in delayCalculations)
            {
                // 線性預測公式
                double delayValue = calc.BaseValue + (od - 35.0) * calc.Slope;
                int roundedValue = (int)Math.Round(delayValue);

                calculatedParams.Add(new ParameterItem
                {
                    Name = "delay",
                    Value = roundedValue.ToString(),
                    Stop = calc.Stop,
                    ChineseName = "拍照延遲時間",
                    Type = TargetType,
                    Zone = ParameterZone.AddedModified
                });
            }

            // 顯示警告訊息（如果有）
            if (warnings.Count > 0)
            {
                MessageBox.Show($"部分參數使用預設值：\n\n" + string.Join("\n", warnings),
                    "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return calculatedParams;
        }

        // 由 GitHub Copilot 產生 - 顯示計算公式的詳細說明
        private string GetCalculationFormula(double od)
        {
            // 計算預測的 delay 值供顯示
            int delay1 = (int)Math.Round(434.0 + (od - 35.0) * 5.2);
            int delay2 = (int)Math.Round(427.0 + (od - 35.0) * 3.6);
            int delay3 = (int)Math.Round(418.0 + (od - 35.0) * 5.8);
            int delay4 = (int)Math.Round(520.0 + (od - 35.0) * 6.2);

            return $@"📐 時間參數計算說明 (當前OD: {od:F2}mm)

🔹 fourTo 系列參數（從來源料號複製）：
   • fourToOK1_time_ms - 站4到OK1出料時間
   • fourToOK2_time_ms - 站4到OK2出料時間
   • fourToNG_time_ms - 站4到NG出料時間
   • fourToNULL_time_ms - 站4到NULL出料時間
   ➤ 這些參數與OD無明顯線性關係，直接從來源料號複製
   ➤ 若來源料號無此參數，則使用預設值

🔹 delay 參數（根據OD線性預測，儲存到camera資料表）：
   • 基準OD：35mm
   • 計算公式：delay = 基準值 + (OD - 35) × 斜率

   站點1：delay_1 = 434 + (OD - 35) × 5.2 = {delay1} ms
   站點2：delay_2 = 427 + (OD - 35) × 3.6 = {delay2} ms
   站點3：delay_3 = 418 + (OD - 35) × 5.8 = {delay3} ms
   站點4：delay_4 = 520 + (OD - 35) × 6.2 = {delay4} ms

💡 說明：
   • delay 參數與 OD 呈線性正相關
   • 斜率基於實際生產資料計算得出
   • 無上下限限制，直接線性外推";
        }
        #endregion
        // 由 GitHub Copilot 產生 - 將計算結果加入到已修改區
        private void AddCalculatedParametersToModified(List<ParameterItem> calculatedParams)
        {
            foreach (var calcParam in calculatedParams)
            {
                // 檢查是否已存在相同的參數（需要同時比對 Name 和 Stop）
                var existingParam = allParameters.FirstOrDefault(p =>
                    p.Name == calcParam.Name && p.Stop == calcParam.Stop);

                if (existingParam != null)
                {
                    // 更新現有參數的值和區域
                    existingParam.Value = calcParam.Value;
                    existingParam.Zone = ParameterZone.AddedModified;
                }
                else
                {
                    // 新增新參數
                    calcParam.Type = TargetType;
                    calcParam.Zone = ParameterZone.AddedModified;
                    calcParam.IsSelected = false;
                    allParameters.Add(calcParam);
                }
            }
        }

        // 【新增】：顯示參考數值對話框
        private void ShowReferenceDialog(string paramPath)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "參考參數數值";
                dialog.Size = new Size(800, 500);
                dialog.StartPosition = FormStartPosition.CenterParent;

                // 讀取參考數據
                var referenceData = LoadReferenceTimingData(paramPath);

                var dgvReference = new DataGridView();
                dgvReference.Location = new Point(10, 10);
                dgvReference.Size = new Size(760, 400);
                dgvReference.DataSource = referenceData;
                dgvReference.ReadOnly = true;
                dgvReference.AllowUserToAddRows = false;

                var btnClose = new Button();
                btnClose.Text = "關閉";
                btnClose.Location = new Point(700, 420);
                btnClose.Size = new Size(80, 30);
                btnClose.Click += (s, args) => dialog.Close();

                dialog.Controls.AddRange(new Control[] { dgvReference, btnClose });
                dialog.ShowDialog();
            }
        }

        // 【新增】：載入參考時間參數數據
        private List<object> LoadReferenceTimingData(string paramPath)
        {
            var referenceData = new List<object>();

            try
            {
                var lines = File.ReadAllLines(paramPath);
                foreach (var line in lines.Skip(1)) // 跳過標題行
                {
                    var parts = line.Split(',');
                    if (parts.Length >= 5)
                    {
                        string paramName = parts[1].Trim();
                        // 只顯示時間相關參數
                        if (paramName.Contains("time") || paramName.Contains("fourTo"))
                        {
                            referenceData.Add(new
                            {
                                料號 = parts[0].Trim(),
                                參數名 = paramName,
                                參數值 = parts[2].Trim(),
                                站點 = parts[3].Trim(),
                                中文名稱 = parts[4].Trim()
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取參考數據時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return referenceData;
        }
        #endregion
        #region 檢測參數智能校正助手
        // 【新增】：建立檢測參數智能校正助手面板
        private void CreateDetectionCalibrationPanel()
        {
            // 主面板
            var grpCalibration = new GroupBox();
            grpCalibration.Text = "🔧 檢測參數智能校正助手";
            grpCalibration.Location = new Point(10, 10);
            grpCalibration.Size = new Size(1550, 40);
            grpCalibration.BackColor = Color.LightCyan;

            // 校正類型下拉選單
            var lblCalibrationType = new Label();
            lblCalibrationType.Text = "校正類型：";
            lblCalibrationType.Location = new Point(10, 15);
            lblCalibrationType.Size = new Size(70, 20);
            lblCalibrationType.Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold);

            var cmbCalibrationType = new ComboBox();
            cmbCalibrationType.Location = new Point(80, 12);
            cmbCalibrationType.Size = new Size(200, 25);
            cmbCalibrationType.DropDownStyle = ComboBoxStyle.DropDownList;

            // 新增校正類型選項
            cmbCalibrationType.Items.AddRange(new object[] {
                new CalibrationItem { Name = "物體偏移量校正 (objBias)", Type = CalibrationType.ObjectBias },
                new CalibrationItem { Name = "開口閾值校正 (gapThreshold)", Type = CalibrationType.gapThresh },
                new CalibrationItem { Name = "檢測精度校正 (Pixel)", Type = CalibrationType.Pixel },
                new CalibrationItem { Name = "對比檢測校正 (Contrast)", Type = CalibrationType.Contrast },
                new CalibrationItem { Name = "白色像素校正 (white)", Type = CalibrationType.White }
            });
            cmbCalibrationType.DisplayMember = "Name";
            cmbCalibrationType.ValueMember = "Type";
            cmbCalibrationType.SelectedIndex = 0;

            /*
            改為選取參數後，進入介面再選站點
            // 站點選擇下拉選單
            var lblStation = new Label();
            lblStation.Text = "站點：";
            lblStation.Location = new Point(290, 15);
            lblStation.Size = new Size(40, 20);
            lblStation.Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold);

            var cmbStation = new ComboBox();
            cmbStation.Location = new Point(330, 12);
            cmbStation.Size = new Size(80, 25);
            cmbStation.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbStation.Items.AddRange(new object[] { "站點1", "站點2", "站點3", "站點4" });
            cmbStation.SelectedIndex = 0;
            */

            // 開始校正按鈕
            var btnStartCalibration = new Button();
            btnStartCalibration.Text = "🎯 開始校正";
            btnStartCalibration.Location = new Point(420, 12);
            btnStartCalibration.Size = new Size(120, 25);
            btnStartCalibration.BackColor = Color.LightGreen;
            btnStartCalibration.Click += (s, e) => StartCalibration(
              (CalibrationType)((CalibrationItem)cmbCalibrationType.SelectedItem).Type);

            // 參數預覽按鈕
            var btnPreviewParameters = new Button();
            btnPreviewParameters.Text = "👁️ 參數預覽";
            btnPreviewParameters.Location = new Point(550, 12);
            btnPreviewParameters.Size = new Size(100, 25);
            btnPreviewParameters.BackColor = Color.LightBlue;
            btnPreviewParameters.Click += (s, e) => PreviewCalibrationParameters(
              (CalibrationType)((CalibrationItem)cmbCalibrationType.SelectedItem).Type);

            // 說明標籤
            var lblHint = new Label();
            lblHint.Text = "💡 選擇校正類型和站點，系統將引導您完成校正流程";
            lblHint.Location = new Point(660, 15);
            lblHint.Size = new Size(400, 20);
            lblHint.ForeColor = Color.DarkBlue;

            // 組裝面板
            grpCalibration.Controls.AddRange(new Control[] {
                lblCalibrationType, cmbCalibrationType,
                btnStartCalibration, btnPreviewParameters, lblHint
            });

            this.tabDetection.Controls.Add(grpCalibration);
        }

        // 【新增】：校正類型列舉
        public enum CalibrationType
        {
            ObjectBias,     // 物體偏移量校正
            gapThresh,      // 閾值參數校正
            Pixel,      // 尺寸檢測校正
            Contrast,        // 輪廓檢測校正
            White          // 白色像素校正
        }

        // 【新增】：校正項目類別
        public class CalibrationItem
        {
            public string Name { get; set; }
            public CalibrationType Type { get; set; }
        }

        // 【新增】：開始校正流程
        private void StartCalibration(CalibrationType calibrationType)
        {
            try
            {
                switch (calibrationType)
                {
                    case CalibrationType.ObjectBias:
                        StartObjectBiasCalibration();
                        break;
                    case CalibrationType.gapThresh:
                        StartGapThreshCalibration();
                        break;
                    case CalibrationType.Pixel:
                        StartPixelCalibration();
                        break;
                    case CalibrationType.Contrast:
                        StartContrastCalibration();
                        break;
                    case CalibrationType.White:
                        StartWhiteCalibration();
                        break;
                    default:
                        MessageBox.Show("此校正類型尚未實作", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                }
            }
            catch (Exception ex)
            {
                // 【詳細錯誤記錄】
                string errorDetails = $"校正類型: {calibrationType}\n" +
                                     $"站點: \n" +
                                     $"料號: {TargetType}\n" +
                                     $"錯誤訊息: {ex.Message}\n" +
                                     $"錯誤類型: {ex.GetType().Name}\n" +
                                     $"堆疊追蹤:\n{ex.StackTrace}";

                // 顯示詳細錯誤
                MessageBox.Show(errorDetails, "詳細錯誤資訊",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                // 【可選】寫入日誌檔案
                WriteErrorLog(errorDetails);
            }
        }
        // 【新增】錯誤日誌方法
        private void WriteErrorLog(string errorMessage)
        {
            try
            {
                string logPath = Path.Combine(Application.StartupPath, "logs");
                if (!Directory.Exists(logPath))
                    Directory.CreateDirectory(logPath);

                string logFile = Path.Combine(logPath, $"calibration_error_{DateTime.Now:yyyyMMdd}.txt");
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {errorMessage}\n" + new string('=', 80) + "\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch
            {
                // 日誌寫入失敗時忽略
            }
        }
        // 【新增】：參數預覽功能
        private void PreviewCalibrationParameters(CalibrationType calibrationType)
        {
            var parameterNames = GetCalibrationParameterNames(calibrationType);

            if (parameterNames.Count == 0)
            {
                MessageBox.Show("此校正類型沒有相關參數", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ShowParameterPreviewDialog(calibrationType, parameterNames);
        }

        // 【修正】：取得校正相關的參數名稱
        private List<string> GetCalibrationParameterNames(CalibrationType calibrationType)
        {
            var parameterNames = new List<string>();

            switch (calibrationType)
            {
                case CalibrationType.ObjectBias:
                    // 顯示所有站點的 objBias 參數
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.AddRange(new[] {
                    $"objBias_x_{station}",
                    $"objBias_y_{station}"
                });
                    }
                    break;

                case CalibrationType.gapThresh:
                    // 顯示所有站點的 gapThresh 參數
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.Add($"gapThresh_{station}");
                    }
                    break;

                case CalibrationType.Pixel:
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.Add($"PixelToMM_{station}");
                    }
                    break;

                case CalibrationType.Contrast:
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.Add($"deepenBrightness_{station}");
                        parameterNames.Add($"deepenContrast_{station}");
                    }
                    break;

                case CalibrationType.White:
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.Add($"white_{station}");
                    }
                    break;
            }

            return parameterNames;
        }

        // 【修正】：顯示參數預覽對話框
        private void ShowParameterPreviewDialog(CalibrationType calibrationType, List<string> parameterNames)
        {
            using (var dialog = new Form())
            {
                dialog.Text = $"{calibrationType} 校正參數預覽 - 所有站點";
                dialog.Size = new Size(700, 500);
                dialog.StartPosition = FormStartPosition.CenterParent;

                // 說明標籤
                var lblDescription = new Label();
                lblDescription.Text = $"以下是 {calibrationType} 校正會影響的參數（所有站點）：";
                lblDescription.Location = new Point(10, 10);
                lblDescription.Size = new Size(500, 25);
                lblDescription.Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold);

                // 參數列表 - 使用 DataGridView 更好顯示
                var dgvParameters = new DataGridView();
                dgvParameters.Location = new Point(10, 40);
                dgvParameters.Size = new Size(660, 350);
                dgvParameters.ReadOnly = true;
                dgvParameters.AllowUserToAddRows = false;
                dgvParameters.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                // 由 GitHub Copilot 產生 - 修正參數查詢邏輯，同時匹配名稱和站點
                var parameterData = parameterNames.Select(paramName => {
                    var (baseName, stop) = GetParameterNameAndStop(paramName);
                    var existingParam = allParameters.FirstOrDefault(p =>
                        p.Name == baseName && p.Stop == stop);
                    return new
                    {
                        參數名稱 = paramName,
                        當前值 = existingParam?.Value ?? "未設定",
                        狀態 = existingParam?.Zone.ToString() ?? "不存在"
                    };
                }).ToList();

                dgvParameters.DataSource = parameterData;

                // 提示標籤
                var lblHint = new Label();
                lblHint.Text = "💡 校正時將在表單內選擇特定站點進行調整";
                lblHint.Location = new Point(10, 400);
                lblHint.Size = new Size(400, 25);
                lblHint.ForeColor = Color.Blue;

                // 關閉按鈕
                var btnClose = new Button();
                btnClose.Text = "關閉";
                btnClose.Location = new Point(590, 430);
                btnClose.Size = new Size(80, 30);
                btnClose.Click += (s, e) => dialog.Close();

                dialog.Controls.AddRange(new Control[] {
            lblDescription, dgvParameters, lblHint, btnClose
        });
                dialog.ShowDialog();
            }
        }

        // 【新增】：從完整參數名稱取得基本名稱
        private string GetParameterBaseName(string fullParameterName)
        {
            // 移除站點後綴，例如 "objBias_x_1" -> "objBias_x"
            if (fullParameterName.Contains("_"))
            {
                var parts = fullParameterName.Split('_');
                if (parts.Length >= 2 && int.TryParse(parts.Last(), out _))
                {
                    // 如果最後一部分是數字（站點），移除它
                    return string.Join("_", parts.Take(parts.Length - 1));
                }
            }
            return fullParameterName;
        }

        // 由 GitHub Copilot 產生 - 從完整參數名稱取得基本名稱和站點號碼
        private (string baseName, int? stop) GetParameterNameAndStop(string fullParameterName)
        {
            // 提取基本名稱和站點，例如 "objBias_x_1" -> ("objBias_x", 1)
            if (fullParameterName.Contains("_"))
            {
                var parts = fullParameterName.Split('_');
                if (parts.Length >= 2 && int.TryParse(parts.Last(), out int stopNum))
                {
                    // 如果最後一部分是數字（站點），提取它
                    string baseName = string.Join("_", parts.Take(parts.Length - 1));
                    return (baseName, stopNum);
                }
            }
            return (fullParameterName, null);
        }

        // 由 GitHub Copilot 產生 - 修正參數查詢邏輯
        private string GetCurrentParameterValues(List<string> parameterNames)
        {
            var values = new List<string>();

            foreach (var paramName in parameterNames)
            {
                var (baseName, stop) = GetParameterNameAndStop(paramName);
                var existingParam = allParameters.FirstOrDefault(p =>
                    p.Name == baseName && p.Stop == stop);
                string value = existingParam?.Value ?? "未設定";
                values.Add($"{paramName}={value}");
            }

            return string.Join(", ", values);
        }
        #endregion

        #region 各種校正功能實作
        // 【實作】：物體偏移量校正
        private void StartObjectBiasCalibration()
        {
            using (var calibrationForm = new ObjectBiasCalibrationForm(TargetType))
            {
                // 附加關閉事件監聽器
                AttachCalibrationFormCloseEvent(calibrationForm);

                // 顯示表單（不等待 DialogResult）
                calibrationForm.ShowDialog();

                // 注意：參數更新會在 FormClosed 事件中自動處理
            }
        }

        // 【框架】：閾值參數校正
        private void StartGapThreshCalibration()
        {
            using (var calibrationForm = new GapThreshCalibrationForm(TargetType))
            {
                AttachCalibrationFormCloseEvent(calibrationForm);
                calibrationForm.ShowDialog();
            }
        }

        // 【框架】：尺寸檢測校正
        private void StartPixelCalibration()
        {
            using (var calibrationForm = new PixelCalibrationForm())
            {
                AttachCalibrationFormCloseEvent(calibrationForm);
                calibrationForm.ShowDialog();
            }
        }

        // 【框架】：對比檢測校正
        // 由 GitHub Copilot 產生 - 對比檢測校正
        private void StartContrastCalibration()
        {
            try
            {
                using (var calibrationForm = new ContrastCalibrationForm(TargetType))
                {
                    AttachCalibrationFormCloseEvent(calibrationForm);
                    calibrationForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"對比度校正時發生錯誤：{ex.Message}";
                WriteErrorLog(errorMsg);
                MessageBox.Show(errorMsg, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 【框架】：白色檢測校正
        private void StartWhiteCalibration()
        {
            using (var calibrationForm = new WhiteCalibrationForm(TargetType))
            {
                AttachCalibrationFormCloseEvent(calibrationForm);
                calibrationForm.ShowDialog();
            }
        }

        // 【通用方法】：從資料庫重新載入指定參數
        private void LoadUpdatedParametersFromDatabase(string[] parameterNames)
        {
            try
            {
                using (var db = new MydbDB())
                {
                    foreach (var paramName in parameterNames)
                    {
                        // 從資料庫查詢最新的參數值
                        var updatedParam = db.@params.FirstOrDefault(p =>
                            p.Type == TargetType && p.Name == paramName);

                        if (updatedParam != null)
                        {
                            // 更新現有參數或新增參數
                            var existingParam = allParameters.FirstOrDefault(p =>
                                p.Name == paramName && p.Stop == updatedParam.Stop);

                            if (existingParam != null)
                            {
                                existingParam.Value = updatedParam.Value;
                                existingParam.Zone = ParameterZone.AddedModified; // 校正完成的參數
                            }
                            else
                            {
                                allParameters.Add(new ParameterItem
                                {
                                    Type = TargetType,
                                    Name = updatedParam.Name,
                                    Value = updatedParam.Value,
                                    Stop = updatedParam.Stop,
                                    ChineseName = updatedParam.ChineseName,
                                    Zone = ParameterZone.AddedModified,
                                    IsSelected = false
                                });
                            }
                        }
                    }

                    // 更新進度和儲存工作階段
                    UpdateCategoryProgress();
                    SaveSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"重新載入參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 由 GitHub Copilot 產生 - 校正介面關閉事件統一處理器
        #region 關閉事件
        private void AttachCalibrationFormCloseEvent<T>(T form) where T : Form
        {
            form.FormClosed += OnCalibrationFormClosed;
        }

        // 由 GitHub Copilot 產生 - 校正介面關閉後自動更新參數
        private void OnCalibrationFormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                // 無論 DialogResult 為何，都要更新參數（因為可能已經套用過參數）
                RefreshParametersAfterCalibration();

                // 顯示更新完成通知
                UpdateStatus("✅ 校正介面已關閉，參數已自動更新");

                // 可選：顯示簡短的通知訊息
                ShowParameterUpdateNotification();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        // 由 GitHub Copilot 產生 - 重新載入校正完成後的參數
        private void RefreshParametersAfterCalibration()
        {
            try
            {
                UpdateModifiedParametersFromDatabase();
                RefreshCurrentTabDisplay(); // 內部會再次呼叫進度更新
                UpdateParameterStatusMap();
                RecalculateProgressAndUpdateUI();
                SaveSession();
                System.Diagnostics.Debug.WriteLine("校正後參數更新完成");
            }
            catch (Exception ex)
            {
                throw new Exception($"重新載入參數失敗：{ex.Message}");
            }
        }
        private void UpdateModifiedParametersFromDatabase()
        {
            using (var db = new MydbDB())
            {
                // 載入最新的目標料號Camera參數
                var updatedCameraParams = db.Cameras
                    .Where(c => c.Type == TargetType)
                    .ToList();

                // 載入最新的目標料號Param參數  
                var updatedParams = db.@params
                    .Where(p => p.Type == TargetType)
                    .ToList();

                // 🔸 更新或新增Camera參數到已修改區
                foreach (var cameraParam in updatedCameraParams)
                {
                    UpdateOrAddModifiedParameter(cameraParam.Name, cameraParam.Value, cameraParam.Stop, cameraParam.ChineseName);
                }

                // 🔸 更新或新增Param參數到已修改區
                foreach (var param in updatedParams)
                {
                    UpdateOrAddModifiedParameter(param.Name, param.Value, param.Stop, param.ChineseName);
                }
            }
        }
        private void UpdateOrAddModifiedParameter(string name, string value, int? stop, string chineseName)
        {
            // 🔸 尋找已修改區中的現有參數
            var existingModifiedParam = allParameters.FirstOrDefault(p =>
                p.Name == name && p.Stop == stop && p.Zone == ParameterZone.AddedModified);

            if (existingModifiedParam != null)
            {
                // 🔸 更新現有已修改參數的值
                existingModifiedParam.Value = value;
                existingModifiedParam.ChineseName = chineseName;
                System.Diagnostics.Debug.WriteLine($"更新已修改參數：{name}_{stop} = {value}");
            }
            else
            {
                // 🔸 檢查是否有參考區的對應參數（判斷是否為新參數）
                var hasReferenceParam = allParameters.Any(p =>
                    p.Name == name && p.Stop == stop && p.Zone == ParameterZone.Reference);

                // 🔸 新增參數到已修改區
                allParameters.Add(new ParameterItem
                {
                    Type = TargetType,
                    Name = name,
                    Value = value,
                    Stop = stop,
                    ChineseName = chineseName,
                    Zone = ParameterZone.AddedModified,
                    IsSelected = false
                });

                System.Diagnostics.Debug.WriteLine($"新增已修改參數：{name}_{stop} = {value} (參考區存在: {hasReferenceParam})");
            }

            // 🔸 同步移除已新增未修改區中的相同參數（如果存在的話）
            var existingUnmodifiedParam = allParameters.FirstOrDefault(p =>
                p.Name == name && p.Stop == stop && p.Zone == ParameterZone.AddedUnmodified);

            if (existingUnmodifiedParam != null)
            {
                allParameters.Remove(existingUnmodifiedParam);
                System.Diagnostics.Debug.WriteLine($"移除未修改區參數：{name}_{stop} (已移至已修改區)");
            }
        }
        // 由 GitHub Copilot 產生 - 從資料庫重新載入所有參數
        private void LoadUpdatedParametersFromDatabase()
        {
            using (var db = new MydbDB())
            {
                // 載入Camera參數
                var updatedCameraParams = db.Cameras
                    .Where(c => c.Type == TargetType)
                    .ToList();

                // 載入Param參數  
                var updatedParams = db.@params
                    .Where(p => p.Type == TargetType)
                    .ToList();

                // 更新現有參數或新增新參數
                foreach (var cameraParam in updatedCameraParams)
                {
                    UpdateOrAddParameter(cameraParam.Name, cameraParam.Value, cameraParam.Stop, cameraParam.ChineseName, true);
                }

                foreach (var param in updatedParams)
                {
                    UpdateOrAddParameter(param.Name, param.Value, param.Stop, param.ChineseName, false);
                }
            }
        }
        // 由 GitHub Copilot 產生 - 更新或新增參數到本地集合
        private void UpdateOrAddParameter(string name, string value, int? stop, string chineseName, bool isCameraParam)
        {
            // 尋找現有參數
            var existingParam = allParameters.FirstOrDefault(p =>
                p.Name == name && p.Stop == stop);

            if (existingParam != null)
            {
                // 更新現有參數
                existingParam.Value = value;
                existingParam.ChineseName = chineseName;

                // 標記為已修改（校正後的參數）
                if (existingParam.Zone == ParameterZone.Reference)
                {
                    // 如果原本在參考區，移動到已新增已修改區
                    existingParam.Zone = ParameterZone.AddedModified;
                }
                else if (existingParam.Zone == ParameterZone.AddedUnmodified)
                {
                    // 如果原本在已新增未修改區，移動到已新增已修改區
                    existingParam.Zone = ParameterZone.AddedModified;
                }
                // 如果已經在已修改區，保持不變
            }
            else
            {
                // 新增新參數（校正新增的參數）
                allParameters.Add(new ParameterItem
                {
                    Type = TargetType,
                    Name = name,
                    Value = value,
                    Stop = stop,
                    ChineseName = chineseName,
                    Zone = ParameterZone.AddedModified, // 新校正的參數直接進入已修改區
                    IsSelected = false
                });
            }
        }

        // 由 GitHub Copilot 產生 - 顯示參數更新通知
        private void ShowParameterUpdateNotification()
        {
            // 建立臨時通知標籤
            var notification = new Label();
            notification.Text = "✅ 參數已更新";
            notification.BackColor = Color.LightGreen;
            notification.ForeColor = Color.DarkGreen;
            notification.Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold);
            notification.Size = new Size(120, 30);
            notification.Location = new Point(this.Width - 150, 50);
            notification.TextAlign = ContentAlignment.MiddleCenter;
            notification.BorderStyle = BorderStyle.FixedSingle;

            this.Controls.Add(notification);
            notification.BringToFront();

            // 3秒後自動移除通知
            var timer = new System.Windows.Forms.Timer();
            timer.Interval = 3000;
            timer.Tick += (s, e) => {
                timer.Stop();
                timer.Dispose();
                this.Controls.Remove(notification);
                notification.Dispose();
            };
            timer.Start();
        }

        // 由 GitHub Copilot 產生 - 更新狀態顯示
        private void UpdateStatus(string message)
        {
            // 如果有狀態標籤，更新它
            if (lblProgress != null)
            {
                lblProgress.Text = $"料號: {TargetType} | {message}";
            }
        }
        #endregion
        #endregion
        #region 部分自動導入功能
        // 由 GitHub Copilot 產生 - 在建構函式中新增部分自動導入按鈕
        private Button btnPartialAutoImport;

        // 由 GitHub Copilot 產生 - 在 InitializeForm 方法中建立按鈕
        private void CreatePartialAutoImportButton()
        {
            // 將「部分自動導入」按鈕加到每個 TabPage，避免被 TabControl 遮蔽
            AddPartialAutoImportButtonToTab(this.tabCamera);
            AddPartialAutoImportButtonToTab(this.tabPosition);
            AddPartialAutoImportButtonToTab(this.tabDetection);
            AddPartialAutoImportButtonToTab(this.tabTiming);
        }
        private void AddPartialAutoImportButtonToTab(TabPage tab)
        {
            if (tab == null) return;

            var btn = new Button();
            btn.Name = "btnPartialAutoImport_" + (tab.Name ?? Guid.NewGuid().ToString("N"));
            btn.Text = "🔄 部分自動導入";
            btn.Size = new System.Drawing.Size(150, 30);

            // 放在每個 TabPage 的右上角（避開你現有的「全選/清除/反選」按鈕）
            var x = Math.Max(10, tab.ClientSize.Width - btn.Width - 15);
            btn.Location = new System.Drawing.Point(x, 12);

            btn.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btn.BackColor = System.Drawing.Color.LightSalmon;
            btn.Font = new System.Drawing.Font("Microsoft JhengHei", 9, System.Drawing.FontStyle.Bold);
            btn.Click += BtnPartialAutoImport_Click;

            var toolTip = new ToolTip();
            toolTip.SetToolTip(btn,
                "根據料號特性自動調整特定參數值\n" +
                "• 無油溝料號：find_NROI 站1 設為 0\n" +
                "• 其他條件式參數自動處理");

            tab.Controls.Add(btn);
            btn.BringToFront(); // 確保在最上層
        }
        // 由 GitHub Copilot 產生 - 部分自動導入主要邏輯
        private void BtnPartialAutoImport_Click(object sender, EventArgs e)
        {
            try
            {
                // 顯示確認對話框
                var confirmResult = MessageBox.Show(
                    $"即將對料號 {TargetType} 執行部分自動導入\n\n" +
                    "將會處理以下條件式參數：\n" +
                    "• find_NROI：根據油溝狀態自動設定\n" +
                    "• 所有相機參數：根據OD、PTFE顏色自動設定\n" +
                    "• 所有位置參數：直接複製來源料號，須後續調整\n" +
                    "• 其他特殊參數依據料號特性調整\n\n" +
                    "確定要執行嗎？",
                    "確認部分自動導入",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirmResult == DialogResult.Yes)
                {
                    ExecutePartialAutoImport();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"部分自動導入時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 由 GitHub Copilot 產生 - 執行部分自動導入的主要邏輯
        private void ExecutePartialAutoImport()
        {
            try
            {
                // 取得當前料號資訊
                var typeInfo = GetTypeInfo(TargetType);
                if (typeInfo == null)
                {
                    MessageBox.Show($"無法取得料號 {TargetType} 的基本資訊", "錯誤",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int processedCount = 0;
                var processedParameters = new List<string>();

                // 處理 find_NROI 參數（油溝相關）
                processedCount += ProcessFindNROIParameters(typeInfo, processedParameters);

                // 處理其他條件式參數
                processedCount += ProcessOtherConditionalParameters(typeInfo, processedParameters);

                // 由 GitHub Copilot 產生 - 處理 delay 參數（根據 OD 線性計算，寫入 Cameras 資料表）
                processedCount += ProcessDelayParameters(typeInfo, processedParameters);

                // 由 GitHub Copilot 產生 - 處理 fourTo 系列參數（從來源料號複製，寫入 params 資料表）
                processedCount += ProcessFourToParameters(typeInfo, processedParameters);

                // 由 GitHub Copilot 產生 - 處理位置參數（known 開頭的所有參數，從來源料號直接複製）
                processedCount += ProcessPositionParameters(typeInfo, processedParameters);

                // 由 GitHub Copilot 產生 - 處理 DefectChecks 資料表（缺陷檢測項目，直接寫入資料庫）
                processedCount += ProcessDefectChecksParameters(typeInfo, processedParameters);

                // 更新顯示
                RefreshCurrentTabDisplay();
                UpdateCategoryProgress();
                SaveSession();

                // 顯示處理結果
                ShowPartialAutoImportResult(processedCount, processedParameters);
            }
            catch (Exception ex)
            {
                throw new Exception($"執行部分自動導入失敗：{ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生 - 取得料號基本資訊
        private TypeInfo GetTypeInfo(string targetType)
        {
            try
            {
                using (var db = new MydbDB())
                {
                    var typeData = db.Types
                        .Where(t => t.TypeColumn == targetType)
                        .FirstOrDefault();

                    if (typeData != null)
                    {
                        return new TypeInfo
                        {
                            TypeColumn = typeData.TypeColumn,
                            Material = typeData.material,
                            Thick = typeData.thick,
                            PTFEColor = typeData.PTFEColor,
                            ID = typeData.ID,
                            OD = typeData.OD,
                            H = typeData.H,
                            HasGroove = typeData.hasgroove,
                            BoxOrPack = typeData.boxorpack,
                            HasYZP = typeData.hasYZP,
                            package = typeData.package
                        };
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception($"取得料號資訊失敗：{ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生 - 處理 find_NROI 參數
        private int ProcessFindNROIParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                // 站1、2：依內圈NROI模型決定
                bool innerModel = app.has_NROI_InnerModel;
                for (int stop = 1; stop <= 2; stop++)
                {
                    // find_NROI
                    string findVal = innerModel ? "1" : "0";
                    UpdateOrAddParameterToModified(
                        "find_NROI",
                        findVal,
                        stop,
                        "尋找非ROI(內圈)",
                        innerModel ? "有內圈NROI模型 → 設為1" : "無內圈NROI模型 → 設為0"
                    );
                    processedParameters.Add($"find_NROI_{stop} = {findVal} (InnerModel={(innerModel ? "true" : "false")})");
                    count++;

                    // expandNROI_* 必定寫入（無模型則為0）
                    string expIn = innerModel ? "15" : "0";
                    string expOut = innerModel ? "30" : "0";

                    UpdateOrAddParameterToModified(
                        "expandNROI_in",
                        expIn,
                        stop,
                        "NROI向內擴展像素(內圈)",
                        innerModel ? "內圈NROI模型 → 15" : "無內圈NROI模型 → 設為0"
                    );
                    UpdateOrAddParameterToModified(
                        "expandNROI_out",
                        expOut,
                        stop,
                        "NROI向外擴展像素(內圈)",
                        innerModel ? "內圈NROI模型 → 30" : "無內圈NROI模型 → 設為0"
                    );
                    processedParameters.Add($"expandNROI_in_{stop} = {expIn}");
                    processedParameters.Add($"expandNROI_out_{stop} = {expOut}");
                    count += 2;
                }

                // 站3、4：依外圈NROI模型決定
                bool outerModel = app.has_NROI_OuterModel;
                for (int stop = 3; stop <= 4; stop++)
                {
                    // find_NROI
                    string findVal = outerModel ? "1" : "0";
                    UpdateOrAddParameterToModified(
                        "find_NROI",
                        findVal,
                        stop,
                        "尋找非ROI(外圈)",
                        outerModel ? "有外圈NROI模型 → 設為1" : "無外圈NROI模型 → 設為0"
                    );
                    processedParameters.Add($"find_NROI_{stop} = {findVal} (OuterModel={(outerModel ? "true" : "false")})");
                    count++;

                    // expandNROI_* 必定寫入（無模型則為0）
                    string expIn = outerModel ? "4" : "0";
                    string expOut = outerModel ? "8" : "0";

                    UpdateOrAddParameterToModified(
                        "expandNROI_in",
                        expIn,
                        stop,
                        "NROI向內擴展像素(外圈)",
                        outerModel ? "外圈NROI模型 → 4" : "無外圈NROI模型 → 設為0"
                    );
                    UpdateOrAddParameterToModified(
                        "expandNROI_out",
                        expOut,
                        stop,
                        "NROI向外擴展像素(外圈)",
                        outerModel ? "外圈NROI模型 → 8" : "無外圈NROI模型 → 設為0"
                    );
                    processedParameters.Add($"expandNROI_in_{stop} = {expIn}");
                    processedParameters.Add($"expandNROI_out_{stop} = {expOut}");
                    count += 2;
                }

                return count;
            }
            catch (Exception ex)
            {
                throw new Exception($"處理 find_NROI/expandNROI 參數時發生錯誤：{ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生 - 處理其他條件式參數
        private int ProcessOtherConditionalParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                // 根據材質處理特定參數
                count += ProcessMaterialBasedParameters(typeInfo, processedParameters);

                // 根據尺寸處理特定參數
                count += ProcessSizeBasedParameters(typeInfo, processedParameters);

                // 根據厚度處理特定參數
                count += ProcessThicknessBasedParameters(typeInfo, processedParameters);

                // 根據有無五彩鋅處理特定參數
                count += ProcessYZPParameters(typeInfo, processedParameters);

                // 依 DefectChecks(outsc) 決定是否寫入 outsc 參數（僅站3/4）
                count += ProcessOutscParametersUsingDefectChecks(typeInfo, processedParameters);

                // 【新增】直接從來源料號複製特定參數（不更動數值）
                // 調整這份清單即可控制要直拷的參數名稱（會套用到 1~4 站）
                var directCopyList = new[] { "objTolerance", "IOU" };
                count += ProcessDirectCopyFromSourceParameters(directCopyList, processedParameters, overwriteExisting: false);

                return count;
            }
            catch (Exception ex)
            {
                throw new Exception($"處理其他條件式參數時發生錯誤：{ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生 - 處理 delay 參數（根據 OD 線性計算，寫入 Cameras 資料表）
        private int ProcessDelayParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                // 確認有有效的 OD 值
                if (typeInfo.OD <= 0)
                {
                    System.Diagnostics.Debug.WriteLine("[ProcessDelayParameters] OD 值無效，無法計算 delay");
                    return 0;
                }

                double od = typeInfo.OD;

                // delay 線性公式：基於 OD=35 的基準值，使用斜率進行線性預測
                // delay_1 = 434 + (OD - 35) × 5.2
                // delay_2 = 427 + (OD - 35) × 3.6
                // delay_3 = 418 + (OD - 35) × 5.8
                // delay_4 = 520 + (OD - 35) × 6.2
                var delayCalculations = new[]
                {
                    new { Stop = 1, BaseValue = 434.0, Slope = 5.2 },
                    new { Stop = 2, BaseValue = 427.0, Slope = 3.6 },
                    new { Stop = 3, BaseValue = 418.0, Slope = 5.8 },
                    new { Stop = 4, BaseValue = 520.0, Slope = 6.2 }
                };

                foreach (var calc in delayCalculations)
                {
                    // 檢查是否已存在此參數
                    var existingParam = allParameters.FirstOrDefault(p =>
                        p.Name == "delay" &&
                        (p.Stop ?? 0) == calc.Stop &&
                        (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified));

                    if (existingParam != null)
                    {
                        // 已存在則跳過
                        continue;
                    }

                    // 線性預測公式
                    double delayValue = calc.BaseValue + (od - 35.0) * calc.Slope;
                    int roundedValue = (int)Math.Round(delayValue);

                    // 新增到「已新增已修改區」（因為是計算得出的新值）
                    allParameters.Add(new ParameterItem
                    {
                        Type = TargetType,
                        Name = "delay",
                        Value = roundedValue.ToString(),
                        Stop = calc.Stop,
                        ChineseName = "拍照延遲時間",
                        Zone = ParameterZone.AddedModified,
                        IsSelected = false
                    });

                    processedParameters.Add($"delay_{calc.Stop} = {roundedValue} ms ← OD線性計算");
                    count++;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProcessDelayParameters] 計算 delay 參數時發生錯誤：{ex.Message}");
            }

            return count;
        }

        // 由 GitHub Copilot 產生 - 處理 fourTo 系列參數（從來源料號複製，寫入 params 資料表）
        private int ProcessFourToParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            // 確認有來源料號
            if (currentSession == null || string.IsNullOrEmpty(currentSession.SourceType))
            {
                System.Diagnostics.Debug.WriteLine("[ProcessFourToParameters] 無來源料號，無法複製 fourTo 參數");
                return 0;
            }

            try
            {
                string[] fourToParamNames = { "fourToOK1_time_ms", "fourToOK2_time_ms", "fourToNG_time_ms", "fourToNULL_time_ms" };
                var fourToDescriptions = new Dictionary<string, string>
                {
                    { "fourToOK1_time_ms", "站4到OK1出料時間" },
                    { "fourToOK2_time_ms", "站4到OK2出料時間" },
                    { "fourToNG_time_ms", "站4到NG出料時間" },
                    { "fourToNULL_time_ms", "站4到NULL出料時間" }
                };

                // 預設值（當來源料號無參數時使用）
                var fourToDefaults = new Dictionary<string, string>
                {
                    { "fourToOK1_time_ms", "1250" },
                    { "fourToOK2_time_ms", "1920" },
                    { "fourToNG_time_ms", "4500" },
                    { "fourToNULL_time_ms", "5000" }
                };

                using (var db = new MydbDB())
                {
                    foreach (var paramName in fourToParamNames)
                    {
                        // 檢查是否已存在此參數
                        var existingParam = allParameters.FirstOrDefault(p =>
                            p.Name == paramName &&
                            (p.Stop ?? 0) == 4 &&
                            (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified));

                        if (existingParam != null)
                        {
                            // 已存在則跳過
                            continue;
                        }

                        // 從來源料號的 params 資料表讀取
                        var sourceParam = db.@params
                            .Where(p => p.Type == currentSession.SourceType && 
                                       p.Name == paramName && 
                                       (p.Stop ?? 0) == 4)
                            .FirstOrDefault();

                        string value;
                        string source;

                        if (sourceParam != null)
                        {
                            value = sourceParam.Value;
                            source = "來源料號複製";
                        }
                        else
                        {
                            // 使用預設值
                            value = fourToDefaults[paramName];
                            source = "預設值";
                        }

                        // 新增到「已新增未修改區」（因為是從來源複製，不是計算）
                        allParameters.Add(new ParameterItem
                        {
                            Type = TargetType,
                            Name = paramName,
                            Value = value,
                            Stop = 4,
                            ChineseName = fourToDescriptions[paramName],
                            Zone = ParameterZone.AddedUnmodified,
                            IsSelected = false
                        });

                        processedParameters.Add($"{paramName} = {value} ← {source}");
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProcessFourToParameters] 複製 fourTo 參數時發生錯誤：{ex.Message}");
            }

            return count;
        }

        // 由 GitHub Copilot 產生 - 處理位置參數（known 開頭的所有參數，從來源料號直接複製）
        private int ProcessPositionParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            if (currentSession == null || string.IsNullOrEmpty(currentSession.SourceType))
            {
                return 0;
            }

            // 從參考區找出所有 known 開頭的參數名稱（不含站點後綴）
            var knownParamBaseNames = allParameters
                .Where(p => p.Zone == ParameterZone.Reference && 
                           p.Name != null && 
                           p.Name.StartsWith("known", StringComparison.OrdinalIgnoreCase))
                .Select(p => p.Name)
                .Distinct()
                .ToList();

            if (knownParamBaseNames.Count == 0)
            {
                return 0;
            }

            // 使用現有的直接複製方法，不覆寫已存在的參數
            foreach (var paramName in knownParamBaseNames)
            {
                // 找出該參數在參考區的所有站點
                var refParams = allParameters
                    .Where(p => p.Zone == ParameterZone.Reference && p.Name == paramName)
                    .ToList();

                foreach (var refParam in refParams)
                {
                    int stop = refParam.Stop ?? 0;

                    // 檢查目標料號是否已經有此參數
                    var existingParam = allParameters.FirstOrDefault(p =>
                        p.Name == paramName &&
                        (p.Stop ?? 0) == stop &&
                        (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified));

                    if (existingParam != null)
                    {
                        // 已存在則跳過（不覆寫）
                        continue;
                    }

                    // 新增到「已新增未修改區」
                    allParameters.Add(new ParameterItem
                    {
                        Type = TargetType,
                        Name = refParam.Name,
                        Value = refParam.Value,
                        Stop = refParam.Stop,
                        ChineseName = refParam.ChineseName,
                        Zone = ParameterZone.AddedUnmodified,
                        IsSelected = false
                    });

                    string stopSuffix = stop > 0 ? $"_{stop}" : "";
                    processedParameters?.Add($"{paramName}{stopSuffix} ← 來源料號直接複製");
                    count++;
                }
            }

            return count;
        }

        // 由 GitHub Copilot 產生 - 處理 DefectChecks 資料表（缺陷檢測項目，從來源料號複製到目標料號）
        private int ProcessDefectChecksParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;
            var skippedItems = new List<string>();

            if (currentSession == null || string.IsNullOrEmpty(currentSession.SourceType))
            {
                return 0;
            }

            try
            {
                // 定義固定導入清單（name 包含 "deform" 或 name = "minArea"）
                var fixedImportNames = new List<string> { "minArea" };

                // 條件導入清單：若 hasYZP == "有"，加入 OTP、OTPratio、blackDot
                var conditionalImportNames = new List<string>();
                if (typeInfo.HasYZP == "有")
                {
                    conditionalImportNames.AddRange(new[] { "OTP", "OTPratio", "blackDot" });
                }

                using (var db = new MydbDB())
                {
                    // 從來源料號讀取 DefectChecks（所有 stop）
                    var sourceDefectChecks = db.DefectChecks
                        .Where(d => d.Type == currentSession.SourceType)
                        .ToList();

                    // 篩選符合條件的項目
                    var itemsToImport = sourceDefectChecks.Where(d =>
                        fixedImportNames.Contains(d.Name, StringComparer.OrdinalIgnoreCase) ||
                        d.Name.IndexOf("deform", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        conditionalImportNames.Contains(d.Name, StringComparer.OrdinalIgnoreCase)
                    ).ToList();

                    // 檢查固定導入項目是否存在於來源
                    foreach (var name in fixedImportNames)
                    {
                        if (!sourceDefectChecks.Any(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                        {
                            skippedItems.Add($"{name} ← 來源料號不存在，跳過");
                        }
                    }

                    // 檢查 deform 項目
                    if (!sourceDefectChecks.Any(d => d.Name.IndexOf("deform", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        skippedItems.Add($"deform* ← 來源料號不存在，跳過");
                    }

                    // 檢查條件導入項目是否存在於來源
                    foreach (var name in conditionalImportNames)
                    {
                        if (!sourceDefectChecks.Any(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                        {
                            skippedItems.Add($"{name} ← 來源料號不存在，跳過");
                        }
                    }

                    // 讀取目標料號已存在的 DefectChecks
                    var existingTargetChecks = db.DefectChecks
                        .Where(d => d.Type == TargetType)
                        .ToList();

                    foreach (var item in itemsToImport)
                    {
                        // 檢查目標料號是否已存在相同 name + stop 的項目
                        bool exists = existingTargetChecks.Any(d =>
                            d.Name == item.Name && d.Stop == item.Stop);

                        if (exists)
                        {
                            // 已存在則跳過
                            continue;
                        }

                        // 使用 update-first, insert-if-none 模式寫入資料庫
                        int updatedRows = db.DefectChecks
                            .Where(d => d.Type == TargetType && d.Name == item.Name && d.Stop == item.Stop)
                            .Set(d => d.Threshold, item.Threshold)
                            .Set(d => d.Yn, item.Yn)
                            .Set(d => d.ChineseName, item.ChineseName)
                            .Update();

                        if (updatedRows == 0)
                        {
                            db.DefectChecks
                                .Value(d => d.Type, TargetType)
                                .Value(d => d.Stop, item.Stop)
                                .Value(d => d.Name, item.Name)
                                .Value(d => d.Threshold, item.Threshold)
                                .Value(d => d.Yn, item.Yn)
                                .Value(d => d.ChineseName, item.ChineseName)
                                .Insert();
                        }

                        processedParameters?.Add($"[DefectCheck] {item.Name}_{item.Stop} ← 來源料號複製");
                        count++;
                    }
                }

                // 將跳過的項目加入處理結果清單
                foreach (var skipped in skippedItems)
                {
                    processedParameters?.Add($"[DefectCheck] {skipped}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ProcessDefectChecksParameters] 處理 DefectChecks 時發生錯誤：{ex.Message}");
            }

            return count;
        }

        // 由 GitHub Copilot 產生 - 根據材質處理參數
        private int ProcessMaterialBasedParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            // 若目標料號與來源料號材質相同，則複製來源料號的相機參數 (exposure, gain)
            if (!string.IsNullOrEmpty(typeInfo.Material) && 
                currentSession != null && 
                !string.IsNullOrEmpty(currentSession.SourceType))
            {
                try
                {
                    // 取得來源料號的 TypeInfo
                    var sourceTypeInfo = GetTypeInfo(currentSession.SourceType);
                    if (sourceTypeInfo != null && 
                        !string.IsNullOrEmpty(sourceTypeInfo.Material) &&
                        typeInfo.Material == sourceTypeInfo.Material)
                    {
                        // 材質相同，從資料庫讀取來源料號的相機參數並複製
                        using (var db = new MydbDB())
                        {
                            var sourceCameraParams = db.Cameras
                                .Where(c => c.Type == currentSession.SourceType &&
                                           (c.Name == "exposure" || c.Name == "gain"))
                                .ToList();

                            foreach (var cam in sourceCameraParams)
                            {
                                // 檢查目標料號是否已有此參數（避免重複新增）
                                var existingParam = allParameters.FirstOrDefault(p =>
                                    p.Name == cam.Name &&
                                    (p.Stop ?? 0) == cam.Stop &&
                                    (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified));

                                if (existingParam != null)
                                {
                                    // 已存在則跳過（不覆寫）
                                    continue;
                                }

                                // 新增為「已新增未修改區」
                                allParameters.Add(new ParameterItem
                                {
                                    Type = TargetType,
                                    Name = cam.Name,
                                    Value = cam.Value,
                                    Stop = cam.Stop,
                                    ChineseName = cam.ChineseName ?? (cam.Name == "exposure" ? "曝光時間" : "增益"),
                                    Zone = ParameterZone.AddedUnmodified,
                                    IsSelected = false
                                });

                                processedParameters.Add($"{cam.Name}_{cam.Stop} ← 來源料號(PTFE相同) 複製相機參數");
                                count++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ProcessMaterialBasedParameters] 複製相機參數時發生錯誤：{ex.Message}");
                }
            }

            return count;
        }

        // 由 GitHub Copilot 產生 - 根據尺寸處理參數
        private int ProcessSizeBasedParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                // 根據 OD 尺寸調整相關參數
                if (typeInfo.OD > 0)
                {
                    // 範例：大尺寸料號可能需要不同的檢測範圍
                    if (typeInfo.OD > 40)
                    {
                        // 處理大尺寸專用參數
                        // 可以在這裡添加具體邏輯
                    }
                    else if (typeInfo.OD < 30)
                    {
                        // 處理小尺寸專用參數
                        // 可以在這裡添加具體邏輯
                    }
                }

                return count;
            }
            catch (Exception ex)
            {
                throw new Exception($"處理尺寸相關參數時發生錯誤：{ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生 - 根據厚度處理參數
        private int ProcessThicknessBasedParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                // 根據厚度調整相關參數
                if (typeInfo.Thick > 0)
                {
                    // 範例：厚料號可能需要不同的檢測參數
                    // 可以在這裡添加具體邏輯
                }

                return count;
            }
            catch (Exception ex)
            {
                throw new Exception($"處理厚度相關參數時發生錯誤：{ex.Message}");
            }
        }
        // 由 GitHub Copilot 產生
        private int ProcessYZPParameters(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                // 判斷是否有五彩鋅（容錯：有/1/true/yes 視為有）
                bool hasYZP = false;
                string v = typeInfo.HasYZP?.Trim();
                if (!string.IsNullOrEmpty(v))
                {
                    hasYZP = v == "有" ||
                             v == "1" ||
                             v.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                             v.Equals("yes", StringComparison.OrdinalIgnoreCase);
                }

                // 站1/2：一律 0；站3/4：依 hasYZP 決定 color
                for (int stop = 1; stop <= 4; stop++)
                {
                    string value;
                    string reason;

                    if (stop <= 2)
                    {
                        value = "0";
                        reason = "站1/2一律灰階(0)";
                    }
                    else
                    {
                        value = hasYZP ? "1" : "0";
                        reason = hasYZP ? "有五彩鋅 → 設為彩色(1)" : "無五彩鋅 → 設為灰階(0)";
                    }

                    UpdateOrAddParameterToModified(
                        "color",
                        value,
                        stop,
                        "ROI色彩模式",
                        reason
                    );

                    processedParameters.Add($"color_{stop} = {value}");
                    count++;
                }

                // 有五彩鋅時，站3/4的 OTPratio 設為 0.1
                if (hasYZP)
                {
                    for (int stop = 3; stop <= 4; stop++)
                    {
                        UpdateOrAddParameterToModified(
                            "OTPratio",
                            "0.1",
                            stop,
                            "OTP占比",
                            "有五彩鋅 → 設為 0.1"
                        );
                        processedParameters.Add($"OTPratio_{stop} = 0.1");
                        count++;
                    }
                }

                return count;
            }
            catch (Exception ex)
            {
                throw new Exception($"處理五彩鋅(color/OTPratio)參數時發生錯誤：{ex.Message}");
            }
        }

        private int ProcessOutscParametersUsingDefectChecks(TypeInfo typeInfo, List<string> processedParameters)
        {
            int count = 0;

            try
            {
                using (var db = new MydbDB())
                {
                    // 個別判斷站3、站4是否在 DefectChecks 內有 outsc
                    bool hasOutsc3 = db.DefectChecks.Any(dc =>
                        dc.Type == TargetType &&
                        dc.Name == "outsc" &&
                        dc.Stop == 3);

                    bool hasOutsc4 = db.DefectChecks.Any(dc =>
                        dc.Type == TargetType &&
                        dc.Name == "outsc" &&
                        dc.Stop == 4);

                    // 站3：有 outsc 才寫入
                    if (hasOutsc3)
                    {
                        UpdateOrAddParameterToModified("outscLen", "800", 3, "金屬面刮痕長度門檻", "依 DefectChecks(outsc, stop=3) 設定");

                        processedParameters.Add("outscLen_3 = 800");

                        count += 1;
                    }

                    // 站4：有 outsc 才寫入
                    if (hasOutsc4)
                    {
                        UpdateOrAddParameterToModified("outscLen", "800", 4, "金屬面刮痕長度門檻", "依 DefectChecks(outsc, stop=4) 設定");

                        processedParameters.Add("outscLen_4 = 800");

                        count += 1;
                    }
                }

                return count;
            }
            catch (Exception ex)
            {
                throw new Exception($"處理 outsc 參數時發生錯誤：{ex.Message}");
            }
        }
        private int ProcessDirectCopyFromSourceParameters(IEnumerable<string> parameterBaseNames,
                                                 List<string> processedParameters,
                                                 bool overwriteExisting = false,
                                                 IEnumerable<int> targetStops = null)
        {
            if (parameterBaseNames == null) return 0;

            int count = 0;
            var stops = (targetStops == null) ? new[] { 1, 2, 3, 4 } : targetStops.ToArray();

            foreach (var name in parameterBaseNames.Where(n => !string.IsNullOrWhiteSpace(n)))
            {
                foreach (var stop in stops)
                {
                    // 找來源料號的參考區參數
                    var refItem = allParameters.FirstOrDefault(p =>
                        p.Zone == ParameterZone.Reference &&
                        p.Name == name &&
                        (p.Stop ?? 0) == stop);

                    if (refItem == null)
                        continue; // 來源沒有這個站點/參數就跳過

                    // 目標料號是否已經有此參數（已新增未修改或已修改）
                    var exist = allParameters.FirstOrDefault(p =>
                        p.Name == name &&
                        (p.Stop ?? 0) == stop &&
                        (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified));

                    if (exist != null && !overwriteExisting)
                        continue; // 已存在且不覆寫 → 跳過

                    if (exist != null && overwriteExisting)
                    {
                        // 覆寫既有值，並標記為已新增未修改（代表是直接複製的）
                        exist.Value = refItem.Value;
                        exist.ChineseName = refItem.ChineseName;
                        exist.Zone = ParameterZone.AddedUnmodified;
                    }
                    else
                    {
                        // 新增為「已新增未修改區」
                        allParameters.Add(new ParameterItem
                        {
                            Type = TargetType,
                            Name = refItem.Name,
                            Value = refItem.Value,
                            Stop = refItem.Stop,
                            ChineseName = refItem.ChineseName,
                            Zone = ParameterZone.AddedUnmodified,
                            IsSelected = false
                        });
                    }

                    processedParameters?.Add($"{name}_{stop} ← 來源料號(參考區) 直接複製");
                    count++;
                }
            }

            return count;
        }
        // 由 GitHub Copilot 產生 - 更新或新增參數到已修改區
        private void UpdateOrAddParameterToModified(string name, string value, int? stop,
            string chineseName, string processingReason = "")
        {
            // 尋找現有參數
            var existingParam = allParameters.FirstOrDefault(p =>
                p.Name == name && p.Stop == stop &&
                (p.Zone == ParameterZone.AddedModified || p.Zone == ParameterZone.AddedUnmodified));

            if (existingParam != null)
            {
                // 更新現有參數
                existingParam.Value = value;
                existingParam.Zone = ParameterZone.AddedModified;
                if (!string.IsNullOrEmpty(processingReason))
                {
                    existingParam.ChineseName = $"{chineseName} ({processingReason})";
                }
            }
            else
            {
                // 新增新參數到已修改區
                allParameters.Add(new ParameterItem
                {
                    Type = TargetType,
                    Name = name,
                    Value = value,
                    Stop = stop,
                    ChineseName = string.IsNullOrEmpty(processingReason) ? chineseName : $"{chineseName} ({processingReason})",
                    Zone = ParameterZone.AddedModified,
                    IsSelected = false
                });
            }
        }

        // 由 GitHub Copilot 產生 - 顯示處理結果
        private void ShowPartialAutoImportResult(int processedCount, List<string> processedParameters)
        {
            if (processedCount > 0)
            {
                string resultMessage = $"✅ 部分自動導入完成\n\n" +
                                      $"處理了 {processedCount} 個參數：\n\n" +
                                      string.Join("\n", processedParameters);

                MessageBox.Show(resultMessage, "部分自動導入完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("沒有找到需要處理的條件式參數", "部分自動導入完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 由 GitHub Copilot 產生 - 料號資訊資料類別
        public class TypeInfo
        {
            public string TypeColumn { get; set; }
            public string Material { get; set; }
            public double Thick { get; set; }
            public string PTFEColor { get; set; }
            public double ID { get; set; }
            public double OD { get; set; }
            public double H { get; set; }
            public string HasGroove { get; set; }
            public string BoxOrPack { get; set; }
            public string HasYZP { get; set; }
            public string package { get; set; }
        }
        #endregion

        public string TargetType { get; private set; }

        private void InitializeForm()
        {
            this.Text = $"新料號參數設定 - {TargetType}";
            this.Size = new Size(1600, 1050);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = true;
            
            // 建立部分自動導入按鈕
            CreatePartialAutoImportButton();

            // 【修正1】：先更新進度顯示，但不要更新Tab狀態（避免TabControl還沒建立完成）
            UpdateProgressDisplayOnly();

            if (File.Exists(sessionFilePath))
            {
                var result = MessageBox.Show(
                    $"發現未完成的參數設定工作階段，是否要繼續？\n\n點擊「是」繼續之前的工作\n點擊「否」重新選擇來源料號（一樣會導入已設定完成的參數）",
                    "發現未完成工作",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    LoadExistingSession();
                }
                else
                {
                    ShowSourceSelectionDialog();
                }
            }
            else
            {
                ShowSourceSelectionDialog();
            }

            // 初始化Tab狀態
            // 確保 TabControl 已經完全初始化後才更新Tab狀態
            this.Load += (s, e) => UpdateTabStates();
        }

        // 由 GitHub Copilot 產生 - 修正進度顯示（AddedUnmodified + AddedModified 都計入）
        private void UpdateProgressDisplayOnly()
        {
            if (allParameters == null)
            {
                lblProgressSummary.Text = "";
                return;
            }

            // 分母：所有參考區參數數量
            int totalReferenceCount = allParameters.Count(p => p.Zone == ParameterZone.Reference);

            // 分子：已新增的參數（未修改 + 已修改）且有對應參考區
            int totalAddedCount = allParameters
                .Where(p => p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified)
                .Count(added => allParameters.Any(r =>
                    r.Zone == ParameterZone.Reference &&
                    r.Name == added.Name &&
                    (r.Stop ?? 0) == (added.Stop ?? 0)));

            int overallProgress = totalReferenceCount > 0
                ? (int)Math.Round((double)totalAddedCount / totalReferenceCount * 100)
                : 0;
            overallProgress = Math.Max(0, Math.Min(100, overallProgress));
            progressBarOverall.Value = overallProgress;

            var parts = new List<string>();
            foreach (ParameterCategory category in Enum.GetValues(typeof(ParameterCategory)))
            {
                var categoryParams = parameterManager.GetParametersByCategory(allParameters, category);
                int referenceCount = categoryParams.Count(p => p.Zone == ParameterZone.Reference);
                if (referenceCount == 0) continue;

                // 已新增的參數（未修改 + 已修改）都計入進度
                int addedCount = categoryParams
                    .Where(p => p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified)
                    .Count(added => categoryParams.Any(r =>
                        r.Zone == ParameterZone.Reference &&
                        r.Name == added.Name &&
                        (r.Stop ?? 0) == (added.Stop ?? 0)));

                int percent = referenceCount > 0
                    ? (int)Math.Round((double)addedCount / referenceCount * 100)
                    : 0;
                string categoryName = GetCategoryDisplayName(category);
                parts.Add($"{categoryName}{percent}% ({addedCount}/{referenceCount})");
            }

            lblProgressSummary.Text = string.Join("  |  ", parts);
            lblProgress.Text = $"料號: {TargetType}  總進度 {overallProgress}% ({totalAddedCount}/{totalReferenceCount})";
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
        private void ShowSourceSelectionDialog()
        {
            using (var dialog = new SourceSelectionDialog(TargetType))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    CreateNewSession(dialog.SelectedSourceType);
                    LoadParametersFromSource(dialog.SelectedSourceType);
                    RefreshCurrentTabDisplay();
                }
                else
                {
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                }
            }
        }

        private void CreateNewSession(string sourceType)
        {
            currentSession = new ParameterSession
            {
                SessionId = $"{DateTime.Now:yyyyMMdd}_{TargetType}",
                CreatedTime = DateTime.Now,
                SourceType = sourceType,
                TargetType = TargetType,
                Parameters = new List<ParameterItem>()
            };

            allParameters = new List<ParameterItem>();
            SaveSession();
        }

        // 由 GitHub Copilot 產生 - 從來源料號載入參數作為參考區
        // 由 GitHub Copilot 產生 - 修正來源參數載入邏輯

        /// <summary>
        /// 從來源料號載入參數
        /// </summary>
        /// <param name="sourceType">來源料號</param>
        private void LoadParametersFromSource(string sourceType)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"========================================");
                System.Diagnostics.Debug.WriteLine($"[開始載入] 來源料號: {sourceType}");
                System.Diagnostics.Debug.WriteLine($"[目標料號] {TargetType}");
                System.Diagnostics.Debug.WriteLine($"========================================");

                using (var db = new MydbDB())
                {
                    // ✅ 步驟1：從來源料號載入 Camera 參數
                    var sourceCameraParams = db.Cameras
                        .Where(c => c.Type == sourceType)
                        .Select(c => new ParameterItem
                        {
                            Type = TargetType,  // 標記為目標料號
                    Name = c.Name,
                            Value = c.Value,
                            Stop = c.Stop,
                            ChineseName = c.ChineseName,
                            Zone = ParameterZone.Reference,  // 放入參考區
                    IsSelected = false
                        }).ToList();

                    System.Diagnostics.Debug.WriteLine($"[載入完成] Camera 參數: {sourceCameraParams.Count} 個");

                    // ✅ 步驟2：從來源料號載入 Param 參數
                    var sourceParamParams = db.@params
                        .Where(p => p.Type == sourceType)
                        .Select(p => new ParameterItem
                        {
                            Type = TargetType,
                            Name = p.Name,
                            Value = p.Value,
                            Stop = p.Stop,
                            ChineseName = p.ChineseName,
                            Zone = ParameterZone.Reference,
                            IsSelected = false
                        }).ToList();

                    System.Diagnostics.Debug.WriteLine($"[載入完成] Param 參數: {sourceParamParams.Count} 個");

                    // ✅ 步驟3：合併參數到 allParameters
                    allParameters = new List<ParameterItem>();
                    allParameters.AddRange(sourceCameraParams);
                    allParameters.AddRange(sourceParamParams);

                    System.Diagnostics.Debug.WriteLine($"[合併完成] 參考區總參數: {allParameters.Count} 個");

                    // ✅ 步驟4：檢查目標料號是否已存在於資料庫（修正關鍵）
                    bool targetExists = IsTargetTypeExistInDatabase();

                    if (targetExists)
                    {
                        System.Diagnostics.Debug.WriteLine($"[檢測到] 目標料號 {TargetType} 已有參數，載入既有資料到已修改區");
                        CheckAndLoadExistingTargetParameters();
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[確認] 目標料號 {TargetType} 為新料號，跳過既有參數檢查");
                    }

                    // ✅ 步驟5：儲存工作階段
                    currentSession.Parameters = allParameters;
                    currentSession.SourceType = sourceType; // 確保記錄來源料號
                    SaveSession();

                    // ✅ 步驟6：重新計算進度
                    RecalculateProgressAndUpdateUI();

                    // ✅ 步驟7：輸出最終統計
                    System.Diagnostics.Debug.WriteLine($"========================================");
                    System.Diagnostics.Debug.WriteLine($"[載入總結]");
                    System.Diagnostics.Debug.WriteLine($"  參考區: {allParameters.Count(p => p.Zone == ParameterZone.Reference)} 個");
                    System.Diagnostics.Debug.WriteLine($"  已新增未修改區: {allParameters.Count(p => p.Zone == ParameterZone.AddedUnmodified)} 個");
                    System.Diagnostics.Debug.WriteLine($"  已新增已修改區: {allParameters.Count(p => p.Zone == ParameterZone.AddedModified)} 個");
                    System.Diagnostics.Debug.WriteLine($"========================================");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[錯誤] 載入來源參數時發生錯誤: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[錯誤堆疊] {ex.StackTrace}");
                MessageBox.Show($"載入來源參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 由 GitHub Copilot 產生 - 檢查目標料號是否已存在於資料庫

        /// <summary>
        /// 檢查目標料號是否已存在於資料庫中
        /// </summary>
        /// <returns>若資料庫中已有此料號的參數則返回 true，否則返回 false</returns>
        private bool IsTargetTypeExistInDatabase()
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 檢查 Cameras 表中是否有此料號
                    bool existsInCameras = db.Cameras
                        .Where(c => c.Type == TargetType)
                        .Any();

                    // 檢查 @params 表中是否有此料號
                    bool existsInParams = db.@params
                        .Where(p => p.Type == TargetType)
                        .Any();

                    bool exists = existsInCameras || existsInParams;

                    System.Diagnostics.Debug.WriteLine($"[資料庫檢查] 料號 {TargetType} 是否存在: {exists}");
                    System.Diagnostics.Debug.WriteLine($"  - Cameras 表: {existsInCameras}");
                    System.Diagnostics.Debug.WriteLine($"  - Params 表: {existsInParams}");

                    return exists;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[錯誤] 檢查料號是否存在時發生錯誤: {ex.Message}");
                // 發生錯誤時，為安全起見返回 false（視為新料號）
                return false;
            }
        }
        // 由 GitHub Copilot 產生 - 修正已存在參數的數值載入
        // 由 GitHub Copilot 產生 - 加強既有參數載入的除錯資訊

        private void CheckAndLoadExistingTargetParameters()
        {
            // 由 GitHub Copilot 產生 - 加入除錯輸出
            System.Diagnostics.Debug.WriteLine($"[CheckAndLoadExistingTargetParameters] 開始檢查目標料號 {TargetType} 的既有參數");

            try
            {
                using (var db = new MydbDB())
                {
                    // 查詢目標料號的 Camera 參數
                    var targetCameraParams = db.Cameras
                        .Where(c => c.Type == TargetType)
                        .ToList();

                    // 由 GitHub Copilot 產生 - 輸出查詢結果
                    System.Diagnostics.Debug.WriteLine($"[查詢結果] 目標料號 Camera 參數: {targetCameraParams.Count} 個");

                    // 查詢目標料號的 Param 參數
                    var targetParamParams = db.@params
                        .Where(p => p.Type == TargetType)
                        .ToList();

                    // 由 GitHub Copilot 產生 - 輸出查詢結果
                    System.Diagnostics.Debug.WriteLine($"[查詢結果] 目標料號 Param 參數: {targetParamParams.Count} 個");

                    int addedToModified = 0;
                    int addedToUnmodified = 0;

                    // 處理 Camera 參數
                    foreach (var targetCamera in targetCameraParams)
                    {
                        // 檢查參考區是否有對應參數
                        var referenceParam = allParameters.FirstOrDefault(p =>
                            p.Name == targetCamera.Name &&
                            p.Stop == targetCamera.Stop &&
                            p.Zone == ParameterZone.Reference);

                        if (referenceParam != null)
                        {
                            // 參考區有對應參數，比較數值
                            if (referenceParam.Value != targetCamera.Value)
                            {
                                // 數值不同，加入「已修改區」
                                allParameters.Add(new ParameterItem
                                {
                                    Type = TargetType,
                                    Name = targetCamera.Name,
                                    Value = targetCamera.Value,
                                    Stop = targetCamera.Stop,
                                    ChineseName = targetCamera.ChineseName,
                                    Zone = ParameterZone.AddedModified,
                                    IsSelected = false
                                });
                                addedToModified++;
                            }
                            else
                            {
                                // 數值相同，加入「已新增未修改區」
                                allParameters.Add(new ParameterItem
                                {
                                    Type = TargetType,
                                    Name = targetCamera.Name,
                                    Value = targetCamera.Value,
                                    Stop = targetCamera.Stop,
                                    ChineseName = targetCamera.ChineseName,
                                    Zone = ParameterZone.AddedUnmodified,
                                    IsSelected = false
                                });
                                addedToUnmodified++;
                            }
                        }
                        else
                        {
                            // 參考區沒有對應參數（目標料號特有），加入「已修改區」
                            allParameters.Add(new ParameterItem
                            {
                                Type = TargetType,
                                Name = targetCamera.Name,
                                Value = targetCamera.Value,
                                Stop = targetCamera.Stop,
                                ChineseName = targetCamera.ChineseName,
                                Zone = ParameterZone.AddedModified,
                                IsSelected = false
                            });
                            addedToModified++;
                        }
                    }

                    // 處理 Param 參數（邏輯相同）
                    foreach (var targetParam in targetParamParams)
                    {
                        var referenceParam = allParameters.FirstOrDefault(p =>
                            p.Name == targetParam.Name &&
                            p.Stop == targetParam.Stop &&
                            p.Zone == ParameterZone.Reference);

                        if (referenceParam != null)
                        {
                            if (referenceParam.Value != targetParam.Value)
                            {
                                allParameters.Add(new ParameterItem
                                {
                                    Type = TargetType,
                                    Name = targetParam.Name,
                                    Value = targetParam.Value,
                                    Stop = targetParam.Stop,
                                    ChineseName = targetParam.ChineseName,
                                    Zone = ParameterZone.AddedModified,
                                    IsSelected = false
                                });
                                addedToModified++;
                            }
                            else
                            {
                                allParameters.Add(new ParameterItem
                                {
                                    Type = TargetType,
                                    Name = targetParam.Name,
                                    Value = targetParam.Value,
                                    Stop = targetParam.Stop,
                                    ChineseName = targetParam.ChineseName,
                                    Zone = ParameterZone.AddedUnmodified,
                                    IsSelected = false
                                });
                                addedToUnmodified++;
                            }
                        }
                        else
                        {
                            allParameters.Add(new ParameterItem
                            {
                                Type = TargetType,
                                Name = targetParam.Name,
                                Value = targetParam.Value,
                                Stop = targetParam.Stop,
                                ChineseName = targetParam.ChineseName,
                                Zone = ParameterZone.AddedModified,
                                IsSelected = false
                            });
                            addedToModified++;
                        }
                    }

                    // 由 GitHub Copilot 產生 - 輸出處理統計
                    System.Diagnostics.Debug.WriteLine($"[處理完成] 加入已修改區: {addedToModified} 個");
                    System.Diagnostics.Debug.WriteLine($"[處理完成] 加入已新增未修改區: {addedToUnmodified} 個");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[錯誤] CheckAndLoadExistingTargetParameters 發生錯誤: {ex.Message}");
                MessageBox.Show($"載入目標料號參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 由 GitHub Copilot 產生 - 修正用戶新增參數載入方法
        private void LoadUserAddedParameters(List<Camera> existingCameras, List<Param> existingParams)
        {
            // 載入用戶新增的Camera參數
            foreach (var camera in existingCameras)
            {
                bool existsInLoaded = allParameters.Any(p =>
                    p.Name == camera.Name && p.Stop == camera.Stop);

                if (!existsInLoaded)
                {
                    allParameters.Add(new ParameterItem
                    {
                        Type = TargetType,
                        Name = camera.Name,
                        Value = camera.Value, // 🔸 確保使用資料庫中的實際值
                        Stop = camera.Stop,
                        ChineseName = camera.ChineseName,
                        Zone = ParameterZone.AddedModified, // 已存在資料庫中，視為已修改
                        IsSelected = false
                    });

                    // 輸出除錯訊息
                    System.Diagnostics.Debug.WriteLine($"新增Camera參數: {camera.Name}_{camera.Stop} = {camera.Value} (僅存在於資料庫)");
                }
            }

            // 載入用戶新增的Param參數
            foreach (var param in existingParams)
            {
                bool existsInLoaded = allParameters.Any(p =>
                    p.Name == param.Name && p.Stop == param.Stop);

                if (!existsInLoaded)
                {
                    allParameters.Add(new ParameterItem
                    {
                        Type = TargetType,
                        Name = param.Name,
                        Value = param.Value, // 🔸 確保使用資料庫中的實際值
                        Stop = param.Stop,
                        ChineseName = param.ChineseName,
                        Zone = ParameterZone.AddedModified, // 已存在資料庫中，視為已修改
                        IsSelected = false
                    });

                    // 輸出除錯訊息
                    System.Diagnostics.Debug.WriteLine($"新增Param參數: {param.Name}_{param.Stop} = {param.Value} (僅存在於資料庫)");
                }
            }
        }

        // 由 GitHub Copilot 產生 - 修正載入既有工作階段的邏輯
        private void LoadExistingSession()
        {
            try
            {
                var sessionJson = File.ReadAllText(sessionFilePath);
                currentSession = JsonConvert.DeserializeObject<ParameterSession>(sessionJson);
                allParameters = currentSession.Parameters ?? new List<ParameterItem>();

                // 🔸 修正：使用新的參數同步邏輯，確保與資料庫狀態一致
                SyncSessionParametersWithDatabase();

                UpdateCategoryProgress();
                RefreshCurrentTabDisplay();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入工作階段失敗：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowSourceSelectionDialog();
            }
        }

        // 由 GitHub Copilot 產生 - 同步工作階段參數與資料庫狀態
        private void SyncSessionParametersWithDatabase()
        {
            try
            {
                using (var db = new MydbDB())
                {
                    var dbCameraParams = db.Cameras.Where(c => c.Type == TargetType).ToList();
                    var dbParams = db.@params.Where(p => p.Type == TargetType).ToList();

                    foreach (var sessionParam in allParameters.Where(p => p.Zone == ParameterZone.AddedModified).ToList())
                    {
                        object dbParam = dbCameraParams.FirstOrDefault(c =>
                                            c.Name == sessionParam.Name && (c.Stop) == (sessionParam.Stop ?? 0))
                                         ?? (object)dbParams.FirstOrDefault(p =>
                                            p.Name == sessionParam.Name && (p.Stop ?? 0) == (sessionParam.Stop ?? 0));

                        if (dbParam is Camera camera)
                        {
                            sessionParam.Value = camera.Value;
                            sessionParam.ChineseName = camera.ChineseName;
                        }
                        else if (dbParam is Param param)
                        {
                            sessionParam.Value = param.Value;
                            sessionParam.ChineseName = param.ChineseName;
                        }
                        else
                        {
                            allParameters.Remove(sessionParam);
                        }
                    }

                    foreach (var cam in dbCameraParams)
                    {
                        if (!allParameters.Any(p => p.Name == cam.Name && (p.Stop ?? 0) == cam.Stop))
                        {
                            allParameters.Add(new ParameterItem
                            {
                                Type = TargetType,
                                Name = cam.Name,
                                Value = cam.Value,
                                Stop = cam.Stop,
                                ChineseName = cam.ChineseName,
                                Zone = ParameterZone.AddedModified
                            });
                        }
                    }
                    foreach (var prm in dbParams)
                    {
                        if (!allParameters.Any(p => p.Name == prm.Name && (p.Stop ?? 0) == (prm.Stop ?? 0)))
                        {
                            allParameters.Add(new ParameterItem
                            {
                                Type = TargetType,
                                Name = prm.Name,
                                Value = prm.Value,
                                Stop = prm.Stop,
                                ChineseName = prm.ChineseName,
                                Zone = ParameterZone.AddedModified
                            });
                        }
                    }

                    currentSession.Parameters = allParameters;
                    SaveSession();

                    // ★ 修正：同步後重算
                    RecalculateProgressAndUpdateUI();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"同步工作階段參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 由 GitHub Copilot 產生 - 確保工作階段儲存來源料號

        private void SaveSession()
        {
            try
            {
                if (!Directory.Exists(settingPath))
                {
                    Directory.CreateDirectory(settingPath);
                }

                // 由 GitHub Copilot 產生 - 確保 SourceType 不為空
                if (string.IsNullOrEmpty(currentSession.SourceType))
                {
                    System.Diagnostics.Debug.WriteLine($"[警告] SourceType 為空，無法追蹤來源料號");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[儲存工作階段] 來源料號: {currentSession.SourceType}");
                }

                string json = JsonConvert.SerializeObject(currentSession, Formatting.Indented);
                File.WriteAllText(sessionFilePath, json);

                System.Diagnostics.Debug.WriteLine($"[儲存成功] 工作階段已儲存至: {sessionFilePath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[錯誤] 儲存工作階段失敗: {ex.Message}");
                MessageBox.Show($"儲存工作階段時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 由 GitHub Copilot 產生 - 修正參數類別進度統計邏輯（AddedUnmodified + AddedModified 都計入）
        private void UpdateCategoryProgress()
        {
            foreach (ParameterCategory category in Enum.GetValues(typeof(ParameterCategory)))
            {
                var categoryParams = parameterManager.GetParametersByCategory(allParameters, category);

                int referenceCount = categoryParams.Count(p => p.Zone == ParameterZone.Reference);

                // 已新增的參數（未修改 + 已修改）都計入進度
                int addedCount = categoryParams
                    .Where(p => p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified)
                    .Count(added => categoryParams.Any(r =>
                        r.Zone == ParameterZone.Reference &&
                        r.Name == added.Name &&
                        (r.Stop ?? 0) == (added.Stop ?? 0)));

                parameterManager.UpdateCategoryProgress(category, addedCount, referenceCount);
                parameterManager.UpdateCategoryStatusFromDatabase(category, TargetType);
            }
        }

        private void UpdateProgressDisplay()
        {
            UpdateProgressDisplayOnly();
            UpdateSaveStatusDisplay();

            // 更新儲存狀態顯示
            UpdateSaveStatusDisplay();
        }

        private void UpdateSaveStatusDisplay()
        {
            // 在TabPage標題後加上儲存狀態
            for (int i = 0; i < tabMain.TabPages.Count && i < 4; i++)
            {
                ParameterCategory category = (ParameterCategory)i;
                bool isSaved = tabSavedStatus.ContainsKey(category) && tabSavedStatus[category];
                string baseTitle = GetTabBaseTitle(category);

                tabMain.TabPages[i].Text = isSaved ? $"{baseTitle} ✅" : baseTitle;
            }
        }

        private string GetTabBaseTitle(ParameterCategory category)
        {
            switch (category)
            {
                case ParameterCategory.Camera: return "相機參數";
                case ParameterCategory.Position: return "位置參數";
                case ParameterCategory.Detection: return "檢測參數";
                case ParameterCategory.Timing: return "時間參數";
                default: return category.ToString();
            }
        }

        private void UpdateTabStates()
        {
            // 【修正1】：確保 TabControl 和 TabPages 都已經建立
            if (tabMain == null || tabMain.TabPages.Count == 0)
                return;

            parameterManager.UpdateTabStates(tabMain);

            // 【修正2】：安全地檢查當前選中的Tab
            int selectedIndex = tabMain.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < tabMain.TabPages.Count)
            {
                if (!tabMain.TabPages[selectedIndex].Enabled)
                {
                    // 尋找第一個可用的Tab
                    for (int i = 0; i < tabMain.TabPages.Count; i++)
                    {
                        if (tabMain.TabPages[i].Enabled)
                        {
                            tabMain.SelectedIndex = i;
                            break;
                        }
                    }
                }
            }
            else if (tabMain.TabPages.Count > 0)
            {
                // 【修正3】：如果當前索引無效，設定為第一個可用的Tab
                for (int i = 0; i < tabMain.TabPages.Count; i++)
                {
                    if (tabMain.TabPages[i].Enabled)
                    {
                        tabMain.SelectedIndex = i;
                        break;
                    }
                }
            }
        }
        private void OnCategoryStatusChanged(object sender, ParameterCategoryCompletedEventArgs e)
        {
            UpdateProgressDisplay();
            UpdateTabStates();
        }

        // 由 GitHub Copilot 產生 - 加強顯示刷新的除錯資訊

        private void RefreshCurrentTabDisplay()
        {
            // 由 GitHub Copilot 產生 - 加入除錯輸出
            System.Diagnostics.Debug.WriteLine($"========================================");
            System.Diagnostics.Debug.WriteLine($"[RefreshCurrentTabDisplay] 當前分類: {currentCategory}");

            if (allParameters == null)
            {
                System.Diagnostics.Debug.WriteLine($"[警告] allParameters 為 null");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[參數總數] {allParameters.Count}");

            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
            System.Diagnostics.Debug.WriteLine($"[當前分類參數] {categoryParams.Count()} 個");

            var (dgvUnmodified, dgvModified, dgvFixed) = GetCurrentTabDataGridViews();

            if (dgvUnmodified == null || dgvModified == null || dgvFixed == null)
            {
                System.Diagnostics.Debug.WriteLine($"[錯誤] DataGridView 控制項為 null");
                return;
            }

            // 篩選各區參數
            var referenceParams = categoryParams.Where(p => p.Zone == ParameterZone.Reference).ToList();
            var unmodifiedParams = categoryParams.Where(p => p.Zone == ParameterZone.AddedUnmodified).ToList();
            var modifiedParams = categoryParams.Where(p => p.Zone == ParameterZone.AddedModified).ToList();

            // 由 GitHub Copilot 產生 - 輸出各區參數數量
            System.Diagnostics.Debug.WriteLine($"[參考區] {referenceParams.Count} 個參數");
            System.Diagnostics.Debug.WriteLine($"[已新增未修改區] {unmodifiedParams.Count} 個參數");
            System.Diagnostics.Debug.WriteLine($"[已新增已修改區] {modifiedParams.Count} 個參數");

            // 由 GitHub Copilot 產生 - 列出參考區的參數（前5個）
            if (referenceParams.Any())
            {
                System.Diagnostics.Debug.WriteLine($"[參考區參數列表]（前5個）:");
                foreach (var p in referenceParams.Take(5))
                {
                    System.Diagnostics.Debug.WriteLine($"  - {p.Name}_{p.Stop} = {p.Value} ({p.ChineseName})");
                }
            }

            // 刷新 DataGridView
            RefreshDataGridView(dgvUnmodified, referenceParams);
            RefreshDataGridView(dgvModified, unmodifiedParams);
            RefreshDataGridView(dgvFixed, modifiedParams);

            UpdateCategoryProgress();
            UpdateProgressDisplay();

            System.Diagnostics.Debug.WriteLine($"[RefreshCurrentTabDisplay] 完成");
            System.Diagnostics.Debug.WriteLine($"========================================");
        }

        private (DataGridView, DataGridView, DataGridView) GetCurrentTabDataGridViews()
        {
            switch (currentCategory)
            {
                case ParameterCategory.Camera:
                    return (dgvCameraUnmodified, dgvCameraModified, dgvCameraFixed);
                case ParameterCategory.Position:
                    return (dgvPositionUnmodified, dgvPositionModified, dgvPositionFixed);
                case ParameterCategory.Detection:
                    return (dgvDetectionUnmodified, dgvDetectionModified, dgvDetectionFixed);
                case ParameterCategory.Timing:
                    return (dgvTimingUnmodified, dgvTimingModified, dgvTimingFixed);
                default:
                    return (dgvCameraUnmodified, dgvCameraModified, dgvCameraFixed);
            }
        }

        private void RefreshDataGridView(DataGridView dgv, List<ParameterItem> parameters)
        {
            var sortedParams = parameters.OrderBy(p => p.Name).ThenBy(p => p.Stop).ToList();

            dgv.DataSource = null;
            dgv.DataSource = sortedParams;

            if (dgv.Columns.Count > 0)
            {
                // 隱藏不需要的欄位
                dgv.Columns["Type"].Visible = false;
                dgv.Columns["Zone"].Visible = false;

                // 重新排列欄位順序，讓選取框在最左邊
                dgv.Columns["IsSelected"].DisplayIndex = 0;
                dgv.Columns["Name"].DisplayIndex = 4;
                dgv.Columns["Stop"].DisplayIndex = 2;
                dgv.Columns["Value"].DisplayIndex = 3;
                dgv.Columns["ChineseName"].DisplayIndex = 1;

                // 設定欄位屬性
                dgv.Columns["IsSelected"].HeaderText = "選取";
                dgv.Columns["IsSelected"].Width = 50;

                dgv.Columns["Name"].HeaderText = "參數名";
                dgv.Columns["Name"].ReadOnly = true;
                dgv.Columns["Name"].Width = 150;

                dgv.Columns["Stop"].HeaderText = "站點";
                dgv.Columns["Stop"].ReadOnly = true;
                dgv.Columns["Stop"].Width = 60;

                dgv.Columns["Value"].HeaderText = "參數值";
                dgv.Columns["Value"].Width = 100;

                dgv.Columns["ChineseName"].HeaderText = "中文名稱";
                dgv.Columns["ChineseName"].ReadOnly = true;
                dgv.Columns["ChineseName"].Width = 120;
            }
        }

        // Tab切換事件
        private void tabMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabMain.SelectedIndex >= 0 && tabMain.SelectedIndex < 4)
            {
                currentCategory = (ParameterCategory)tabMain.SelectedIndex;
                RefreshCurrentTabDisplay();

                // 位置參數Tab的特殊處理
                if (btnOpenPositionCalibration != null)
                {
                    btnOpenPositionCalibration.Visible = (currentCategory == ParameterCategory.Position);
                }
                if (lblPositionHint != null)
                {
                    lblPositionHint.Visible = (currentCategory == ParameterCategory.Position);
                }
            }
        }

        // 【新邏輯】：複製參考區參數到已新增未修改區
        private void btnCopyToAddedUnmodified_Click(object sender, EventArgs e)
        {
            CopyParametersFromReference();
        }

        // 【新邏輯】：移動已新增未修改區參數到已新增已修改區
        private void btnMoveToAddedModified_Click(object sender, EventArgs e)
        {
            MoveParametersFromSpecificZone(ParameterZone.AddedUnmodified, ParameterZone.AddedModified);
        }

        // 【新邏輯】：復原已新增已修改區參數到已新增未修改區
        private void btnRevertToAddedUnmodified_Click(object sender, EventArgs e)
        {
            MoveParametersFromSpecificZone(ParameterZone.AddedModified, ParameterZone.AddedUnmodified);
        }

        // 【新邏輯】：從已新增區域移除參數（回到未複製狀態）
        private void btnRemoveFromAdded_Click(object sender, EventArgs e)
        {
            RemoveParametersFromAddedZones();
        }

        // 【新增方法】：複製參考區參數到已新增未修改區
        private void CopyParametersFromReference()
        {
            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
            var selectedRefParams = categoryParams.Where(p => p.Zone == ParameterZone.Reference && p.IsSelected).ToList();

            if (selectedRefParams.Count == 0)
            {
                MessageBox.Show("請先在「參考區」中選擇要複製的參數", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 顯示確認對話框
            var result = MessageBox.Show(
                $"確定要複製 {selectedRefParams.Count} 個參數到「已新增未修改區」嗎？\n\n" +
                $"參數清單：\n{string.Join("\n", selectedRefParams.Select(p => $"• {p.Name}_{p.Stop}"))}",
                "確認複製",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                foreach (var refParam in selectedRefParams)
                {
                    // 檢查是否已存在相同參數
                    var existingParam = allParameters.FirstOrDefault(p =>
                        p.Name == refParam.Name && p.Stop == refParam.Stop && 
                        (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified));

                    if (existingParam == null)
                    {
                        // 建立新的參數副本
                        allParameters.Add(new ParameterItem
                        {
                            Type = TargetType,
                            Name = refParam.Name,
                            Value = refParam.Value,
                            Stop = refParam.Stop,
                            ChineseName = refParam.ChineseName,
                            Zone = ParameterZone.AddedUnmodified,
                            IsSelected = false
                        });
                    }

                    // 取消參考區參數的選取
                    refParam.IsSelected = false;
                }

                RefreshCurrentTabDisplay();
                SaveSession();

                // 標記當前Tab未儲存
                tabSavedStatus[currentCategory] = false;
                UpdateSaveStatusDisplay();

                MessageBox.Show($"成功複製 {selectedRefParams.Count} 個參數", "複製完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 【新增方法】：從已新增區域移除參數
        private void RemoveParametersFromAddedZones()
        {
            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
            var selectedAddedParams = categoryParams.Where(p => 
                (p.Zone == ParameterZone.AddedUnmodified || p.Zone == ParameterZone.AddedModified) && 
                p.IsSelected).ToList();

            if (selectedAddedParams.Count == 0)
            {
                MessageBox.Show("請先在「已新增」區域中選擇要移除的參數", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 顯示確認對話框
            var result = MessageBox.Show(
                $"確定要移除 {selectedAddedParams.Count} 個參數嗎？\n這些參數將從目標料號中移除。\n\n" +
                $"參數清單：\n{string.Join("\n", selectedAddedParams.Select(p => $"• {p.Name}_{p.Stop}"))}",
                "確認移除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                foreach (var param in selectedAddedParams)
                {
                    allParameters.Remove(param);
                }

                RefreshCurrentTabDisplay();
                SaveSession();

                // 標記當前Tab未儲存
                tabSavedStatus[currentCategory] = false;
                UpdateSaveStatusDisplay();

                MessageBox.Show($"成功移除 {selectedAddedParams.Count} 個參數", "移除完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void MoveSelectedParameters(ParameterZone targetZone)
        {
            var selectedParams = GetSelectedParametersFromCurrentTab();

            if (selectedParams.Count == 0)
            {
                MessageBox.Show("請先選擇要移動的參數", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 顯示預覽對話框
            string zoneName = GetZoneDisplayName(targetZone);
            var result = MessageBox.Show(
                $"確定要將 {selectedParams.Count} 個參數移動到「{zoneName}」嗎？\n\n" +
                $"參數清單：\n{string.Join("\n", selectedParams.Select(p => $"• {p.Name}_{p.Stop}"))}",
                "確認移動",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                foreach (var param in selectedParams)
                {
                    param.Zone = targetZone;
                    param.IsSelected = false;
                }

                RefreshCurrentTabDisplay();
                SaveSession();

                // 標記當前Tab未儲存
                tabSavedStatus[currentCategory] = false;
                UpdateSaveStatusDisplay();
            }
        }

        // 【新增方法】：只移動特定區域的選中參數
        private void MoveParametersFromSpecificZone(ParameterZone sourceZone, ParameterZone targetZone)
        {
            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
            var selectedParams = categoryParams.Where(p => p.Zone == sourceZone && p.IsSelected).ToList();

            if (selectedParams.Count == 0)
            {
                string sourceZoneName = GetZoneDisplayName(sourceZone);
                MessageBox.Show($"請先在「{sourceZoneName}」中選擇要移動的參數", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 顯示預覽對話框
            string sourceDisplayName = GetZoneDisplayName(sourceZone);
            string targetDisplayName = GetZoneDisplayName(targetZone);
            var result = MessageBox.Show(
                $"確定要將 {selectedParams.Count} 個參數從「{sourceDisplayName}」移動到「{targetDisplayName}」嗎？\n\n" +
                $"參數清單：\n{string.Join("\n", selectedParams.Select(p => $"• {p.Name}_{p.Stop}"))}",
                "確認移動",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                foreach (var param in selectedParams)
                {
                    param.Zone = targetZone;
                    param.IsSelected = false;
                }

                RefreshCurrentTabDisplay();
                SaveSession();

                // 標記當前Tab未儲存
                tabSavedStatus[currentCategory] = false;
                UpdateSaveStatusDisplay();
            }
        }
        private List<ParameterItem> GetSelectedParametersFromCurrentTab()
        {
            var selected = new List<ParameterItem>();
            var (dgvUnmodified, dgvModified, dgvFixed) = GetCurrentTabDataGridViews();

            // 從三個DataGridView中取得選中的參數
            AddSelectedParametersFromGrid(dgvUnmodified, selected);
            AddSelectedParametersFromGrid(dgvModified, selected);
            AddSelectedParametersFromGrid(dgvFixed, selected);

            return selected;
        }

        private void AddSelectedParametersFromGrid(DataGridView dgv, List<ParameterItem> selectedList)
        {
            if (dgv.DataSource is List<ParameterItem> parameters)
            {
                selectedList.AddRange(parameters.Where(p => p.IsSelected));
            }
        }

        private string GetZoneDisplayName(ParameterZone zone)
        {
            switch (zone)
            {
                case ParameterZone.Reference: return "參考區";
                case ParameterZone.AddedUnmodified: return "已新增未修改區";
                case ParameterZone.AddedModified: return "已新增已修改區";
                default: return zone.ToString();
            }
        }

        // 選擇按鈕事件
        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            SelectParametersInCurrentTab(true);
        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            SelectParametersInCurrentTab(false);
        }

        private void btnSelectInvert_Click(object sender, EventArgs e)
        {
            InvertSelectionInCurrentTab();
        }

        private void SelectParametersInCurrentTab(bool selected)
        {
            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
            foreach (var param in categoryParams)
            {
                param.IsSelected = selected;
            }
            RefreshCurrentTabDisplay();
        }

        private void InvertSelectionInCurrentTab()
        {
            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
            foreach (var param in categoryParams)
            {
                param.IsSelected = !param.IsSelected;
            }
            RefreshCurrentTabDisplay();
        }

        // 位置校正工具按鈕
        private void btnOpenPositionCalibration_Click(object sender, EventArgs e)
        {
            try
            {
                using (var calibrationForm = new CircleCalibrationForm())
                {
                    // 由 GitHub Copilot 產生
                    // 移除：不再訂閱 FormClosed，避免未儲存也回載參數
                    // calibrationForm.FormClosed += OnPositionCalibrationClosed;

                    // 顯示為對話框
                    calibrationForm.ShowDialog(this);

                    // 僅當使用者在校正畫面內「有成功儲存到資料庫」時才回載
                    if (calibrationForm.HasSavedToDb)
                    {
                        LoadPositionParametersFromDatabase();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"開啟位置校正工具時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnPositionCalibrationClosed(object sender, FormClosedEventArgs e)
        {
            // 檢查是否有新的位置參數
            LoadPositionParametersFromDatabase();
        }

        // 【重構方法】：修正位置參數載入邏輯
        private void LoadPositionParametersFromDatabase()
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 載入最新的位置相關參數
                    var positionParams = db.@params
                        .Where(p => p.Type == TargetType &&
                               (p.Name.Contains("center") || p.Name.Contains("radius") ||
                                p.Name.Contains("chamfer") || p.Name.Contains("position")))
                        .ToList();

                    // 【修正】：只移除校正工具產生的特定參數，保留參考區和其他參數
                    allParameters.RemoveAll(p =>
                        p.Zone == ParameterZone.AddedModified &&
                        (p.Name.Contains("center") || p.Name.Contains("radius") ||
                         p.Name.Contains("chamfer") || p.Name.Contains("position")));

                    // 新增載入的位置參數到已新增已修改區
                    foreach (var param in positionParams)
                    {
                        // 檢查是否已存在相同參數
                        var existingParam = allParameters.FirstOrDefault(p =>
                            p.Name == param.Name && p.Stop == param.Stop);

                        if (existingParam != null)
                        {
                            // 更新現有參數
                            existingParam.Value = param.Value;
                            existingParam.Zone = ParameterZone.AddedModified; // 校正完成的參數
                        }
                        else
                        {
                            // 新增新參數
                            allParameters.Add(new ParameterItem
                            {
                                Type = TargetType,
                                Name = param.Name,
                                Value = param.Value,
                                Stop = param.Stop,
                                ChineseName = param.ChineseName,
                                Zone = ParameterZone.AddedModified, // 校正完成的參數
                                IsSelected = false
                            });
                        }
                    }

                    // 更新顯示
                    if (currentCategory == ParameterCategory.Position)
                    {
                        RefreshCurrentTabDisplay();
                    }

                    UpdateCategoryProgress();
                    SaveSession();

                    if (positionParams.Any())
                    {
                        MessageBox.Show($"已載入 {positionParams.Count} 個位置參數到已新增已修改區", "載入完成",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入位置參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 儲存按鈕事件
        private void btnSaveCurrentTab_Click(object sender, EventArgs e)
        {
            SaveCurrentTabParameters();
        }

        private void btnSaveAllAndComplete_Click(object sender, EventArgs e)
        {
            SaveAllParameters();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "確定要取消設定嗎？未儲存的變更將會遺失。",
                "確認取消",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void SaveCurrentTabParameters()
        {
            try
            {
                var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);
                var paramsToSave = categoryParams.Where(p => p.Zone == ParameterZone.AddedModified || p.Zone == ParameterZone.AddedUnmodified).ToList();

                if (paramsToSave.Count == 0)
                {
                    MessageBox.Show("當前頁面沒有需要儲存的參數", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var result = MessageBox.Show(
                    $"確定要儲存當前頁面的 {paramsToSave.Count} 個參數嗎？\n\n" +
                    $"參數清單：\n{string.Join("\n", paramsToSave.Select(p => $"• {p.Name}_{p.Stop} = {p.Value}"))}",
                    "確認儲存",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    SaveParametersToDatabase(paramsToSave);
                    tabSavedStatus[currentCategory] = true;
                    UpdateSaveStatusDisplay();

                    MessageBox.Show($"成功儲存 {paramsToSave.Count} 個參數", "儲存完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveAllParameters()
        {
            try
            {
                var allParamsToSave = allParameters.Where(p => p.Zone == ParameterZone.AddedModified || p.Zone == ParameterZone.AddedUnmodified).ToList();

                if (allParamsToSave.Count == 0)
                {
                    MessageBox.Show("沒有需要儲存的參數", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var result = MessageBox.Show(
                    $"確定要儲存全部 {allParamsToSave.Count} 個參數並完成設定嗎？\n\n" +
                    "這將會：\n• 儲存所有已修改和固定的參數到資料庫\n• 刪除工作階段檔案\n• 關閉此視窗",
                    "確認完成",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    SaveParametersToDatabase(allParamsToSave);

                    // 刪除session檔案
                    if (File.Exists(sessionFilePath))
                    {
                        File.Delete(sessionFilePath);
                    }

                    MessageBox.Show($"參數設定完成！\n成功儲存 {allParamsToSave.Count} 個參數", "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveParametersToDatabase(List<ParameterItem> parameters)
        {
            using (var db = new MydbDB())
            {
                foreach (var param in parameters)
                {
                    // 判斷參數應該存到哪個資料表
                    if (IsCameraParameter(param.Name))
                    {
                        // 先刪除舊資料
                        var existingCamera = db.Cameras.FirstOrDefault(c =>
                            c.Type == param.Type && c.Name == param.Name && c.Stop == param.Stop);
                        if (existingCamera != null)
                        {
                            db.Delete(existingCamera);
                        }

                        // 新增新資料
                        db.Insert(new Camera
                        {
                            Type = param.Type,
                            Name = param.Name,
                            Value = param.Value,
                            Stop = param.Stop ?? 0,
                            ChineseName = param.ChineseName
                        });
                    }
                    else
                    {
                        // 先刪除舊資料
                        var existingParam = db.@params.FirstOrDefault(p =>
                            p.Type == param.Type && p.Name == param.Name && p.Stop == param.Stop);
                        if (existingParam != null)
                        {
                            db.Delete(existingParam);
                        }

                        // 新增新資料
                        db.Insert(new Param
                        {
                            Type = param.Type,
                            Name = param.Name,
                            Value = param.Value,
                            Stop = param.Stop,
                            ChineseName = param.ChineseName
                        });
                    }
                }
            }
        }

        private bool IsCameraParameter(string parameterName)
        {
            return parameterName.Equals("exposure", StringComparison.OrdinalIgnoreCase) ||
                   parameterName.Equals("gain", StringComparison.OrdinalIgnoreCase) ||
                   parameterName.Equals("delay", StringComparison.OrdinalIgnoreCase);
        }
        private void RecalculateProgressAndUpdateUI()
        {
            UpdateCategoryProgress();          // 重新統計各分類 (分母=參考區, 分子=已修改區)
            UpdateProgressDisplayOnly();       // 更新上方文字/總進度
        }
        private void ApplyReferenceZoneColors()
        {
            // 由 GitHub Copilot 產生 - 先建立快取，避免每列都執行 LINQ 查詢
            BuildParameterZoneCache();

            foreach (var dgv in new[] { dgvCameraUnmodified, dgvPositionUnmodified, dgvDetectionUnmodified, dgvTimingUnmodified })
            {
                if (dgv?.DataSource is List<ParameterItem> list)
                {
                    for (int i = 0; i < list.Count; i++)
                    {
                        var item = list[i];
                        if (item.Zone == ParameterZone.Reference)
                        {
                            UpdateReferenceRowVisual(dgv.Rows[i], item);
                        }
                    }
                }
            }
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSession();
            base.OnFormClosing(e);
        }
    }
}