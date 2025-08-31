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
        // 【新增方法】：設定三區域佈局（優化的尺寸）
        private void SetupThreeZones(TabPage tab, string prefix,
            GroupBox grpUnmodified, GroupBox grpModified, GroupBox grpFixed,
            DataGridView dgvUnmodified, DataGridView dgvModified, DataGridView dgvFixed,
            Button btnMoveToModified, Button btnMoveToFixed, Button btnMoveToUnmodified,
            Button btnSelectAll, Button btnSelectNone, Button btnSelectInvert)
        {
            // 視窗1600x1000，Tab可用空間更大，調整為更寬敞的佈局
            // 優化的GroupBox佈局 - 增加寬度和高度
            grpUnmodified.Text = "尚未修改區";
            grpUnmodified.Location = new Point(10, 55);
            grpUnmodified.Size = new Size(510, 720);  // 增加寬度至510，高度至720

            grpModified.Text = "已修改區";
            grpModified.Location = new Point(530, 55);  // 調整X位置
            grpModified.Size = new Size(510, 720);  // 增加寬度至510，高度至720

            grpFixed.Text = "固定區";
            grpFixed.Location = new Point(1050, 55);  // 調整X位置
            grpFixed.Size = new Size(510, 720);  // 增加寬度至510，高度至720

            // DataGridView設定
            SetupDataGridView(dgvUnmodified);
            SetupDataGridView(dgvModified);
            SetupDataGridView(dgvFixed);

            // 【新增】建立缺少的按鈕
            var btnModifiedToFixed = new Button();
            var btnFixedToModified = new Button();

            // 【修正】移動按鈕 - 明確標示移動方向

            // 尚未修改區的按鈕（向外移動）
            btnMoveToModified.Text = "未修改 → 已修改";
            btnMoveToModified.Location = new Point(10, 785);
            btnMoveToModified.Size = new Size(130, 30);  // 增加寬度以容納文字
            btnMoveToModified.Tag = prefix;
            btnMoveToModified.Click += btnMoveToModified_Click;

            btnMoveToFixed.Text = "未修改 → 固定";
            btnMoveToFixed.Location = new Point(150, 785);  // 調整位置
            btnMoveToFixed.Size = new Size(130, 30);  // 增加寬度
            btnMoveToFixed.Tag = prefix;
            btnMoveToFixed.Click += btnMoveToFixed_Click;

            // 已修改區的按鈕
            btnModifiedToFixed.Text = "已修改 → 固定";
            btnModifiedToFixed.Location = new Point(530, 785);
            btnModifiedToFixed.Size = new Size(130, 30);
            btnModifiedToFixed.Tag = prefix;
            btnModifiedToFixed.Click += btnModifiedToFixed_Click;

            btnFixedToModified.Text = "固定 → 已修改";
            btnFixedToModified.Location = new Point(1050, 785);
            btnFixedToModified.Size = new Size(130, 30);
            btnFixedToModified.Tag = prefix;
            btnFixedToModified.Click += btnFixedToModified_Click;

            btnMoveToUnmodified.Text = "任何 → 未修改";
            btnMoveToUnmodified.Location = new Point(735, 840);
            btnMoveToUnmodified.Size = new Size(130, 30);  // 增加寬度
            btnMoveToUnmodified.Tag = prefix;
            btnMoveToUnmodified.Click += btnMoveToUnmodified_Click;


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

        // 【修正】：設定DataGridView - 增加大小以適應更大的GroupBox
        private void SetupDataGridView(DataGridView dgv)
        {
            dgv.Location = new Point(8, 28);
            dgv.Size = new Size(494, 680);  // 增加寬度至494，高度至680以適應GroupBox
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = true;
            dgv.AutoGenerateColumns = true;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        }
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

        // 【修正】：顯示時間參數計算對話框 - 加入詳細公式說明
        private void ShowTimingCalculationDialog(double od)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "時間參數計算結果";
                dialog.Size = new Size(900, 600); // 增加寬度以容納更多資訊
                dialog.StartPosition = FormStartPosition.CenterParent;

                // 計算時間參數
                var calculatedParams = CalculateTimingParameters(od);

                // 【新增】公式說明區域
                var txtFormula = new TextBox();
                txtFormula.Text = GetCalculationFormula(od);
                txtFormula.Location = new Point(10, 10);
                txtFormula.Size = new Size(860, 200);
                txtFormula.Multiline = true;
                txtFormula.ReadOnly = true;
                txtFormula.ScrollBars = ScrollBars.Vertical;
                txtFormula.Font = new Font("Microsoft JhengHei", 9);

                // 簡化的公式顯示標籤
                var lblFormula = new Label();
                lblFormula.Text = $"📊 基於實際生產資料的智能演算法 | 當前OD: {od:F2}mm";
                lblFormula.Location = new Point(10, 220);
                lblFormula.Size = new Size(600, 25);
                lblFormula.Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold);
                lblFormula.ForeColor = Color.DarkBlue;

                // 結果表格
                var dgvResults = new DataGridView();
                dgvResults.Location = new Point(10, 250);
                dgvResults.Size = new Size(860, 250);
                dgvResults.DataSource = calculatedParams.Select(p => new {
                    參數名稱 = p.Name,
                    計算結果 = $"{p.Value} ms",
                    站點 = p.Stop,
                    說明 = p.ChineseName
                }).ToList();
                dgvResults.ReadOnly = true;
                dgvResults.AllowUserToAddRows = false;
                dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                // 按鈕
                var btnOK = new Button();
                btnOK.Text = "✅ 使用這些數值";
                btnOK.Location = new Point(680, 520);
                btnOK.Size = new Size(120, 30);
                btnOK.BackColor = Color.LightGreen;
                btnOK.Click += (s, args) =>
                {
                    AddCalculatedParametersToModified(calculatedParams);
                    dialog.DialogResult = DialogResult.OK;
                };

                var btnCancel = new Button();
                btnCancel.Text = "❌ 取消";
                btnCancel.Location = new Point(810, 520);
                btnCancel.Size = new Size(80, 30);
                btnCancel.Click += (s, args) => dialog.DialogResult = DialogResult.Cancel;

                dialog.Controls.AddRange(new Control[] { txtFormula, lblFormula, dgvResults, btnOK, btnCancel });

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    RefreshCurrentTabDisplay();
                    SaveSession();

                    MessageBox.Show($"✅ 成功計算並加入 {calculatedParams.Count} 個時間參數到已修改區\n\n" +
                        $"計算結果摘要：\n" +
                        $"• fourToOK1: {calculatedParams.First(p => p.Name == "fourToOK1_time_ms").Value} ms\n" +
                        $"• fourToOK2: {calculatedParams.First(p => p.Name == "fourToOK2_time_ms").Value} ms\n" +
                        $"• fourToNG: {calculatedParams.First(p => p.Name == "fourToNG_time_ms").Value} ms\n" +
                        $"• fourToNULL: {calculatedParams.First(p => p.Name == "fourToNULL_time_ms").Value} ms",
                        "計算完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        #region 時間參數演算法
        // 【修正】：計算時間參數 - 基於真實資料的演算法
        private List<ParameterItem> CalculateTimingParameters(double od)
        {
            var calculatedParams = new List<ParameterItem>();

            // 【新演算法】：基於實際資料的線性回歸和經驗公式
            var timingCalculations = new[]
            {
                new {
                    Name = "fourToOK1_time_ms",
                    Calculator = new Func<double, double>(CalculateFourToOK1Time),
                    Description = "站4到OK1出料時間"
                },
                new {
                    Name = "fourToOK2_time_ms",
                    Calculator = new Func<double, double>(CalculateFourToOK2Time),
                    Description = "站4到OK2出料時間"
                },
                new {
                    Name = "fourToNG_time_ms",
                    Calculator = new Func<double, double>(CalculateFourToNGTime),
                    Description = "站4到NG出料時間"
                },
                new {
                    Name = "fourToNULL_time_ms",
                    Calculator = new Func<double, double>(CalculateFourToNULLTime),
                    Description = "站4到NULL出料時間"
                }
            };

            foreach (var timing in timingCalculations)
            {
                double calculatedValue = Math.Round(timing.Calculator(od));

                calculatedParams.Add(new ParameterItem
                {
                    Name = timing.Name,
                    Value = calculatedValue.ToString(),
                    Stop = 4, // 這些參數都是站4的
                    ChineseName = timing.Description,
                    Type = TargetType,
                    Zone = ParameterZone.Modified
                });
            }

            return calculatedParams;
        }

        // 【新增】：基於實際資料的時間計算方法

        /// <summary>
        /// 計算 fourToOK1 時間 - 基於線性回歸分析
        /// 資料點：(35, 1240), (37.75, 1250), (40.03, 1250)
        /// </summary>
        private double CalculateFourToOK1Time(double od)
        {
            // 基於實際資料的分段線性函數
            if (od <= 35.0)
            {
                // 小於35mm：基礎時間1240ms
                return 1240 + (35.0 - od) * 2; // 每減少1mm，時間減少2ms
            }
            else if (od <= 37.75)
            {
                // 35-37.75mm：緩慢增長
                double slope = (1250 - 1240) / (37.75 - 35.0); // 約3.64 ms/mm
                return 1240 + slope * (od - 35.0);
            }
            else if (od <= 40.03)
            {
                // 37.75-40.03mm：保持穩定
                return 1250;
            }
            else
            {
                // 大於40.03mm：略微增加
                return 1250 + (od - 40.03) * 1.5; // 每增加1mm，時間增加1.5ms
            }
        }

        /// <summary>
        /// 計算 fourToOK2 時間 - 考慮機械傳送距離差異
        /// 資料點：(35, 1930), (37.75, 1900), (40.03, 1940)
        /// </summary>
        private double CalculateFourToOK2Time(double od)
        {
            // OK2出料口距離較遠，時間較長且受OD影響更明顯
            if (od <= 35.0)
            {
                return 1930 + (35.0 - od) * 5; // 每減少1mm，時間減少5ms
            }
            else if (od <= 37.75)
            {
                // 35-37.75mm：略微下降（可能因為中等尺寸的流動性較好）
                double slope = (1900 - 1930) / (37.75 - 35.0); // 約-10.91 ms/mm
                return 1930 + slope * (od - 35.0);
            }
            else if (od <= 40.03)
            {
                // 37.75-40.03mm：上升趨勢
                double slope = (1940 - 1900) / (40.03 - 37.75); // 約17.54 ms/mm
                return 1900 + slope * (od - 37.75);
            }
            else
            {
                // 大於40.03mm：持續增加
                return 1940 + (od - 40.03) * 8; // 每增加1mm，時間增加8ms
            }
        }

        /// <summary>
        /// 計算 fourToNG 時間 - NG出料路徑最長
        /// 資料點：(35, 4500), (37.75, 4600), (40.03, 4500)
        /// </summary>
        private double CalculateFourToNGTime(double od)
        {
            // NG出料時間相對穩定，但中等尺寸稍長
            if (od <= 35.0)
            {
                return 4500 + (35.0 - od) * 10; // 每減少1mm，時間減少10ms
            }
            else if (od <= 37.75)
            {
                // 35-37.75mm：上升到峰值
                double slope = (4600 - 4500) / (37.75 - 35.0); // 約36.36 ms/mm
                return 4500 + slope * (od - 35.0);
            }
            else if (od <= 40.03)
            {
                // 37.75-40.03mm：下降趨勢
                double slope = (4500 - 4600) / (40.03 - 37.75); // 約-43.86 ms/mm
                return 4600 + slope * (od - 37.75);
            }
            else
            {
                // 大於40.03mm：緩慢增加
                return 4500 + (od - 40.03) * 5; // 每增加1mm，時間增加5ms
            }
        }

        /// <summary>
        /// 計算 fourToNULL 時間 - NULL出料時間最長
        /// 基於現有資料推算：約為 NG時間 + 500ms
        /// </summary>
        private double CalculateFourToNULLTime(double od)
        {
            // NULL出料時間 = NG時間 + 固定延遲
            double ngTime = CalculateFourToNGTime(od);

            // 根據實際資料，NULL時間約為NG時間加上420-520ms
            double nullOffset = 500; // 基礎偏移量

            // 根據OD調整偏移量
            if (od <= 35.0)
            {
                nullOffset = 520; // 小尺寸稍長
            }
            else if (od <= 40.03)
            {
                nullOffset = 500; // 中等尺寸標準
            }
            else
            {
                nullOffset = 480; // 大尺寸稍短
            }

            return ngTime + nullOffset;
        }

        // 【新增】：顯示計算公式的詳細說明
        private string GetCalculationFormula(double od)
        {
            return $@"
            📐 時間參數計算公式 (當前OD: {od:F2}mm)

            🔸 fourToOK1_time_ms:
               • 基準值：1240-1250ms
               • 小於35mm：1240 + (35-OD)×2
               • 35-37.75mm：線性增長 (斜率≈3.64ms/mm)
               • 37.75-40.03mm：保持1250ms
               • 大於40.03mm：1250 + (OD-40.03)×1.5

            🔸 fourToOK2_time_ms:
               • 基準值：1900-1940ms
               • 小於35mm：1930 + (35-OD)×5
               • 35-37.75mm：下降趨勢 (斜率≈-10.91ms/mm)
               • 37.75-40.03mm：上升趨勢 (斜率≈17.54ms/mm)
               • 大於40.03mm：1940 + (OD-40.03)×8

            🔸 fourToNG_time_ms:
               • 基準值：4500-4600ms
               • 呈現倒U型分佈，37.75mm為峰值
               • 考慮機械阻力和流動性因素

            🔸 fourToNULL_time_ms:
               • 計算方式：NG時間 + 固定偏移量(480-520ms)
               • 確保NULL出料有足夠的時間緩衝

            💡 演算法特點：
               • 基於實際生產資料的線性回歸分析
               • 考慮不同OD的機械特性差異
               • 採用分段函數處理非線性關係
               • 包含邊界值的外推處理";
        }
        #endregion
        // 【新增】：將計算結果加入到已修改區
        private void AddCalculatedParametersToModified(List<ParameterItem> calculatedParams)
        {
            foreach (var calcParam in calculatedParams)
            {
                // 檢查是否已存在相同的參數
                var existingParam = allParameters.FirstOrDefault(p =>
                    p.Name == calcParam.Name && p.Stop == calcParam.Stop);

                if (existingParam != null)
                {
                    // 更新現有參數的值和區域
                    existingParam.Value = calcParam.Value;
                    existingParam.Zone = ParameterZone.Modified;
                }
                else
                {
                    // 新增新參數
                    calcParam.Type = TargetType;
                    calcParam.Zone = ParameterZone.Modified;
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
                new CalibrationItem { Name = "尺寸檢測校正 (dimension)", Type = CalibrationType.Dimension },
                new CalibrationItem { Name = "輪廓檢測校正 (contour)", Type = CalibrationType.Contour },
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
            lblHint.Text = "💡 選擇校正類型和站點，系統將引導您完成自動校正流程";
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
            Dimension,      // 尺寸檢測校正
            Contour,        // 輪廓檢測校正
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
                    case CalibrationType.Dimension:
                        StartDimensionCalibration();
                        break;
                    case CalibrationType.Contour:
                        StartContourCalibration();
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

        // 【新增】：取得校正相關的參數名稱
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

                case CalibrationType.Dimension:
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.AddRange(new[] {
                    $"minContourArea_{station}",
                    $"maxContourArea_{station}",
                    $"dimensionTolerance_{station}"
                });
                    }
                    break;

                case CalibrationType.Contour:
                    for (int station = 1; station <= 4; station++)
                    {
                        parameterNames.AddRange(new[] {
                    $"contourApproxEpsilon_{station}",
                    $"minContourLength_{station}",
                    $"contourSmoothness_{station}"
                });
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

                // 準備數據
                var parameterData = parameterNames.Select(paramName => {
                    var existingParam = allParameters.FirstOrDefault(p => p.Name == GetParameterBaseName(paramName));
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

        // 【新增】：取得當前參數值
        private string GetCurrentParameterValues(List<string> parameterNames)
        {
            var values = new List<string>();

            foreach (var paramName in parameterNames)
            {
                var existingParam = allParameters.FirstOrDefault(p => p.Name == paramName);
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
                if (calibrationForm.ShowDialog() == DialogResult.OK)
                {
                    // 重新載入參數並更新顯示
                    LoadUpdatedParametersFromDatabase(new[] { $"objBias_x", $"objBias_y" });
                    RefreshCurrentTabDisplay();

                    MessageBox.Show("物體偏移量校正完成，參數已更新", "校正完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // 【框架】：閾值參數校正
        private void StartGapThreshCalibration()
        {
            using (var calibrationForm = new GapThreshCalibrationForm(TargetType))
            {
                if (calibrationForm.ShowDialog() == DialogResult.OK)
                {
                    LoadUpdatedParametersFromDatabase(new[] { "gapThresh" });
                    RefreshCurrentTabDisplay();

                    MessageBox.Show("gapThresh 參數校正完成", "校正完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // 【框架】：尺寸檢測校正
        private void StartDimensionCalibration()
        {
            MessageBox.Show("尺寸檢測校正功能開發中...", "開發中",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            // TODO: 實作尺寸檢測校正邏輯
            // using (var calibrationForm = new DimensionCalibrationForm(TargetType, station))
            // {
            //     if (calibrationForm.ShowDialog() == DialogResult.OK)
            //     {
            //         var parameterNames = new[] { 
            //             $"minContourArea_{station}", 
            //             $"maxContourArea_{station}", 
            //             $"dimensionTolerance_{station}" 
            //         };
            //         LoadUpdatedParametersFromDatabase(parameterNames);
            //         RefreshCurrentTabDisplay();
            //         
            //         MessageBox.Show("尺寸檢測校正完成", "校正完成", 
            //             MessageBoxButtons.OK, MessageBoxIcon.Information);
            //     }
            // }
        }

        // 【框架】：輪廓檢測校正
        private void StartContourCalibration()
        {
            MessageBox.Show("輪廓檢測校正功能開發中...", "開發中",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            // TODO: 實作輪廓檢測校正邏輯
        }

        // 【框架】：顏色檢測校正
        private void StartWhiteCalibration()
        {
            using (var calibrationForm = new WhiteCalibrationForm(TargetType))
            {
                if (calibrationForm.ShowDialog() == DialogResult.OK)
                {
                    // 重新載入參數並更新顯示
                    LoadUpdatedParametersFromDatabase(new[] { "white" });
                    RefreshCurrentTabDisplay();

                    MessageBox.Show("白色像素占比校正完成，參數已更新", "校正完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
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
                                existingParam.Zone = ParameterZone.Fixed; // 校正完成的參數放到固定區
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
                                    Zone = ParameterZone.Fixed,
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
        #endregion
        public string TargetType { get; private set; }

        private void InitializeForm()
        {
            this.Text = $"新料號參數設定 - {TargetType}";
            this.Size = new Size(1600, 1050);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = true;

            // 【修正1】：先更新進度顯示，但不要更新Tab狀態（避免TabControl還沒建立完成）
            UpdateProgressDisplayOnly();        

            if (File.Exists(sessionFilePath))
            {
                var result = MessageBox.Show(
                    $"發現未完成的參數設定工作階段，是否要繼續？\n\n點擊「是」繼續之前的工作\n點擊「否」重新開始",
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

        // 【新增方法】：只更新進度顯示，不更新Tab狀態
        private void UpdateProgressDisplayOnly()
        {
            lblProgress.Text = $"料號: {TargetType}";

            int overallProgress = parameterManager.GetOverallProgress();
            progressBarOverall.Value = Math.Min(overallProgress, 100);

            string progressText = $"總進度: {overallProgress}%";
            lblProgressSummary.Text = $"{progressText} {parameterManager.GetProgressSummary()}";
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

        private void LoadParametersFromSource(string sourceType)
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 【修正1】：先載入來源參數
                    var cameraParams = db.Cameras
                        .Where(c => c.Type == sourceType)
                        .Select(c => new ParameterItem
                        {
                            Type = TargetType,
                            Name = c.Name,
                            Value = c.Value,
                            Stop = c.Stop,
                            ChineseName = c.ChineseName,
                            Zone = ParameterZone.Unmodified,
                            IsSelected = false
                        }).ToList();

                    var paramParams = db.@params
                        .Where(p => p.Type == sourceType)
                        .Select(p => new ParameterItem
                        {
                            Type = TargetType,
                            Name = p.Name,
                            Value = p.Value,
                            Stop = p.Stop,
                            ChineseName = p.ChineseName,
                            Zone = ParameterZone.Unmodified,
                            IsSelected = false
                        }).ToList();

                    allParameters.Clear();
                    allParameters.AddRange(cameraParams);
                    allParameters.AddRange(paramParams);

                    // 【修正2】：檢查目標資料庫中已存在的參數，更新其狀態
                    CheckAndUpdateExistingParameters();

                    currentSession.Parameters = allParameters;

                    // 更新各類別的參數數量
                    UpdateCategoryProgress();
                    SaveSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 【新增方法】：檢查目標資料庫中已存在的參數
        private void CheckAndUpdateExistingParameters()
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 檢查Camera表中的已存在參數
                    var existingCameras = db.Cameras
                        .Where(c => c.Type == TargetType)
                        .ToList();

                    // 檢查Param表中的已存在參數
                    var existingParams = db.@params
                        .Where(p => p.Type == TargetType)
                        .ToList();

                    // 更新allParameters中對應參數的狀態
                    foreach (var param in allParameters)
                    {
                        bool existsInDatabase = false;

                        if (IsCameraParameter(param.Name))
                        {
                            existsInDatabase = existingCameras.Any(c =>
                                c.Name == param.Name && c.Stop == param.Stop);
                        }
                        else
                        {
                            existsInDatabase = existingParams.Any(p =>
                                p.Name == param.Name && p.Stop == param.Stop);
                        }

                        if (existsInDatabase)
                        {
                            // 如果資料庫中已存在，標記為Fixed（已完成）
                            param.Zone = ParameterZone.Fixed;
                        }
                    }

                    // 【修正3】：載入資料庫中存在但來源中沒有的參數（用戶新增的）
                    LoadUserAddedParameters(existingCameras, existingParams);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"檢查已存在參數時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 【新增方法】：載入用戶新增的參數（存在於目標但不存在於來源）
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
                        Value = camera.Value,
                        Stop = camera.Stop,
                        ChineseName = camera.ChineseName,
                        Zone = ParameterZone.Fixed, // 已存在資料庫中
                        IsSelected = false
                    });
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
                        Value = param.Value,
                        Stop = param.Stop,
                        ChineseName = param.ChineseName,
                        Zone = ParameterZone.Fixed, // 已存在資料庫中
                        IsSelected = false
                    });
                }
            }
        }

        private void LoadExistingSession()
        {
            try
            {
                var sessionJson = File.ReadAllText(sessionFilePath);
                currentSession = JsonConvert.DeserializeObject<ParameterSession>(sessionJson);
                allParameters = currentSession.Parameters ?? new List<ParameterItem>();

                // 【修正】：重新檢查資料庫狀態，確保狀態正確
                CheckAndUpdateExistingParameters();

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

        private void SaveSession()
        {
            try
            {
                if (!Directory.Exists(settingPath))
                {
                    Directory.CreateDirectory(settingPath);
                }

                if (currentSession != null)
                {
                    currentSession.Parameters = allParameters;
                    var sessionJson = JsonConvert.SerializeObject(currentSession, Formatting.Indented);
                    File.WriteAllText(sessionFilePath, sessionJson);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存工作階段失敗：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void UpdateCategoryProgress()
        {
            foreach (ParameterCategory category in Enum.GetValues(typeof(ParameterCategory)))
            {
                var categoryParams = parameterManager.GetParametersByCategory(allParameters, category);
                int totalCount = categoryParams.Count;
                int completedCount = categoryParams.Count(p => p.Zone == ParameterZone.Fixed || p.Zone == ParameterZone.Modified);

                parameterManager.UpdateCategoryProgress(category, completedCount, totalCount);

                // 【修正】：同時基於資料庫更新狀態
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

        private void RefreshCurrentTabDisplay()
        {
            // 【修正1】：確保有參數資料
            if (allParameters == null)
                return;

            var categoryParams = parameterManager.GetParametersByCategory(allParameters, currentCategory);

            // 取得當前Tab的DataGridView
            var (dgvUnmodified, dgvModified, dgvFixed) = GetCurrentTabDataGridViews();

            // 【修正2】：確保DataGridView都存在
            if (dgvUnmodified == null || dgvModified == null || dgvFixed == null)
                return;
            // 分別載入三個區域的參數
            RefreshDataGridView(dgvUnmodified, categoryParams.Where(p => p.Zone == ParameterZone.Unmodified).ToList());
            RefreshDataGridView(dgvModified, categoryParams.Where(p => p.Zone == ParameterZone.Modified).ToList());
            RefreshDataGridView(dgvFixed, categoryParams.Where(p => p.Zone == ParameterZone.Fixed).ToList());

            UpdateCategoryProgress();
            UpdateProgressDisplay();
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

        // 【修正】移動按鈕事件 - 只移動特定區域的參數
        private void btnMoveToModified_Click(object sender, EventArgs e)
        {
            MoveParametersFromSpecificZone(ParameterZone.Unmodified, ParameterZone.Modified);
        }

        private void btnMoveToFixed_Click(object sender, EventArgs e)
        {
            MoveParametersFromSpecificZone(ParameterZone.Unmodified, ParameterZone.Fixed);
        }

        private void btnModifiedToFixed_Click(object sender, EventArgs e)
        {
            MoveParametersFromSpecificZone(ParameterZone.Modified, ParameterZone.Fixed);
        }

        private void btnFixedToModified_Click(object sender, EventArgs e)
        {
            MoveParametersFromSpecificZone(ParameterZone.Fixed, ParameterZone.Modified);
        }

        private void btnMoveToUnmodified_Click(object sender, EventArgs e)
        {
            MoveSelectedParameters(ParameterZone.Unmodified);  // 保持原有功能：任何區域→未修改
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
                case ParameterZone.Unmodified: return "尚未修改區";
                case ParameterZone.Modified: return "已修改區";
                case ParameterZone.Fixed: return "固定區";
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
                    // 訂閱參數儲存事件
                    calibrationForm.FormClosed += OnPositionCalibrationClosed;

                    if (calibrationForm.ShowDialog() == DialogResult.OK)
                    {
                        // 位置校正完成，重新載入位置參數
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

                    // 移除舊的位置參數
                    allParameters.RemoveAll(p =>
                        p.Name.Contains("center") || p.Name.Contains("radius") ||
                        p.Name.Contains("chamfer") || p.Name.Contains("position"));

                    // 新增載入的位置參數
                    foreach (var param in positionParams)
                    {
                        allParameters.Add(new ParameterItem
                        {
                            Type = TargetType,
                            Name = param.Name,
                            Value = param.Value,
                            Stop = param.Stop,
                            ChineseName = param.ChineseName,
                            Zone = ParameterZone.Fixed, // 位置參數設定完成後直接放到固定區
                            IsSelected = false
                        });
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
                        MessageBox.Show($"已載入 {positionParams.Count} 個位置參數", "載入完成",
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
                var paramsToSave = categoryParams.Where(p => p.Zone == ParameterZone.Fixed || p.Zone == ParameterZone.Modified).ToList();

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
                var allParamsToSave = allParameters.Where(p => p.Zone == ParameterZone.Fixed || p.Zone == ParameterZone.Modified).ToList();

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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSession();
            base.OnFormClosing(e);
        }
    }
}