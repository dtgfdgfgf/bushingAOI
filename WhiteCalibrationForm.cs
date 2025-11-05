using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using OpenCvSharp;
using LinqToDB;

namespace peilin
{
    public partial class WhiteCalibrationForm : Form
    {
        #region 私有成員變數
        private List<string> selectedImagePaths = new List<string>();
        private List<WhitePixelResult> whitePixelResults = new List<WhitePixelResult>();
        // 由 GitHub Copilot 產生
        // 誤觸樣品組相關變數
        private List<string> selectedFalseImagePaths = new List<string>();
        private List<WhitePixelResult> falsePixelResults = new List<WhitePixelResult>();
        private int currentFalseImageIndex = 0;
        
        private string targetType;
        private int targetStation;
        private float recommendedWhiteValue;
        private float recommendedWhiteNullValue;
        private float recommendedTolerance;
        private int currentImageIndex = 0;
        private bool isAnalyzing = false;
        private bool eventsInitialized = false;

        // 控件成員變數
        private ComboBox cmbStation;
        private PictureBox picPreviewControl;
        private TextBox lblImageInfoControl;
        private Button btnPreviousControl;
        private Button btnNextControl;
        private DataGridView dgvResultsControl;
        private ProgressBar progressBarControl;
        private Button btnAnalyzeControl;
        private Button btnApplyControl;
        private Label lblRecommendationControl;
        private GroupBox grpStatisticsControl;
        private Chart chartWhitePixelRatio;
        private ToolTip toolTip;
        
        // 由 GitHub Copilot 產生 - 警告提示標籤
        private Label lblWarningControl;
        
        // 由 GitHub Copilot 產生
        // 誤觸樣品組控件成員變數
        private PictureBox picFalsePreviewControl;
        private PictureBox picFalseBinaryPreviewControl;
        private TextBox lblFalseImageInfoControl;
        private Button btnFalsePreviousControl;
        private Button btnFalseNextControl;
        private Button btnSelectFalseImagesControl;
        
        // 由 GitHub Copilot 產生
        // 正常樣品二值化預覽控件
        private PictureBox picBinaryPreviewControl;
        
        // 二值化閾值控件
        private TrackBar trackBarThreshold;
        private Label lblThresholdValue;
        private System.Threading.Timer _thresholdUpdateTimer;
        #endregion

        #region 白色像素分析結果結構
        // 由 GitHub Copilot 產生
        // 樣品類型列舉
        public enum SampleType
        {
            Normal,  // 正常樣品
            False    // 誤觸樣品
        }

        // 由 GitHub Copilot 產生
        // 離群值嚴重程度列舉
        public enum OutlierSeverity
        {
            Normal,          // 正常
            Mild,            // 輕微異常（1σ-2σ）
            GroupOutlier,    // 組內離群（IQR 檢測）
            CrossSample      // 疑似誤放（組間交叉）
        }

        public class WhitePixelResult
        {
            public string ImagePath { get; set; }
            public float WhitePixelRatio { get; set; }
            public int TotalPixels { get; set; }
            public int WhitePixels { get; set; }
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; }
            public SampleType Type { get; set; } = SampleType.Normal; // 樣品類型
            //public Mat ResultImage { get; set; }
            private byte[] _compressedImageData; // 儲存壓縮後的原圖標註資料
            private byte[] _compressedBinaryImage; // 儲存壓縮後的二值化圖像資料
            public bool IsOutlier1Sigma { get; set; }
            public bool IsOutlier2Sigma { get; set; }
            // 由 GitHub Copilot 產生 - 離群值嚴重程度（用於混合檢測）
            public OutlierSeverity Severity { get; set; } = OutlierSeverity.Normal;
            // 設定壓縮後的原圖標註資料
            public void SetResultImageData(Mat resultImage)
            {
                if (resultImage != null && !resultImage.Empty())
                {
                    try
                    {
                        // 壓縮圖像為 JPEG 格式以節省記憶體
                        var compressedBytes = resultImage.ImEncode(".jpg", new int[] { (int)ImwriteFlags.JpegQuality, 85 });
                        _compressedImageData = compressedBytes;
                    }
                    catch
                    {
                        _compressedImageData = null;
                    }
                }
            }
            public bool HasImageData()
            {
                return _compressedImageData != null && _compressedImageData.Length > 0;
            }
            
            // 按需取得原圖標註
            public Mat GetResultImage()
            {
                if (_compressedImageData != null && _compressedImageData.Length > 0)
                {
                    try
                    {
                        return Cv2.ImDecode(_compressedImageData, ImreadModes.Color);
                    }
                    catch
                    {
                        return new Mat();
                    }
                }
                return new Mat();
            }

            // 由 GitHub Copilot 產生
            // 設定壓縮後的二值化圖像資料
            public void SetBinaryImageData(Mat binaryImage)
            {
                if (binaryImage != null && !binaryImage.Empty())
                {
                    try
                    {
                        // 使用 PNG 格式壓縮二值化圖像（黑白圖壓縮率高）
                        var compressedBytes = binaryImage.ImEncode(".png");
                        _compressedBinaryImage = compressedBytes;
                    }
                    catch
                    {
                        _compressedBinaryImage = null;
                    }
                }
            }

            // 按需取得二值化圖像
            public Mat GetBinaryImage()
            {
                if (_compressedBinaryImage != null && _compressedBinaryImage.Length > 0)
                {
                    try
                    {
                        return Cv2.ImDecode(_compressedBinaryImage, ImreadModes.Grayscale);
                    }
                    catch
                    {
                        return new Mat();
                    }
                }
                return new Mat();
            }

            public bool HasBinaryImageData()
            {
                return _compressedBinaryImage != null && _compressedBinaryImage.Length > 0;
            }

            // 釋放資源
            public void Dispose()
            {
                _compressedImageData = null;
                _compressedBinaryImage = null;
            }
        }
        #endregion

        #region 建構函數和初始化
        public WhiteCalibrationForm(string type)
        {
            targetType = type;
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = $"白色像素占比校正 - {targetType}";
            this.Size = new System.Drawing.Size(1400, 1150); // 由 GitHub Copilot 產生 - 優化視窗高度以容納完整介面
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            CreateControls();
        }
        #endregion

        #region 控件建立和布局
        private void CreateControls()
        {
            // 初始化工具提示
            toolTip = new ToolTip();

            // 說明標籤
            var lblDescription = new Label
            {
                Text = "📋 請選擇多張照片分析白色像素占比，系統將計算統計數據並提供推薦的 white 參數值",
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(1160, 40),
                Font = new Font("Microsoft JhengHei", 10),
                ForeColor = Color.Blue
            };
            
            // 站點選擇
            var lblStation = new Label
            {
                Text = "選擇站點：",
                Location = new System.Drawing.Point(10, 55),
                Size = new System.Drawing.Size(80, 25),
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            cmbStation = new ComboBox
            {
                Location = new System.Drawing.Point(90, 55),
                Size = new System.Drawing.Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbStation.Items.AddRange(new object[] { "請選擇","站點1", "站點2", "站點3", "站點4" });
            cmbStation.SelectedIndex = 0;
            targetStation = 0;
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged;

            // 由 GitHub Copilot 產生
            // 選擇正常樣品照片按鈕
            var btnSelectImages = new Button
            {
                Text = "📂 選擇正常照片",
                Location = new System.Drawing.Point(210, 55),
                Size = new System.Drawing.Size(130, 30),
                BackColor = Color.LightBlue
            };
            btnSelectImages.Click += BtnSelectImages_Click;

            // 選擇誤觸樣品照片按鈕
            btnSelectFalseImagesControl = new Button
            {
                Text = "📂 選擇誤觸照片",
                Location = new System.Drawing.Point(350, 55),
                Size = new System.Drawing.Size(130, 30),
                BackColor = Color.LightCoral
            };
            btnSelectFalseImagesControl.Click += BtnSelectFalseImages_Click;

            // 二值化閾值滑桿標籤
            var lblThreshold = new Label
            {
                Text = "二值化閾值：",
                Location = new System.Drawing.Point(210, 95),
                Size = new System.Drawing.Size(100, 25),
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            // 二值化閾值滑桿
            trackBarThreshold = new TrackBar
            {
                Minimum = 0,
                Maximum = 255,
                TickFrequency = 25,
                Value = 180,
                Size = new System.Drawing.Size(200, 45),
                Location = new System.Drawing.Point(310, 90)
            };
            trackBarThreshold.ValueChanged += TrackBarThreshold_ValueChanged;

            // 顯示當前閾值的標籤
            lblThresholdValue = new Label
            {
                Text = "180",
                Location = new System.Drawing.Point(520, 95),
                Size = new System.Drawing.Size(50, 25),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            toolTip.SetToolTip(trackBarThreshold, "調整二值化閾值 (0-255)，影響白色像素的判定標準");

            // 開始分析按鈕
            btnAnalyzeControl = new Button
            {
                Text = "🔍 開始分析",
                Location = new System.Drawing.Point(580, 90),
                Size = new System.Drawing.Size(180, 30),
                BackColor = Color.LightGreen,
                Enabled = false
            };
            btnAnalyzeControl.Click += BtnAnalyze_Click;

            // 進度條
            progressBarControl = new ProgressBar
            {
                Location = new System.Drawing.Point(770, 93),
                Size = new System.Drawing.Size(250, 25),
                Visible = false
            };

            // 結果顯示區域
            dgvResultsControl = new DataGridView
            {
                Location = new System.Drawing.Point(10, 140),
                Size = new System.Drawing.Size(580, 445), // 由 GitHub Copilot 產生 - 增加高度使底部與折線圖對齊（原305）
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            // 統計資訊面板
            grpStatisticsControl = new GroupBox
            {
                Text = "📊 統計分析結果",
                Location = new System.Drawing.Point(600, 140),
                Size = new System.Drawing.Size(760, 285) // 由 GitHub Copilot 產生 - 增加高度以容納完整的雙組統計資訊
            };
            CreateStatisticsLabels(grpStatisticsControl);

            // 由 GitHub Copilot 產生 - 警告提示標籤（用於顯示疑似問題的照片）
            lblWarningControl = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(600, 430),
                Size = new System.Drawing.Size(760, 35),
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold),
                ForeColor = Color.DarkRed,
                BackColor = Color.LightYellow,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(5, 0, 5, 0),
                Visible = false, // 初始隱藏，有警告時才顯示
                AutoEllipsis = true // 文字過長時顯示省略號
            };
            toolTip.SetToolTip(lblWarningControl, "點擊圖片列表中對應的照片查看詳情");

            // 折線圖
            chartWhitePixelRatio = new Chart
            {
                Location = new System.Drawing.Point(600, 470), // 由 GitHub Copilot 產生 - 調整位置為警告標籤留出空間（原435）
                Size = new System.Drawing.Size(760, 115), // 由 GitHub Copilot 產生 - 調整高度以維持整體布局（原150）
                BackColor = Color.White
            };
            InitializeChart();

            // 由 GitHub Copilot 產生
            // 正常樣品預覽區域
            var grpNormalPreview = new GroupBox
            {
                Text = "📷 正常樣品預覽",
                Location = new System.Drawing.Point(10, 595), // 由 GitHub Copilot 產生 - 調整位置配合折線圖高度增加（原510）
                Size = new System.Drawing.Size(670, 300),
                ForeColor = Color.Blue
            };

            // 正常樣品原圖標註預覽
            picPreviewControl = new PictureBox
            {
                Location = new System.Drawing.Point(10, 20),
                Size = new System.Drawing.Size(280, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };
            var lblNormalAnnotated = new Label
            {
                Text = "原圖標註",
                Location = new System.Drawing.Point(10, 195),
                Size = new System.Drawing.Size(280, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            // 正常樣品二值化預覽
            picBinaryPreviewControl = new PictureBox
            {
                Location = new System.Drawing.Point(300, 20),
                Size = new System.Drawing.Size(280, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };
            var lblNormalBinary = new Label
            {
                Text = "二值化圖",
                Location = new System.Drawing.Point(300, 195),
                Size = new System.Drawing.Size(280, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            // 正常樣品導航按鈕
            btnPreviousControl = new Button
            {
                Text = "◀ 上一張",
                Location = new System.Drawing.Point(590, 30),
                Size = new System.Drawing.Size(70, 30),
                Enabled = false
            };
            btnPreviousControl.Click += BtnPrevious_Click;

            btnNextControl = new Button
            {
                Text = "下一張 ▶",
                Location = new System.Drawing.Point(590, 70),
                Size = new System.Drawing.Size(70, 30),
                Enabled = false
            };
            btnNextControl.Click += BtnNext_Click;

            // 全螢幕預覽按鈕
            var btnFullScreen = new Button
            {
                Text = "🔍 全螢幕",
                Location = new System.Drawing.Point(590, 110),
                Size = new System.Drawing.Size(70, 30),
                BackColor = Color.LightBlue
            };
            btnFullScreen.Click += BtnFullScreen_Click;

            // 正常樣品圖片資訊標籤
            lblImageInfoControl = new TextBox
            {
                Text = "尚未選擇正常樣品",
                Location = new System.Drawing.Point(10, 230),
                Size = new System.Drawing.Size(650, 60),
                Font = new Font("Microsoft JhengHei", 9),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                WordWrap = true,
                TabStop = false
            };

            grpNormalPreview.Controls.AddRange(new Control[] {
                picPreviewControl, lblNormalAnnotated, 
                picBinaryPreviewControl, lblNormalBinary,
                btnPreviousControl, btnNextControl, btnFullScreen, lblImageInfoControl
            });

            // 誤觸樣品預覽區域
            var grpFalsePreview = new GroupBox
            {
                Text = "📷 誤觸樣品預覽",
                Location = new System.Drawing.Point(690, 595), // 由 GitHub Copilot 產生 - 調整位置配合折線圖高度增加（原510）
                Size = new System.Drawing.Size(670, 300),
                ForeColor = Color.Red
            };

            // 誤觸樣品原圖標註預覽
            picFalsePreviewControl = new PictureBox
            {
                Location = new System.Drawing.Point(10, 20),
                Size = new System.Drawing.Size(280, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };
            var lblFalseAnnotated = new Label
            {
                Text = "原圖標註",
                Location = new System.Drawing.Point(10, 195),
                Size = new System.Drawing.Size(280, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            // 誤觸樣品二值化預覽
            picFalseBinaryPreviewControl = new PictureBox
            {
                Location = new System.Drawing.Point(300, 20),
                Size = new System.Drawing.Size(280, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };
            var lblFalseBinary = new Label
            {
                Text = "二值化圖",
                Location = new System.Drawing.Point(300, 195),
                Size = new System.Drawing.Size(280, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            // 誤觸樣品導航按鈕
            btnFalsePreviousControl = new Button
            {
                Text = "◀ 上一張",
                Location = new System.Drawing.Point(590, 30),
                Size = new System.Drawing.Size(70, 30),
                Enabled = false
            };
            btnFalsePreviousControl.Click += BtnFalsePrevious_Click;

            btnFalseNextControl = new Button
            {
                Text = "下一張 ▶",
                Location = new System.Drawing.Point(590, 70),
                Size = new System.Drawing.Size(70, 30),
                Enabled = false
            };
            btnFalseNextControl.Click += BtnFalseNext_Click;

            // 誤觸樣品圖片資訊標籤
            lblFalseImageInfoControl = new TextBox
            {
                Text = "尚未選擇誤觸樣品",
                Location = new System.Drawing.Point(10, 230),
                Size = new System.Drawing.Size(650, 60),
                Font = new Font("Microsoft JhengHei", 9),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                WordWrap = true,
                TabStop = false
            };

            grpFalsePreview.Controls.AddRange(new Control[] {
                picFalsePreviewControl, lblFalseAnnotated,
                picFalseBinaryPreviewControl, lblFalseBinary,
                btnFalsePreviousControl, btnFalseNextControl, lblFalseImageInfoControl
            });

            // 建議值顯示
            lblRecommendationControl = new Label
            {
                Text = "💡 推薦的參數值將在分析完成後顯示",
                Location = new System.Drawing.Point(10, 905), // 由 GitHub Copilot 產生 - 調整位置配合預覽區位置變更（原835）
                Size = new System.Drawing.Size(700, 50),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.DarkGreen
            };

            // 套用按鈕
            btnApplyControl = new Button
            {
                Text = "✅ 套用推薦值",
                Location = new System.Drawing.Point(720, 910), // 由 GitHub Copilot 產生 - 調整位置配合預覽區位置變更（原840）
                Size = new System.Drawing.Size(120, 40),
                BackColor = Color.Orange,
                Enabled = false,
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold)
            };
            btnApplyControl.Click += BtnApply_Click;

            // 取消按鈕
            var btnCancel = new Button
            {
                Text = "❌ 取消",
                Location = new System.Drawing.Point(850, 910), // 由 GitHub Copilot 產生 - 調整位置配合預覽區位置變更（原840）
                Size = new System.Drawing.Size(80, 40),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            // 匯出數據按鈕
            var btnExport = new Button
            {
                Text = "📊 匯出數據",
                Location = new System.Drawing.Point(940, 910), // 由 GitHub Copilot 產生 - 調整位置配合預覽區位置變更（原840）
                Size = new System.Drawing.Size(100, 40),
                BackColor = Color.LightCyan
            };
            btnExport.Click += BtnExport_Click;

            // 清除結果按鈕
            var btnClearResults = new Button
            {
                Text = "🗑️ 清除結果",
                Location = new System.Drawing.Point(1050, 910), // 由 GitHub Copilot 產生 - 調整位置配合預覽區位置變更（原840）
                Size = new System.Drawing.Size(100, 40),
                BackColor = System.Drawing.Color.LightPink
            };
            btnClearResults.Click += BtnClearResults_Click;
            toolTip.SetToolTip(btnClearResults, "清除目前的分析結果，可重新選擇照片分析");

            // 加入所有控件
            this.Controls.AddRange(new Control[] {
                lblDescription, lblStation, cmbStation, btnSelectImages, btnSelectFalseImagesControl,
                lblThreshold, trackBarThreshold, lblThresholdValue, btnAnalyzeControl, progressBarControl,
                dgvResultsControl, grpStatisticsControl, lblWarningControl, chartWhitePixelRatio, grpNormalPreview, grpFalsePreview,
                lblRecommendationControl, btnApplyControl, btnCancel, btnExport, btnClearResults
            });
        }
        // 由 GitHub Copilot 產生
        // 新增清除分析結果的事件處理
        private void BtnClearResults_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0)
            {
                MessageBox.Show("目前沒有分析結果需要清除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show("確定要清除目前的分析結果嗎？\n這將清除所有已分析的照片資料。",
                "確認清除", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                ResetAnalysisResults();

                // 重置UI狀態
                dgvResultsControl.DataSource = null;
                lblImageInfoControl.Text = "尚未分析圖片";
                lblRecommendationControl.Text = "💡 推薦的 white 參數值將在分析完成後顯示";
                lblRecommendationControl.ForeColor = Color.DarkGreen;
                
                // 由 GitHub Copilot 產生 - 隱藏警告標籤
                lblWarningControl.Visible = false;
                lblWarningControl.Text = "";

                // 重置統計標籤
                for (int i = 0; i < 10; i++)
                {
                    var lblValue = grpStatisticsControl.Controls.Find($"lblStatValue{i}", false).FirstOrDefault() as Label;
                    if (lblValue != null)
                    {
                        lblValue.Text = "---";
                    }
                }

                // 重置按鈕狀態
                btnApplyControl.Enabled = false;
                btnApplyControl.Text = "✅ 套用推薦值";
                btnApplyControl.BackColor = System.Drawing.Color.Orange;
                btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
                UpdateNavigationButtons();

                MessageBox.Show("已清除所有分析結果", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void CreateStatisticsLabels(GroupBox parent)
        {
            // 由 GitHub Copilot 產生 - 分為正常樣品和誤觸樣品兩組統計標籤
            var normalLabels = new[] {
                "正常樣品數：", "有效樣品：", "平均占比：", "中位數占比：",
                "標準差：", "最小值：", "最大值：", "1σ異常：", "2σ異常：", "推薦值："
            };

            var falseLabels = new[] {
                "誤觸樣品數：", "有效樣品：", "平均占比：", "中位數占比：",
                "標準差：", "最小值：", "最大值：", "1σ異常：", "2σ異常：", "推薦值："
            };

            // 正常樣品統計 (上半部)
            for (int i = 0; i < normalLabels.Length; i++)
            {
                var lbl = new Label
                {
                    Text = normalLabels[i],
                    Location = new System.Drawing.Point(10, 25 + (i % 5) * 25),
                    Size = new System.Drawing.Size(100, 25),
                    Name = $"lblNormalStat{i}",
                    ForeColor = Color.Blue
                };

                var lblValue = new Label
                {
                    Text = "---",
                    Location = new System.Drawing.Point(120, 25 + (i % 5) * 25),
                    Size = new System.Drawing.Size(120, 25),
                    Name = $"lblStatValue{i}",
                    Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
                };

                // 第二列
                if (i >= 5)
                {
                    lbl.Location = new System.Drawing.Point(270, 25 + (i - 5) * 25);
                    lblValue.Location = new System.Drawing.Point(380, 25 + (i - 5) * 25);
                }

                parent.Controls.AddRange(new Control[] { lbl, lblValue });
            }

            // 由 GitHub Copilot 產生 - 誤觸樣品統計（下半部，留10px間距）
            int falseStartY = 155; // 正常樣品結束於 25 + 5*25 = 150，留5px間距
            for (int i = 0; i < falseLabels.Length; i++)
            {
                var lbl = new Label
                {
                    Text = falseLabels[i],
                    Location = new System.Drawing.Point(10, falseStartY + (i % 5) * 25),
                    Size = new System.Drawing.Size(100, 25),
                    Name = $"lblFalseStat{i}",
                    ForeColor = Color.Red
                };

                var lblValue = new Label
                {
                    Text = "---",
                    Location = new System.Drawing.Point(120, falseStartY + (i % 5) * 25),
                    Size = new System.Drawing.Size(120, 25),
                    Name = $"lblFalseStatValue{i}",
                    Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold),
                    ForeColor = Color.Red // 由 GitHub Copilot 產生 - 誤觸樣品數據用紅色顯示
                };

                // 第二列
                if (i >= 5)
                {
                    lbl.Location = new System.Drawing.Point(270, falseStartY + (i - 5) * 25);
                    lblValue.Location = new System.Drawing.Point(380, falseStartY + (i - 5) * 25);
                }

                parent.Controls.AddRange(new Control[] { lbl, lblValue });
            }
        }

        // 新增圖表初始化標誌
        private bool chartInitialized = false;

        private void InitializeChart()
        {
            if (chartInitialized) return; // 避免重複初始化

            chartWhitePixelRatio.Series.Clear();
            chartWhitePixelRatio.ChartAreas.Clear();
            chartWhitePixelRatio.Legends.Clear();
            chartWhitePixelRatio.Titles.Clear();

            var chartArea = new ChartArea("MainArea")
            {
                AxisX = { Title = "圖片順序", MajorGrid = { Enabled = true, LineColor = Color.LightGray } },
                AxisY = { Title = "白色像素占比 (%)", MajorGrid = { Enabled = true, LineColor = Color.LightGray } }
            };
            chartWhitePixelRatio.ChartAreas.Add(chartArea);

            // 主要數據系列
            var mainSeries = new Series("白色像素占比")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Blue,
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 6,
                BorderWidth = 2
            };
            chartWhitePixelRatio.Series.Add(mainSeries);

            // 1σ異常點系列
            var outlier1Series = new Series("1σ異常")
            {
                ChartType = SeriesChartType.Point,
                Color = Color.Orange,
                MarkerStyle = MarkerStyle.Triangle,
                MarkerSize = 8
            };
            chartWhitePixelRatio.Series.Add(outlier1Series);

            // 2σ異常點系列
            var outlier2Series = new Series("2σ異常")
            {
                ChartType = SeriesChartType.Point,
                Color = Color.Red,
                MarkerStyle = MarkerStyle.Star4,
                MarkerSize = 10
            };
            chartWhitePixelRatio.Series.Add(outlier2Series);

            // 由 GitHub Copilot 產生 - 新增誤觸樣品系列
            var falseSeries = new Series("誤觸樣品")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Red,
                MarkerStyle = MarkerStyle.Diamond,
                MarkerSize = 6,
                BorderWidth = 2
            };
            chartWhitePixelRatio.Series.Add(falseSeries);

            // 平均值線
            var avgSeries = new Series("平均值")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Green,
                BorderDashStyle = ChartDashStyle.Dash,
                BorderWidth = 2
            };
            chartWhitePixelRatio.Series.Add(avgSeries);

            // ±1σ線
            var sigma1UpperSeries = new Series("+1σ")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Orange,
                BorderDashStyle = ChartDashStyle.Dot,
                BorderWidth = 1
            };
            chartWhitePixelRatio.Series.Add(sigma1UpperSeries);

            var sigma1LowerSeries = new Series("-1σ")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Orange,
                BorderDashStyle = ChartDashStyle.Dot,
                BorderWidth = 1
            };
            chartWhitePixelRatio.Series.Add(sigma1LowerSeries);

            // ±2σ線
            var sigma2UpperSeries = new Series("+2σ")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Red,
                BorderDashStyle = ChartDashStyle.DashDot,
                BorderWidth = 1
            };
            chartWhitePixelRatio.Series.Add(sigma2UpperSeries);

            var sigma2LowerSeries = new Series("-2σ")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Red,
                BorderDashStyle = ChartDashStyle.DashDot,
                BorderWidth = 1
            };
            chartWhitePixelRatio.Series.Add(sigma2LowerSeries);

            // 設定圖例
            var legend = new Legend("MainLegend")
            {
                Docking = Docking.Top,
                Alignment = StringAlignment.Center
            };
            chartWhitePixelRatio.Legends.Add(legend);

            chartInitialized = true;
        }

        #endregion

        #region 事件處理器
        // 由 GitHub Copilot 產生
        // 二值化閾值滑桿值變更事件（使用防抖動機制防止介面卡死）
        private void TrackBarThreshold_ValueChanged(object sender, EventArgs e)
        {
            // 立即更新顯示的數值
            lblThresholdValue.Text = trackBarThreshold.Value.ToString();

            // 使用 Timer 實現防抖動，避免快速滑動時介面卡死
            // 每次滑動都會重置計時器，只有停止滑動 300ms 後才會觸發更新
            if (_thresholdUpdateTimer != null)
            {
                _thresholdUpdateTimer.Change(300, System.Threading.Timeout.Infinite);
            }
            else
            {
                _thresholdUpdateTimer = new System.Threading.Timer(
                    callback: _ =>
                    {
                        try
                        {
                            // 在 UI 執行緒上更新提示訊息
                            this.BeginInvoke(new Action(() =>
                            {
                                if (whitePixelResults.Count > 0)
                                {
                                    // 提示使用者需要重新分析
                                    lblRecommendationControl.Text = $"⚠️ 閾值已變更為 {trackBarThreshold.Value}，請重新分析以更新結果";
                                    lblRecommendationControl.ForeColor = Color.OrangeRed;
                                }
                            }));
                        }
                        catch { }
                    },
                    state: null,
                    dueTime: 300,
                    period: System.Threading.Timeout.Infinite
                );
            }
        }

        // 由 GitHub Copilot 產生
        // 修正站點選擇事件，重置套用按鈕狀態並載入閾值
        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            targetStation = cmbStation.SelectedIndex; // 0 = 未選
            if (targetStation <= 0)
            {
                this.Text = $"白色像素占比校正 - {targetType} (請選擇站點)";
                btnAnalyzeControl.Enabled = false;
                btnApplyControl.Enabled = false;
                trackBarThreshold.Enabled = false;
                return;
            }

            this.Text = $"白色像素占比校正 - {targetType} 站點{targetStation}";
            trackBarThreshold.Enabled = true;

            // 從資料庫讀取該站點的 whiteThresh 設定值
            int defaultThresh = GetDefaultThreshold(targetStation);
            int savedThresh = Form1.GetIntParam(app.param, $"whiteThresh_{targetStation}", defaultThresh);
            trackBarThreshold.Value = savedThresh;
            lblThresholdValue.Text = savedThresh.ToString();

            // 若已經選擇了圖片才允許分析
            btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
            // 推薦值出來後才會再啟用套用；此處只重置
            if (recommendedWhiteValue > 0)
            {
                btnApplyControl.Text = "✅ 套用推薦值";
                btnApplyControl.BackColor = System.Drawing.Color.Orange;
            }
        }

        // 由 GitHub Copilot 產生
        // 取得站點預設閾值
        private int GetDefaultThreshold(int station)
        {
            if (station == 1 || station == 4)
                return 180;
            else if (station == 2)
                return 250;
            else if (station == 3)
                return 170;
            else
                return 180; // 預設值
        }

        private void BtnSelectImages_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "選擇樣品照片";
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedImagePaths = openFileDialog.FileNames.ToList();

                    if (selectedImagePaths.Count < 10)
                    {
                        MessageBox.Show("建議至少選擇10張照片以獲得更準確的統計結果", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (selectedImagePaths.Count > 100)
                    {
                        MessageBox.Show("選擇的照片過多，建議選擇30-50張代表性照片", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    // 由 GitHub Copilot 產生 - 更新正常照片按鈕文字
                    var btnSelectImages = sender as Button;
                    if (btnSelectImages != null)
                    {
                        btnSelectImages.Text = $"正常照片 ({selectedImagePaths.Count})";
                    }

                    // 由 GitHub Copilot 產生 - 更新正常樣品預覽區域提示文字
                    if (lblImageInfoControl != null)
                    {
                        lblImageInfoControl.Text = $"已選擇 {selectedImagePaths.Count} 張正常樣品照片，點擊「開始分析」進行分析";
                        lblImageInfoControl.ForeColor = Color.Blue;
                    }

                    // 檢查是否已選擇兩組照片，如果是則啟用分析按鈕
                    btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0 && selectedFalseImagePaths.Count > 0;
                    if (btnAnalyzeControl.Enabled)
                    {
                        btnAnalyzeControl.Text = $"🔍 分析 (正常:{selectedImagePaths.Count} / 誤觸:{selectedFalseImagePaths.Count})";
                    }
                    else if (selectedImagePaths.Count > 0)
                    {
                        btnAnalyzeControl.Text = $"🔍 開始分析 (需選擇誤觸照片)";
                    }
                }
            }
        }

        private async void BtnAnalyze_Click(object sender, EventArgs e)
        {
            if (targetStation == 0)
            {
                MessageBox.Show("請先選擇站點", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (isAnalyzing)
            {
                MessageBox.Show("正在分析中，請稍等...", "提示");
                return;
            }

            // 由 GitHub Copilot 產生 - 檢查雙組照片
            if (selectedImagePaths.Count == 0)
            {
                MessageBox.Show("請先選擇正常樣品照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            if (selectedFalseImagePaths.Count == 0)
            {
                MessageBox.Show("請先選擇誤觸樣品照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            isAnalyzing = true;
            btnAnalyzeControl.Enabled = false;
            btnAnalyzeControl.Text = "🔄 分析中...";

            try
            {
                // 由 GitHub Copilot 產生 - 處理雙組圖片
                progressBarControl.Maximum = selectedImagePaths.Count + selectedFalseImagePaths.Count;
                progressBarControl.Value = 0;
                progressBarControl.Visible = true;

                whitePixelResults.Clear();
                falsePixelResults.Clear();

                // 分析正常樣品
                await ProcessImages(selectedImagePaths, whitePixelResults, SampleType.Normal);
                
                // 分析誤觸樣品
                await ProcessImages(selectedFalseImagePaths, falsePixelResults, SampleType.False);
                
                AnalyzeOutliers();
                
                // 由 GitHub Copilot 產生 - 執行混合檢測（組內離群 + 組間交叉污染）
                DetectOutliersAndCrossSamples();
                
                // 由 GitHub Copilot 產生 - 先計算推薦參數，再顯示結果（避免顯示時 recommendedTolerance 尚未計算）
                CalculateRecommendedParameters();

                Console.WriteLine("000");
                DisplayAnalysisResults();
                Console.WriteLine("111");
                UpdateChart();
                Console.WriteLine("222");
                MessageBox.Show($"白色像素占比分析完成！\n\n" +
                               $"正常樣品：{whitePixelResults.Count} 張 (推薦值: {recommendedWhiteValue:F1}%)\n" +
                               $"誤觸樣品：{falsePixelResults.Count} 張 (基準值: {recommendedWhiteNullValue:F1}%)\n" +
                               $"推薦容忍度：{recommendedTolerance:F1}%\n\n" +
                               $"正常樣品 1σ異常：{whitePixelResults.Count(r => r.IsOutlier1Sigma)} 張\n" +
                               $"正常樣品 2σ異常：{whitePixelResults.Count(r => r.IsOutlier2Sigma)} 張",
                               "分析完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"分析錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                isAnalyzing = false;
                btnAnalyzeControl.Enabled = true;
                btnAnalyzeControl.Text = $"🔍 分析 (正常:{selectedImagePaths.Count} / 誤觸:{selectedFalseImagePaths.Count})";
                progressBarControl.Visible = false;
            }
        }

        private bool _isNavigating = false; // 防護標誌

        private void BtnPrevious_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0 || _isNavigating) return;

            try
            {
                _isNavigating = true;

                // 立即禁用所有導航按鈕
                btnPreviousControl.Enabled = false;
                btnNextControl.Enabled = false;

                currentImageIndex = (currentImageIndex - 1 + whitePixelResults.Count) % whitePixelResults.Count;

                // 使用異步方式更新圖像
                Task.Run(() =>
                {
                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            DisplayCurrentImage();
                        }));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"導航錯誤: {ex.Message}");
                    }
                    finally
                    {
                        // 延遲重新啟用按鈕
                        Task.Delay(200).ContinueWith(t =>
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                _isNavigating = false;
                                UpdateNavigationButtons();
                            }));
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                _isNavigating = false;
                MessageBox.Show($"導航錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateNavigationButtons();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0 || _isNavigating) return;

            try
            {
                _isNavigating = true;

                // 立即禁用所有導航按鈕
                btnPreviousControl.Enabled = false;
                btnNextControl.Enabled = false;

                currentImageIndex = (currentImageIndex + 1) % whitePixelResults.Count;

                // 使用異步方式更新圖像
                Task.Run(() =>
                {
                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            DisplayCurrentImage();
                        }));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"導航錯誤: {ex.Message}");
                    }
                    finally
                    {
                        // 延遲重新啟用按鈕
                        Task.Delay(200).ContinueWith(t =>
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                _isNavigating = false;
                                UpdateNavigationButtons();
                            }));
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                _isNavigating = false;
                MessageBox.Show($"導航錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateNavigationButtons();
            }
        }
        private void BtnFullScreen_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0 || currentImageIndex >= whitePixelResults.Count || _isNavigating) return;

            try
            {
                var currentResult = whitePixelResults[currentImageIndex];

                // 檢查是否有圖像資料
                if (!currentResult.HasImageData())
                {
                    MessageBox.Show("此圖像沒有可用的預覽資料", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ShowFullScreenPreview(currentResult);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"預覽錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"全螢幕預覽錯誤: {ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生
        // 修正套用推薦值按鈕事件，讓介面保持開啟狀態
        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (targetStation == 0)
            {
                MessageBox.Show("尚未選擇站點，無法套用。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // 由 GitHub Copilot 產生 - 檢查是否已計算推薦參數
            if (recommendedWhiteValue == 0 || recommendedWhiteNullValue == 0)
            {
                MessageBox.Show("請先完成正常樣品和誤觸樣品的分析", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                using (var db = new MydbDB())
                {
                    // 由 GitHub Copilot 產生 - 套用四個參數
                    UpdateParameter(db, "white", recommendedWhiteValue.ToString("F1"));
                    UpdateParameter(db, "whiteNULL", recommendedWhiteNullValue.ToString("F1"));
                    UpdateParameter(db, "whiteThresh", trackBarThreshold.Value.ToString());
                    UpdateParameter(db, "whiteTolerance", recommendedTolerance.ToString("F1"));
                }

                MessageBox.Show($"成功套用參數設定！\n\n" +
                               $"white_{targetStation} = {recommendedWhiteValue:F1}%\n" +
                               $"whiteNULL_{targetStation} = {recommendedWhiteNullValue:F1}%\n" +
                               $"whiteThresh_{targetStation} = {trackBarThreshold.Value}\n" +
                               $"whiteTolerance_{targetStation} = {recommendedTolerance:F1}%\n\n" +
                               $"判定區間：[{(recommendedWhiteValue - recommendedTolerance):F1}%, {(recommendedWhiteValue + recommendedTolerance):F1}%]\n\n" +
                               $"您可以繼續選擇其他照片進行校正，或關閉視窗。",
                               "設定完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 移除自動關閉視窗的程式碼
                // this.DialogResult = DialogResult.OK;

                // 可選：重新啟用分析按鈕讓使用者可以重新分析
                btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0 && selectedFalseImagePaths.Count > 0;

                // 可選：更新按鈕文字提示已套用
                btnApplyControl.Text = "✅ 已套用 (可重新套用)";
                btnApplyControl.BackColor = System.Drawing.Color.LightGreen;

                // 重置提示訊息顏色
                lblRecommendationControl.ForeColor = Color.DarkGreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"套用設定時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0)
            {
                MessageBox.Show("沒有分析結果可以匯出", "提示");
                return;
            }

            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV 檔案|*.csv";
                saveFileDialog.FileName = $"white_calibration_{targetType}_station{targetStation}_{DateTime.Now:yyyyMMdd_HHmmss}.csv";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportResults(saveFileDialog.FileName);
                    MessageBox.Show("數據匯出成功！", "完成");
                }
            }
        }

        // 由 GitHub Copilot 產生
        // 修正 DataGridView 選擇變更事件（支援正常樣品和誤觸樣品連動預覽）
        private void DgvResults_SelectionChanged(object sender, EventArgs e)
        {
            if (isAnalyzing || _isNavigating) return;
            if (dgvResultsControl.SelectedRows.Count == 0) return;

            try
            {
                _isNavigating = true;
                int selectedRowIndex = dgvResultsControl.SelectedRows[0].Index;

                // 由 GitHub Copilot 產生 - 判斷選中的是正常樣品還是誤觸樣品
                // DataGridView 顯示順序：先正常樣品，後誤觸樣品
                if (selectedRowIndex < whitePixelResults.Count)
                {
                    // 選中的是正常樣品
                    int newNormalIndex = selectedRowIndex;
                    
                    // 避免重複更新同一張圖片
                    if (newNormalIndex == currentImageIndex)
                    {
                        _isNavigating = false;
                        return;
                    }

                    currentImageIndex = newNormalIndex;

                    // 使用異步方式更新圖像，避免阻塞UI
                    Task.Run(() =>
                    {
                        try
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                DisplayCurrentImage();
                                UpdateNavigationButtons();
                                _isNavigating = false;
                            }));
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"正常樣品預覽錯誤: {ex.Message}");
                            _isNavigating = false;
                        }
                    });
                }
                else
                {
                    // 選中的是誤觸樣品
                    int newFalseIndex = selectedRowIndex - whitePixelResults.Count;
                    
                    // 避免重複更新同一張圖片
                    if (newFalseIndex == currentFalseImageIndex)
                    {
                        _isNavigating = false;
                        return;
                    }

                    currentFalseImageIndex = newFalseIndex;

                    // 使用異步方式更新圖像，避免阻塞UI
                    Task.Run(() =>
                    {
                        try
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                DisplayCurrentFalseImage();
                                UpdateFalseNavigationButtons();
                                _isNavigating = false;
                            }));
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"誤觸樣品預覽錯誤: {ex.Message}");
                            _isNavigating = false;
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DgvResults_SelectionChanged 錯誤: {ex.Message}");
                _isNavigating = false;
            }
        }

        // 由 GitHub Copilot 產生
        // 誤觸樣品事件處理器
        private void BtnSelectFalseImages_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "選擇誤觸樣品照片";
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFalseImagePaths = openFileDialog.FileNames.ToList();

                    if (selectedFalseImagePaths.Count < 10)
                    {
                        MessageBox.Show("建議至少選擇10張誤觸照片以獲得更準確的統計結果", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    btnSelectFalseImagesControl.Text = $"📂 誤觸照片 ({selectedFalseImagePaths.Count})";
                    
                    // 由 GitHub Copilot 產生 - 更新誤觸樣品預覽區域提示文字
                    if (lblFalseImageInfoControl != null)
                    {
                        lblFalseImageInfoControl.Text = $"已選擇 {selectedFalseImagePaths.Count} 張誤觸樣品照片，點擊「🔍 開始分析」進行分析";
                        lblFalseImageInfoControl.ForeColor = Color.Blue;
                    }
                    
                    // 檢查是否已選擇兩組照片，如果是則啟用分析按鈕
                    btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0 && selectedFalseImagePaths.Count > 0;
                    if (btnAnalyzeControl.Enabled)
                    {
                        btnAnalyzeControl.Text = $"🔍 分析 (正常:{selectedImagePaths.Count} / 誤觸:{selectedFalseImagePaths.Count})";
                    }
                    else if (selectedFalseImagePaths.Count > 0)
                    {
                        btnAnalyzeControl.Text = $"🔍 開始分析 (需選擇正常照片)";
                    }
                }
            }
        }

        private void BtnFalsePrevious_Click(object sender, EventArgs e)
        {
            if (falsePixelResults.Count == 0 || _isNavigating) return;

            try
            {
                _isNavigating = true;
                btnFalsePreviousControl.Enabled = false;
                btnFalseNextControl.Enabled = false;

                currentFalseImageIndex--;
                if (currentFalseImageIndex < 0)
                    currentFalseImageIndex = falsePixelResults.Count - 1;

                Task.Run(() =>
                {
                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            DisplayCurrentFalseImage();
                        }));
                    }
                    finally
                    {
                        Task.Delay(200).ContinueWith(t =>
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                _isNavigating = false;
                                UpdateFalseNavigationButtons();
                            }));
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                _isNavigating = false;
                MessageBox.Show($"導航錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateFalseNavigationButtons();
            }
        }

        private void BtnFalseNext_Click(object sender, EventArgs e)
        {
            if (falsePixelResults.Count == 0 || _isNavigating) return;

            try
            {
                _isNavigating = true;
                btnFalsePreviousControl.Enabled = false;
                btnFalseNextControl.Enabled = false;

                currentFalseImageIndex = (currentFalseImageIndex + 1) % falsePixelResults.Count;

                Task.Run(() =>
                {
                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            DisplayCurrentFalseImage();
                        }));
                    }
                    finally
                    {
                        Task.Delay(200).ContinueWith(t =>
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                _isNavigating = false;
                                UpdateFalseNavigationButtons();
                            }));
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                _isNavigating = false;
                MessageBox.Show($"導航錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateFalseNavigationButtons();
            }
        }
        #endregion

        #region 圖像處理方法
        // 由 GitHub Copilot 產生 - 支援雙組處理的方法
        private async Task ProcessImages(List<string> imagePaths, List<WhitePixelResult> results, SampleType sampleType)
        {
            await Task.Run(() =>
            {
                for (int i = 0; i < imagePaths.Count; i++)
                {
                    string imagePath = imagePaths[i];

                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            progressBarControl.Value += 1;
                            string typeText = sampleType == SampleType.Normal ? "正常" : "誤觸";
                            btnAnalyzeControl.Text = $"🔄 處理{typeText}樣品 {i + 1}/{imagePaths.Count}";
                        }));

                        var result = ProcessSingleImage(imagePath);
                        result.Type = sampleType; // 設定樣品類型
                        results.Add(result);

                        System.Threading.Thread.Sleep(50);
                    }
                    catch (Exception ex)
                    {
                        results.Add(new WhitePixelResult
                        {
                            ImagePath = Path.GetFileName(imagePath),
                            IsValid = false,
                            ErrorMessage = ex.Message
                        });
                    }
                }

                this.BeginInvoke(new Action(() =>
                {
                    progressBarControl.Value = selectedImagePaths.Count;
                }));
            });
        }

        private WhitePixelResult ProcessSingleImage(string imagePath)
        {
            // 在 UI 執行緒外先取得閾值，避免跨執行緒存取控件
            // 使用 Invoke 而非 BeginInvoke，因為我們需要立即取得值才能繼續處理
            int currentThreshold = 0;
            this.Invoke(new Action(() => currentThreshold = trackBarThreshold.Value));

            using (var image = Cv2.ImRead(imagePath))
            {
                if (image.Empty())
                {
                    throw new Exception("無法載入圖片");
                }

                // 執行白色像素分析
                using (var resultImage = new Mat())
                using (var binaryImage = new Mat()) // 由 GitHub Copilot 產生 - 新增二值化圖像
                {
                    var (whitePixelRatio, totalPixels, whitePixels) = AnalyzeWhitePixels(image, resultImage, binaryImage, currentThreshold);

                    var result = new WhitePixelResult
                    {
                        ImagePath = System.IO.Path.GetFileName(imagePath),
                        WhitePixelRatio = whitePixelRatio,
                        TotalPixels = totalPixels,
                        WhitePixels = whitePixels,
                        IsValid = true,
                        Type = SampleType.Normal // 預設為正常樣品
                    };

                    // 由 GitHub Copilot 產生 - 設定壓縮後的圖像資料
                    result.SetResultImageData(resultImage);
                    result.SetBinaryImageData(binaryImage);

                    return result;
                }
            }
        } //演算法

        // 由 GitHub Copilot 產生 - 修改簽章以輸出二值化圖像
        private (float ratio, int totalPixels, int whitePixels) AnalyzeWhitePixels(Mat image, Mat resultImage, Mat binaryImage, int minthresh)
        {
            using (var gray = new Mat())
            using (var binary = new Mat())
            {
                // 轉為灰階
                Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                // 二值化處理
                Cv2.Threshold(gray, binary, minthresh, 255, ThresholdTypes.Binary);

                // 由 GitHub Copilot 產生 - 複製二值化結果供外部使用
                binary.CopyTo(binaryImage);

                // 計算白色像素
                int totalPixels = image.Width * image.Height;
                int whitePixels = Cv2.CountNonZero(binary);
                float ratio = (float)whitePixels / totalPixels * 100.0f;

                // 創建結果圖像（原圖標註版本）
                image.CopyTo(resultImage);

                // 在圖像上標註統計信息
                string info = $"White: {whitePixels}/{totalPixels} ({ratio:F2}%)";
                Cv2.PutText(resultImage, info, new OpenCvSharp.Point(10, 30),
                           HersheyFonts.HersheySimplex, 1.0, new Scalar(0, 255, 0), 2);

                // 標註使用的閾值
                string threshInfo = $"Threshold: {minthresh}";
                Cv2.PutText(resultImage, threshInfo, new OpenCvSharp.Point(10, 70),
                           HersheyFonts.HersheySimplex, 1.0, new Scalar(255, 255, 0), 2);

                // 可視化白色區域（可選）
                using (var colorBinary = new Mat())
                {
                    Cv2.CvtColor(binary, colorBinary, ColorConversionCodes.GRAY2BGR);
                    Cv2.AddWeighted(resultImage, 0.7, colorBinary, 0.3, 0, resultImage);
                }

                return (ratio, totalPixels, whitePixels);
            }
        }
        #endregion

        #region 統計分析方法
        private void AnalyzeOutliers()
        {
            // 由 GitHub Copilot 產生 - 分析正常樣品的離群值
            var validResults = whitePixelResults.Where(r => r.IsValid).ToList();
            if (validResults.Count < 3) return;

            var ratios = validResults.Select(r => r.WhitePixelRatio).ToList();
            double mean = ratios.Average();
            double stdDev = CalculateStandardDeviation(ratios, mean);

            // 標記異常值
            foreach (var result in validResults)
            {
                double deviation = Math.Abs(result.WhitePixelRatio - mean);

                if (deviation > 2 * stdDev)
                {
                    result.IsOutlier2Sigma = true;
                    result.IsOutlier1Sigma = true; // 2σ 也包含 1σ
                }
                else if (deviation > stdDev)
                {
                    result.IsOutlier1Sigma = true;
                }
            }

            // 計算推薦值（中位數）
            var sortedRatios = ratios.OrderBy(x => x).ToList();
            recommendedWhiteValue = GetMedian(sortedRatios);

            // 由 GitHub Copilot 產生 - 分析誤觸樣品的離群值
            var validFalseResults = falsePixelResults.Where(r => r.IsValid).ToList();
            if (validFalseResults.Count >= 3)
            {
                var falseRatios = validFalseResults.Select(r => r.WhitePixelRatio).ToList();
                double falseMean = falseRatios.Average();
                double falseStdDev = CalculateStandardDeviation(falseRatios, falseMean);

                // 標記誤觸樣品的異常值
                foreach (var result in validFalseResults)
                {
                    double deviation = Math.Abs(result.WhitePixelRatio - falseMean);

                    if (deviation > 2 * falseStdDev)
                    {
                        result.IsOutlier2Sigma = true;
                        result.IsOutlier1Sigma = true;
                    }
                    else if (deviation > falseStdDev)
                    {
                        result.IsOutlier1Sigma = true;
                    }
                }
            }
        }

        // 由 GitHub Copilot 產生
        // 混合檢測：組內離群值檢測 + 組間交叉污染檢測（改良版：避免誤判分離良好的樣品組）
        private void DetectOutliersAndCrossSamples()
        {
            var validNormal = whitePixelResults.Where(r => r.IsValid).ToList();
            var validFalse = falsePixelResults.Where(r => r.IsValid).ToList();

            if (validNormal.Count < 3 || validFalse.Count < 3)
            {
                return; // 樣品數量不足，無法進行檢測
            }

            var normalRatios = validNormal.Select(r => r.WhitePixelRatio).OrderBy(x => x).ToList();
            var falseRatios = validFalse.Select(r => r.WhitePixelRatio).OrderBy(x => x).ToList();

            // 計算正常樣品的 IQR 邊界
            float q1_normal = GetPercentile(normalRatios, 25);
            float q3_normal = GetPercentile(normalRatios, 75);
            float iqr_normal = q3_normal - q1_normal;
            float lower_normal = q1_normal - 1.5f * iqr_normal;
            float upper_normal = q3_normal + 1.5f * iqr_normal;

            // 計算誤觸樣品的 IQR 邊界
            float q1_false = GetPercentile(falseRatios, 25);
            float q3_false = GetPercentile(falseRatios, 75);
            float iqr_false = q3_false - q1_false;
            float lower_false = q1_false - 1.5f * iqr_false;
            float upper_false = q3_false + 1.5f * iqr_false;

            // 計算組間中界線
            float median_normal = GetMedian(normalRatios);
            float median_false = GetMedian(falseRatios);
            float boundary = (median_normal + median_false) / 2.0f;

            // 由 GitHub Copilot 產生 - 判斷是否需要啟用組間交叉檢測
            // 計算兩組中位數的距離
            float distance = Math.Abs(median_normal - median_false);
            
            // 計算兩組的範圍（用3σ或IQR範圍，取較大值）
            double std_normal = CalculateStandardDeviation(normalRatios, normalRatios.Average());
            double std_false = CalculateStandardDeviation(falseRatios, falseRatios.Average());
            
            float normal_range = Math.Max((float)(std_normal * 3.0), iqr_normal * 1.5f);
            float false_range = Math.Max((float)(std_false * 3.0), iqr_false * 1.5f);
            
            // 只有當距離小於兩組範圍之和時，才啟用交叉檢測
            // 這樣可以避免誤判分離良好的樣品組（例如：正常10%、誤觸25%）
            bool enableCrossCheck = distance < (normal_range + false_range);

            // 由 GitHub Copilot 產生 - 記錄診斷資訊
            System.Diagnostics.Debug.WriteLine($"=== 混合檢測診斷 ===");
            System.Diagnostics.Debug.WriteLine($"正常樣品中位數：{median_normal:F2}%，3σ範圍：{std_normal * 3.0:F2}%，IQR：{iqr_normal:F2}%");
            System.Diagnostics.Debug.WriteLine($"誤觸樣品中位數：{median_false:F2}%，3σ範圍：{std_false * 3.0:F2}%，IQR：{iqr_false:F2}%");
            System.Diagnostics.Debug.WriteLine($"兩組距離：{distance:F2}%");
            System.Diagnostics.Debug.WriteLine($"檢測閾值：{(normal_range + false_range):F2}%");
            System.Diagnostics.Debug.WriteLine($"組間交叉檢測：{(enableCrossCheck ? "✅ 啟用（分布接近）" : "❌ 停用（分離良好）")}");

            // 檢測正常樣品
            foreach (var sample in validNormal)
            {
                // 第一層：組內離群值檢測（始終啟用）
                if (sample.WhitePixelRatio < lower_normal || sample.WhitePixelRatio > upper_normal)
                {
                    sample.Severity = OutlierSeverity.GroupOutlier;
                }
                // 第二層：組間交叉污染檢測（只在分布接近時啟用）
                else if (enableCrossCheck && sample.WhitePixelRatio < boundary)
                {
                    sample.Severity = OutlierSeverity.CrossSample;
                }
                // 正常或輕微異常（1σ-2σ）不額外標記
                else if (sample.IsOutlier1Sigma)
                {
                    sample.Severity = OutlierSeverity.Mild;
                }
                else
                {
                    sample.Severity = OutlierSeverity.Normal;
                }
            }

            // 檢測誤觸樣品
            foreach (var sample in validFalse)
            {
                // 第一層：組內離群值檢測（始終啟用）
                if (sample.WhitePixelRatio < lower_false || sample.WhitePixelRatio > upper_false)
                {
                    sample.Severity = OutlierSeverity.GroupOutlier;
                }
                // 第二層：組間交叉污染檢測（只在分布接近時啟用）
                else if (enableCrossCheck && sample.WhitePixelRatio > boundary)
                {
                    sample.Severity = OutlierSeverity.CrossSample;
                }
                // 正常或輕微異常（1σ-2σ）不額外標記
                else if (sample.IsOutlier1Sigma)
                {
                    sample.Severity = OutlierSeverity.Mild;
                }
                else
                {
                    sample.Severity = OutlierSeverity.Normal;
                }
            }
        }

        // 由 GitHub Copilot 產生
        // 計算百分位數（用於 IQR 計算）
        private float GetPercentile(List<float> sortedValues, int percentile)
        {
            if (sortedValues.Count == 0) return 0;
            
            double index = (percentile / 100.0) * (sortedValues.Count - 1);
            int lowerIndex = (int)Math.Floor(index);
            int upperIndex = (int)Math.Ceiling(index);
            
            if (lowerIndex == upperIndex)
            {
                return sortedValues[lowerIndex];
            }
            
            double fraction = index - lowerIndex;
            return sortedValues[lowerIndex] + (float)fraction * (sortedValues[upperIndex] - sortedValues[lowerIndex]);
        }

        // 由 GitHub Copilot 產生
        // 生成警告訊息文字
        private string GenerateWarningMessage()
        {
            var warnings = new List<string>();

            // 收集正常樣品的問題
            var normalCrossSamples = whitePixelResults.Where(r => r.IsValid && r.Severity == OutlierSeverity.CrossSample).ToList();
            var normalOutliers = whitePixelResults.Where(r => r.IsValid && r.Severity == OutlierSeverity.GroupOutlier).ToList();

            if (normalCrossSamples.Count > 0)
            {
                var indices = string.Join("、", normalCrossSamples.Select((r, i) => $"#{whitePixelResults.IndexOf(r) + 1}"));
                warnings.Add($"正常樣品 {indices}（疑似誤放）");
            }

            if (normalOutliers.Count > 0)
            {
                var indices = string.Join("、", normalOutliers.Select((r, i) => $"#{whitePixelResults.IndexOf(r) + 1}"));
                warnings.Add($"正常樣品 {indices}（組內離群）");
            }

            // 收集誤觸樣品的問題
            var falseCrossSamples = falsePixelResults.Where(r => r.IsValid && r.Severity == OutlierSeverity.CrossSample).ToList();
            var falseOutliers = falsePixelResults.Where(r => r.IsValid && r.Severity == OutlierSeverity.GroupOutlier).ToList();

            if (falseCrossSamples.Count > 0)
            {
                var indices = string.Join("、", falseCrossSamples.Select((r, i) => $"#{falsePixelResults.IndexOf(r) + 1}"));
                warnings.Add($"誤觸樣品 {indices}（疑似誤放）");
            }

            if (falseOutliers.Count > 0)
            {
                var indices = string.Join("、", falseOutliers.Select((r, i) => $"#{falsePixelResults.IndexOf(r) + 1}"));
                warnings.Add($"誤觸樣品 {indices}（組內離群）");
            }

            if (warnings.Count > 0)
            {
                return "⚠️ 發現疑似問題：" + string.Join(" | ", warnings);
            }

            return "";
        }

        // 由 GitHub Copilot 產生
        // 計算推薦參數（white、whiteNULL、whiteTolerance）
        private void CalculateRecommendedParameters()
        {
            var validNormalResults = whitePixelResults.Where(r => r.IsValid).ToList();
            var validFalseResults = falsePixelResults.Where(r => r.IsValid).ToList();

            if (validNormalResults.Count < 3 || validFalseResults.Count < 3)
            {
                MessageBox.Show("樣品數量不足，無法計算推薦參數（建議至少各3張）", "警告",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 計算正常樣品統計
            var normalRatios = validNormalResults.Select(r => r.WhitePixelRatio).ToList();
            float white_normal = GetMedian(normalRatios);
            double std_normal = CalculateStandardDeviation(normalRatios, normalRatios.Average());

            // 計算誤觸樣品統計
            var falseRatios = validFalseResults.Select(r => r.WhitePixelRatio).ToList();
            float white_null = GetMedian(falseRatios);
            double std_null = CalculateStandardDeviation(falseRatios, falseRatios.Average());

            // 由 GitHub Copilot 產生 - 處理誤觸樣品標準差為零的特殊情況
            // 當誤觸樣品數值完全一致時，使用最小標準差 0.3% 作為基準
            const double MIN_STD_DEV = 0.3;
            if (std_null < 0.01) // 標準差接近零（數值完全一致）
            {
                std_null = MIN_STD_DEV;
            }

            // 由 GitHub Copilot 產生 - 混合策略計算 whiteTolerance
            // 結合距離、統計、經驗值三種方法，確保容忍度既能容納偏離樣品又保持安全距離

            // 1. 計算距離
            float distance = Math.Abs(white_normal - white_null);

            // 2. 基於距離的動態容忍度（50% 距離）
            const float DYNAMIC_RATIO = 0.5f;
            float dynamic_tolerance = distance * DYNAMIC_RATIO;

            // 3. 基於統計的容忍度（3σ）
            float statistical_tolerance = (float)(std_normal * 3.0);

            // 4. 基於經驗的絕對最小容忍度（2%）
            const float ABSOLUTE_MIN = 2.0f;

            // 5. 三者取最大值作為基礎容忍度
            float tolerance_base = Math.Max(Math.Max(dynamic_tolerance, statistical_tolerance), ABSOLUTE_MIN);

            // 6. 上限保護：不超過距離的 60%，且絕對不超過 5%
            const float MAX_RATIO = 0.6f;
            const float ABSOLUTE_MAX = 5.0f;
            float max_tolerance_by_distance = distance * MAX_RATIO;
            float max_tolerance = Math.Min(max_tolerance_by_distance, ABSOLUTE_MAX);
            float tolerance_normal = Math.Min(tolerance_base, max_tolerance);

            // 由 GitHub Copilot 產生 - 記錄容忍度計算資訊供顯示
            string toleranceCalculationInfo = $"容忍度計算：\n" +
                                             $"  距離：{distance:F2}%\n" +
                                             $"  動態容忍度 (50%距離)：{dynamic_tolerance:F2}%\n" +
                                             $"  統計容忍度 (3σ)：{statistical_tolerance:F2}%\n" +
                                             $"  絕對最小值：{ABSOLUTE_MIN:F2}%\n" +
                                             $"  基礎容忍度：{tolerance_base:F2}%\n" +
                                             $"  上限 (60%距離，最大5%)：{max_tolerance:F2}%\n" +
                                             $"  最終容忍度：{tolerance_normal:F2}%";

            // 計算正常樣品的判定區間
            float normal_lower = white_normal - tolerance_normal;
            float normal_upper = white_normal + tolerance_normal;

            // 計算誤觸樣品的判定區間
            float null_lower = white_null - (float)(std_null * 3.0);
            float null_upper = white_null + (float)(std_null * 3.0);

            // 由 GitHub Copilot 產生 - 檢查兩個區間是否重疊（不假設誰大誰小）
            bool hasOverlap = !(normal_upper < null_lower || normal_lower > null_upper);

            string warningMessage = "";

            if (hasOverlap)
            {
                // 計算重疊範圍
                float overlap_start = Math.Max(normal_lower, null_lower);
                float overlap_end = Math.Min(normal_upper, null_upper);
                float overlap_range = overlap_end - overlap_start;

                warningMessage = $"⚠️ 警告：正常與誤觸樣品的判定區間重疊 {overlap_range:F2}%\n\n" +
                               $"正常判定區間：[{normal_lower:F2}%, {normal_upper:F2}%]\n" +
                               $"誤觸判定區間：[{null_lower:F2}%, {null_upper:F2}%]\n" +
                               $"重疊區間：[{overlap_start:F2}%, {overlap_end:F2}%]\n\n" +
                               toleranceCalculationInfo + "\n\n" +
                               $"建議調整 whiteThresh 閾值後重新分析以增加區分度。";
            }

            // 設定推薦值
            recommendedWhiteValue = white_normal;
            recommendedWhiteNullValue = white_null;
            recommendedTolerance = tolerance_normal;

            // 由 GitHub Copilot 產生 - 顯示詳細的區分度資訊
            if (!string.IsNullOrEmpty(warningMessage))
            {
                MessageBox.Show(warningMessage, "區分度檢查", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                // 計算安全距離
                float safety_distance = 0;
                if (white_normal > white_null)
                {
                    safety_distance = normal_lower - null_upper;
                }
                else
                {
                    safety_distance = null_lower - normal_upper;
                }

                string successMessage = $"✅ 區分度良好：兩組樣品的判定區間未重疊\n\n" +
                                      $"正常判定區間：[{normal_lower:F2}%, {normal_upper:F2}%]\n" +
                                      $"誤觸判定區間：[{null_lower:F2}%, {null_upper:F2}%]\n" +
                                      $"安全距離：{safety_distance:F2}%\n\n" +
                                      toleranceCalculationInfo;

                MessageBox.Show(successMessage, "區分度檢查", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void DisplayAnalysisResults()
        {
            // 由 GitHub Copilot 產生 - 合併正常樣品和誤觸樣品到列表中
            var allResults = new List<object>();

            // 加入正常樣品
            foreach (var (r, index) in whitePixelResults.Select((r, i) => (r, i)))
            {
                allResults.Add(new
                {
                    樣品類型 = "正常",
                    順序 = index + 1,
                    照片名稱 = r.ImagePath,
                    是否有效 = r.IsValid ? "✅" : "❌",
                    白色像素數 = r.IsValid ? r.WhitePixels.ToString() : "---",
                    總像素數 = r.IsValid ? r.TotalPixels.ToString() : "---",
                    占比百分比 = r.IsValid ? r.WhitePixelRatio.ToString("F2") + "%" : "---",
                    異常狀態 = GetOutlierStatus(r),
                    錯誤訊息 = r.ErrorMessage ?? ""
                });
            }
            
            // 加入誤觸樣品
            foreach (var (r, index) in falsePixelResults.Select((r, i) => (r, i)))
            {
                allResults.Add(new
                {
                    樣品類型 = "誤觸",
                    順序 = index + 1,
                    照片名稱 = r.ImagePath,
                    是否有效 = r.IsValid ? "✅" : "❌",
                    白色像素數 = r.IsValid ? r.WhitePixels.ToString() : "---",
                    總像素數 = r.IsValid ? r.TotalPixels.ToString() : "---",
                    占比百分比 = r.IsValid ? r.WhitePixelRatio.ToString("F2") + "%" : "---",
                    異常狀態 = GetOutlierStatus(r),
                    錯誤訊息 = r.ErrorMessage ?? ""
                });
            }

            dgvResultsControl.DataSource = allResults;

            // 由 GitHub Copilot 產生 - 為疑似問題的列表行著色
            foreach (DataGridViewRow row in dgvResultsControl.Rows)
            {
                var status = row.Cells["異常狀態"].Value?.ToString();

                if (status?.Contains("疑似誤放") == true)
                {
                    row.DefaultCellStyle.BackColor = Color.LightCoral; // 紅底：疑似誤放
                    row.DefaultCellStyle.ForeColor = Color.DarkRed;

                    // 由 GitHub Copilot 產生 - 修復 Font 為 null 的問題
                    // 先檢查當前 Font 是否為 null，如果是則使用 DataGridView 的預設字型
                    Font baseFont = row.DefaultCellStyle.Font ?? dgvResultsControl.DefaultCellStyle.Font ?? this.Font;
                    row.DefaultCellStyle.Font = new Font(baseFont, FontStyle.Bold);
                }
                else if (status?.Contains("組內離群") == true)
                {
                    row.DefaultCellStyle.BackColor = Color.LightYellow; // 黃底：組內離群
                    row.DefaultCellStyle.ForeColor = Color.DarkOrange;
                }
            }

            if (!eventsInitialized)
            {
                dgvResultsControl.SelectionChanged += DgvResults_SelectionChanged;
                eventsInitialized = true;
            }

            var validResults = whitePixelResults.Where(r => r.IsValid).ToList();

            if (validResults.Count > 0)
            {
                CalculateAndDisplayStatistics(validResults);
                currentImageIndex = 0;
                DisplayCurrentImage();
            }
            else
            {
                MessageBox.Show("沒有有效的分析結果，請檢查照片品質", "警告",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // 由 GitHub Copilot 產生 - 顯示誤觸樣品圖像
            if (falsePixelResults.Count > 0)
            {
                currentFalseImageIndex = 0;
                DisplayCurrentFalseImage();
            }

            // 由 GitHub Copilot 產生 - 顯示警告訊息
            string warningMessage = GenerateWarningMessage();
            if (!string.IsNullOrEmpty(warningMessage))
            {
                lblWarningControl.Text = warningMessage;
                lblWarningControl.Visible = true;
            }
            else
            {
                lblWarningControl.Visible = false;
            }
        }

        private string GetOutlierStatus(WhitePixelResult result)
        {
            if (!result.IsValid) return "無效";
            
            // 由 GitHub Copilot 產生 - 優先顯示混合檢測的嚴重程度
            if (result.Severity == OutlierSeverity.CrossSample) return "🔴 疑似誤放";
            if (result.Severity == OutlierSeverity.GroupOutlier) return "🟠 組內離群";
            
            // 原有的 σ 檢測標記（輕微異常不額外警告）
            if (result.IsOutlier2Sigma) return "🔴 2σ異常";
            if (result.IsOutlier1Sigma) return "🟡 1σ異常";
            
            return "✅ 正常";
        }

        // 由 GitHub Copilot 產生
        // 修改為顯示雙組統計數據
        private void CalculateAndDisplayStatistics(List<WhitePixelResult> validResults)
        {
            // 正常樣品統計
            int totalCount = whitePixelResults.Count;
            int validCount = validResults.Count;
            var ratios = validResults.Select(r => r.WhitePixelRatio).ToList();
            double avgRatio = ratios.Average();
            double medianRatio = GetMedian(ratios);
            double stdDev = CalculateStandardDeviation(ratios, avgRatio);
            float minRatio = ratios.Min();
            float maxRatio = ratios.Max();
            int outliers1Sigma = validResults.Count(r => r.IsOutlier1Sigma);
            int outliers2Sigma = validResults.Count(r => r.IsOutlier2Sigma);

            // 由 GitHub Copilot 產生 - 更新正常樣品統計標籤
            var statValues = new string[] {
                totalCount.ToString(),
                validCount.ToString(),
                avgRatio.ToString("F2") + "%",
                medianRatio.ToString("F2") + "%",
                stdDev.ToString("F2") + "%",
                minRatio.ToString("F2") + "%",
                maxRatio.ToString("F2") + "%",
                outliers1Sigma.ToString(),
                outliers2Sigma.ToString(),
                recommendedWhiteValue.ToString("F1") + "%"
            };

            for (int i = 0; i < statValues.Length; i++)
            {
                var lblValue = grpStatisticsControl.Controls.Find($"lblStatValue{i}", false).FirstOrDefault() as Label;
                if (lblValue != null)
                {
                    lblValue.Text = statValues[i];
                }
            }

            // 由 GitHub Copilot 產生 - 誤觸樣品統計（完整顯示所有10項統計資訊）
            var validFalseResults = falsePixelResults.Where(r => r.IsValid).ToList();
            if (validFalseResults.Count > 0)
            {
                var falseRatios = validFalseResults.Select(r => r.WhitePixelRatio).ToList();
                double falseAvg = falseRatios.Average();
                double falseMedian = GetMedian(falseRatios);
                double falseStdDev = CalculateStandardDeviation(falseRatios, falseAvg);
                float falseMin = falseRatios.Min();
                float falseMax = falseRatios.Max();
                int falseOutliers1Sigma = validFalseResults.Count(r => r.IsOutlier1Sigma);
                int falseOutliers2Sigma = validFalseResults.Count(r => r.IsOutlier2Sigma);

                // 由 GitHub Copilot 產生 - 更新誤觸樣品的所有統計標籤（包含平均、中位數、標準差）
                var falseStatValues = new string[] {
                    falsePixelResults.Count.ToString(),
                    validFalseResults.Count.ToString(),
                    falseAvg.ToString("F2") + "%",
                    falseMedian.ToString("F2") + "%",
                    falseStdDev.ToString("F2") + "%",
                    falseMin.ToString("F2") + "%",
                    falseMax.ToString("F2") + "%",
                    falseOutliers1Sigma.ToString(),
                    falseOutliers2Sigma.ToString(),
                    recommendedWhiteNullValue.ToString("F1") + "%"
                };

                for (int i = 0; i < falseStatValues.Length; i++)
                {
                    var lblFalseValue = grpStatisticsControl.Controls.Find($"lblFalseStatValue{i}", false).FirstOrDefault() as Label;
                    if (lblFalseValue != null)
                    {
                        lblFalseValue.Text = falseStatValues[i];
                    }
                }
            }
            else
            {
                // 由 GitHub Copilot 產生 - 若無誤觸樣品資料，將所有標籤設為 "---"
                for (int i = 0; i < 10; i++)
                {
                    var lblFalseValue = grpStatisticsControl.Controls.Find($"lblFalseStatValue{i}", false).FirstOrDefault() as Label;
                    if (lblFalseValue != null)
                    {
                        lblFalseValue.Text = "---";
                    }
                }
            }

            // 由 GitHub Copilot 產生 - 計算並顯示區分度資訊（不假設誰大誰小）
            var validFalse = falsePixelResults.Where(r => r.IsValid).ToList();
            string safetyInfo = "";
            if (validFalse.Count > 0)
            {
                var falseRatios = validFalse.Select(r => r.WhitePixelRatio).ToList();
                double falseStdDev = CalculateStandardDeviation(falseRatios, falseRatios.Average());
                
                // 由 GitHub Copilot 產生 - 處理誤觸標準差為零的情況
                const double MIN_STD_DEV = 0.3;
                if (falseStdDev < 0.01)
                {
                    falseStdDev = MIN_STD_DEV;
                }

                // 計算正常樣品判定區間
                float normal_lower = recommendedWhiteValue - recommendedTolerance;
                float normal_upper = recommendedWhiteValue + recommendedTolerance;

                // 計算誤觸樣品判定區間
                float null_lower = recommendedWhiteNullValue - (float)(falseStdDev * 3.0);
                float null_upper = recommendedWhiteNullValue + (float)(falseStdDev * 3.0);

                // 檢查是否重疊
                bool hasOverlap = !(normal_upper < null_lower || normal_lower > null_upper);

                if (hasOverlap)
                {
                    float overlap_start = Math.Max(normal_lower, null_lower);
                    float overlap_end = Math.Min(normal_upper, null_upper);
                    float overlap_range = overlap_end - overlap_start;
                    
                    safetyInfo = $"\n⚠️ 警告：判定區間重疊 {overlap_range:F2}%！" +
                               $"\n正常區間：[{normal_lower:F2}%, {normal_upper:F2}%]" +
                               $"\n誤觸區間：[{null_lower:F2}%, {null_upper:F2}%]";
                }
                else
                {
                    // 計算安全距離（不假設大小關係）
                    float safety_distance = 0;
                    if (recommendedWhiteValue > recommendedWhiteNullValue)
                    {
                        safety_distance = normal_lower - null_upper;
                    }
                    else
                    {
                        safety_distance = null_lower - normal_upper;
                    }
                    
                    safetyInfo = $"\n✅ 安全距離：{safety_distance:F2}%" +
                               $"\n正常區間：[{normal_lower:F2}%, {normal_upper:F2}%]" +
                               $"\n誤觸區間：[{null_lower:F2}%, {null_upper:F2}%]";
                }
            }

            lblRecommendationControl.Text = $"💡 推薦參數：\n" +
                $"white_{targetStation} = {recommendedWhiteValue:F1}%  |  " +
                $"whiteNULL_{targetStation} = {recommendedWhiteNullValue:F1}%  |  " +
                $"whiteTolerance_{targetStation} = {recommendedTolerance:F1}%  |  " +
                $"whiteThresh_{targetStation} = {trackBarThreshold.Value}" +
                safetyInfo;

            // 重置套用按鈕為未套用狀態
            btnApplyControl.Enabled = true;
            btnApplyControl.Text = "✅ 套用推薦值";
            btnApplyControl.BackColor = System.Drawing.Color.Orange;
        }

        private void UpdateChart()
        {
            var validResults = whitePixelResults.Where(r => r.IsValid).ToList();
            if (validResults.Count == 0) return;

            // 確保圖表已初始化
            if (!chartInitialized)
            {
                InitializeChart();
            }

            // 只清除數據點，保留圖表結構
            chartWhitePixelRatio.Series["白色像素占比"].Points.Clear();
            chartWhitePixelRatio.Series["1σ異常"].Points.Clear();
            chartWhitePixelRatio.Series["2σ異常"].Points.Clear();
            chartWhitePixelRatio.Series["誤觸樣品"].Points.Clear(); // 由 GitHub Copilot 產生
            chartWhitePixelRatio.Series["平均值"].Points.Clear();
            chartWhitePixelRatio.Series["+1σ"].Points.Clear();
            chartWhitePixelRatio.Series["-1σ"].Points.Clear();
            chartWhitePixelRatio.Series["+2σ"].Points.Clear();
            chartWhitePixelRatio.Series["-2σ"].Points.Clear();

            var ratios = validResults.Select(r => r.WhitePixelRatio).ToList();
            double mean = ratios.Average();
            double stdDev = CalculateStandardDeviation(ratios, mean);

            // 暫停圖表更新以提高效能
            chartWhitePixelRatio.SuspendLayout();

            try
            {
                // 取得系列參考，避免重複查詢
                var mainSeries = chartWhitePixelRatio.Series["白色像素占比"];
                var outlier1Series = chartWhitePixelRatio.Series["1σ異常"];
                var outlier2Series = chartWhitePixelRatio.Series["2σ異常"];
                var avgSeries = chartWhitePixelRatio.Series["平均值"];
                var sigma1UpperSeries = chartWhitePixelRatio.Series["+1σ"];
                var sigma1LowerSeries = chartWhitePixelRatio.Series["-1σ"];
                var sigma2UpperSeries = chartWhitePixelRatio.Series["+2σ"];
                var sigma2LowerSeries = chartWhitePixelRatio.Series["-2σ"];

                // 批次新增數據點
                for (int i = 0; i < validResults.Count; i++)
                {
                    var result = validResults[i];
                    int xValue = i + 1;

                    // 主要數據點分類
                    if (result.IsOutlier2Sigma)
                    {
                        outlier2Series.Points.AddXY(xValue, result.WhitePixelRatio);
                    }
                    else if (result.IsOutlier1Sigma)
                    {
                        outlier1Series.Points.AddXY(xValue, result.WhitePixelRatio);
                    }
                    else
                    {
                        mainSeries.Points.AddXY(xValue, result.WhitePixelRatio);
                    }

                    // 統計線（批次新增）
                    avgSeries.Points.AddXY(xValue, mean);
                    sigma1UpperSeries.Points.AddXY(xValue, mean + stdDev);
                    sigma1LowerSeries.Points.AddXY(xValue, mean - stdDev);
                    sigma2UpperSeries.Points.AddXY(xValue, mean + 2 * stdDev);
                    sigma2LowerSeries.Points.AddXY(xValue, mean - 2 * stdDev);
                }

                // 由 GitHub Copilot 產生 - 新增誤觸樣品數據
                var validFalseResults = falsePixelResults.Where(r => r.IsValid).ToList();
                if (validFalseResults.Count > 0)
                {
                    var falseSeries = chartWhitePixelRatio.Series["誤觸樣品"];
                    int offsetX = validResults.Count + 1; // 誤觸樣品從正常樣品後面開始

                    for (int i = 0; i < validFalseResults.Count; i++)
                    {
                        var result = validFalseResults[i];
                        int xValue = offsetX + i;
                        falseSeries.Points.AddXY(xValue, result.WhitePixelRatio);

                        // 為誤觸樣品也繪製統計線
                        avgSeries.Points.AddXY(xValue, mean);
                        sigma1UpperSeries.Points.AddXY(xValue, mean + stdDev);
                        sigma1LowerSeries.Points.AddXY(xValue, mean - stdDev);
                        sigma2UpperSeries.Points.AddXY(xValue, mean + 2 * stdDev);
                        sigma2LowerSeries.Points.AddXY(xValue, mean - 2 * stdDev);
                    }
                }

                // 更新圖表標題
                chartWhitePixelRatio.Titles.Clear();
                string title = $"白色像素占比分析 - {targetType} 站點{targetStation}";
                if (falsePixelResults.Count > 0)
                {
                    title += $" (正常:{validResults.Count} / 誤觸:{falsePixelResults.Count(r => r.IsValid)})";
                }
                chartWhitePixelRatio.Titles.Add(title);
            }
            finally
            {
                // 恢復圖表更新
                chartWhitePixelRatio.ResumeLayout(true);
            }
        }
        #endregion

        #region 顯示方法
        // 由 GitHub Copilot 產生
        // 顯示當前正常樣品圖像（原圖標註 + 二值化圖）
        private void DisplayCurrentImage()
        {
            if (whitePixelResults.Count == 0 || currentImageIndex >= whitePixelResults.Count) return;

            var currentResult = whitePixelResults[currentImageIndex];

            // 顯示原圖標註
            DisplayImageInPictureBox(currentResult, picPreviewControl, true);
            
            // 顯示二值化圖
            DisplayImageInPictureBox(currentResult, picBinaryPreviewControl, false);

            // 更新資訊
            UpdateImageInfo(currentResult);
            UpdateNavigationButtons();
        }

        private void UpdateImageInfo(WhitePixelResult currentResult)
        {
            string info = $"📷 圖片 {currentImageIndex + 1}/{whitePixelResults.Count}: {currentResult.ImagePath}\r\n\r\n";
            info += $"🔍 分析狀態: {(currentResult.IsValid ? "✅ 成功分析" : "❌ 分析失敗")}\r\n\r\n";

            if (currentResult.IsValid)
            {
                info += $"📊 白色像素數: {currentResult.WhitePixels:N0}\r\n";
                info += $"📊 總像素數: {currentResult.TotalPixels:N0}\r\n";
                info += $"📊 占比: {currentResult.WhitePixelRatio:F2}%\r\n\r\n";

                // 異常狀態
                if (currentResult.IsOutlier2Sigma)
                    info += "🔴 異常狀態: 超過2個標準差 (嚴重異常)\r\n";
                else if (currentResult.IsOutlier1Sigma)
                    info += "🟡 異常狀態: 超過1個標準差 (輕微異常)\r\n";
                else
                    info += "✅ 異常狀態: 正常範圍內\r\n";

                info += "\r\n💡 此圖片的白色像素占比：\r\n";
                if (currentResult.WhitePixelRatio < 5.0f)
                    info += "• 偏低 - 可能光線不足或物體較暗";
                else if (currentResult.WhitePixelRatio > 30.0f)
                    info += "• 偏高 - 可能過度曝光或背景過亮";
                else
                    info += "• 適中 - 符合一般預期範圍";
            }
            else
            {
                info += $"❌ 錯誤原因: {currentResult.ErrorMessage ?? "未知錯誤"}";
            }

            lblImageInfoControl.Text = info;
        }

        private void UpdateNavigationButtons()
        {
            btnPreviousControl.Enabled = whitePixelResults.Count > 1;
            btnNextControl.Enabled = whitePixelResults.Count > 1;

            if (dgvResultsControl.Rows.Count > currentImageIndex)
            {
                dgvResultsControl.ClearSelection();
                dgvResultsControl.Rows[currentImageIndex].Selected = true;
            }
        }

        // 由 GitHub Copilot 產生
        // 顯示當前誤觸樣品圖像
        private void DisplayCurrentFalseImage()
        {
            if (falsePixelResults.Count == 0 || currentFalseImageIndex >= falsePixelResults.Count) return;

            var currentResult = falsePixelResults[currentFalseImageIndex];

            // 顯示原圖標註
            DisplayImageInPictureBox(currentResult, picFalsePreviewControl, true);
            
            // 顯示二值化圖
            DisplayImageInPictureBox(currentResult, picFalseBinaryPreviewControl, false);

            // 更新資訊
            UpdateFalseImageInfo(currentResult);
            UpdateFalseNavigationButtons();
        }

        // 顯示圖像到指定的 PictureBox（通用方法）
        private void DisplayImageInPictureBox(WhitePixelResult result, PictureBox pictureBox, bool isAnnotated)
        {
            // 釋放舊圖像
            if (pictureBox.Image != null)
            {
                var oldImage = pictureBox.Image;
                pictureBox.Image = null;
                oldImage.Dispose();
            }

            Mat sourceImage = null;
            System.Drawing.Bitmap bitmap = null;

            try
            {
                // 載入圖像
                sourceImage = isAnnotated ? result.GetResultImage() : result.GetBinaryImage();

                if (sourceImage == null || sourceImage.Empty())
                {
                    pictureBox.Image = null;
                    return;
                }

                // 創建縮圖
                using (var resizedImage = new Mat())
                {
                    double scaleX = 280.0 / sourceImage.Width;
                    double scaleY = 170.0 / sourceImage.Height;
                    double scale = Math.Min(Math.Min(scaleX, scaleY), 1.0);

                    if (scale < 1.0)
                    {
                        var newSize = new OpenCvSharp.Size(
                            (int)(sourceImage.Width * scale),
                            (int)(sourceImage.Height * scale)
                        );
                        Cv2.Resize(sourceImage, resizedImage, newSize, 0, 0, InterpolationFlags.Area);
                    }
                    else
                    {
                        sourceImage.CopyTo(resizedImage);
                    }

                    bitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(resizedImage);
                    pictureBox.Image = bitmap;
                    bitmap = null;
                }
            }
            catch (Exception ex)
            {
                pictureBox.Image = null;
                System.Diagnostics.Debug.WriteLine($"DisplayImageInPictureBox 錯誤: {ex.Message}");
            }
            finally
            {
                sourceImage?.Dispose();
                bitmap?.Dispose();
            }
        }

        private void UpdateFalseImageInfo(WhitePixelResult currentResult)
        {
            string info = $"📷 圖片 {currentFalseImageIndex + 1}/{falsePixelResults.Count}: {currentResult.ImagePath}\r\n\r\n";
            info += $"🔍 分析狀態: {(currentResult.IsValid ? "✅ 成功分析" : "❌ 分析失敗")}\r\n\r\n";

            if (currentResult.IsValid)
            {
                info += $"📊 白色像素數: {currentResult.WhitePixels:N0}\r\n";
                info += $"📊 總像素數: {currentResult.TotalPixels:N0}\r\n";
                info += $"📊 占比: {currentResult.WhitePixelRatio:F2}%\r\n\r\n";

                // 異常狀態
                if (currentResult.IsOutlier2Sigma)
                    info += "🔴 異常狀態: 超過2個標準差 (嚴重異常)\r\n";
                else if (currentResult.IsOutlier1Sigma)
                    info += "🟡 異常狀態: 超過1個標準差 (輕微異常)\r\n";
                else
                    info += "✅ 異常狀態: 正常範圍內\r\n";

                info += "\r\n💡 此圖片的白色像素占比：\r\n";
                if (currentResult.WhitePixelRatio < 5.0f)
                    info += "• 偏低 - 可能光線不足或物體較暗";
                else if (currentResult.WhitePixelRatio > 30.0f)
                    info += "• 偏高 - 可能過度曝光或背景過亮";
                else
                    info += "• 適中 - 符合一般預期範圍";
            }
            else
            {
                info += $"❌ 錯誤原因: {currentResult.ErrorMessage ?? "未知錯誤"}";
            }

            lblFalseImageInfoControl.Text = info;
        }

        private void UpdateFalseNavigationButtons()
        {
            btnFalsePreviousControl.Enabled = falsePixelResults.Count > 1;
            btnFalseNextControl.Enabled = falsePixelResults.Count > 1;
        }
        #endregion

        #region 輔助方法
        private float GetMedian(List<float> values)
        {
            var sorted = values.OrderBy(x => x).ToList();
            int count = sorted.Count;

            if (count % 2 == 0)
            {
                return (sorted[count / 2 - 1] + sorted[count / 2]) / 2.0f;
            }
            else
            {
                return sorted[count / 2];
            }
        }

        private double CalculateStandardDeviation(List<float> values, double mean)
        {
            if (values.Count <= 1) return 0;

            double sumSquaredDiffs = values.Sum(x => Math.Pow(x - mean, 2));
            return Math.Sqrt(sumSquaredDiffs / (values.Count - 1));
        }

        // 由 GitHub Copilot 產生
        // 修改 UpdateParameter 方法，使用 Value 語法避免主鍵問題
        private void UpdateParameter(MydbDB db, string paramName, string value)
        {
            try
            {
                // 根據參數名稱設定中文名稱
                string chineseName = paramName == "white" ? "白色像素占比" : "白色像素閾值";

                // 嘗試更新現有記錄
                int updatedRows = db.@params
                    .Where(p => p.Type == targetType && p.Name == paramName && p.Stop == targetStation)
                    .Set(p => p.Value, value)
                    .Update();

                // 如果沒有更新任何記錄，表示記錄不存在，需要插入
                if (updatedRows == 0)
                {
                    db.@params
                        .Value(p => p.Type, targetType)           // 類型
                        .Value(p => p.Name, paramName)            // 參數名稱
                        .Value(p => p.Value, value)               // 參數值
                        .Value(p => p.Stop, targetStation)        // 站點
                        .Value(p => p.ChineseName, chineseName)   // 中文名稱
                        .Insert();

                    System.Diagnostics.Debug.WriteLine($"新增參數：{targetType} - {paramName} (站點{targetStation}) = {value}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"更新參數：{targetType} - {paramName} (站點{targetStation}) = {value}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateParameter 錯誤: {ex.Message}");
                throw new Exception($"參數更新失敗：{ex.Message}");
            }
        }

        private void ShowFullScreenPreview(WhitePixelResult result)
        {
            Mat resultImage = null;
            System.Drawing.Bitmap bitmap = null;
            Form fullScreenForm = null;
            PictureBox picFullScreen = null;

            try
            {
                resultImage = result.GetResultImage();

                if (resultImage == null || resultImage.Empty())
                {
                    MessageBox.Show("無法載入圖像", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(resultImage);

                fullScreenForm = new Form
                {
                    Text = $"白色像素分析結果 - {result.ImagePath} (按ESC或點擊關閉)",
                    WindowState = FormWindowState.Maximized,
                    KeyPreview = true,
                    BackColor = System.Drawing.Color.Black,
                    FormBorderStyle = FormBorderStyle.None
                };

                picFullScreen = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = System.Drawing.Color.Black,
                    Image = bitmap
                };

                // 事件處理
                fullScreenForm.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Escape)
                        fullScreenForm.Close();
                };

                picFullScreen.Click += (s, e) => fullScreenForm.Close();

                // 表單關閉時清理資源
                fullScreenForm.FormClosed += (s, e) =>
                {
                    try
                    {
                        if (picFullScreen?.Image != null)
                        {
                            var img = picFullScreen.Image;
                            picFullScreen.Image = null;
                            img.Dispose();
                        }
                    }
                    catch { }
                };

                fullScreenForm.Controls.Add(picFullScreen);

                // 顯示表單
                fullScreenForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"預覽錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"ShowFullScreenPreview 錯誤: {ex.Message}");
            }
            finally
            {
                // 確保資源釋放
                resultImage?.Dispose();
                if (bitmap != null && picFullScreen?.Image != bitmap)
                {
                    bitmap.Dispose();
                }
                // fullScreenForm 會在 ShowDialog 結束後自動釋放
            }
        }

        private void ExportResults(string filePath)
        {
            using (var writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("順序,照片名稱,是否有效,白色像素數,總像素數,占比百分比,1σ異常,2σ異常,錯誤訊息");

                for (int i = 0; i < whitePixelResults.Count; i++)
                {
                    var result = whitePixelResults[i];
                    writer.WriteLine($"{i + 1},{result.ImagePath},{(result.IsValid ? "是" : "否")}," +
                                   $"{(result.IsValid ? result.WhitePixels.ToString() : "")}," +
                                   $"{(result.IsValid ? result.TotalPixels.ToString() : "")}," +
                                   $"{(result.IsValid ? result.WhitePixelRatio.ToString("F2") : "")}," +
                                   $"{(result.IsOutlier1Sigma ? "是" : "否")}," +
                                   $"{(result.IsOutlier2Sigma ? "是" : "否")}," +
                                   $"{result.ErrorMessage ?? ""}");
                }

                // 統計摘要
                var validResults = whitePixelResults.Where(r => r.IsValid).ToList();
                if (validResults.Count > 0)
                {
                    var ratios = validResults.Select(r => r.WhitePixelRatio).ToList();
                    writer.WriteLine();
                    writer.WriteLine("統計摘要:");
                    writer.WriteLine($"樣品總數,{whitePixelResults.Count}");
                    writer.WriteLine($"有效樣品,{validResults.Count}");
                    writer.WriteLine($"平均占比,{ratios.Average():F2}%");
                    writer.WriteLine($"中位數占比,{GetMedian(ratios):F2}%");
                    writer.WriteLine($"標準差,{CalculateStandardDeviation(ratios, ratios.Average()):F2}%");
                    writer.WriteLine($"最小值,{ratios.Min():F2}%");
                    writer.WriteLine($"最大值,{ratios.Max():F2}%");
                    writer.WriteLine($"1σ異常數,{validResults.Count(r => r.IsOutlier1Sigma)}");
                    writer.WriteLine($"2σ異常數,{validResults.Count(r => r.IsOutlier2Sigma)}");
                    writer.WriteLine($"推薦值,{recommendedWhiteValue:F1}%");
                }
            }
        }

        // 由 GitHub Copilot 產生
        // 在資源釋放時也要釋放工具提示
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            try
            {
                // 釋放 Timer
                if (_thresholdUpdateTimer != null)
                {
                    _thresholdUpdateTimer.Dispose();
                    _thresholdUpdateTimer = null;
                }

                // 釋放所有結果的資源
                foreach (var result in whitePixelResults)
                {
                    result?.Dispose();
                }
                whitePixelResults.Clear();

                // 釋放預覽圖片
                if (picPreviewControl?.Image != null)
                {
                    picPreviewControl.Image.Dispose();
                    picPreviewControl.Image = null;
                }

                // 清理圖表資源
                if (chartWhitePixelRatio != null)
                {
                    foreach (var series in chartWhitePixelRatio.Series)
                    {
                        series.Points.Clear();
                    }
                    chartWhitePixelRatio.Series.Clear();
                    chartWhitePixelRatio.ChartAreas.Clear();
                    chartWhitePixelRatio.Legends.Clear();
                    chartWhitePixelRatio.Titles.Clear();
                }

                // 釋放工具提示
                toolTip?.Dispose();

                // 強制執行垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"資源釋放錯誤: {ex.Message}");
            }
        }

        // 新增重置分析結果的方法
        private void ResetAnalysisResults()
        {
            // 釋放現有結果
            foreach (var result in whitePixelResults)
            {
                result?.Dispose();
            }
            whitePixelResults.Clear();

            // 釋放預覽圖片
            if (picPreviewControl?.Image != null)
            {
                picPreviewControl.Image.Dispose();
                picPreviewControl.Image = null;
            }

            // 重置圖表
            if (chartInitialized && chartWhitePixelRatio != null)
            {
                foreach (var series in chartWhitePixelRatio.Series)
                {
                    series.Points.Clear();
                }
            }

            currentImageIndex = 0;
            recommendedWhiteValue = 0;
        }
        #endregion
    }
}