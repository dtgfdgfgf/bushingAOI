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
        private string targetType;
        private int targetStation;
        private float recommendedWhiteValue;
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
        
        // 二值化閾值控件
        private TrackBar trackBarThreshold;
        private Label lblThresholdValue;
        private System.Threading.Timer _thresholdUpdateTimer;
        #endregion

        #region 白色像素分析結果結構
        public class WhitePixelResult
        {
            public string ImagePath { get; set; }
            public float WhitePixelRatio { get; set; }
            public int TotalPixels { get; set; }
            public int WhitePixels { get; set; }
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; }
            //public Mat ResultImage { get; set; }
            private byte[] _compressedImageData; // 儲存壓縮後的圖像資料
            public bool IsOutlier1Sigma { get; set; }
            public bool IsOutlier2Sigma { get; set; }
            // 設定壓縮後的圖像資料
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
            // 按需取得結果圖像
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

            // 釋放資源
            public void Dispose()
            {
                _compressedImageData = null;
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
            this.Size = new System.Drawing.Size(1200, 900);
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

            // 選擇照片按鈕
            var btnSelectImages = new Button
            {
                Text = "📂 選擇樣品照片",
                Location = new System.Drawing.Point(210, 55),
                Size = new System.Drawing.Size(120, 30)
            };
            btnSelectImages.Click += BtnSelectImages_Click;

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
                Size = new System.Drawing.Size(250, 45),
                Location = new System.Drawing.Point(310, 90)
            };
            trackBarThreshold.ValueChanged += TrackBarThreshold_ValueChanged;

            // 顯示當前閾值的標籤
            lblThresholdValue = new Label
            {
                Text = "180",
                Location = new System.Drawing.Point(570, 95),
                Size = new System.Drawing.Size(50, 25),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            toolTip.SetToolTip(trackBarThreshold, "調整二值化閾值 (0-255)，影響白色像素的判定標準");

            // 開始分析按鈕
            btnAnalyzeControl = new Button
            {
                Text = "🔍 開始分析",
                Location = new System.Drawing.Point(630, 90),
                Size = new System.Drawing.Size(120, 30),
                BackColor = Color.LightGreen,
                Enabled = false
            };
            btnAnalyzeControl.Click += BtnAnalyze_Click;

            // 進度條
            progressBarControl = new ProgressBar
            {
                Location = new System.Drawing.Point(760, 93),
                Size = new System.Drawing.Size(300, 25),
                Visible = false
            };

            // 結果顯示區域
            dgvResultsControl = new DataGridView
            {
                Location = new System.Drawing.Point(10, 140),
                Size = new System.Drawing.Size(580, 305),
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
                Size = new System.Drawing.Size(580, 200)
            };
            CreateStatisticsLabels(grpStatisticsControl);

            // 折線圖
            chartWhitePixelRatio = new Chart
            {
                Location = new System.Drawing.Point(600, 350),
                Size = new System.Drawing.Size(580, 255),
                BackColor = Color.White
            };
            InitializeChart();

            // 圖片預覽區域
            var grpImagePreview = new GroupBox
            {
                Text = "📷 分析結果預覽",
                Location = new System.Drawing.Point(10, 455),
                Size = new System.Drawing.Size(580, 200)
            };

            // 預覽圖片
            picPreviewControl = new PictureBox
            {
                Location = new System.Drawing.Point(10, 20),
                Size = new System.Drawing.Size(280, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };

            // 導航按鈕
            btnPreviousControl = new Button
            {
                Text = "◀ 上一張",
                Location = new System.Drawing.Point(300, 30),
                Size = new System.Drawing.Size(80, 30),
                Enabled = false
            };
            btnPreviousControl.Click += BtnPrevious_Click;

            btnNextControl = new Button
            {
                Text = "下一張 ▶",
                Location = new System.Drawing.Point(300, 70),
                Size = new System.Drawing.Size(80, 30),
                Enabled = false
            };
            btnNextControl.Click += BtnNext_Click;

            // 全螢幕預覽按鈕
            var btnFullScreen = new Button
            {
                Text = "🔍 全螢幕預覽",
                Location = new System.Drawing.Point(300, 110),
                Size = new System.Drawing.Size(80, 30),
                BackColor = Color.LightBlue
            };
            btnFullScreen.Click += BtnFullScreen_Click;

            // 圖片資訊標籤
            lblImageInfoControl = new TextBox
            {
                Text = "尚未分析圖片",
                Location = new System.Drawing.Point(390, 20),
                Size = new System.Drawing.Size(180, 170),
                Font = new Font("Microsoft JhengHei", 9),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                WordWrap = true,
                TabStop = false
            };

            grpImagePreview.Controls.AddRange(new Control[] {
                picPreviewControl, btnPreviousControl, btnNextControl, btnFullScreen, lblImageInfoControl
            });

            // 建議值顯示
            lblRecommendationControl = new Label
            {
                Text = "💡 推薦的 white 參數值將在分析完成後顯示",
                Location = new System.Drawing.Point(10, 665),
                Size = new System.Drawing.Size(500, 30),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.DarkGreen
            };

            // 套用按鈕
            btnApplyControl = new Button
            {
                Text = "✅ 套用推薦值",
                Location = new System.Drawing.Point(520, 665),
                Size = new System.Drawing.Size(120, 30),
                BackColor = Color.Orange,
                Enabled = false
            };
            btnApplyControl.Click += BtnApply_Click;

            // 取消按鈕
            var btnCancel = new Button
            {
                Text = "❌ 取消",
                Location = new System.Drawing.Point(650, 665),
                Size = new System.Drawing.Size(80, 30)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            // 匯出數據按鈕
            var btnExport = new Button
            {
                Text = "📊 匯出數據",
                Location = new System.Drawing.Point(740, 665),
                Size = new System.Drawing.Size(100, 30),
                BackColor = Color.LightCyan
            };
            btnExport.Click += BtnExport_Click;

            // 清除結果按鈕（加在匯出數據按鈕旁邊）
            var btnClearResults = new Button
            {
                Text = "🗑️ 清除結果",
                Location = new System.Drawing.Point(850, 665),
                Size = new System.Drawing.Size(100, 30),
                BackColor = System.Drawing.Color.LightPink,
                //ToolTipText = "清除目前的分析結果，可重新選擇照片分析"
            };
            btnClearResults.Click += BtnClearResults_Click;
            toolTip.SetToolTip(btnClearResults, "清除目前的分析結果，可重新選擇照片分析");

            // 加入所有控件
            this.Controls.AddRange(new Control[] {
                lblDescription, lblStation, cmbStation, btnSelectImages, 
                lblThreshold, trackBarThreshold, lblThresholdValue, btnAnalyzeControl, progressBarControl,
                dgvResultsControl, grpStatisticsControl, chartWhitePixelRatio, grpImagePreview,
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
            var labels = new[] {
                "樣品數量：", "有效樣品：", "平均占比：", "中位數占比：",
                "標準差：", "最小值：", "最大值：", "1σ異常數：", "2σ異常數：", "推薦值："
            };

            for (int i = 0; i < labels.Length; i++)
            {
                var lbl = new Label
                {
                    Text = labels[i],
                    Location = new System.Drawing.Point(10, 25 + (i % 5) * 30),
                    Size = new System.Drawing.Size(100, 25),
                    Name = $"lblStat{i}"
                };

                var lblValue = new Label
                {
                    Text = "---",
                    Location = new System.Drawing.Point(120, 25 + (i % 5) * 30),
                    Size = new System.Drawing.Size(120, 25),
                    Name = $"lblStatValue{i}",
                    Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
                };

                // 第二列
                if (i >= 5)
                {
                    lbl.Location = new System.Drawing.Point(270, 25 + (i - 5) * 30);
                    lblValue.Location = new System.Drawing.Point(380, 25 + (i - 5) * 30);
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

                    btnAnalyzeControl.Text = $"🔍 分析 {selectedImagePaths.Count} 張照片";
                    btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
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

            if (selectedImagePaths.Count == 0)
            {
                MessageBox.Show("請先選擇照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            isAnalyzing = true;
            btnAnalyzeControl.Enabled = false;
            btnAnalyzeControl.Text = "🔄 分析中...";

            try
            {
                progressBarControl.Maximum = selectedImagePaths.Count;
                progressBarControl.Value = 0;
                progressBarControl.Visible = true;

                whitePixelResults.Clear();
                await ProcessImages();
                AnalyzeOutliers();
                DisplayAnalysisResults();
                UpdateChart();

                MessageBox.Show($"白色像素占比分析完成！\n" +
                               $"總計 {whitePixelResults.Count} 張照片\n" +
                               $"成功分析 {whitePixelResults.Count(r => r.IsValid)} 張\n" +
                               $"1σ異常 {whitePixelResults.Count(r => r.IsOutlier1Sigma)} 張\n" +
                               $"2σ異常 {whitePixelResults.Count(r => r.IsOutlier2Sigma)} 張\n" +
                               $"推薦值：{recommendedWhiteValue:F1}%",
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
                btnAnalyzeControl.Text = $"🔍 分析 {selectedImagePaths.Count} 張照片";
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
            try
            {
                using (var db = new MydbDB())
                {
                    // 套用 white 參數
                    UpdateParameter(db, "white", recommendedWhiteValue.ToString("F1"));
                    
                    // 套用 whiteThresh 參數
                    UpdateParameter(db, "whiteThresh", trackBarThreshold.Value.ToString());
                }

                MessageBox.Show($"成功套用參數設定！\n" +
                               $"white (站點{targetStation}) = {recommendedWhiteValue:F1}%\n" +
                               $"whiteThresh (站點{targetStation}) = {trackBarThreshold.Value}\n\n" +
                               $"您可以繼續選擇其他照片進行校正，或關閉視窗。",
                               "設定完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 移除自動關閉視窗的程式碼
                // this.DialogResult = DialogResult.OK;

                // 可選：重新啟用分析按鈕讓使用者可以重新分析
                btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;

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
        // 修正 DataGridView 選擇變更事件
        private void DgvResults_SelectionChanged(object sender, EventArgs e)
        {
            if (isAnalyzing || _isNavigating) return;
            if (dgvResultsControl.SelectedRows.Count == 0) return;

            try
            {
                int newIndex = dgvResultsControl.SelectedRows[0].Index;

                // 避免重複更新同一張圖片
                if (newIndex == currentImageIndex) return;

                currentImageIndex = newIndex;

                // 使用異步方式更新圖像，避免阻塞UI
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
                        System.Diagnostics.Debug.WriteLine($"選擇變更錯誤: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DgvResults_SelectionChanged 錯誤: {ex.Message}");
            }
        }
        #endregion

        #region 圖像處理方法
        private async Task ProcessImages()
        {
            await Task.Run(() =>
            {
                for (int i = 0; i < selectedImagePaths.Count; i++)
                {
                    string imagePath = selectedImagePaths[i];

                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            progressBarControl.Value = i;
                            btnAnalyzeControl.Text = $"🔄 處理 {i + 1}/{selectedImagePaths.Count}";
                        }));

                        var result = ProcessSingleImage(imagePath);
                        whitePixelResults.Add(result);

                        System.Threading.Thread.Sleep(50);
                    }
                    catch (Exception ex)
                    {
                        whitePixelResults.Add(new WhitePixelResult
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
                {
                    var (whitePixelRatio, totalPixels, whitePixels) = AnalyzeWhitePixels(image, resultImage, currentThreshold);

                    var result = new WhitePixelResult
                    {
                        ImagePath = System.IO.Path.GetFileName(imagePath),
                        WhitePixelRatio = whitePixelRatio,
                        TotalPixels = totalPixels,
                        WhitePixels = whitePixels,
                        IsValid = true
                    };

                    // 設定壓縮後的圖像資料
                    result.SetResultImageData(resultImage);

                    return result;
                }
            }
        } //演算法

        private (float ratio, int totalPixels, int whitePixels) AnalyzeWhitePixels(Mat image, Mat resultImage, int minthresh)
        {
            using (var gray = new Mat())
            using (var binary = new Mat())
            {
                // 轉為灰階
                Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                // 二值化處理
                Cv2.Threshold(gray, binary, minthresh, 255, ThresholdTypes.Binary);

                // 計算白色像素
                int totalPixels = image.Width * image.Height;
                int whitePixels = Cv2.CountNonZero(binary);
                float ratio = (float)whitePixels / totalPixels * 100.0f;

                // 創建結果圖像
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
        }

        private void DisplayAnalysisResults()
        {
            var resultsData = whitePixelResults.Select((r, index) => new
            {
                順序 = index + 1,
                照片名稱 = r.ImagePath,
                是否有效 = r.IsValid ? "✅" : "❌",
                白色像素數 = r.IsValid ? r.WhitePixels.ToString() : "---",
                總像素數 = r.IsValid ? r.TotalPixels.ToString() : "---",
                占比百分比 = r.IsValid ? r.WhitePixelRatio.ToString("F2") + "%" : "---",
                異常狀態 = GetOutlierStatus(r),
                錯誤訊息 = r.ErrorMessage ?? ""
            }).ToList();

            dgvResultsControl.DataSource = resultsData;

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
        }

        private string GetOutlierStatus(WhitePixelResult result)
        {
            if (!result.IsValid) return "無效";
            if (result.IsOutlier2Sigma) return "🔴 2σ異常";
            if (result.IsOutlier1Sigma) return "🟡 1σ異常";
            return "✅ 正常";
        }

        // 由 GitHub Copilot 產生
        // 修正分析完成後的按鈕狀態設定
        private void CalculateAndDisplayStatistics(List<WhitePixelResult> validResults)
        {
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

            lblRecommendationControl.Text = $"💡 推薦設定：white (站點{targetStation}) = {recommendedWhiteValue:F1}%";

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

                // 更新圖表標題
                chartWhitePixelRatio.Titles.Clear();
                chartWhitePixelRatio.Titles.Add($"白色像素占比分析 - {targetType} 站點{targetStation}");
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
        // 完全重寫圖像顯示方法，解決卡死問題
        private void DisplayCurrentImage()
        {
            if (whitePixelResults.Count == 0 || currentImageIndex >= whitePixelResults.Count) return;

            var currentResult = whitePixelResults[currentImageIndex];

            // 強制釋放舊的圖像資源
            if (picPreviewControl.Image != null)
            {
                var oldImage = picPreviewControl.Image;
                picPreviewControl.Image = null;
                oldImage.Dispose();

                // 強制垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            Mat resultImage = null;
            System.Drawing.Bitmap bitmap = null;

            try
            {
                // 檢查是否有圖像資料
                if (!currentResult.HasImageData())
                {
                    picPreviewControl.Image = null;
                    lblImageInfoControl.Text = "無可用的圖像資料";
                    return;
                }

                // 載入結果圖像
                resultImage = currentResult.GetResultImage();

                if (resultImage == null || resultImage.Empty())
                {
                    picPreviewControl.Image = null;
                    lblImageInfoControl.Text = "圖像載入失敗";
                    return;
                }

                // 創建適當大小的縮圖以節省記憶體
                using (var resizedImage = new Mat())
                {
                    // 計算適合的大小（最大280x170，保持比例）
                    double scaleX = 280.0 / resultImage.Width;
                    double scaleY = 170.0 / resultImage.Height;
                    double scale = Math.Min(Math.Min(scaleX, scaleY), 1.0); // 不要放大

                    if (scale < 1.0)
                    {
                        var newSize = new OpenCvSharp.Size(
                            (int)(resultImage.Width * scale),
                            (int)(resultImage.Height * scale)
                        );
                        Cv2.Resize(resultImage, resizedImage, newSize, 0, 0, InterpolationFlags.Area);
                    }
                    else
                    {
                        resultImage.CopyTo(resizedImage);
                    }

                    // 轉換為 Bitmap
                    bitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(resizedImage);
                    picPreviewControl.Image = bitmap;
                    bitmap = null; // 避免在 finally 中重複釋放
                }

                UpdateImageInfo(currentResult);
            }
            catch (Exception ex)
            {
                lblImageInfoControl.Text = $"圖像顯示錯誤：{ex.Message}";
                picPreviewControl.Image = null;
                System.Diagnostics.Debug.WriteLine($"DisplayCurrentImage 錯誤: {ex.Message}");
            }
            finally
            {
                // 確保資源釋放
                resultImage?.Dispose();
                bitmap?.Dispose();
            }

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