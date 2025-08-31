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
            public Mat ResultImage { get; set; }
            public bool IsOutlier1Sigma { get; set; }
            public bool IsOutlier2Sigma { get; set; }
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
            cmbStation.Items.AddRange(new object[] { "站點1", "站點2", "站點3", "站點4" });
            cmbStation.SelectedIndex = 0;
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged;

            // 選擇照片按鈕
            var btnSelectImages = new Button
            {
                Text = "📂 選擇樣品照片",
                Location = new System.Drawing.Point(210, 55),
                Size = new System.Drawing.Size(120, 30)
            };
            btnSelectImages.Click += BtnSelectImages_Click;

            // 開始分析按鈕
            btnAnalyzeControl = new Button
            {
                Text = "🔍 開始分析",
                Location = new System.Drawing.Point(340, 55),
                Size = new System.Drawing.Size(120, 30),
                BackColor = Color.LightGreen,
                Enabled = false
            };
            btnAnalyzeControl.Click += BtnAnalyze_Click;

            // 進度條
            progressBarControl = new ProgressBar
            {
                Location = new System.Drawing.Point(470, 58),
                Size = new System.Drawing.Size(300, 25),
                Visible = false
            };

            // 結果顯示區域
            dgvResultsControl = new DataGridView
            {
                Location = new System.Drawing.Point(10, 95),
                Size = new System.Drawing.Size(580, 350),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            // 統計資訊面板
            grpStatisticsControl = new GroupBox
            {
                Text = "📊 統計分析結果",
                Location = new System.Drawing.Point(600, 95),
                Size = new System.Drawing.Size(580, 200)
            };
            CreateStatisticsLabels(grpStatisticsControl);

            // 折線圖
            chartWhitePixelRatio = new Chart
            {
                Location = new System.Drawing.Point(600, 305),
                Size = new System.Drawing.Size(580, 300),
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

            // 加入所有控件
            this.Controls.AddRange(new Control[] {
                lblDescription, lblStation, cmbStation, btnSelectImages, btnAnalyzeControl, progressBarControl,
                dgvResultsControl, grpStatisticsControl, chartWhitePixelRatio, grpImagePreview,
                lblRecommendationControl, btnApplyControl, btnCancel, btnExport
            });
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

        private void InitializeChart()
        {
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
        }
        #endregion

        #region 事件處理器
        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            targetStation = cmbStation.SelectedIndex + 1;
            this.Text = $"白色像素占比校正 - {targetType} 站點{targetStation}";
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
            if (isAnalyzing)
            {
                MessageBox.Show("正在分析中，請稍等...", "提示");
                return;
            }

            if (targetStation == 0)
            {
                MessageBox.Show("請先選擇站點", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void BtnPrevious_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0) return;
            currentImageIndex = (currentImageIndex - 1 + whitePixelResults.Count) % whitePixelResults.Count;
            DisplayCurrentImage();
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0) return;
            currentImageIndex = (currentImageIndex + 1) % whitePixelResults.Count;
            DisplayCurrentImage();
        }

        private void BtnFullScreen_Click(object sender, EventArgs e)
        {
            if (whitePixelResults.Count == 0 || currentImageIndex >= whitePixelResults.Count) return;

            var currentResult = whitePixelResults[currentImageIndex];
            if (currentResult.ResultImage != null)
            {
                ShowFullScreenPreview(currentResult);
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            try
            {
                using (var db = new MydbDB())
                {
                    UpdateParameter(db, "white", recommendedWhiteValue.ToString("F1"));
                }

                MessageBox.Show($"成功套用 white 參數設定！\n" +
                               $"white (站點{targetStation}) = {recommendedWhiteValue:F1}%",
                               "設定完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
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

        private void DgvResults_SelectionChanged(object sender, EventArgs e)
        {
            if (isAnalyzing) return;

            if (dgvResultsControl.SelectedRows.Count > 0)
            {
                currentImageIndex = dgvResultsControl.SelectedRows[0].Index;
                DisplayCurrentImage();
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
                        this.Invoke(new Action(() =>
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

                this.Invoke(new Action(() =>
                {
                    progressBarControl.Value = selectedImagePaths.Count;
                }));
            });
        }

        private WhitePixelResult ProcessSingleImage(string imagePath)
        {
            using (var image = Cv2.ImRead(imagePath))
            {
                if (image.Empty())
                {
                    throw new Exception("無法載入圖片");
                }

                // 執行白色像素分析（參考 Form1.cs 中的 button46_Click 方法）
                var (whitePixelRatio, totalPixels, whitePixels, resultImage) = AnalyzeWhitePixels(image);

                return new WhitePixelResult
                {
                    ImagePath = Path.GetFileName(imagePath),
                    WhitePixelRatio = whitePixelRatio,
                    TotalPixels = totalPixels,
                    WhitePixels = whitePixels,
                    IsValid = true,
                    ResultImage = resultImage
                };
            }
        }

        private (float ratio, int totalPixels, int whitePixels, Mat resultImage) AnalyzeWhitePixels(Mat image)
        {
            // 參考 Form1.cs 中的 CheckWhitePixelRatio 方法
            using (var gray = new Mat())
            using (var binary = new Mat())
            {
                // 轉為灰階
                Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                // 二值化處理（白色像素閾值設為240）
                Cv2.Threshold(gray, binary, 240, 255, ThresholdTypes.Binary);

                // 計算白色像素
                int totalPixels = image.Width * image.Height;
                int whitePixels = Cv2.CountNonZero(binary);
                float ratio = (float)whitePixels / totalPixels * 100.0f;

                // 創建結果圖像
                Mat resultImage = image.Clone();

                // 在圖像上標註統計信息
                string info = $"White: {whitePixels}/{totalPixels} ({ratio:F2}%)";
                Cv2.PutText(resultImage, info, new OpenCvSharp.Point(10, 30),
                           HersheyFonts.HersheySimplex, 1.0, new Scalar(0, 255, 0), 2);

                // 可視化白色區域（可選）
                using (var colorBinary = new Mat())
                {
                    Cv2.CvtColor(binary, colorBinary, ColorConversionCodes.GRAY2BGR);
                    Cv2.AddWeighted(resultImage, 0.7, colorBinary, 0.3, 0, resultImage);
                }

                return (ratio, totalPixels, whitePixels, resultImage);
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
            btnApplyControl.Enabled = true;
        }

        private void UpdateChart()
        {
            var validResults = whitePixelResults.Where(r => r.IsValid).ToList();
            if (validResults.Count == 0) return;

            // 【修正】：完全清除圖表的所有元素
            chartWhitePixelRatio.Series.Clear();
            chartWhitePixelRatio.ChartAreas.Clear();
            chartWhitePixelRatio.Legends.Clear();
            chartWhitePixelRatio.Titles.Clear();

            InitializeChart();

            var ratios = validResults.Select(r => r.WhitePixelRatio).ToList();
            double mean = ratios.Average();
            double stdDev = CalculateStandardDeviation(ratios, mean);

            // 主要數據點
            var mainSeries = chartWhitePixelRatio.Series["白色像素占比"];
            var outlier1Series = chartWhitePixelRatio.Series["1σ異常"];
            var outlier2Series = chartWhitePixelRatio.Series["2σ異常"];

            for (int i = 0; i < validResults.Count; i++)
            {
                var result = validResults[i];

                if (result.IsOutlier2Sigma)
                {
                    outlier2Series.Points.AddXY(i + 1, result.WhitePixelRatio);
                }
                else if (result.IsOutlier1Sigma)
                {
                    outlier1Series.Points.AddXY(i + 1, result.WhitePixelRatio);
                }
                else
                {
                    mainSeries.Points.AddXY(i + 1, result.WhitePixelRatio);
                }
            }

            // 統計線
            var avgSeries = chartWhitePixelRatio.Series["平均值"];
            var sigma1UpperSeries = chartWhitePixelRatio.Series["+1σ"];
            var sigma1LowerSeries = chartWhitePixelRatio.Series["-1σ"];
            var sigma2UpperSeries = chartWhitePixelRatio.Series["+2σ"];
            var sigma2LowerSeries = chartWhitePixelRatio.Series["-2σ"];

            for (int i = 1; i <= validResults.Count; i++)
            {
                avgSeries.Points.AddXY(i, mean);
                sigma1UpperSeries.Points.AddXY(i, mean + stdDev);
                sigma1LowerSeries.Points.AddXY(i, mean - stdDev);
                sigma2UpperSeries.Points.AddXY(i, mean + 2 * stdDev);
                sigma2LowerSeries.Points.AddXY(i, mean - 2 * stdDev);
            }

            // 設定圖表標題
            chartWhitePixelRatio.Titles.Clear();
            chartWhitePixelRatio.Titles.Add($"白色像素占比分析 - {targetType} 站點{targetStation}");
        }
        #endregion

        #region 顯示方法
        private void DisplayCurrentImage()
        {
            if (whitePixelResults.Count == 0 || currentImageIndex >= whitePixelResults.Count) return;

            var currentResult = whitePixelResults[currentImageIndex];

            if (currentResult.ResultImage != null)
            {
                try
                {
                    var bitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(currentResult.ResultImage);

                    if (picPreviewControl.Image != null)
                    {
                        picPreviewControl.Image.Dispose();
                    }
                    picPreviewControl.Image = bitmap;

                    UpdateImageInfo(currentResult);
                }
                catch (Exception ex)
                {
                    lblImageInfoControl.Text = $"圖像顯示錯誤：{ex.Message}";
                }
            }
            else
            {
                UpdateImageInfo(currentResult);
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

        private void UpdateParameter(MydbDB db, string paramName, string value)
        {
            var existingParam = db.@params.FirstOrDefault(p =>
                p.Type == targetType && p.Name == paramName && p.Stop == targetStation);

            if (existingParam != null)
            {
                existingParam.Value = value;
                db.Update(existingParam);
            }
            else
            {
                db.Insert(new Param
                {
                    Type = targetType,
                    Name = paramName,
                    Value = value,
                    Stop = targetStation,
                    ChineseName = "白色像素占比"
                });
            }
        }

        private void ShowFullScreenPreview(WhitePixelResult result)
        {
            using (var fullScreenForm = new Form())
            {
                fullScreenForm.Text = $"白色像素分析結果 - {result.ImagePath}";
                fullScreenForm.WindowState = FormWindowState.Maximized;
                fullScreenForm.KeyPreview = true;
                fullScreenForm.BackColor = Color.Black;

                var picFullScreen = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = Color.Black
                };

                var bitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(result.ResultImage);
                picFullScreen.Image = bitmap;

                fullScreenForm.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Escape)
                        fullScreenForm.Close();
                };

                picFullScreen.Click += (s, e) => fullScreenForm.Close();

                fullScreenForm.Controls.Add(picFullScreen);
                fullScreenForm.ShowDialog();
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

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            foreach (var result in whitePixelResults)
            {
                result.ResultImage?.Dispose();
            }

            picPreviewControl?.Image?.Dispose();
        }
        #endregion
    }
}