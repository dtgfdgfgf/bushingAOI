using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OpenCvSharp;
using LinqToDB;

namespace peilin
{
    public partial class GapThreshCalibrationForm : Form
    {
        

        #region 私有成員變數
        private List<string> selectedImagePaths = new List<string>();
        private List<GapAnalysisResult> gapAnalysisResults = new List<GapAnalysisResult>();
        private string targetType;
        private int targetStation;
        private int currentGapThresh;
        private int currentImageIndex = 0;
        private bool isAnalyzing = false;

        // 控件成員變數
        private ComboBox cmbStation;
        private PictureBox picOriginalControl;
        private PictureBox picResultControl;
        private TextBox lblImageInfoControl;
        private Button btnPreviousControl;
        private Button btnNextControl;
        private TrackBar trackBarGapThresh;
        private Label lblCurrentValue;
        private Button btnAnalyzeControl;
        private Button btnApplyControl;
        private Label lblRecommendationControl;
        #endregion

        #region Gap分析結果結構
        public class GapAnalysisResult
        {
            public string ImagePath { get; set; }
            public bool IsNG { get; set; }
            // 由 GitHub Copilot 產生 - 改為返回開口大小數據
            public double OuterGapMm { get; set; }          // 外擴開口大小 (mm)
            public double OuterGapAngleDeg { get; set; }    // 外擴開口角度 (度)
            public double InwardBendMm { get; set; }        // 內彎弧長 (mm)
            public double InwardBendAngleDeg { get; set; }  // 內彎角度 (度)
            public Mat OriginalImage { get; set; }
            public Mat ResultImage { get; set; }
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; }
            public int UsedGapThresh { get; set; }
        }
        #endregion

        #region 建構函數和初始化
        public GapThreshCalibrationForm(string type)
        {
            targetType = type;
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = $"gapThresh 參數視覺化調校 - {targetType}";
            this.Size = new System.Drawing.Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            CreateControls();
            LoadCurrentGapThreshFromDatabase();
        }

        private void LoadCurrentGapThreshFromDatabase()
        {
            if (targetStation <= 0) return;
            try
            {
                using (var db = new MydbDB())
                {
                    var param = db.@params.FirstOrDefault(p =>
                        p.Type == targetType && p.Name == "gapThresh" && p.Stop == targetStation);

                    if (param != null && int.TryParse(param.Value, out int value))
                    {
                        currentGapThresh = Math.Max(0, Math.Min(254, value));
                    }
                    else
                    {
                        currentGapThresh = 127; // 預設值
                    }

                    trackBarGapThresh.Value = currentGapThresh;
                    lblCurrentValue.Text = currentGapThresh.ToString();
                }
            }
            catch
            {
                currentGapThresh = 127; // 預設值
                trackBarGapThresh.Value = currentGapThresh;
                lblCurrentValue.Text = currentGapThresh.ToString();
            }
        }
        #endregion

        #region 控件建立和布局
        private void CreateControls()
        {
            // 說明標籤
            var lblDescription = new Label
            {
                Text = "📋 請選擇測試圖片，調整 gapThresh 參數並觀察 Gap 檢測效果",
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(1160, 30),
                Font = new Font("Microsoft JhengHei", 10),
                ForeColor = Color.Blue
            };

            // 站點選擇
            var lblStation = new Label
            {
                Text = "選擇站點：",
                Location = new System.Drawing.Point(10, 50),
                Size = new System.Drawing.Size(80, 25),
                Font = new Font("Microsoft JhengHei", 9, FontStyle.Bold)
            };

            cmbStation = new ComboBox
            {
                Location = new System.Drawing.Point(90, 50),
                Size = new System.Drawing.Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbStation.Items.AddRange(new object[] { "請選擇", "站點1", "站點2", "站點3", "站點4" });
            cmbStation.SelectedIndex = 0;
            targetStation = 0;
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged;

            // 選擇照片按鈕
            var btnSelectImages = new Button
            {
                Text = "📂 選擇測試圖片",
                Location = new System.Drawing.Point(210, 50),
                Size = new System.Drawing.Size(120, 30)
            };
            btnSelectImages.Click += BtnSelectImages_Click;

            // gapThresh 調整區域
            var grpGapThreshControl = new GroupBox
            {
                Text = "🎛️ gapThresh 參數調整",
                Location = new System.Drawing.Point(10, 90),
                Size = new System.Drawing.Size(1160, 80)
            };

            var lblGapThresh = new Label
            {
                Text = "gapThresh 值：",
                Location = new System.Drawing.Point(20, 25),
                Size = new System.Drawing.Size(100, 25)
            };

            trackBarGapThresh = new TrackBar
            {
                Location = new System.Drawing.Point(120, 20),
                Size = new System.Drawing.Size(800, 45),
                Minimum = 0,
                Maximum = 254,
                Value = 127,
                TickFrequency = 25,
                LargeChange = 10,
                SmallChange = 1
            };
            trackBarGapThresh.ValueChanged += TrackBarGapThresh_ValueChanged;

            lblCurrentValue = new Label
            {
                Text = "127",
                Location = new System.Drawing.Point(930, 25),
                Size = new System.Drawing.Size(50, 25),
                Font = new Font("Microsoft JhengHei", 12, FontStyle.Bold),
                ForeColor = Color.Red
            };

            btnAnalyzeControl = new Button
            {
                Text = "🔄 重新分析",
                Location = new System.Drawing.Point(1000, 20),
                Size = new System.Drawing.Size(100, 35),
                BackColor = Color.LightGreen,
                Enabled = false
            };
            btnAnalyzeControl.Click += BtnAnalyze_Click;

            grpGapThreshControl.Controls.AddRange(new Control[] {
                lblGapThresh, trackBarGapThresh, lblCurrentValue, btnAnalyzeControl
            });

            // 圖片顯示區域
            var grpImageDisplay = new GroupBox
            {
                Text = "📷 圖片預覽與分析結果",
                Location = new System.Drawing.Point(10, 180),
                Size = new System.Drawing.Size(1160, 450)
            };

            // 原始圖片
            var lblOriginal = new Label
            {
                Text = "原始圖片",
                Location = new System.Drawing.Point(20, 25),
                Size = new System.Drawing.Size(280, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold)
            };

            picOriginalControl = new PictureBox
            {
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(280, 350),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };

            // Gap檢測結果
            var lblResult = new Label
            {
                Text = "Gap檢測結果",
                Location = new System.Drawing.Point(320, 25),
                Size = new System.Drawing.Size(280, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold)
            };

            picResultControl = new PictureBox
            {
                Location = new System.Drawing.Point(320, 50),
                Size = new System.Drawing.Size(280, 350),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };

            // 導航按鈕
            btnPreviousControl = new Button
            {
                Text = "◀ 上一張",
                Location = new System.Drawing.Point(620, 50),
                Size = new System.Drawing.Size(80, 35),
                Enabled = false
            };
            btnPreviousControl.Click += BtnPrevious_Click;

            btnNextControl = new Button
            {
                Text = "下一張 ▶",
                Location = new System.Drawing.Point(620, 95),
                Size = new System.Drawing.Size(80, 35),
                Enabled = false
            };
            btnNextControl.Click += BtnNext_Click;

            // 全螢幕預覽按鈕
            var btnFullScreen = new Button
            {
                Text = "🔍 全螢幕預覽",
                Location = new System.Drawing.Point(620, 140),
                Size = new System.Drawing.Size(80, 35),
                BackColor = Color.LightBlue
            };
            btnFullScreen.Click += BtnFullScreen_Click;

            // 資訊顯示區域
            lblImageInfoControl = new TextBox
            {
                Text = "請選擇圖片並調整 gapThresh 參數",
                Location = new System.Drawing.Point(720, 50),
                Size = new System.Drawing.Size(420, 350),
                Font = new Font("Microsoft JhengHei", 9),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                WordWrap = true,
                TabStop = false
            };

            grpImageDisplay.Controls.AddRange(new Control[] {
                lblOriginal, picOriginalControl, lblResult, picResultControl,
                btnPreviousControl, btnNextControl, btnFullScreen, lblImageInfoControl
            });

            // 底部控制區域
            lblRecommendationControl = new Label
            {
                Text = "💡 調整 gapThresh 參數觀察檢測效果，找到最適合的數值",
                Location = new System.Drawing.Point(10, 640),
                Size = new System.Drawing.Size(600, 30),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.DarkGreen
            };

            // 套用按鈕
            btnApplyControl = new Button
            {
                Text = "✅ 套用當前參數值",
                Location = new System.Drawing.Point(620, 640),
                Size = new System.Drawing.Size(140, 30),
                BackColor = Color.Orange,
                Enabled = false
            };
            btnApplyControl.Click += BtnApply_Click;

            // 取消按鈕
            var btnCancel = new Button
            {
                Text = "❌ 取消",
                Location = new System.Drawing.Point(770, 640),
                Size = new System.Drawing.Size(80, 30)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            // 加入所有控件
            this.Controls.AddRange(new Control[] {
                lblDescription, lblStation, cmbStation, btnSelectImages,
                grpGapThreshControl, grpImageDisplay,
                lblRecommendationControl, btnApplyControl, btnCancel
            });
        }
        #endregion

        #region 事件處理器
        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            targetStation = cmbStation.SelectedIndex;
            if (targetStation <= 0)
            {
                this.Text = $"gapThresh 參數視覺化調校 - {targetType} (請選擇站點)";
                btnAnalyzeControl.Enabled = false;
                btnApplyControl.Enabled = false;
                return;
            }
            this.Text = $"gapThresh 參數視覺化調校 - {targetType} 站點{targetStation}";
            LoadCurrentGapThreshFromDatabase();
            btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
            btnApplyControl.Enabled = selectedImagePaths.Count > 0;
        }

        private void BtnSelectImages_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "選擇測試圖片";
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedImagePaths = openFileDialog.FileNames.ToList();

                    if (selectedImagePaths.Count > 20)
                    {
                        MessageBox.Show("選擇的圖片過多，建議選擇5-10張代表性圖片進行調校", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
                    btnApplyControl.Enabled = selectedImagePaths.Count > 0;

                    // 自動進行首次分析
                    if (selectedImagePaths.Count > 0)
                    {
                        PerformAnalysis();
                    }
                }
            }
        }

        private void TrackBarGapThresh_ValueChanged(object sender, EventArgs e)
        {
            currentGapThresh = trackBarGapThresh.Value;
            lblCurrentValue.Text = currentGapThresh.ToString();

            lblRecommendationControl.Text = $"💡 當前 gapThresh = {currentGapThresh}，點擊「重新分析」查看效果";
        }

        private void BtnAnalyze_Click(object sender, EventArgs e)
        {
            if (selectedImagePaths.Count == 0)
            {
                MessageBox.Show("請先選擇測試圖片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PerformAnalysis();
        }

        private void BtnPrevious_Click(object sender, EventArgs e)
        {
            if (gapAnalysisResults.Count == 0) return;
            currentImageIndex = (currentImageIndex - 1 + gapAnalysisResults.Count) % gapAnalysisResults.Count;
            DisplayCurrentImage();
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (gapAnalysisResults.Count == 0) return;
            currentImageIndex = (currentImageIndex + 1) % gapAnalysisResults.Count;
            DisplayCurrentImage();
        }

        private void BtnFullScreen_Click(object sender, EventArgs e)
        {
            if (gapAnalysisResults.Count == 0 || currentImageIndex >= gapAnalysisResults.Count) return;

            var currentResult = gapAnalysisResults[currentImageIndex];
            if (currentResult.ResultImage != null)
            {
                ShowFullScreenPreview(currentResult);
            }
        }

        // 由 GitHub Copilot 產生 - 修正套用按鈕不關閉視窗
        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (targetStation <= 0)
            {
                MessageBox.Show("尚未選擇站點，無法套用。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                using (var db = new MydbDB())
                {
                    UpdateParameter(db, "gapThresh", currentGapThresh.ToString());
                }

                MessageBox.Show(
                    $"成功套用 gapThresh！\n站點: {targetStation}\n" +
                    $"gapThresh = {currentGapThresh}",
                    "設定完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                btnApplyControl.Text = "✅ 已套用 (可重新套用)";
                btnApplyControl.BackColor = System.Drawing.Color.LightGreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"套用設定時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region 圖像處理方法
        private void PerformAnalysis()
        {
            if (targetStation <= 0)
            {
                MessageBox.Show("請先選擇站點", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (isAnalyzing) return;

            isAnalyzing = true;
            btnAnalyzeControl.Enabled = false;
            btnAnalyzeControl.Text = "🔄 分析中...";

            try
            {
                gapAnalysisResults.Clear();

                foreach (string imagePath in selectedImagePaths)
                {
                    try
                    {
                        var result = ProcessSingleImage(imagePath);
                        gapAnalysisResults.Add(result);
                    }
                    catch (Exception ex)
                    {
                        gapAnalysisResults.Add(new GapAnalysisResult
                        {
                            ImagePath = Path.GetFileName(imagePath),
                            IsValid = false,
                            ErrorMessage = ex.Message,
                            UsedGapThresh = currentGapThresh
                        });
                    }
                }

                currentImageIndex = 0;
                DisplayCurrentImage();
                UpdateNavigationButtons();

                var validCount = gapAnalysisResults.Count(r => r.IsValid);
                var ngCount = gapAnalysisResults.Count(r => r.IsValid && r.IsNG);

                lblRecommendationControl.Text = $"💡 分析完成: {validCount}/{gapAnalysisResults.Count} 張成功, {ngCount} 張檢出Gap (gapThresh={currentGapThresh})";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"分析錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                isAnalyzing = false;
                btnAnalyzeControl.Enabled = true;
                btnAnalyzeControl.Text = "🔄 重新分析";
            }
        }

        private GapAnalysisResult ProcessSingleImage(string imagePath)
        {
            using (var image = Cv2.ImRead(imagePath))
            {
                if (image.Empty())
                {
                    throw new Exception("無法載入圖片");
                }

                // 由 GitHub Copilot 產生 - 更新為接收開口大小數據
                var (isNG, resultImage, outerGapMm, outerGapAngleDeg, inwardBendMm, inwardBendAngleDeg) =
                    AnalyzeGapWithThreshold(image, targetStation, currentGapThresh);

                return new GapAnalysisResult
                {
                    ImagePath = Path.GetFileName(imagePath),
                    IsNG = isNG,
                    OuterGapMm = outerGapMm,
                    OuterGapAngleDeg = outerGapAngleDeg,
                    InwardBendMm = inwardBendMm,
                    InwardBendAngleDeg = inwardBendAngleDeg,
                    OriginalImage = image.Clone(),
                    ResultImage = resultImage,
                    IsValid = true,
                    UsedGapThresh = currentGapThresh
                };
            }
        }

        private (bool isNG, Mat resultImage, List<OpenCvSharp.Point> gapPositions) AnalyzeGapWithThreshold_old(Mat image, int gapThresh)
        {
            // 參考 Form1.cs 中的 findGapWidth 方法實作
            var gapPositions = new List<OpenCvSharp.Point>();
            Mat resultImage = image.Clone();
            bool isNG = false;

            try
            {
                using (var gray = new Mat())
                using (var binary = new Mat())
                using (var morphKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3)))
                {
                    // 轉換為灰階
                    Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                    // 使用調整的 gapThresh 進行二值化
                    Cv2.Threshold(gray, binary, gapThresh, 255, ThresholdTypes.Binary);

                    // 形態學處理
                    using (var opened = new Mat())
                    {
                        Cv2.MorphologyEx(binary, opened, MorphTypes.Open, morphKernel);
                        Cv2.MorphologyEx(opened, binary, MorphTypes.Close, morphKernel);
                    }

                    // 尋找輪廓
                    Cv2.FindContours(binary, out OpenCvSharp.Point[][] contours, out HierarchyIndex[] hierarchy,
                        RetrievalModes.External, ContourApproximationModes.ApproxSimple);

                    // 分析輪廓，找出可能的 Gap
                    foreach (var contour in contours)
                    {
                        double area = Cv2.ContourArea(contour);

                        // Gap 檢測條件（根據實際需求調整）
                        if (area > 50 && area < 5000)
                        {
                            var rect = Cv2.BoundingRect(contour);

                            // 檢查長寬比是否符合 Gap 特徵
                            double aspectRatio = (double)rect.Width / rect.Height;
                            if (aspectRatio > 0.2 && aspectRatio < 5.0)
                            {
                                var center = new OpenCvSharp.Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                gapPositions.Add(center);

                                // 在結果圖像上繪製檢測到的 Gap
                                Cv2.Rectangle(resultImage, rect, new Scalar(0, 0, 255), 2);
                                Cv2.Circle(resultImage, center, 5, new Scalar(255, 0, 0), -1);

                                // 標註 Gap 編號
                                Cv2.PutText(resultImage, $"Gap{gapPositions.Count}",
                                           new OpenCvSharp.Point(rect.X, rect.Y - 5),
                                           HersheyFonts.HersheySimplex, 0.5, new Scalar(0, 255, 0), 1);
                            }
                        }
                    }

                    // 判斷是否為 NG（根據實際需求調整閾值）
                    isNG = gapPositions.Count > 0;

                    // 在圖像上添加資訊
                    string info = $"gapThresh: {gapThresh}, Gaps: {gapPositions.Count}, Result: {(isNG ? "NG" : "OK")}";
                    Cv2.PutText(resultImage, info, new OpenCvSharp.Point(10, 30),
                               HersheyFonts.HersheySimplex, 0.8, new Scalar(0, 255, 255), 2);

                    return (isNG, resultImage, gapPositions);
                }
            }
            catch (Exception ex)
            {
                // 在錯誤時返回原始圖像並標記錯誤
                Cv2.PutText(resultImage, $"Error: {ex.Message}", new OpenCvSharp.Point(10, 30),
                           HersheyFonts.HersheySimplex, 0.8, new Scalar(0, 0, 255), 2);

                return (false, resultImage, gapPositions);
            }
        }
        // 由 GitHub Copilot 產生 - 修改返回值為開口大小數據
        public (bool isNG, Mat img, double outerGapMm, double outerGapAngleDeg, double inwardBendMm, double inwardBendAngleDeg)
            AnalyzeGapWithThreshold(Mat img, int stop, int gapThresh)
        {
            // 無效站別的thresh會預設為0

            // 由 GitHub Copilot 產生
            // 修正: 將 gray 和 ringThresh 納入 using 管理，確保記憶體釋放
            using (Mat ori = img.Clone())
            using (Mat visualImg = img.Clone())
            using (Mat gray = new Mat())
            using (Mat ringThresh = new Mat())
            {
                int minthresh = gapThresh;
                if (minthresh == 0)
                {
                    return (false, visualImg.Clone(), 0, 0, 0, 0); // 0表示不做阈值，需要 Clone 因為會被 Dispose
                }

                Cv2.CvtColor(visualImg, gray, ColorConversionCodes.BGR2GRAY);
                Cv2.Threshold(gray, ringThresh, minthresh, 255, ThresholdTypes.Binary);

                #region 不霍夫
                /*
                int centerTolerance = 80;
                bool matched = false;
                Mat roi_blurred = null;
                CircleSegment? bestInner = null;

                // 1) 用霍夫圓找外圈+內圈 看是否匹配 匹配意義在於確認樣本沒有非常嚴重的變形
                // 2) 沒有非常嚴重變形再去計算內彎和外擴
                if (outerCircles.Length == 0) Console.WriteLine("外圈沒圓");
                if (innerCircles.Length == 0) Console.WriteLine("內圈沒圓");

                if (outerCircles.Length > 0 && innerCircles.Length > 0) //皆存在
                {
                    foreach (var outer in outerCircles)
                    {
                        foreach (var inner in innerCircles)
                        {
                            double radiusDistance = outer.Radius - inner.Radius;
                            double centerDistance = Math.Sqrt(
                                Math.Pow(outer.Center.X - inner.Center.X, 2) +
                                Math.Pow(outer.Center.Y - inner.Center.Y, 2));
                            if (centerDistance <= centerTolerance)
                            {
                                // 视为匹配
                                bestInner = inner;
                                // 建 mask
                                Mat mask = new Mat(visualImg.Size(), MatType.CV_8UC1, Scalar.Black);
                                // 外圈=white
                                Cv2.Circle(mask, (Point)outer.Center, (int)outer.Radius, Scalar.White, -1);
                                // 内圈=black
                                Cv2.Circle(mask, (Point)inner.Center, (int)inner.Radius, Scalar.Black, -1);

                                Mat roi_full = new Mat();
                                Cv2.BitwiseAnd(visualImg, visualImg, roi_full, mask);

                                // 转灰阶+模糊
                                Mat roi_gray = new Mat();
                                Cv2.CvtColor(roi_full, roi_gray, ColorConversionCodes.BGR2GRAY);
                                roi_blurred = new Mat();
                                Cv2.GaussianBlur(roi_gray, roi_blurred, new Size(5, 5), 0);

                                Cv2.CvtColor(roi_blurred, roi_blurred, ColorConversionCodes.GRAY2BGR);

                                matched = true;
                                break;
                            }
                        }
                        if (matched) break;
                    }
                }

                if (!matched || bestInner == null)
                {
                    Console.WriteLine("內外圈未成功匹配, 判斷變形嚴重.");
                    return (true, visualImg);
                }
                */
                #endregion

                int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);
                int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
                int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
                int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);

                Point2f innerCenter = new Point2f(knownInnerCenterX, knownInnerCenterY);
                double innerRadius = knownInnerRadius;
                Point2f outerCenter = new Point2f(knownOuterCenterX, knownOuterCenterY);
                double outerRadius = knownOuterRadius;

                // 4) 對 ringThresh 執行極坐標掃描
                double angleStep = 0.5;
                int nSteps = (int)(360.0 / angleStep);

                bool[] inner_isHole_outward = new bool[nSteps]; // 內圓向外擴掃描的結果
                bool[] inner_isHole_inward = new bool[nSteps];  // 內圓向內彎掃描的結果

                double pixeltomm = double.Parse(app.param[$"PixelToMM_{stop}"]);
                double roiTolerance = (1 / pixeltomm) / (1 / 0.6) - 5; // 0.5mm, 基本落在15~23px
                                                                       //double roiTolerance = 10;
                                                                       // 向外擴掃描的半徑 (原來邏輯)
                double outerScanRadius = innerRadius + roiTolerance;

                // 向內彎掃描的半徑 (新增邏輯：內圓半徑向內縮小)
                double inwardScanRadius = innerRadius - roiTolerance;

                // =========== 向外擴掃描 (內圓向外) ===========
                // 可视化: 在 src 上画点(绿=環, 红=hole)
                for (int i = 0; i < nSteps; i++)
                {
                    double angleDeg = i * angleStep;
                    double rad = angleDeg * Math.PI / 180.0;

                    double rx = innerCenter.X + outerScanRadius * Math.Cos(rad);
                    double ry = innerCenter.Y + outerScanRadius * Math.Sin(rad);

                    int px = (int)Math.Round(rx);
                    int py = (int)Math.Round(ry);

                    if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                    {
                        // hole
                        inner_isHole_outward[i] = true;
                        Cv2.Circle(visualImg, new OpenCvSharp.Point(px, py), 2, Scalar.Red, -1);
                    }
                    else
                    {
                        byte val = ringThresh.Get<byte>(py, px);
                        if (val < 127)
                        {
                            // ring
                            inner_isHole_outward[i] = false;
                            Cv2.Circle(visualImg, new OpenCvSharp.Point(px, py), 1, Scalar.Green, -1);
                        }
                        else
                        {
                            // hole
                            inner_isHole_outward[i] = true;
                            Cv2.Circle(visualImg, new OpenCvSharp.Point(px, py), 2, Scalar.Red, -1);
                        }
                    }
                }

                // =========== 向內彎掃描 (內圓向內) ===========
                for (int i = 0; i < nSteps; i++)
                {
                    double angleDeg = i * angleStep;
                    double rad = angleDeg * Math.PI / 180.0;

                    double rx = innerCenter.X + inwardScanRadius * Math.Cos(rad);
                    double ry = innerCenter.Y + inwardScanRadius * Math.Sin(rad);

                    int px = (int)Math.Round(rx);
                    int py = (int)Math.Round(ry);

                    if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                    {
                        // 超出圖像邊界，標記為hole
                        inner_isHole_inward[i] = true;
                        Cv2.Circle(visualImg, new OpenCvSharp.Point(px, py), 2, Scalar.Blue, -1); // 藍色標記
                    }
                    else
                    {
                        byte val = ringThresh.Get<byte>(py, px);
                        if (val < 127)
                        {
                            // 檢測到環狀物體，說明內彎！這是異常情況
                            inner_isHole_inward[i] = false; // 非hole代表檢測到內彎
                            Cv2.Circle(visualImg, new OpenCvSharp.Point(px, py), 1, Scalar.Yellow, -1); // 黃色標記
                        }
                        else
                        {
                            // 正常情況，內側應該是空白的
                            inner_isHole_inward[i] = true;
                            Cv2.Circle(visualImg, new OpenCvSharp.Point(px, py), 2, Scalar.Blue, -1); // 藍色標記
                        }
                    }
                }

                // 5) 計算外擴尺寸 - 找最大連續缺口
                double outerMaxGapAngleDeg = 0;
                int outerStartIdx = -1;
                for (int idx = 0; idx < nSteps * 2; idx++)
                {
                    int realIdx = idx % nSteps;
                    if (inner_isHole_outward[realIdx])
                    {
                        if (outerStartIdx < 0) outerStartIdx = idx;
                    }
                    else
                    {
                        if (outerStartIdx >= 0)
                        {
                            int length = idx - outerStartIdx;
                            double gapDeg = length * angleStep;
                            if (gapDeg > outerMaxGapAngleDeg) outerMaxGapAngleDeg = gapDeg;
                            outerStartIdx = -1;
                        }
                    }
                }
                if (outerStartIdx >= 0)
                {
                    int length = (nSteps * 2) - outerStartIdx;
                    double gapDeg = length * angleStep;
                    if (gapDeg > outerMaxGapAngleDeg) outerMaxGapAngleDeg = gapDeg;
                }

                // 6) 計算向內彎的尺寸 - 找連續檢測到"非空白"區域的長度
                double inwardBendAngleDeg = 0;
                int inwardStartIdx = -1;
                for (int idx = 0; idx < nSteps * 2; idx++)
                {
                    int realIdx = idx % nSteps;
                    // 如果檢測到非空白，即內彎
                    if (!inner_isHole_inward[realIdx])
                    {
                        if (inwardStartIdx < 0) inwardStartIdx = idx;
                    }
                    else
                    {
                        if (inwardStartIdx >= 0)
                        {
                            int length = idx - inwardStartIdx;
                            double bendDeg = length * angleStep;
                            if (bendDeg > inwardBendAngleDeg) inwardBendAngleDeg = bendDeg;
                            inwardStartIdx = -1;
                        }
                    }
                }
                if (inwardStartIdx >= 0)
                {
                    int length = (nSteps * 2) - inwardStartIdx;
                    double bendDeg = length * angleStep;
                    if (bendDeg > inwardBendAngleDeg) inwardBendAngleDeg = bendDeg;
                }

                // 7) 將角度轉換為物理尺寸 (mm)
                double outerGapArcPx = outerScanRadius * (outerMaxGapAngleDeg * Math.PI / 180.0);
                double outerGapArcMm = outerGapArcPx * pixeltomm;

                double inwardBendArcPx = inwardScanRadius * (inwardBendAngleDeg * Math.PI / 180.0);
                double inwardBendArcMm = inwardBendArcPx * pixeltomm;

                //Console.WriteLine($"外擴情況: 最大開口角度={outerMaxGapAngleDeg:F2} deg => 弧長={outerGapArcMm:F2} mm");
                //Console.WriteLine($"內彎情況: 最大內彎角度={inwardBendAngleDeg:F2} deg => 弧長={inwardBendArcMm:F2} mm");

                // 由 GitHub Copilot 產生
                // 修正: 移除手動 Dispose（using 會自動處理）
                // gray.Dispose();
                // ringThresh.Dispose();

                // 8) 判斷是否NG (內彎或外擴超出閾值)
                bool isOutwardNG = outerGapArcMm >= 1.5;
                bool isInwardNG = inwardBendAngleDeg > 0; // 只要有內彎就算NG

                // 由 GitHub Copilot 產生 - 返回開口大小數據而非位置列表
                if (isOutwardNG || isInwardNG)
                {
                    // ori 會被 using 自動釋放
                    return (true, visualImg.Clone(), outerGapArcMm, outerMaxGapAngleDeg, inwardBendArcMm, inwardBendAngleDeg);
                }
                else
                {
                    // ori, visualImg 都會被 using 自動釋放
                    return (false, visualImg.Clone(), outerGapArcMm, outerMaxGapAngleDeg, inwardBendArcMm, inwardBendAngleDeg);
                }
            } // using 結束，ori, visualImg, gray, ringThresh 自動 Dispose
        }
        #endregion

        #region 顯示方法
        private void DisplayCurrentImage()
        {
            if (gapAnalysisResults.Count == 0 || currentImageIndex >= gapAnalysisResults.Count) return;

            var currentResult = gapAnalysisResults[currentImageIndex];

            try
            {
                // 顯示原始圖像
                if (currentResult.OriginalImage != null)
                {
                    var originalBitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(currentResult.OriginalImage);
                    if (picOriginalControl.Image != null)
                        picOriginalControl.Image.Dispose();
                    picOriginalControl.Image = originalBitmap;
                }

                // 顯示檢測結果圖像
                if (currentResult.ResultImage != null)
                {
                    var resultBitmap = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(currentResult.ResultImage);
                    if (picResultControl.Image != null)
                        picResultControl.Image.Dispose();
                    picResultControl.Image = resultBitmap;
                }

                UpdateImageInfo(currentResult);
            }
            catch (Exception ex)
            {
                lblImageInfoControl.Text = $"圖像顯示錯誤：{ex.Message}";
            }
        }

        private void UpdateImageInfo(GapAnalysisResult currentResult)
        {
            string info = $"📷 圖片 {currentImageIndex + 1}/{gapAnalysisResults.Count}: {currentResult.ImagePath}\r\n\r\n";
            info += $"🔍 分析狀態: {(currentResult.IsValid ? "✅ 成功分析" : "❌ 分析失敗")}\r\n\r\n";

            if (currentResult.IsValid)
            {
                info += $"⚙️ 使用 gapThresh: {currentResult.UsedGapThresh}\r\n";
                info += $"🔍 檢測結果: {(currentResult.IsNG ? "❌ NG (檢出Gap)" : "✅ OK (無Gap)")}\r\n\r\n";

                // 由 GitHub Copilot 產生 - 顯示開口大小數據而非位置
                info += $"📏 開口尺寸測量:\r\n";
                info += $"  外擴開口:\r\n";
                info += $"    弧長: {currentResult.OuterGapMm:F3} mm\r\n";
                info += $"    角度: {currentResult.OuterGapAngleDeg:F2}°\r\n";
                info += $"    判定: {(currentResult.OuterGapMm >= 1.5 ? "❌ NG (≥1.5mm)" : "✅ OK (<1.5mm)")}\r\n\r\n";

                info += $"  內彎檢測:\r\n";
                info += $"    弧長: {currentResult.InwardBendMm:F3} mm\r\n";
                info += $"    角度: {currentResult.InwardBendAngleDeg:F2}°\r\n";
                info += $"    判定: {(currentResult.InwardBendAngleDeg > 0 ? "❌ NG (有內彎)" : "✅ OK (無內彎)")}\r\n";
            }
            else
            {
                info += $"❌ 錯誤原因: {currentResult.ErrorMessage ?? "未知錯誤"}";
            }

            lblImageInfoControl.Text = info;
        }

        private void UpdateNavigationButtons()
        {
            btnPreviousControl.Enabled = gapAnalysisResults.Count > 1;
            btnNextControl.Enabled = gapAnalysisResults.Count > 1;
        }
        #endregion

        #region 輔助方法
        private void UpdateParameter(MydbDB db, string paramName, string value)
        {
            string chineseName = "Gap檢測閾值";

            try
            {
                int updatedRows = db.@params
                    .Where(p => p.Type == targetType && p.Name == paramName && p.Stop == targetStation)
                    .Set(p => p.Value, value)
                    .Set(p => p.ChineseName, chineseName)
                    .Update();

                if (updatedRows == 0)
                {
                    db.@params
                      .Value(p => p.Type, targetType)
                      .Value(p => p.Name, paramName)
                      .Value(p => p.Value, value)
                      .Value(p => p.Stop, targetStation)
                      .Value(p => p.ChineseName, chineseName)
                      .Insert();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"更新 {paramName} (站點 {targetStation}) 時發生錯誤: {ex.Message}");
            }
        }

        private void ShowFullScreenPreview(GapAnalysisResult result)
        {
            using (var fullScreenForm = new Form())
            {
                fullScreenForm.Text = $"Gap檢測結果 - {result.ImagePath} (gapThresh={result.UsedGapThresh})";
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

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);

            foreach (var result in gapAnalysisResults)
            {
                result.OriginalImage?.Dispose();
                result.ResultImage?.Dispose();
            }

            picOriginalControl?.Image?.Dispose();
            picResultControl?.Image?.Dispose();
        }
        #endregion
    }
}