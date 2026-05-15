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
            // 改為返回開口大小數據
            public double OuterGapMm { get; set; }          // 外擴開口弦長 (mm)
            public double OuterGapAngleDeg { get; set; }    // 外擴開口角度 (度)
            public double InwardBendMm { get; set; }        // 內彎弧長 (mm)
            public double InwardBendAngleDeg { get; set; }  // 內彎角度 (度)
            // 新增欄位以正確反映檢測邏輯
            public int InwardBendingCount { get; set; }           // 內彎像素點數（與閾值比較用）
            public int OutwardDeformationCount { get; set; }      // 外凸像素點數（與閾值比較用）
            public double MaxGapWidthMm { get; set; }             // 本次分析實際使用的開口容許閾值 (mm)
            public bool HasGap { get; set; }                      // 是否偵測到開口
            public List<string> NgReasons { get; set; }           // NG 原因清單
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

        // 更新視窗標題預設顯示站點2
        private void InitializeUI()
        {
            this.Text = $"gapThresh 參數視覺化調校 - {targetType} 站點2";
            this.Size = new System.Drawing.Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            targetStation = 2;  // 預設站點2
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
            // 說明標籤（更新為標示僅站點2使用）
            var lblDescription = new Label
            {
                Text = "📋 請選擇測試圖片，調整 gapThresh 參數並觀察 Gap 檢測效果（此參數僅適用於站點 2）",
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
                Size = new System.Drawing.Size(90, 25),
                Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold)
            };

            // 修改為只顯示站點2並預設選中
            cmbStation = new ComboBox
            {
                Location = new System.Drawing.Point(100, 50),
                Size = new System.Drawing.Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft JhengHei", 10F)
            };
            cmbStation.Items.AddRange(new object[] { "站點2" });
            cmbStation.SelectedIndex = 0;
            targetStation = 2;  // 直接設定為站點2
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged;

            // 新增醒目提示標籤
            var lblStation2Only = new Label
            {
                Text = "此參數僅需設定站點 2",
                Location = new System.Drawing.Point(210, 50),
                Size = new System.Drawing.Size(180, 25),
                Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.OrangeRed,
                TextAlign = ContentAlignment.MiddleCenter
            };

            // 選擇照片按鈕
            var btnSelectImages = new Button
            {
                Text = "📂 載入圖片",
                Location = new System.Drawing.Point(400, 50),
                Size = new System.Drawing.Size(130, 30),
                Font = new Font("Microsoft JhengHei", 10F)
            };
            btnSelectImages.Click += BtnSelectImages_Click;

            // gapThresh 調整區域
            var grpGapThreshControl = new GroupBox
            {
                Text = "🎛️ gapThresh 參數調整",
                Location = new System.Drawing.Point(10, 90),
                Size = new System.Drawing.Size(1160, 80),
                Font = new Font("Microsoft JhengHei", 10F)
            };

            var lblGapThresh = new Label
            {
                Text = "gapThresh 值：",
                Location = new System.Drawing.Point(20, 25),
                Size = new System.Drawing.Size(110, 25),
                Font = new Font("Microsoft JhengHei", 10F)
            };

            trackBarGapThresh = new TrackBar
            {
                Location = new System.Drawing.Point(130, 20),
                Size = new System.Drawing.Size(790, 45),
                Minimum = 0,
                Maximum = 254,
                Value = 127,
                TickFrequency = 25,
                LargeChange = 10,
                SmallChange = 1,
                Font = new Font("Microsoft JhengHei", 10F)
            };
            trackBarGapThresh.ValueChanged += TrackBarGapThresh_ValueChanged;

            lblCurrentValue = new Label
            {
                Text = "127",
                Location = new System.Drawing.Point(930, 25),
                Size = new System.Drawing.Size(50, 25),
                Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold),
                ForeColor = Color.Red
            };

            btnAnalyzeControl = new Button
            {
                Text = "🔄 開始分析",
                Location = new System.Drawing.Point(1000, 20),
                Size = new System.Drawing.Size(100, 35),
                BackColor = Color.LightGreen,
                Enabled = false,
                Font = new Font("Microsoft JhengHei", 10F)
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
                Size = new System.Drawing.Size(1160, 450),
                Font = new Font("Microsoft JhengHei", 10F)
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
                Size = new System.Drawing.Size(85, 35),
                Enabled = false,
                Font = new Font("Microsoft JhengHei", 10F)
            };
            btnPreviousControl.Click += BtnPrevious_Click;

            btnNextControl = new Button
            {
                Text = "下一張 ▶",
                Location = new System.Drawing.Point(620, 95),
                Size = new System.Drawing.Size(85, 35),
                Enabled = false,
                Font = new Font("Microsoft JhengHei", 10F)
            };
            btnNextControl.Click += BtnNext_Click;

            // 全螢幕預覽按鈕
            var btnFullScreen = new Button
            {
                Text = "🔍 全螢幕預覽",
                Location = new System.Drawing.Point(620, 140),
                Size = new System.Drawing.Size(85, 35),
                BackColor = Color.LightBlue,
                Font = new Font("Microsoft JhengHei", 10F)
            };
            btnFullScreen.Click += BtnFullScreen_Click;

            // 資訊顯示區域
            lblImageInfoControl = new TextBox
            {
                Text = "請選擇圖片並調整 gapThresh 參數",
                Location = new System.Drawing.Point(720, 50),
                Size = new System.Drawing.Size(420, 350),
                Font = new Font("Microsoft JhengHei", 10F),
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
                Text = "✅ 套用推薦值",
                Location = new System.Drawing.Point(620, 640),
                Size = new System.Drawing.Size(145, 30),
                BackColor = Color.Orange,
                Enabled = false,
                Font = new Font("Microsoft JhengHei", 10F)
            };
            btnApplyControl.Click += BtnApply_Click;

            // 取消按鈕
            var btnCancel = new Button
            {
                Text = "❌ 取消",
                Location = new System.Drawing.Point(775, 640),
                Size = new System.Drawing.Size(80, 30),
                Font = new Font("Microsoft JhengHei", 10F)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            // 加入所有控件（包含新增的站點2提示標籤）
            this.Controls.AddRange(new Control[] {
                lblDescription, lblStation, cmbStation, lblStation2Only, btnSelectImages,
                grpGapThreshControl, grpImageDisplay,
                lblRecommendationControl, btnApplyControl, btnCancel
            });
        }
        #endregion

        #region 事件處理器
        // 修改事件處理邏輯以配合只有站點2的選項結構
        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 由於下拉選單只有「站點2」一個選項，索引0對應站點2
            targetStation = 2;
            this.Text = $"gapThresh 參數視覺化調校 - {targetType} 站點2";
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
                    // 載入新圖片時恢復按鈕文字
                    btnApplyControl.Text = "✅ 套用推薦值";

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

        // 修正套用按鈕不關閉視窗
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

                btnApplyControl.Text = "✅ 已套用";
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

                // 更新為接收完整檢測詳情
                var (isNG, resultImage, outerGapMm, outerGapAngleDeg, inwardBendMm, inwardBendAngleDeg, inwardBendingCount, outwardDeformationCount, maxGapWidthMm, hasGap, ngReasons) =
                    AnalyzeGapWithThreshold(image, targetStation, currentGapThresh);

                return new GapAnalysisResult
                {
                    ImagePath = Path.GetFileName(imagePath),
                    IsNG = isNG,
                    OuterGapMm = outerGapMm,
                    OuterGapAngleDeg = outerGapAngleDeg,
                    InwardBendMm = inwardBendMm,
                    InwardBendAngleDeg = inwardBendAngleDeg,
                    InwardBendingCount = inwardBendingCount,
                    OutwardDeformationCount = outwardDeformationCount,
                    MaxGapWidthMm = maxGapWidthMm,
                    HasGap = hasGap,
                    NgReasons = ngReasons,
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
        // 修改返回值為開口大小數據
        public (bool isNG, Mat img, double outerGapMm, double outerGapAngleDeg, double inwardBendMm, double inwardBendAngleDeg, int inwardBendingCount, int outwardDeformationCount, double maxGapWidthMm, bool hasGap, List<string> ngReasons)
            AnalyzeGapWithThreshold(Mat img, int stop, int gapThresh)
        {
            // 改用 findGapWidth_Single 邏輯：動態擬合圓心 + 變形檢測

            Mat visualImg = img.Clone();

            // ==========================================
            // 1. 參數讀取 (統一於函數開頭)
            // ==========================================

            // --- A. 既有參數 (維持原讀取方式) ---
            int minthresh = gapThresh;
            if (minthresh == 0)
            {
                return (false, visualImg, 0, 0, 0, 0, 0, 0, 0.0, false, new List<string>());
            }

            double pixeltomm = 1.0;
            if (app.param.ContainsKey($"PixelToMM_{stop}"))
                double.TryParse(app.param[$"PixelToMM_{stop}"], out pixeltomm);

            int knownCenterX = 0, knownCenterY = 0, knownRadius = 0;
            if (app.param.ContainsKey($"known_inner_center_x_{stop}")) int.TryParse(app.param[$"known_inner_center_x_{stop}"], out knownCenterX);
            if (app.param.ContainsKey($"known_inner_center_y_{stop}")) int.TryParse(app.param[$"known_inner_center_y_{stop}"], out knownCenterY);
            if (app.param.ContainsKey($"known_inner_radius_{stop}")) int.TryParse(app.param[$"known_inner_radius_{stop}"], out knownRadius);

            // --- B. 新增參數 (使用 _threshold 結尾格式) ---

            // 1. deform_inward -> inwardThreshold
            int inwardThreshold = 40;
            if (app.param.TryGetValue($"deform_inward{stop}_threshold", out string inwardStr))
            {
                int.TryParse(inwardStr, out inwardThreshold);
            }

            // 2. deform_outward -> outwardThreshold
            int outwardThreshold = 10;
            if (app.param.TryGetValue($"deform_outward{stop}_threshold", out string outwardStr))
            {
                int.TryParse(outwardStr, out outwardThreshold);
            }

            // 3. deform_gapTolerance -> maxGapWidthMm
            double maxGapWidthMm = 10.0;
            if (app.param.TryGetValue($"deform_gapTolerance{stop}_threshold", out string gapTolStr))
            {
                double.TryParse(gapTolStr, out maxGapWidthMm);
            }

            // 4. deform_misalign -> tolerancePx
            double tolerancePx = 6.0;
            if (app.param.TryGetValue($"deform_misalign{stop}_threshold", out string misalignStr))
            {
                double.TryParse(misalignStr, out tolerancePx);
            }

            // 5. deform_gapClusterSpan -> gapClusterSpanThreshold (無開口時外凸群集角度容許值，單位：度)
            double gapClusterSpanThreshold = 5.0;
            if (app.param.TryGetValue($"deform_gapClusterSpan{stop}_threshold", out string gapClusterStr))
            {
                double.TryParse(gapClusterStr, out gapClusterSpanThreshold);
            }

            using (Mat gray = new Mat())
            using (Mat binary = new Mat())
            {
                if (img.Channels() == 3)
                    Cv2.CvtColor(img, gray, ColorConversionCodes.BGR2GRAY);
                else
                    img.CopyTo(gray);

                // White = Gap (255), Black = Object (0)
                Cv2.Threshold(gray, binary, minthresh, 255, ThresholdTypes.Binary);

                OpenCvSharp.Point[][] contours;
                HierarchyIndex[] hierarchy;
                Cv2.FindContours(binary, out contours, out hierarchy, RetrievalModes.List, ContourApproximationModes.ApproxNone);

                if (contours.Length == 0) return (false, visualImg, 0, 0, 0, 0, 0, 0, 0.0, false, new List<string>());

                var bushingContour = contours.OrderByDescending(c => Cv2.ContourArea(c)).First();

                Point2f tempCenter = new Point2f(knownCenterX, knownCenterY);
                if (knownRadius == 0)
                {
                    Moments m = Cv2.Moments(bushingContour);
                    if (m.M00 != 0)
                    {
                        tempCenter = new Point2f((float)(m.M10 / m.M00), (float)(m.M01 / m.M00));
                        knownRadius = (int)bushingContour.Min(p => Math.Sqrt(Math.Pow(p.X - tempCenter.X, 2) + Math.Pow(p.Y - tempCenter.Y, 2)));
                    }
                }

                // Extract Inner Points
                double distThreshold = knownRadius + 30;
                List<OpenCvSharp.Point> innerPoints = new List<OpenCvSharp.Point>();
                foreach (var p in bushingContour)
                {
                    double d = Math.Sqrt(Math.Pow(p.X - tempCenter.X, 2) + Math.Pow(p.Y - tempCenter.Y, 2));
                    if (d < distThreshold) innerPoints.Add(p);
                }

                if (innerPoints.Count < 10) return (false, visualImg, 0, 0, 0, 0, 0, 0, 0.0, false, new List<string>());

                // Fit Circle (Least Squares)
                var fitResult = FitCircle(innerPoints);
                Point2f center = fitResult.center;
                float fittedRadius = fitResult.radius;

                // Draw Fitted Circle
                Cv2.Circle(visualImg, (int)center.X, (int)center.Y, 5, Scalar.Blue, -1);
                Cv2.Circle(visualImg, (int)center.X, (int)center.Y, (int)fittedRadius, Scalar.Blue, 1);
                // Y=120: 與 Gap(Y=80) 保持足夠間距，避免重疊
                Cv2.PutText(visualImg, $"R_fit: {fittedRadius:F1}px", new OpenCvSharp.Point(10, 120), HersheyFonts.HersheySimplex, 0.8, Scalar.Yellow, 2);

                // Pixel Scanning for Gap
                int scanRadius = (int)fittedRadius + 20;
                if (scanRadius <= 0) scanRadius = 1;

                int totalSteps = 3600;
                double stepAngle = 360.0 / totalSteps;
                List<List<OpenCvSharp.Point>> allGaps = new List<List<OpenCvSharp.Point>>();
                List<OpenCvSharp.Point> currentGap = new List<OpenCvSharp.Point>();

                for (int i = 0; i < totalSteps; i++)
                {
                    double angle = i * stepAngle;
                    double rad = angle * Math.PI / 180.0;
                    int x = (int)(center.X + scanRadius * Math.Cos(rad));
                    int y = (int)(center.Y + scanRadius * Math.Sin(rad));

                    if (x >= 0 && x < binary.Width && y >= 0 && y < binary.Height)
                    {
                        if (binary.At<byte>(y, x) > 225) // White = Gap
                        {
                            currentGap.Add(new OpenCvSharp.Point(x, y));
                        }
                        else
                        {
                            if (currentGap.Count > 0)
                            {
                                allGaps.Add(new List<OpenCvSharp.Point>(currentGap));
                                currentGap.Clear();
                            }
                        }
                    }
                }
                // Handle wrap-around
                if (currentGap.Count > 0)
                {
                    if (allGaps.Count > 0 && binary.At<byte>((int)(center.Y + scanRadius * Math.Sin(0)), (int)(center.X + scanRadius * Math.Cos(0))) > 225)
                    {
                        allGaps[0].InsertRange(0, currentGap);
                    }
                    else
                    {
                        allGaps.Add(currentGap);
                    }
                }

                var mainGap = allGaps.OrderByDescending(g => g.Count).FirstOrDefault();
                bool hasGap = false;
                OpenCvSharp.Point pStart = new OpenCvSharp.Point();
                OpenCvSharp.Point pEnd = new OpenCvSharp.Point();
                double gapWidthMm = 0;

                if (mainGap != null && mainGap.Count > 1)
                {
                    hasGap = true;
                    pStart = mainGap.First();
                    pEnd = mainGap.Last();
                    double gapWidthPx = Math.Sqrt(Math.Pow(pStart.X - pEnd.X, 2) + Math.Pow(pStart.Y - pEnd.Y, 2));
                    gapWidthMm = gapWidthPx * pixeltomm;
                }

                // 邏輯修正：若有開口，外凸容許值為 outwardThreshold (參數值)；若無開口，容許值為 2
                //outwardThreshold = hasGap ? outwardThreshold : 2;

                // Check Defects
                List<string> ngReasons = new List<string>();
                bool isGapTooWide = gapWidthMm > maxGapWidthMm;

                if (isGapTooWide) ngReasons.Add($"GapBig({gapWidthMm:F1}>{maxGapWidthMm})");

                Scalar gapColor = isGapTooWide ? Scalar.Red : Scalar.Cyan;

                if (hasGap)
                {
                    Cv2.Line(visualImg, pStart, pEnd, gapColor, 2);
                    // Y=80: 顯示開口寬度(mm)，scale 0.8 避免與 R_fit(Y=120) 重疊
                    Cv2.PutText(visualImg, $"Gap: {gapWidthMm:F2}mm", new OpenCvSharp.Point(10, 80), HersheyFonts.HersheySimplex, 0.8, gapColor, 2);
                }
                else
                {
                    Cv2.PutText(visualImg, "No Gap", new OpenCvSharp.Point(10, 80), HersheyFonts.HersheySimplex, 0.8, Scalar.Red, 2);
                }

                int inwardBendingCount = 0;
                int outwardDeformationCount = 0;
                double gapExclusionAngle = 1.0;

                // 無開口時：預先收集所有外凸點角度，計算角度跨度以判斷是否為開口邊緣假陽性
                bool suppressOutward = false;
                if (!hasGap)
                {
                    List<double> outwardAngles = new List<double>();
                    foreach (var q in innerPoints)
                    {
                        double dx2 = q.X - center.X;
                        double dy2 = q.Y - center.Y;
                        double r2 = Math.Sqrt(dx2 * dx2 + dy2 * dy2);
                        if (r2 > fittedRadius + tolerancePx)
                        {
                            double ang = Math.Atan2(dy2, dx2) * 180.0 / Math.PI; // -180 ~ 180
                            outwardAngles.Add(ang);
                        }
                    }
                    if (outwardAngles.Count > 0)
                    {
                        outwardAngles.Sort();
                        // 計算最小包含弧 = 360 - 最大間隔
                        double maxGapBetween = 0;
                        for (int i = 1; i < outwardAngles.Count; i++)
                        {
                            double g = outwardAngles[i] - outwardAngles[i - 1];
                            if (g > maxGapBetween) maxGapBetween = g;
                        }
                        // 考慮跨越 -180/+180 邊界的間隔
                        double wrapGap = (outwardAngles[0] + 360.0) - outwardAngles[outwardAngles.Count - 1];
                        if (wrapGap > maxGapBetween) maxGapBetween = wrapGap;
                        double clusterSpan = 360.0 - maxGapBetween;
                        // 角度跨度 < 門檻值 → 視為開口邊緣幾何假陽性，全部忽略
                        suppressOutward = (clusterSpan < gapClusterSpanThreshold);
                    }
                }

                foreach (var p in innerPoints)
                {
                    double dx = p.X - center.X;
                    double dy = p.Y - center.Y;
                    double actualRadius = Math.Sqrt(dx * dx + dy * dy);
                    double ux = dx / actualRadius;
                    double uy = dy / actualRadius;

                    // Draw Green Corridor
                    OpenCvSharp.Point pOut = new OpenCvSharp.Point((int)(center.X + ux * (fittedRadius + tolerancePx)), (int)(center.Y + uy * (fittedRadius + tolerancePx)));
                    OpenCvSharp.Point pIn = new OpenCvSharp.Point((int)(center.X + ux * (fittedRadius - tolerancePx)), (int)(center.Y + uy * (fittedRadius - tolerancePx)));
                    if (pOut.X >= 0 && pOut.X < visualImg.Width && pOut.Y >= 0 && pOut.Y < visualImg.Height) visualImg.At<Vec3b>(pOut.Y, pOut.X) = new Vec3b(0, 255, 0);
                    if (pIn.X >= 0 && pIn.X < visualImg.Width && pIn.Y >= 0 && pIn.Y < visualImg.Height) visualImg.At<Vec3b>(pIn.Y, pIn.X) = new Vec3b(0, 255, 0);

                    if (actualRadius < (fittedRadius - tolerancePx))
                    {
                        Cv2.Circle(visualImg, p, 2, Scalar.Red, -1);
                        inwardBendingCount++;
                    }
                    else if (actualRadius > (fittedRadius + tolerancePx))
                    {
                        bool isDefect = true;
                        if (hasGap)
                        {
                            double angleP = Math.Atan2(p.Y - center.Y, p.X - center.X) * 180.0 / Math.PI;
                            double angleStart = Math.Atan2(pStart.Y - center.Y, pStart.X - center.X) * 180.0 / Math.PI;
                            double angleEnd = Math.Atan2(pEnd.Y - center.Y, pEnd.X - center.X) * 180.0 / Math.PI;

                            if (GetAngleDiff(angleP, angleStart) < gapExclusionAngle || GetAngleDiff(angleP, angleEnd) < gapExclusionAngle)
                            {
                                isDefect = false;
                            }
                        }

                        // 無開口且群集跨度 < 門檻值時，忽略此外凸點
                        if (isDefect && !suppressOutward)
                        {
                            Cv2.Circle(visualImg, p, 2, Scalar.Magenta, -1);
                            outwardDeformationCount++;
                        }
                    }
                }

                if (inwardBendingCount > inwardThreshold) ngReasons.Add($"Inward({inwardBendingCount})");
                if (outwardDeformationCount > outwardThreshold) ngReasons.Add($"Outward({outwardDeformationCount})");

                bool isDeformed = ngReasons.Count > 0;
                if (isDeformed)
                {
                    // 每條原因分行顯示，從 Y=150 開始，間距 23px，避免單行過長或與 R_fit(Y=120) 重疊
                    int reasonY = 150;
                    foreach (var r in ngReasons)
                    {
                        Cv2.PutText(visualImg, r, new OpenCvSharp.Point(10, reasonY), HersheyFonts.HersheySimplex, 0.6, Scalar.Red, 1);
                        reasonY += 23;
                    }
                }

                string statusText = isDeformed ? "NG" : "OK";
                Scalar statusColor = isDeformed ? Scalar.Red : Scalar.Green;
                Cv2.PutText(visualImg, $"Result: {statusText}", new OpenCvSharp.Point(10, 30), HersheyFonts.HersheySimplex, 1.2, statusColor, 3);

                // 計算回傳值以保持與 UI 相容
                // outerGapMm: 直接使用 gapWidthMm (開口寬度)
                // outerGapAngleDeg: 根據主要開口的點數計算角度
                double outerGapAngleDeg = 0;
                if (mainGap != null && mainGap.Count > 0)
                {
                    outerGapAngleDeg = (mainGap.Count / (double)totalSteps) * 360.0;
                }

                // inwardBendMm: 根據內彎點數計算弧長
                double inwardBendAngleDeg = 0;
                double inwardBendMm = 0;
                if (inwardBendingCount > 0)
                {
                    // 估算內彎弧長：假設內彎點均勻分布
                    inwardBendAngleDeg = (inwardBendingCount / (double)innerPoints.Count) * 360.0;
                    inwardBendMm = inwardBendAngleDeg * Math.PI / 180.0 * fittedRadius * pixeltomm;
                }

                // 擴充回傳元組以包含檢測詳情
                return (isDeformed, visualImg, gapWidthMm, outerGapAngleDeg, inwardBendMm, inwardBendAngleDeg, inwardBendingCount, outwardDeformationCount, maxGapWidthMm, hasGap, new List<string>(ngReasons));
            }
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
            // 修正所有顯示邏輯與實際檢測邏輯的出入
            string info = $"📷 圖片 {currentImageIndex + 1}/{gapAnalysisResults.Count}: {currentResult.ImagePath}\r\n\r\n";
            info += $"🔍 分析狀態: {(currentResult.IsValid ? "✅ 成功分析" : "❌ 分析失敗")}\r\n\r\n";

            if (currentResult.IsValid)
            {
                info += $"⚙️ 使用 gapThresh: {currentResult.UsedGapThresh}\r\n";

                // 修正出入一：NG 描述不再硬寫「檢出Gap」，改由 NgReasons 顯示實際原因
                info += $"🔍 檢測結果: {(currentResult.IsNG ? "❌ NG" : "✅ OK")}\r\n";
                if (currentResult.IsNG && currentResult.NgReasons != null && currentResult.NgReasons.Count > 0)
                {
                    info += $"  NG 原因: {string.Join(", ", currentResult.NgReasons)}\r\n";
                }
                info += "\r\n";

                info += $"📏 開口尺寸測量:\r\n";
                // 補充開口狀態
                info += $"  開口狀態: {(currentResult.HasGap ? "有開口" : "無開口")}\r\n";
                info += $"  外擴開口:\r\n";
                // 修正出入四：「弧長」改為「開口弦長」（計算方式為兩端點直線距離）
                info += $"    開口弦長: {currentResult.OuterGapMm:F3} mm\r\n";
                info += $"    角度: {currentResult.OuterGapAngleDeg:F2}°\r\n";
                // 修正出入二：判定閾值使用實際參數值 MaxGapWidthMm，而非硬編碼 1.5mm
                info += $"    判定: {(currentResult.OuterGapMm > currentResult.MaxGapWidthMm ? $"❌ NG (>{currentResult.MaxGapWidthMm:F1}mm)" : $"✅ OK (≤{currentResult.MaxGapWidthMm:F1}mm)")}\r\n\r\n";

                // 修正出入三：內彎判定依 NgReasons 決定 NG/OK，顯示實際點數
                bool inwardNG = currentResult.NgReasons != null && currentResult.NgReasons.Any(r => r.StartsWith("Inward"));
                info += $"  內彎檢測:\r\n";
                info += $"    像素點數: {currentResult.InwardBendingCount}\r\n";
                info += $"    弧長: {currentResult.InwardBendMm:F3} mm\r\n";
                info += $"    角度: {currentResult.InwardBendAngleDeg:F2}°\r\n";
                if (inwardNG)
                    info += $"    判定: ❌ NG (超過閾值)\r\n\r\n";
                else if (currentResult.InwardBendingCount > 0)
                    info += $"    判定: ✅ OK ({currentResult.InwardBendingCount} 點，未達閾值)\r\n\r\n";
                else
                    info += $"    判定: ✅ OK (無內彎)\r\n\r\n";

                // 修正出入五：新增外凸檢測區塊
                bool outwardNG = currentResult.NgReasons != null && currentResult.NgReasons.Any(r => r.StartsWith("Outward"));
                info += $"  外凸檢測:\r\n";
                info += $"    像素點數: {currentResult.OutwardDeformationCount}\r\n";
                if (outwardNG)
                    info += $"    判定: ❌ NG (超過閾值)\r\n";
                else if (currentResult.OutwardDeformationCount > 0)
                    info += $"    判定: ✅ OK ({currentResult.OutwardDeformationCount} 點，未達閾值)\r\n";
                else
                    info += $"    判定: ✅ OK (無外凸)\r\n";
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

        #region 輔助函數
        // 複製自 Form1.cs - 最小平方法圓擬合
        private (Point2f center, float radius) FitCircle(List<OpenCvSharp.Point> points)
        {
            int numPoints = points.Count;
            if (numPoints < 3) return (new Point2f(0, 0), 0);

            // 使用 OpenCV Solve 求解線性方程組
            // 2x_i * A + 2y_i * B + C = x_i^2 + y_i^2
            // Center(A, B), Radius = Sqrt(C + A^2 + B^2)

            using (Mat matA = new Mat(numPoints, 3, MatType.CV_32F))
            using (Mat matB = new Mat(numPoints, 1, MatType.CV_32F))
            {
                var indexerA = matA.GetGenericIndexer<float>();
                var indexerB = matB.GetGenericIndexer<float>();

                for (int i = 0; i < numPoints; i++)
                {
                    float x = points[i].X;
                    float y = points[i].Y;
                    indexerA[i, 0] = 2 * x;
                    indexerA[i, 1] = 2 * y;
                    indexerA[i, 2] = 1;
                    indexerB[i, 0] = x * x + y * y;
                }

                using (Mat matX = new Mat())
                {
                    // SVD 分解求解
                    Cv2.Solve(matA, matB, matX, DecompTypes.SVD);

                    float A = matX.At<float>(0);
                    float B = matX.At<float>(1);
                    float C = matX.At<float>(2);

                    float r = (float)Math.Sqrt(C + A * A + B * B);
                    return (new Point2f(A, B), r);
                }
            }
        }

        // 複製自 Form1.cs - 計算角度差
        private double GetAngleDiff(double a1, double a2)
        {
            double diff = Math.Abs(a1 - a2);
            if (diff > 180) diff = 360 - diff;
            return diff;
        }
        #endregion

        #endregion
    }
}