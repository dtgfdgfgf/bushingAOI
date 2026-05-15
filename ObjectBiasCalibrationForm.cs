using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using LinqToDB;

namespace peilin
{
    public partial class ObjectBiasCalibrationForm : Form
    {
        #region 私有成員變數
        private List<string> selectedImagePaths = new List<string>();
        private List<ObjectCenterResult> centerResults = new List<ObjectCenterResult>();
        private string targetType;
        private int targetStation;
        private int recommendedBiasX, recommendedBiasY;
        private int currentImageIndex = 0;
        private bool isAnalyzing = false;
        private bool eventsInitialized = false;
        private bool isUpdatingImage = false; // 防止遞迴更新
        
        // 統計分析結果
        private double _overallStdX = 0;
        private double _overallStdY = 0;
        private int _medianX = 0;
        private int _medianY = 0;
        private string _stabilityLevel = "---";  // 整體穩定性等級文字
        private double _maxStd = 0;               // 最大標準差值

        // 控件成員變數
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
        #endregion

        #region 物體中心計算結果結構
        public class ObjectCenterResult
        {
            public string ImagePath { get; set; }
            public OpenCvSharp.Point DetectedCenter { get; set; }
            public OpenCvSharp.Point ImageCenter { get; set; }
            public OpenCvSharp.Point Offset { get; set; }
            public double Distance { get; set; }
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; }
            public Mat ResultImage { get; set; }
        }
        #endregion

        #region 建構函數和初始化
        public ObjectBiasCalibrationForm(string type)
        {
            targetType = type;
            targetStation = 0;
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = $"物體偏移校正 - {targetType} (請選擇站點)";
            this.Size = new System.Drawing.Size(1005, 805);
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
                Text = "📋 請選擇8-15張品質良好的樣品照片，系統將自動計算最佳的物體偏移量(objBias)",
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(960, 25),
                Font = new Font("Microsoft JhengHei", 10),
                ForeColor = Color.Blue
            };
            var lblDescription2 = new Label
            {
                Text = "📋 請選擇已經過前處理的樣品照片",
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(960, 25),
                Font = new Font("Microsoft JhengHei", 10),
                ForeColor = Color.Red
            };
            // 顯示此介面只須設定站點2的提示
            var lblStationNotice = new Label
            {
                Text = "⚠️ 此介面只須設定站點2",
                Location = new System.Drawing.Point(410, 68),
                Size = new System.Drawing.Size(200, 29),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.Orange,
                TextAlign = ContentAlignment.MiddleLeft
            };
            // 【新增】：站點選擇
            var lblStation = new Label
            {
                Text = "選擇站點：",
                Location = new System.Drawing.Point(10, 70),
                Size = new System.Drawing.Size(70, 25),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold)
            };

            // 修正只保留站點2
            var cmbStation = new ComboBox
            {
                Location = new System.Drawing.Point(85, 70),
                Size = new System.Drawing.Size(80, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft JhengHei", 10)
            };
            cmbStation.Items.AddRange(new object[] { "請選擇", "站點2" });
            cmbStation.SelectedIndexChanged += (s, e) =>
            {
                // 站點2 對應 targetStation = 2
                targetStation = cmbStation.SelectedIndex == 1 ? 2 : 0;
                if (targetStation <= 0)
                {
                    this.Text = $"物體偏移校正 - {targetType} (請選擇站點)";
                    btnAnalyzeControl.Enabled = false;
                    btnApplyControl.Enabled = false;
                    return;
                }
                this.Text = $"物體偏移校正 - {targetType} 站點{targetStation}";
                btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
            };

            // 選擇照片按鈕
            var btnSelectImages = new Button
            {
                Text = "📂 載入圖片",
                Location = new System.Drawing.Point(175, 68),
                Size = new System.Drawing.Size(110, 29),
                Font = new Font("Microsoft JhengHei", 10)
            };
            btnSelectImages.Click += BtnSelectImages_Click;

            // 開始分析按鈕
            btnAnalyzeControl = new Button
            {
                Text = "🔍 開始分析",
                Location = new System.Drawing.Point(290, 68),
                Size = new System.Drawing.Size(100, 29),
                BackColor = Color.LightGreen,
                Font = new Font("Microsoft JhengHei", 10),
                Enabled = false
            };
            btnAnalyzeControl.Click += BtnAnalyze_Click;

            // 預設選擇站點2（必須在 btnAnalyzeControl 初始化之後）
            cmbStation.SelectedIndex = 1;

            // 進度條 - 修正位置移到按鈕右側
            progressBarControl = new ProgressBar
            {
                Location = new System.Drawing.Point(400, 68),
                Size = new System.Drawing.Size(150, 29),
                Visible = false
            };

            // 結果顯示區域
            dgvResultsControl = new DataGridView
            {
                Location = new System.Drawing.Point(10, 110),
                Size = new System.Drawing.Size(600, 400),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Microsoft JhengHei", 10)
            };

            // 統計資訊面板
            grpStatisticsControl = new GroupBox
            {
                Text = "📊 統計分析結果",
                Location = new System.Drawing.Point(620, 95),
                Size = new System.Drawing.Size(360, 425),
                Font = new Font("Microsoft JhengHei", 10)
            };
            CreateStatisticsLabels(grpStatisticsControl);

            // 圖片預覽區域
            var grpImagePreview = new GroupBox
            {
                Text = "📷 檢測結果預覽",
                Location = new System.Drawing.Point(10, 520),
                Size = new System.Drawing.Size(970, 200),
                Font = new Font("Microsoft JhengHei", 10)
            };

            // 預覽圖片
            picPreviewControl = new PictureBox
            {
                Location = new System.Drawing.Point(10, 20),
                Size = new System.Drawing.Size(320, 170),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };

            // 導航按鈕
            btnPreviousControl = new Button
            {
                Text = "◀ 上一張",
                Location = new System.Drawing.Point(340, 30),
                Size = new System.Drawing.Size(90, 35),
                Font = new Font("Microsoft JhengHei", 10),
                Enabled = false
            };
            btnPreviousControl.Click += BtnPrevious_Click;

            btnNextControl = new Button
            {
                Text = "下一張 ▶",
                Location = new System.Drawing.Point(340, 75),
                Size = new System.Drawing.Size(90, 35),
                Font = new Font("Microsoft JhengHei", 10),
                Enabled = false
            };
            btnNextControl.Click += BtnNext_Click;

            // 全螢幕預覽按鈕
            var btnFullScreen = new Button
            {
                Text = "🔍 全螢幕預覽",
                Location = new System.Drawing.Point(340, 120),
                Size = new System.Drawing.Size(90, 35),
                Font = new Font("Microsoft JhengHei", 10),
                BackColor = Color.LightBlue
            };
            btnFullScreen.Click += BtnFullScreen_Click;

            // 圖片資訊標籤
            // 【修正後】：改為 TextBox（支援捲動）
            lblImageInfoControl = new TextBox
            {
                Text = "尚未分析圖片",
                Location = new System.Drawing.Point(450, 20),
                Size = new System.Drawing.Size(500, 170),
                Font = new Font("Microsoft JhengHei", 10),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Multiline = true,              // 【關鍵】：支援多行
                ScrollBars = ScrollBars.Vertical, // 【關鍵】：垂直捲動條
                ReadOnly = true,               // 【關鍵】：唯讀模式
                WordWrap = true,               // 自動換行
                TabStop = false                // 不參與 Tab 順序
            };

            grpImagePreview.Controls.AddRange(new Control[] {
                picPreviewControl, btnPreviousControl, btnNextControl, btnFullScreen, lblImageInfoControl
            });

            // 建議值顯示
            lblRecommendationControl = new Label
            {
                Text = "💡 建議的 objBias 值將在分析完成後顯示",
                Location = new System.Drawing.Point(10, 730),
                Size = new System.Drawing.Size(500, 30),
                Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold),
                ForeColor = Color.DarkGreen
            };

            // 套用按鈕
            btnApplyControl = new Button
            {
                Text = "✅ 套用推薦值",
                Location = new System.Drawing.Point(520, 730),
                Size = new System.Drawing.Size(120, 30),
                Font = new Font("Microsoft JhengHei", 10),
                BackColor = Color.Orange,
                Enabled = false
            };
            btnApplyControl.Click += BtnApply_Click;

            // 取消按鈕
            var btnCancel = new Button
            {
                Text = "❌ 取消",
                Location = new System.Drawing.Point(650, 730),
                Size = new System.Drawing.Size(80, 30),
                Font = new Font("Microsoft JhengHei", 10)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            // 加入所有控件
            this.Controls.AddRange(new Control[] {
                lblDescription,  lblDescription2, lblStationNotice, lblStation, cmbStation, btnSelectImages, btnAnalyzeControl, progressBarControl,
                dgvResultsControl, grpStatisticsControl, grpImagePreview,
                lblRecommendationControl, btnApplyControl, btnCancel
            }); 
        }

        private void CreateStatisticsLabels(GroupBox parent)
        {
            var labels = new[] {
                "樣品數量：", "有效樣品：", "平均X偏移：", "平均Y偏移：",
                "X標準差：", "Y標準差：", "最大偏移距離：", "建議objBias_X：", "建議objBias_Y：",
                "整體穩定性："
            };

            for (int i = 0; i < labels.Length; i++)
            {
                var lbl = new Label
                {
                    Text = labels[i],
                    Location = new System.Drawing.Point(10, 25 + i * 35),
                    Size = new System.Drawing.Size(120, 25),
                    Name = $"lblStat{i}",
                    Font = new Font("Microsoft JhengHei", 10)
                };

                // 整體穩定性欄位需要較大高度以顯示多行文字
                bool isStabilityRow = (i == labels.Length - 1);
                var lblValue = new Label
                {
                    Text = "---",
                    Location = new System.Drawing.Point(140, 25 + i * 35),
                    Size = isStabilityRow ? new System.Drawing.Size(210, 75) : new System.Drawing.Size(200, 25),
                    Name = $"lblStatValue{i}",
                    Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold)
                };

                parent.Controls.AddRange(new Control[] { lbl, lblValue });
            }
        }
        #endregion

        #region 事件處理器
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

                    if (selectedImagePaths.Count < 5)
                    {
                        MessageBox.Show("建議至少選擇5張照片以獲得更準確的結果", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (selectedImagePaths.Count > 20)
                    {
                        MessageBox.Show("選擇的照片過多，建議選擇8-15張代表性照片", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    btnAnalyzeControl.Text = $"🔍 分析 {selectedImagePaths.Count} 張照片";
                    btnAnalyzeControl.Enabled = selectedImagePaths.Count > 0;
                    // 載入新圖片時恢復按鈕文字
                    btnApplyControl.Text = "✅ 套用推薦值";
                }
            }
        }

        private async void BtnAnalyze_Click(object sender, EventArgs e)
        {
            if (targetStation <= 0)
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

                centerResults.Clear();
                await ProcessImagesSimple();
                DisplayAnalysisResults();

                MessageBox.Show($"分析完成！\n" +
                               $"總計 {centerResults.Count} 張照片\n" +
                               $"成功檢測 {centerResults.Count(r => r.IsValid)} 張\n" +
                               $"失敗 {centerResults.Count(r => !r.IsValid)} 張",
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
            if (centerResults.Count == 0 || isUpdatingImage) return;
            currentImageIndex = (currentImageIndex - 1 + centerResults.Count) % centerResults.Count;
            DisplayCurrentImageAsync(); // 【修改】改用非同步方法
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (centerResults.Count == 0 || isUpdatingImage) return;
            currentImageIndex = (currentImageIndex + 1) % centerResults.Count;
            DisplayCurrentImageAsync(); // 【修改】改用非同步方法
        }

        private void BtnFullScreen_Click(object sender, EventArgs e)
        {
            if (centerResults.Count == 0 || currentImageIndex >= centerResults.Count) return;

            var currentResult = centerResults[currentImageIndex];
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
                    UpdateParameter(db, "objBias_x", recommendedBiasX.ToString());
                    UpdateParameter(db, "objBias_y", recommendedBiasY.ToString());
                }

                MessageBox.Show(
                    $"成功套用 objBias 參數！\n站點: {targetStation}\n" +
                    $"objBias_x = {recommendedBiasX}\nobjBias_y = {recommendedBiasY}",
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

        private void DgvResults_SelectionChanged(object sender, EventArgs e)
        {
            if (isAnalyzing || isUpdatingImage) return; // 【修改】加入 isUpdatingImage 檢查

            if (dgvResultsControl.SelectedRows.Count > 0)
            {
                currentImageIndex = dgvResultsControl.SelectedRows[0].Index;
                DisplayCurrentImageAsync(); // 【修改】改用非同步方法
            }
        }
        #endregion

        #region 圖像處理方法
        private async Task ProcessImagesSimple()
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
                        centerResults.Add(result);

                        System.Threading.Thread.Sleep(50);
                    }
                    catch (Exception ex)
                    {
                        centerResults.Add(new ObjectCenterResult
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

        private ObjectCenterResult ProcessSingleImage(string imagePath)
        {
            using (var image = Cv2.ImRead(imagePath))
            {
                if (image.Empty())
                {
                    throw new Exception("無法載入圖片");
                }

                var (isValid, detectedCenter, resultImage) = DetectObjectCenter(image);

                var imageCenter = new OpenCvSharp.Point(image.Width / 2, image.Height / 2);
                var offset = new OpenCvSharp.Point(detectedCenter.X - imageCenter.X, detectedCenter.Y - imageCenter.Y);
                var distance = Math.Sqrt(Math.Pow(offset.X, 2) + Math.Pow(offset.Y, 2));

                return new ObjectCenterResult
                {
                    ImagePath = Path.GetFileName(imagePath),
                    DetectedCenter = detectedCenter,
                    ImageCenter = imageCenter,
                    Offset = offset,
                    Distance = distance,
                    IsValid = isValid,
                    ResultImage = resultImage
                };
            }
        }

        private (bool isValid, OpenCvSharp.Point center, Mat resultImage) DetectObjectCenter(Mat image)
        {
            try
            {
                using (var gray = new Mat())
                using (var binary = new Mat())
                {
                    Mat resultImg = image.Clone();
                    var imageCenterPoint = new OpenCvSharp.Point(image.Width / 2, image.Height / 2);

                    Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);
                    Cv2.Threshold(gray, binary, 250, 255, ThresholdTypes.Binary);

                    Cv2.FindContours(binary, out OpenCvSharp.Point[][] contours, out HierarchyIndex[] hierarchy,
                        RetrievalModes.External, ContourApproximationModes.ApproxSimple);

                    if (contours.Length == 0)
                    {
                        Cv2.Circle(resultImg, imageCenterPoint, 10, new Scalar(0, 0, 255), -1);
                        Cv2.PutText(resultImg, "No contours found",
                                   new OpenCvSharp.Point(50, 50), HersheyFonts.HersheySimplex, 1, new Scalar(0, 0, 255), 2);
                        return (false, new OpenCvSharp.Point(0, 0), resultImg);
                    }

                    OpenCvSharp.Point[] largestContour = null;
                    double maxArea = 0;

                    foreach (var contour in contours)
                    {
                        double area = Cv2.ContourArea(contour);
                        if (200000 <= area && area <= 1000000 && area > maxArea)
                        {
                            maxArea = area;
                            largestContour = contour;
                        }
                    }

                    if (largestContour == null)
                    {
                        Cv2.Circle(resultImg, imageCenterPoint, 10, new Scalar(0, 0, 255), -1);
                        Cv2.PutText(resultImg, "No valid size contour",
                                   new OpenCvSharp.Point(50, 50), HersheyFonts.HersheySimplex, 1, new Scalar(0, 0, 255), 2);
                        return (false, new OpenCvSharp.Point(0, 0), resultImg);
                    }

                    var moments = Cv2.Moments(largestContour);
                    if (moments.M00 == 0)
                    {
                        Cv2.Circle(resultImg, imageCenterPoint, 10, new Scalar(0, 0, 255), -1);
                        Cv2.DrawContours(resultImg, new[] { largestContour }, -1, new Scalar(0, 255, 255), 2);
                        Cv2.PutText(resultImg, "Cannot calculate moments",
                                   new OpenCvSharp.Point(50, 50), HersheyFonts.HersheySimplex, 1, new Scalar(0, 0, 255), 2);
                        return (false, new OpenCvSharp.Point(0, 0), resultImg);
                    }

                    var centerX = (int)(moments.M10 / moments.M00);
                    var centerY = (int)(moments.M01 / moments.M00);
                    var detectedCenter = new OpenCvSharp.Point(centerX, centerY);

                    // 視覺化
                    Cv2.DrawContours(resultImg, new[] { largestContour }, -1, new Scalar(0, 255, 0), 2);
                    Cv2.Circle(resultImg, imageCenterPoint, 10, new Scalar(0, 0, 255), -1);
                    Cv2.Circle(resultImg, detectedCenter, 10, new Scalar(255, 0, 0), -1);
                    Cv2.Line(resultImg, detectedCenter, imageCenterPoint, new Scalar(0, 255, 255), 2);

                    var offset = new OpenCvSharp.Point(detectedCenter.X - imageCenterPoint.X, detectedCenter.Y - imageCenterPoint.Y);
                    var distance = Math.Sqrt(Math.Pow(offset.X, 2) + Math.Pow(offset.Y, 2));

                    string offsetInfo = $"Offset: X={offset.X}, Y={offset.Y}, Dist={distance:F1}";
                    Cv2.PutText(resultImg, offsetInfo, new OpenCvSharp.Point(10, 30),
                               HersheyFonts.HersheySimplex, 0.8, new Scalar(255, 255, 255), 2);

                    return (true, detectedCenter, resultImg);
                }
            }
            catch (Exception ex)
            {
                Mat resultImg = image.Clone();
                var imageCenterPoint = new OpenCvSharp.Point(image.Width / 2, image.Height / 2);
                Cv2.Circle(resultImg, imageCenterPoint, 10, new Scalar(0, 0, 255), -1);
                Cv2.PutText(resultImg, $"Error: {ex.Message}",
                           new OpenCvSharp.Point(50, 50), HersheyFonts.HersheySimplex, 0.8, new Scalar(0, 0, 255), 2);

                return (false, new OpenCvSharp.Point(0, 0), resultImg);
            }
        }
        #endregion

        #region 結果顯示和統計
        private void DisplayAnalysisResults()
        {
            var resultsData = centerResults.Select(r => new
            {
                照片名稱 = r.ImagePath,
                是否有效 = r.IsValid ? "✅" : "❌",
                檢測中心X = r.IsValid ? r.DetectedCenter.X.ToString() : "---",
                檢測中心Y = r.IsValid ? r.DetectedCenter.Y.ToString() : "---",
                X偏移量 = r.IsValid ? r.Offset.X.ToString() : "---",
                Y偏移量 = r.IsValid ? r.Offset.Y.ToString() : "---",
                偏移距離 = r.IsValid ? r.Distance.ToString("F1") : "---",
                錯誤訊息 = r.ErrorMessage ?? ""
            }).ToList();

            dgvResultsControl.DataSource = resultsData;

            if (!eventsInitialized)
            {
                dgvResultsControl.SelectionChanged += DgvResults_SelectionChanged;
                eventsInitialized = true;
            }

            var validResults = centerResults.Where(r => r.IsValid).ToList();

            if (validResults.Count > 0)
            {
                CalculateAndDisplayStatistics(validResults);
                currentImageIndex = 0;
                DisplayCurrentImageAsync(); // 【修改】改用非同步方法
            }
            else
            {
                MessageBox.Show("沒有有效的分析結果，請檢查照片品質或參數設定", "警告",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private async void DisplayCurrentImageAsync()
        {
            if (centerResults.Count == 0 || currentImageIndex >= centerResults.Count || isUpdatingImage)
                return;

            isUpdatingImage = true; // 【關鍵】防止重複觸發

            try
            {
                var currentResult = centerResults[currentImageIndex];

                // 先更新文字資訊（不會造成延遲）
                UpdateImageInfo(currentResult);
                UpdateNavigationButtons();

                if (currentResult.ResultImage != null)
                {
                    // 【關鍵修正】在背景執行緒進行圖像轉換
                    System.Drawing.Bitmap bitmap = await Task.Run(() =>
                    {
                        try
                        {
                            return OpenCvSharp.Extensions.BitmapConverter.ToBitmap(currentResult.ResultImage);
                        }
                        catch
                        {
                            return null;
                        }
                    });

                    // 回到 UI 執行緒更新圖片
                    if (bitmap != null)
                    {
                        if (picPreviewControl.Image != null)
                        {
                            var oldImage = picPreviewControl.Image;
                            picPreviewControl.Image = null;
                            oldImage.Dispose();
                        }

                        picPreviewControl.Image = bitmap;
                    }
                    else
                    {
                        lblImageInfoControl.Text += "\r\n\r\n❌ 圖像顯示錯誤：無法轉換圖像";

                        if (picPreviewControl.Image != null)
                        {
                            var oldImage = picPreviewControl.Image;
                            picPreviewControl.Image = null;
                            oldImage.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lblImageInfoControl.Text += $"\r\n\r\n❌ 圖像顯示錯誤：{ex.Message}";

                if (picPreviewControl.Image != null)
                {
                    var oldImage = picPreviewControl.Image;
                    picPreviewControl.Image = null;
                    oldImage.Dispose();
                }
            }
            finally
            {
                isUpdatingImage = false; // 【關鍵】解除鎖定
            }
        }
        

        // 修正穩定性評估邏輯
        // 穩定性應基於「整體標準差大小」，而非單一樣本離群程度
        // 標準差小 = 所有樣本位置一致 = 穩定；標準差大 = 樣本間位置變化大 = 不穩定
        private void UpdateImageInfo(ObjectCenterResult currentResult)
        {
            string info = $"📷 圖片 {currentImageIndex + 1}/{centerResults.Count}: {currentResult.ImagePath}\r\n\r\n";
            info += $"🔍 檢測狀態: {(currentResult.IsValid ? "✅ 成功檢測" : "❌ 檢測失敗")}\r\n\r\n";

            if (currentResult.IsValid)
            {
                info += $"📍 物體中心座標: ({currentResult.DetectedCenter.X}, {currentResult.DetectedCenter.Y})\r\n";
                info += $"🎯 圖像中心座標: ({currentResult.ImageCenter.X}, {currentResult.ImageCenter.Y})\r\n\r\n";
                info += $"📏 X軸偏移量: {currentResult.Offset.X} 像素\r\n";
                info += $"📏 Y軸偏移量: {currentResult.Offset.Y} 像素\r\n";
                info += $"📐 總偏移距離: {currentResult.Distance:F1} 像素\r\n\r\n";

                // 顯示整體穩定性結論（所有圖片共用同一評估結果）
                info += $"📊 整體穩定性: {_stabilityLevel}";
            }
            else
            {
                info += $"❌ 錯誤原因: {currentResult.ErrorMessage ?? "未知錯誤"}";
            }

            lblImageInfoControl.Text = info;
        }

        private void UpdateNavigationButtons()
        {
            btnPreviousControl.Enabled = centerResults.Count > 1;
            btnNextControl.Enabled = centerResults.Count > 1;

            if (dgvResultsControl.Rows.Count > currentImageIndex)
            {
                dgvResultsControl.ClearSelection();
                dgvResultsControl.Rows[currentImageIndex].Selected = true;
            }
        }

        private void CalculateAndDisplayStatistics(List<ObjectCenterResult> validResults)
        {
            int totalCount = centerResults.Count;
            int validCount = validResults.Count;

            var xOffsets = validResults.Select(r => r.Offset.X).ToList();
            var yOffsets = validResults.Select(r => r.Offset.Y).ToList();

            double avgX = xOffsets.Average();
            double avgY = yOffsets.Average();
            double stdX = CalculateStandardDeviation(xOffsets, avgX);
            double stdY = CalculateStandardDeviation(yOffsets, avgY);
            double maxDistance = validResults.Max(r => r.Distance);

            int medianX = GetMedian(xOffsets);
            int medianY = GetMedian(yOffsets);

            // 儲存統計結果供 UpdateImageInfo 使用
            _overallStdX = stdX;
            _overallStdY = stdY;
            _medianX = medianX;
            _medianY = medianY;

            // 計算整體穩定性等級（基於最大標準差）
            _maxStd = Math.Max(stdX, stdY);
            if (_maxStd < 1)
                _stabilityLevel = $"位置非常穩定 (std={_maxStd:F1}) - \n結果可信，直接套用";
            else if (_maxStd < 2)
                _stabilityLevel = $"位置正常 (std={_maxStd:F1}) - \n結果可信，直接套用";
            else if (_maxStd < 3)
                _stabilityLevel = $"樣品太少，或位置有些不穩定 (std={_maxStd:F1}) - 可以套用，建議增加樣品";
            else
                _stabilityLevel = $"樣品太少，或位置很不穩定 (std={_maxStd:F1}) - 可以套用，建議增加樣品";

            var statValues = new string[] {
                totalCount.ToString(),
                validCount.ToString(),
                avgX.ToString("F1"),
                avgY.ToString("F1"),
                stdX.ToString("F1"),
                stdY.ToString("F1"),
                maxDistance.ToString("F1"),
                medianX.ToString(),
                medianY.ToString(),
                _stabilityLevel
            };

            for (int i = 0; i < statValues.Length; i++)
            {
                var lblValue = grpStatisticsControl.Controls.Find($"lblStatValue{i}", false).FirstOrDefault() as Label;
                if (lblValue != null)
                {
                    lblValue.Text = statValues[i];
                }
            }

            lblRecommendationControl.Text = $"💡 建議設定：objBias_{targetStation}_X = {medianX}, objBias_{targetStation}_Y = {medianY}";
            btnApplyControl.Enabled = true;

            recommendedBiasX = medianX;
            recommendedBiasY = medianY;
        }

        private int GetMedian(List<int> values)
        {
            var sorted = values.OrderBy(x => x).ToList();
            int count = sorted.Count;

            if (count % 2 == 0)
            {
                return (sorted[count / 2 - 1] + sorted[count / 2]) / 2;
            }
            else
            {
                return sorted[count / 2];
            }
        }

        private double CalculateStandardDeviation(List<int> values, double mean)
        {
            if (values.Count <= 1) return 0;

            double sumSquaredDiffs = values.Sum(x => Math.Pow(x - mean, 2));
            return Math.Sqrt(sumSquaredDiffs / (values.Count - 1));
        }
        #endregion

        #region 輔助方法
        private void UpdateParameter(MydbDB db, string paramName, string value)
        {
            // 參數中文名稱
            string chineseName = paramName == "objBias_x" ? "物體X軸偏移補償" : "物體Y軸偏移補償";

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

        private void ShowFullScreenPreview(ObjectCenterResult result)
        {
            using (var fullScreenForm = new Form())
            {
                fullScreenForm.Text = $"檢測結果預覽 - {result.ImagePath}";
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

            foreach (var result in centerResults)
            {
                result.ResultImage?.Dispose();
            }

            picPreviewControl?.Image?.Dispose();
        }
        #endregion
    }
}