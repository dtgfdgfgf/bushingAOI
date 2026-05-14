// 由 GitHub Copilot 產生 - 對比度校正表單
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Runtime.InteropServices;
using LinqToDB;

namespace peilin
{
    public partial class ContrastCalibrationForm : Form
    {
        private Mat originalImage;
        private Mat processedImage;
        private string currentImagePath = "";
        private int selectedStation = 0;
        private int currentContrast = 0;
        private int currentBrightness = 0;
        private bool hasValidImage = false;
        private string targetType;

        // 由 GitHub Copilot 產生 - 修正預覽功能
        private float zoomFactor = 1.0f;
        private System.Drawing.Point imageOffset = System.Drawing.Point.Empty;
        private bool isDragging = false;
        private System.Drawing.Point dragStartPoint = System.Drawing.Point.Empty;
        private System.Drawing.Point lastImageOffset = System.Drawing.Point.Empty;
        private Button btnShowOriginal;

        // UI 控制項
        private ComboBox cmbStation;
        private PictureBox imagePreview;
        private Button btnLoadImage;
        private Button btnApplyProcessing;
        private Button btnApplyValues;
        private Button btnCancel;
        private TrackBar trackContrast;
        private TrackBar trackBrightness;
        private Label lblContrastValue;
        private Label lblBrightnessValue;
        private Label lblContrastDisplay;
        private Label lblBrightnessDisplay;
        private Label lblStatus;

        public ContrastCalibrationForm(string targetType)
        {
            // 由 GitHub Copilot 產生 - 初始化視窗
            this.targetType = targetType;
            InitializeFormProperties();
            SetupUI();
            LoadCurrentParameters();
        }

        private void InitializeFormProperties()
        {
            // 由 GitHub Copilot 產生 - 基本視窗屬性設定
            this.Text = "對比度校正";
            this.Size = new System.Drawing.Size(1600, 1000);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Microsoft JhengHei", 10F);
        }

        private void SetupUI()
        {
            // 由 GitHub Copilot 產生 - 建立所有UI控制項
            CreateStationAndControlArea();
            CreateImagePreviewArea();
            CreateParameterDisplayArea();
            CreateActionButtons();
        }

        private void CreateStationAndControlArea()
        {
            // 由 GitHub Copilot 產生 - 站點選擇與控制區域
            var grpStationAndControl = new GroupBox();
            grpStationAndControl.Text = "站點選擇與參數調整";
            grpStationAndControl.Location = new System.Drawing.Point(20, 20);
            grpStationAndControl.Size = new System.Drawing.Size(820, 120);
            grpStationAndControl.Font = new Font("Microsoft JhengHei", 10F);

            // 站點選擇
            var lblStation = new Label();
            lblStation.Text = "選擇站點：";
            lblStation.Location = new System.Drawing.Point(15, 30);
            lblStation.Size = new System.Drawing.Size(90, 22);
            lblStation.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);

            cmbStation = new ComboBox();
            cmbStation.Location = new System.Drawing.Point(110, 28);
            cmbStation.Size = new System.Drawing.Size(90, 25);
            cmbStation.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbStation.Font = new Font("Microsoft JhengHei", 10F);
            cmbStation.Items.AddRange(new object[] { "請選擇", "站點1", "站點2" });
            cmbStation.SelectedIndex = 0;
            selectedStation = 0;
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged;

            // 對比度調整
            var lblContrast = new Label();
            lblContrast.Text = "對比度(Contrast)：";
            lblContrast.Location = new System.Drawing.Point(215, 30);
            lblContrast.Size = new System.Drawing.Size(150, 22);
            lblContrast.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);

            trackContrast = new TrackBar();
            trackContrast.Location = new System.Drawing.Point(370, 25);
            trackContrast.Size = new System.Drawing.Size(280, 45);
            trackContrast.Minimum = -100;
            trackContrast.Maximum = 100;
            trackContrast.Value = 0;
            trackContrast.TickFrequency = 20;
            trackContrast.ValueChanged += TrackContrast_ValueChanged;

            lblContrastValue = new Label();
            lblContrastValue.Text = "0";
            lblContrastValue.Location = new System.Drawing.Point(655, 30);
            lblContrastValue.Size = new System.Drawing.Size(50, 22);
            lblContrastValue.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);
            lblContrastValue.ForeColor = System.Drawing.Color.Blue;

            // 亮度調整
            var lblBrightness = new Label();
            lblBrightness.Text = "亮度(Brightness)：";
            lblBrightness.Location = new System.Drawing.Point(215, 75);
            lblBrightness.Size = new System.Drawing.Size(150, 22);
            lblBrightness.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);

            trackBrightness = new TrackBar();
            trackBrightness.Location = new System.Drawing.Point(370, 70);
            trackBrightness.Size = new System.Drawing.Size(280, 45);
            trackBrightness.Minimum = -100;
            trackBrightness.Maximum = 100;
            trackBrightness.Value = 0;
            trackBrightness.TickFrequency = 20;
            trackBrightness.ValueChanged += TrackBrightness_ValueChanged;

            lblBrightnessValue = new Label();
            lblBrightnessValue.Text = "0";
            lblBrightnessValue.Location = new System.Drawing.Point(655, 75);
            lblBrightnessValue.Size = new System.Drawing.Size(50, 22);
            lblBrightnessValue.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);
            lblBrightnessValue.ForeColor = System.Drawing.Color.Green;

            // 導入按鈕
            btnApplyProcessing = new Button();
            btnApplyProcessing.Text = "導入";
            btnApplyProcessing.Location = new System.Drawing.Point(720, 45);
            btnApplyProcessing.Size = new System.Drawing.Size(80, 35);
            btnApplyProcessing.BackColor = System.Drawing.Color.LightBlue;
            btnApplyProcessing.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);
            btnApplyProcessing.Enabled = false;
            btnApplyProcessing.Click += BtnApplyProcessing_Click;

            grpStationAndControl.Controls.AddRange(new Control[] {
                lblStation, cmbStation,
                lblContrast, trackContrast, lblContrastValue,
                lblBrightness, trackBrightness, lblBrightnessValue,
                btnApplyProcessing
            });

            this.Controls.Add(grpStationAndControl);
        }

        private void CreateImagePreviewArea()
        {
            // 由 GitHub Copilot 產生 - 圖片預覽區域
            var grpImagePreview = new GroupBox();
            grpImagePreview.Text = "圖片預覽";
            grpImagePreview.Location = new System.Drawing.Point(20, 160);
            grpImagePreview.Size = new System.Drawing.Size(1200, 700);
            grpImagePreview.Font = new Font("Microsoft JhengHei", 10F);

            // 載入圖片按鈕
            btnLoadImage = new Button();
            btnLoadImage.Text = "載入圖片";
            btnLoadImage.Location = new System.Drawing.Point(15, 25);
            btnLoadImage.Size = new System.Drawing.Size(120, 35);
            btnLoadImage.BackColor = System.Drawing.Color.LightGreen;
            btnLoadImage.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);
            btnLoadImage.Click += BtnLoadImage_Click;

            // 顯示原圖按鈕
            btnShowOriginal = new Button();
            btnShowOriginal.Text = "顯示原圖";
            btnShowOriginal.Location = new System.Drawing.Point(145, 25);
            btnShowOriginal.Size = new System.Drawing.Size(120, 35);
            btnShowOriginal.BackColor = System.Drawing.Color.LightYellow;
            btnShowOriginal.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold);
            btnShowOriginal.Enabled = false;
            btnShowOriginal.MouseDown += BtnShowOriginal_MouseDown;
            btnShowOriginal.MouseUp += BtnShowOriginal_MouseUp;
            btnShowOriginal.MouseLeave += BtnShowOriginal_MouseLeave;

            // 圖片預覽控制項
            imagePreview = new PictureBox();
            imagePreview.Location = new System.Drawing.Point(15, 70);
            imagePreview.Size = new System.Drawing.Size(1170, 615);
            imagePreview.SizeMode = PictureBoxSizeMode.Normal;
            imagePreview.BorderStyle = BorderStyle.FixedSingle;
            imagePreview.BackColor = System.Drawing.Color.LightGray;
            imagePreview.Paint += ImagePreview_Paint;
            imagePreview.MouseWheel += ImagePreview_MouseWheel;
            imagePreview.MouseDown += ImagePreview_MouseDown;
            imagePreview.MouseMove += ImagePreview_MouseMove;
            imagePreview.MouseUp += ImagePreview_MouseUp;

            grpImagePreview.Controls.Add(btnLoadImage);
            grpImagePreview.Controls.Add(btnShowOriginal);
            grpImagePreview.Controls.Add(imagePreview);
            this.Controls.Add(grpImagePreview);
        }

        private void CreateParameterDisplayArea()
        {
            // 由 GitHub Copilot 產生 - 參數顯示區域
            var grpParameters = new GroupBox();
            grpParameters.Text = "當前參數值";
            grpParameters.Location = new System.Drawing.Point(1240, 160);
            grpParameters.Size = new System.Drawing.Size(330, 400);
            grpParameters.Font = new Font("Microsoft JhengHei", 10F);

            // 對比度顯示
            var lblContrastTitle = new Label();
            lblContrastTitle.Text = "對比度(Contrast)：";
            lblContrastTitle.Location = new System.Drawing.Point(15, 40);
            lblContrastTitle.Size = new System.Drawing.Size(165, 28);
            lblContrastTitle.Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold);

            lblContrastDisplay = new Label();
            lblContrastDisplay.Text = "0";
            lblContrastDisplay.Location = new System.Drawing.Point(185, 38);
            lblContrastDisplay.Size = new System.Drawing.Size(130, 32);
            lblContrastDisplay.Font = new Font("Microsoft JhengHei", 16F, FontStyle.Bold);
            lblContrastDisplay.ForeColor = System.Drawing.Color.Blue;

            // 亮度顯示
            var lblBrightnessTitle = new Label();
            lblBrightnessTitle.Text = "亮度(Brightness)：";
            lblBrightnessTitle.Location = new System.Drawing.Point(15, 95);
            lblBrightnessTitle.Size = new System.Drawing.Size(165, 28);
            lblBrightnessTitle.Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold);

            lblBrightnessDisplay = new Label();
            lblBrightnessDisplay.Text = "0";
            lblBrightnessDisplay.Location = new System.Drawing.Point(185, 93);
            lblBrightnessDisplay.Size = new System.Drawing.Size(130, 32);
            lblBrightnessDisplay.Font = new Font("Microsoft JhengHei", 16F, FontStyle.Bold);
            lblBrightnessDisplay.ForeColor = System.Drawing.Color.Green;

            // 狀態顯示
            lblStatus = new Label();
            lblStatus.Text = "請載入圖片並調整參數";
            lblStatus.Location = new System.Drawing.Point(15, 145);
            lblStatus.Size = new System.Drawing.Size(300, 80);
            lblStatus.Font = new Font("Microsoft JhengHei", 10F);
            lblStatus.ForeColor = System.Drawing.Color.Gray;

            grpParameters.Controls.AddRange(new Control[] {
                lblContrastTitle, lblContrastDisplay,
                lblBrightnessTitle, lblBrightnessDisplay,
                lblStatus
            });

            this.Controls.Add(grpParameters);
        }

        private void CreateActionButtons()
        {
            // 由 GitHub Copilot 產生 - 動作按鈕
            btnApplyValues = new Button();
            btnApplyValues.Text = "套用推薦值";
            btnApplyValues.Location = new System.Drawing.Point(1280, 900);
            btnApplyValues.Size = new System.Drawing.Size(130, 40);
            btnApplyValues.BackColor = System.Drawing.Color.LightGreen;
            btnApplyValues.Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold);
            btnApplyValues.Enabled = false;
            btnApplyValues.Click += BtnApplyValues_Click;

            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.Location = new System.Drawing.Point(1425, 900);
            btnCancel.Size = new System.Drawing.Size(110, 40);
            btnCancel.BackColor = System.Drawing.Color.LightCoral;
            btnCancel.Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold);
            btnCancel.Click += BtnCancel_Click;

            this.Controls.Add(btnApplyValues);
            this.Controls.Add(btnCancel);
        }

        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedStation = cmbStation.SelectedIndex; // 0 = 未選
            if (selectedStation <= 0)
            {
                UpdateStatus("請先選擇站點");
                btnApplyProcessing.Enabled = false;
                btnApplyValues.Enabled = false;
                return;
            }

            UpdateStatus($"已選擇站點 {selectedStation}");
            // 由 GitHub Copilot 產生 - 切換站點時恢復按鈕文字
            btnApplyValues.Text = "套用推薦值";
            LoadCurrentParameters(); // 只在有效站點時載入
        }

        private void TrackContrast_ValueChanged(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 對比度拖拉桿變更事件
            currentContrast = trackContrast.Value;
            lblContrastValue.Text = currentContrast.ToString();
            lblContrastDisplay.Text = currentContrast.ToString();
        }

        private void TrackBrightness_ValueChanged(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 亮度拖拉桿變更事件
            currentBrightness = trackBrightness.Value;
            lblBrightnessValue.Text = currentBrightness.ToString();
            lblBrightnessDisplay.Text = currentBrightness.ToString();
        }

        private void BtnLoadImage_Click(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 載入圖片
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "選擇校正圖片";
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp;*.tiff|所有檔案|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        LoadCalibrationImage(openFileDialog.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"載入圖片失敗：{ex.Message}", "錯誤",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void LoadCalibrationImage(string imagePath)
        {
            // 由 GitHub Copilot 產生 - 載入校正圖片
            try
            {
                // 由 GitHub Copilot 產生 - 先設為 false 防止 Paint 事件存取已丟棄物件
                hasValidImage = false;

                // 釋放舊圖片並設為 null
                if (originalImage != null)
                {
                    originalImage.Dispose();
                    originalImage = null;
                }
                if (processedImage != null)
                {
                    processedImage.Dispose();
                    processedImage = null;
                }

                // 載入新圖片
                originalImage = new Mat(imagePath);
                currentImagePath = imagePath;
                hasValidImage = true;

                // 重置縮放和偏移
                zoomFactor = 1.0f;
                imageOffset = System.Drawing.Point.Empty;

                // 啟用按鈕
                btnApplyProcessing.Enabled = true;
                btnShowOriginal.Enabled = true;

                // 重繪圖片
                imagePreview.Invalidate();

                UpdateStatus($"圖片載入成功：{System.IO.Path.GetFileName(imagePath)}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入圖片時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                hasValidImage = false;
                btnApplyProcessing.Enabled = false;
                btnShowOriginal.Enabled = false;
            }
        }

        private void BtnApplyProcessing_Click(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 套用對比度與亮度處理
            if (!hasValidImage || originalImage == null)
            {
                MessageBox.Show("請先載入圖片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 由 GitHub Copilot 產生 - 釋放舊的處理結果並設為 null
                if (processedImage != null)
                {
                    processedImage.Dispose();
                    processedImage = null;
                }

                // 由 GitHub Copilot 產生 - 用 using 包裝 grayImage 防止記憶體洩漏
                using (Mat grayImage = new Mat())
                {
                    Cv2.CvtColor(originalImage, grayImage, ColorConversionCodes.BGR2GRAY);

                    // 使用 Contrast 函數處理圖片
                    processedImage = Contrast(grayImage, currentContrast, currentBrightness);
                }

                // 轉回 BGR 以便顯示和保存
                if (!(app.param.ContainsKey($"color_{selectedStation}") && app.param[$"color_{selectedStation}"] == "1"))
                {
                    Cv2.CvtColor(processedImage, processedImage, ColorConversionCodes.GRAY2BGR);
                }

                // 啟用套用按鈕
                btnApplyValues.Enabled = true;

                // 重繪圖片預覽
                imagePreview.Invalidate();

                UpdateStatus($"參數已導入：對比度 {currentContrast}, 亮度 {currentBrightness}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"處理圖片時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private Mat Contrast(Mat src, int contrast, int brightness)
        {
            if (src == null)
                throw new ArgumentNullException(nameof(src), "輸入圖像不能為空。");
            if (contrast < -100 || contrast > 100)
                throw new ArgumentOutOfRangeException(nameof(contrast), "對比度應在 -100 到 100 之間。");
            if (brightness < -100 || brightness > 100)
                throw new ArgumentOutOfRangeException(nameof(brightness), "亮度應在 -100 到 100 之間。");

            // 创建查找表
            byte[] lookupTableData = new byte[256];
            double delta, a, b;
            int brightnessOffset = brightness;
            int contrastOffset = contrast;

            if (contrastOffset > 0)
            {
                delta = 127 * contrastOffset / 100.0;
                a = 255.0 / (255.0 - delta * 2);
                b = a * (brightnessOffset - delta);
            }
            else
            {
                delta = -128 * contrastOffset / 100.0;
                a = (256.0 - delta * 2) / 255.0;
                b = a * brightnessOffset + delta;
            }

            for (int i = 0; i < 256; i++)
            {
                int y = (int)(a * i + b);
                if (y < 0)
                    y = 0;
                if (y > 255)
                    y = 255;
                lookupTableData[i] = (byte)y;
            }

            // 创建查找表矩阵
            Mat lookupTableMatrix = new Mat(256, 1, MatType.CV_8UC1);

            // 将数据复制到矩阵中
            Marshal.Copy(lookupTableData, 0, lookupTableMatrix.Data, lookupTableData.Length);

            // 应用查找表
            Mat dst = new Mat();
            Cv2.LUT(src, lookupTableMatrix, dst);

            lookupTableMatrix.Dispose();

            
            Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
            Cv2.MorphologyEx(dst, dst, MorphTypes.Close, kernel);
            kernel.Dispose();
            

            return dst;
        }
        private void LoadCurrentParameters()
        {
            if (selectedStation <= 0) return;
            // 由 GitHub Copilot 產生 - 從資料庫載入當前站點的參數
            try
            {
                if (string.IsNullOrEmpty(targetType))
                {
                    UpdateStatus("目標料號未設定");
                    return;
                }

                using (var db = new MydbDB())
                {
                    // 載入對比度參數
                    var contrastParam = db.@params
                        .Where(p => p.Type == targetType && p.Name == "deepenContrast" && p.Stop == selectedStation)
                        .FirstOrDefault();

                    if (contrastParam != null && int.TryParse(contrastParam.Value, out int contrast))
                    {
                        currentContrast = contrast;
                        trackContrast.Value = Math.Max(-100, Math.Min(100, contrast));
                    }
                    else
                    {
                        currentContrast = 0;
                        trackContrast.Value = 0;
                    }

                    // 載入亮度參數
                    var brightnessParam = db.@params
                        .Where(p => p.Type == targetType && p.Name == "deepenBrightness" && p.Stop == selectedStation)
                        .FirstOrDefault();

                    if (brightnessParam != null && int.TryParse(brightnessParam.Value, out int brightness))
                    {
                        currentBrightness = brightness;
                        trackBrightness.Value = Math.Max(-100, Math.Min(100, brightness));
                    }
                    else
                    {
                        currentBrightness = 0;
                        trackBrightness.Value = 0;
                    }

                    // 更新顯示
                    lblContrastValue.Text = currentContrast.ToString();
                    lblBrightnessValue.Text = currentBrightness.ToString();
                    lblContrastDisplay.Text = currentContrast.ToString();
                    lblBrightnessDisplay.Text = currentBrightness.ToString();

                    UpdateStatus($"已載入站點 {selectedStation} 的參數");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"載入參數時發生錯誤：{ex.Message}");
            }
        }

        private void SaveParametersToDatabase()
        {
            if (selectedStation <= 0) throw new Exception("尚未選擇站點，無法儲存參數。");
            // 由 GitHub Copilot 產生 - 儲存參數到資料庫使用 Set().Update() 方式
            if (string.IsNullOrEmpty(targetType))
            {
                throw new Exception("目標料號未設定");
            }

            using (var db = new MydbDB())
            {
                // 儲存對比度參數
                int contrastUpdatedRows = db.@params
                    .Where(p => p.Type == targetType &&
                               p.Name == "deepenContrast" &&
                               p.Stop == selectedStation)
                    .Set(p => p.Value, currentContrast.ToString())
                    .Update();

                // 如果沒有更新任何記錄，表示記錄不存在，需要插入
                if (contrastUpdatedRows == 0)
                {
                    db.@params
                        .Value(p => p.Type, targetType)
                        .Value(p => p.Name, "deepenContrast")
                        .Value(p => p.Value, currentContrast.ToString())
                        .Value(p => p.Stop, selectedStation)
                        .Value(p => p.ChineseName, $"站點{selectedStation}對比度")
                        .Insert();
                }

                // 儲存亮度參數
                int brightnessUpdatedRows = db.@params
                    .Where(p => p.Type == targetType &&
                               p.Name == "deepenBrightness" &&
                               p.Stop == selectedStation)
                    .Set(p => p.Value, currentBrightness.ToString())
                    .Update();

                // 如果沒有更新任何記錄，表示記錄不存在，需要插入
                if (brightnessUpdatedRows == 0)
                {
                    db.@params
                        .Value(p => p.Type, targetType)
                        .Value(p => p.Name, "deepenBrightness")
                        .Value(p => p.Value, currentBrightness.ToString())
                        .Value(p => p.Stop, selectedStation)
                        .Value(p => p.ChineseName, $"站點{selectedStation}亮度")
                        .Insert();
                }
            }
        }

        private void UpdateStatus(string message)
        {
            // 由 GitHub Copilot 產生 - 更新狀態顯示
            lblStatus.Text = message;
            lblStatus.ForeColor = System.Drawing.Color.DarkBlue;
        }

        private void BtnApplyValues_Click(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 套用參數到資料庫但保持視窗開啟
            try
            {
                SaveParametersToDatabase();
                MessageBox.Show($"對比度校正參數已成功儲存\n站點：{selectedStation}\n對比度：{currentContrast}\n亮度：{currentBrightness}",
                    "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 不關閉視窗，只更新狀態
                UpdateStatus($"參數已儲存 - 站點{selectedStation} 對比度:{currentContrast} 亮度:{currentBrightness}");
                // 由 GitHub Copilot 產生 - 套用成功後更新按鈕文字
                btnApplyValues.Text = "✅ 已套用";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 取消操作
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 清理資源並設為 null
            hasValidImage = false;
            if (originalImage != null)
            {
                originalImage.Dispose();
                originalImage = null;
            }
            if (processedImage != null)
            {
                processedImage.Dispose();
                processedImage = null;
            }
            if (imagePreview.Image != null)
            {
                imagePreview.Image.Dispose();
                imagePreview.Image = null;
            }
            base.OnFormClosing(e);
        }
        private void BtnShowOriginal_MouseDown(object sender, MouseEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 按下顯示原圖按鈕
            if (hasValidImage)
            {
                btnShowOriginal.BackColor = System.Drawing.Color.Orange;
                imagePreview.Invalidate();
            }
        }

        private void BtnShowOriginal_MouseUp(object sender, MouseEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 放開顯示原圖按鈕
            if (hasValidImage)
            {
                btnShowOriginal.BackColor = System.Drawing.Color.LightYellow;
                imagePreview.Invalidate();
            }
        }

        private void BtnShowOriginal_MouseLeave(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 滑鼠離開顯示原圖按鈕
            if (hasValidImage)
            {
                btnShowOriginal.BackColor = System.Drawing.Color.LightYellow;
                imagePreview.Invalidate();
            }
        }
        private void ImagePreview_Paint(object sender, PaintEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 自訂繪製圖片預覽
            // 由 GitHub Copilot 產生 - 加入 IsDisposed 檢查防止存取已丟棄物件
            if (!hasValidImage || originalImage == null || originalImage.IsDisposed)
                return;

            try
            {
                // 決定要顯示的圖片
                Mat displayImage;
                bool showingOriginal = btnShowOriginal.BackColor == System.Drawing.Color.Orange;

                if (showingOriginal || processedImage == null || processedImage.IsDisposed)
                {
                    displayImage = originalImage;
                }
                else
                {
                    displayImage = processedImage;
                }

                // 由 GitHub Copilot 產生 - 再次檢查 displayImage 是否有效
                if (displayImage == null || displayImage.IsDisposed)
                    return;

                // 轉換為 Bitmap
                var bitmap = displayImage.ToBitmap();

                // 計算縮放後的尺寸
                int scaledWidth = (int)(bitmap.Width * zoomFactor);
                int scaledHeight = (int)(bitmap.Height * zoomFactor);

                // 繪製圖片
                System.Drawing.Rectangle destRect = new System.Drawing.Rectangle(
                    imageOffset.X, imageOffset.Y, scaledWidth, scaledHeight);

                e.Graphics.DrawImage(bitmap, destRect);
                bitmap.Dispose();
            }
            catch (Exception ex)
            {
                UpdateStatus($"繪製圖片時發生錯誤：{ex.Message}");
            }
        }

        private void ImagePreview_MouseWheel(object sender, MouseEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 滑鼠滾輪縮放功能
            if (!hasValidImage || (Control.ModifierKeys & Keys.Control) == 0)
                return;

            float oldZoomFactor = zoomFactor;

            // 計算縮放倍率
            if (e.Delta > 0)
                zoomFactor *= 1.1f;
            else
                zoomFactor /= 1.1f;

            // 限制縮放範圍
            zoomFactor = Math.Max(0.1f, Math.Min(10.0f, zoomFactor));

            // 計算縮放中心點的偏移調整
            var mousePos = imagePreview.PointToClient(Cursor.Position);
            float scaleChange = zoomFactor / oldZoomFactor;

            imageOffset.X = (int)(mousePos.X - (mousePos.X - imageOffset.X) * scaleChange);
            imageOffset.Y = (int)(mousePos.Y - (mousePos.Y - imageOffset.Y) * scaleChange);

            imagePreview.Invalidate();
        }

        private void ImagePreview_MouseDown(object sender, MouseEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 開始拖曳
            if (e.Button == MouseButtons.Left && hasValidImage)
            {
                isDragging = true;
                dragStartPoint = e.Location;
                lastImageOffset = imageOffset;
                imagePreview.Cursor = Cursors.Hand;
            }
        }

        private void ImagePreview_MouseMove(object sender, MouseEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 拖曳移動視野
            if (isDragging && hasValidImage)
            {
                int deltaX = e.X - dragStartPoint.X;
                int deltaY = e.Y - dragStartPoint.Y;

                imageOffset = new System.Drawing.Point(
                    lastImageOffset.X + deltaX,
                    lastImageOffset.Y + deltaY);

                imagePreview.Invalidate();
            }
        }

        private void ImagePreview_MouseUp(object sender, MouseEventArgs e)
        {
            // 由 GitHub Copilot 產生 - 結束拖曳
            if (e.Button == MouseButtons.Left)
            {
                isDragging = false;
                imagePreview.Cursor = Cursors.Default;
            }
        }
    }
}