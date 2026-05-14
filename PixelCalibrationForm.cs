// 由 GitHub Copilot 產生
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
using LinqToDB;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace peilin
{
    public partial class PixelCalibrationForm : Form
    {
        // 影像與量測相關
        private Mat calibrationImage;                  // OpenCV 原圖 (須記得 Dispose)
        private string currentImagePath = "";
        private double currentOD = 0;
        private int selectedStation = 0;
        private double currentLineLength = 0;          // 目前選取線段像素長度
        private double currentDetectionPrecision = 0;  // PixelToMM (mm / pixel)
        private bool hasValidParameters = false;

        // UI 控制項
        private ComboBox cmbStation;
        private PictureBox imagePreview;
        private Button btnLoadImage;
        private Button btnApplyValues;
        private Button btnCancel;
        private Label lblLineLength;
        private Label lblOD;
        private Label lblDetectionPrecision;
        private Label lblStatus;

        // 線段繪製狀態（直接整合自原 FullscreenPreviewForm）
        private bool isDrawingLine = false;
        private bool lineCompleted = false;
        private bool isDraggingStart = false;
        private bool isDraggingEnd = false;
        private bool isDraggingLine = false;
        private System.Drawing.PointF lineStartRelative = System.Drawing.PointF.Empty; // 0~1
        private System.Drawing.PointF lineEndRelative = System.Drawing.PointF.Empty;   // 0~1
        private System.Drawing.PointF dragOffsetRelative;
        private bool isShiftPressed = false;
        private System.Drawing.Point currentMousePoint;

        public PixelCalibrationForm()
        {
            // 由 GitHub Copilot 產生 - 初始化
            InitializeFormProperties();
            SetupUI();
            LoadCurrentTypeOD();
            this.KeyPreview = true;
            this.KeyDown += PixelCalibrationForm_KeyDown;
        }

        // 由 GitHub Copilot 產生 - 視窗屬性
        private void InitializeFormProperties()
        {
            // 由 GitHub Copilot 產生 - 設定主視窗大小與屬性
            this.Text = "檢測精度校正";
            this.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);  // 統一字體
            this.Size = new System.Drawing.Size(1500, 1050);    // 螢幕適配
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        // 由 GitHub Copilot 產生 - 建立 UI
        private void SetupUI()
        {
            CreateImagePreviewArea();
            CreateRightPanel(); // 右側合併站點/數值/操作
        }

        // 由 GitHub Copilot 產生 - 左側圖片區
        private void CreateImagePreviewArea()
        {
            var grpImagePreview = new GroupBox();
            grpImagePreview.Text = "圖片預覽（左鍵畫線 / Shift直線 / 拖端點 / 拖線 / Delete刪除）";
            grpImagePreview.Location = new System.Drawing.Point(15, 15);
            grpImagePreview.Size = new System.Drawing.Size(1200, 1000);

            imagePreview = new PictureBox();
            imagePreview.Location = new System.Drawing.Point(15, 30);
            imagePreview.Size = new System.Drawing.Size(1170, 955);
            imagePreview.SizeMode = PictureBoxSizeMode.Zoom;
            imagePreview.BorderStyle = BorderStyle.FixedSingle;
            imagePreview.BackColor = System.Drawing.Color.LightGray;

            // 滑鼠事件（線段繪製）
            imagePreview.MouseDown += ImagePreview_MouseDown;
            imagePreview.MouseMove += ImagePreview_MouseMove;
            imagePreview.MouseUp += ImagePreview_MouseUp;
            imagePreview.Paint += ImagePreview_Paint;

            grpImagePreview.Controls.Add(imagePreview);
            this.Controls.Add(grpImagePreview);
        }

        // 由 GitHub Copilot 產生 - 右側站點/數值/操作
        private void CreateRightPanel()
        {
            var grpRight = new GroupBox();
            grpRight.Text = "校正參數";
            grpRight.Location = new System.Drawing.Point(1230, 15);
            grpRight.Size = new System.Drawing.Size(250, 1000);

            // 載入圖片按鈕
            btnLoadImage = new Button();
            btnLoadImage.Text = "載入圖片";
            btnLoadImage.Location = new System.Drawing.Point(40, 35);
            btnLoadImage.Size = new System.Drawing.Size(160, 35);
            btnLoadImage.BackColor = System.Drawing.Color.LightBlue;
            btnLoadImage.Click += BtnLoadImage_Click;

            // 站點
            var lblStation = new Label();
            lblStation.Text = "站點：";
            lblStation.Location = new System.Drawing.Point(20, 95);
            lblStation.Size = new System.Drawing.Size(55, 20);

            cmbStation = new ComboBox();
            cmbStation.Location = new System.Drawing.Point(80, 92);
            cmbStation.Size = new System.Drawing.Size(130, 24);
            cmbStation.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbStation.Items.AddRange(new object[] { "請選擇", "站點1", "站點2", "站點3", "站點4" });
            cmbStation.SelectedIndex = 0;
            selectedStation = 0;
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged;

            // 像素長度
            var lblLineLenTitle = new Label();
            lblLineLenTitle.Text = "線段像素：";
            lblLineLenTitle.Location = new System.Drawing.Point(20, 140);
            lblLineLenTitle.Size = new System.Drawing.Size(80, 20);

            lblLineLength = new Label();
            lblLineLength.Text = "0.0000";
            lblLineLength.Location = new System.Drawing.Point(105, 138);
            lblLineLength.Size = new System.Drawing.Size(110, 22);
            lblLineLength.Font = new System.Drawing.Font("Microsoft JhengHei", 10F, System.Drawing.FontStyle.Bold);
            lblLineLength.ForeColor = System.Drawing.Color.Blue;

            // OD
            var lblODTitle = new Label();
            lblODTitle.Text = "OD(mm)：";
            lblODTitle.Location = new System.Drawing.Point(20, 180);
            lblODTitle.Size = new System.Drawing.Size(80, 20);

            lblOD = new Label();
            lblOD.Text = "0.00";
            lblOD.Location = new System.Drawing.Point(105, 178);
            lblOD.Size = new System.Drawing.Size(110, 22);
            lblOD.Font = new System.Drawing.Font("Microsoft JhengHei", 10F, System.Drawing.FontStyle.Bold);
            lblOD.ForeColor = System.Drawing.Color.DarkGreen;

            // 檢測精度
            var lblDpTitle = new Label();
            lblDpTitle.Text = "PixelToMM：";
            lblDpTitle.Location = new System.Drawing.Point(20, 220);
            lblDpTitle.Size = new System.Drawing.Size(85, 20);

            lblDetectionPrecision = new Label();
            lblDetectionPrecision.Text = "0.0000";
            lblDetectionPrecision.Location = new System.Drawing.Point(105, 218);
            lblDetectionPrecision.Size = new System.Drawing.Size(110, 22);
            lblDetectionPrecision.Font = new System.Drawing.Font("Microsoft JhengHei", 10F, System.Drawing.FontStyle.Bold);
            lblDetectionPrecision.ForeColor = System.Drawing.Color.Red;

            // 狀態
            lblStatus = new Label();
            lblStatus.Text = "請載入圖片並畫線";
            lblStatus.Location = new System.Drawing.Point(20, 260);
            lblStatus.Size = new System.Drawing.Size(210, 60);
            lblStatus.ForeColor = System.Drawing.Color.Gray;

            // 套用按鈕
            btnApplyValues = new Button();
            btnApplyValues.Text = "套用推薦值";
            btnApplyValues.Location = new System.Drawing.Point(25, 340);
            btnApplyValues.Size = new System.Drawing.Size(90, 35);
            btnApplyValues.BackColor = System.Drawing.Color.LightGreen;
            btnApplyValues.Enabled = false;
            btnApplyValues.Click += BtnApplyValues_Click;

            // 取消按鈕
            btnCancel = new Button();
            btnCancel.Text = "關閉";
            btnCancel.Location = new System.Drawing.Point(130, 340);
            btnCancel.Size = new System.Drawing.Size(90, 35);
            btnCancel.BackColor = System.Drawing.Color.LightCoral;
            btnCancel.Click += BtnCancel_Click;

            grpRight.Controls.AddRange(new Control[] {
            btnLoadImage,
            lblStation, cmbStation,
            lblLineLenTitle, lblLineLength,
            lblODTitle, lblOD,
            lblDpTitle, lblDetectionPrecision,
            lblStatus,
            btnApplyValues, btnCancel
        });

            this.Controls.Add(grpRight);
        }

        // 由 GitHub Copilot 產生 - 站點切換事件
        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedStation = cmbStation.SelectedIndex;
            if (selectedStation <= 0)
            {
                UpdateStatus("請先選擇站點");
                btnApplyValues.Enabled = false;
                return;
            }
            UpdateStatus($"已選擇站點 {selectedStation}");
            // 由 GitHub Copilot 產生 - 切換站點時恢復按鈕文字
            btnApplyValues.Text = "套用推薦值";
            // 若已有量測結果才允許套用
            if (currentDetectionPrecision > 0)
                btnApplyValues.Enabled = true;
        }

        // 由 GitHub Copilot 產生 - 載入圖片
        private void BtnLoadImage_Click(object sender, EventArgs e)
        {
            using (var fileDialog = new CommonOpenFileDialog())
            {
                fileDialog.Title = "選擇校正圖片";
                fileDialog.Filters.Add(new CommonFileDialogFilter("圖片檔案", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff"));
                fileDialog.EnsureFileExists = true;
                fileDialog.Multiselect = false;

                if (fileDialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    try
                    {
                        LoadCalibrationImage(fileDialog.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"載入圖片失敗：{ex.Message}", "錯誤",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // 由 GitHub Copilot 產生 - 實際載入影像
        private void LoadCalibrationImage(string path)
        {
            try
            {
                calibrationImage?.Dispose();
                // 釋放舊 bitmap
                if (imagePreview.Image != null)
                {
                    var old = imagePreview.Image;
                    imagePreview.Image = null;
                    old.Dispose();
                }

                calibrationImage = new Mat(path);
                currentImagePath = path;

                using (var bmp = calibrationImage.ToBitmap())
                {
                    // 建立一份新的 Bitmap (避免 using 釋放)
                    imagePreview.Image = new System.Drawing.Bitmap(bmp);
                }

                hasValidParameters = true;
                ResetValues();
                ClearLine();
                UpdateStatus($"圖片載入成功：{System.IO.Path.GetFileName(path)}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入圖片時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 由 GitHub Copilot 產生 - 讀取 OD
        private void LoadCurrentTypeOD()
        {
            try
            {
                if (string.IsNullOrEmpty(app.produce_No))
                {
                    UpdateStatus("無法取得當前料號");
                    lblOD.Text = "未設定";
                    return;
                }

                using (var db = new MydbDB())
                {
                    var info = db.Types
                        .Where(t => t.TypeColumn == app.produce_No)
                        .FirstOrDefault();

                    if (info != null && info.OD > 0)
                    {
                        currentOD = info.OD;
                        lblOD.Text = currentOD.ToString("F2");
                        UpdateStatus($"料號 {app.produce_No} OD={currentOD:F2}mm");
                    }
                    else
                    {
                        currentOD = 0;
                        lblOD.Text = "未設定";
                        UpdateStatus("OD 未設定");
                    }
                }
            }
            catch (Exception ex)
            {
                currentOD = 0;
                lblOD.Text = "錯誤";
                UpdateStatus($"載入OD錯誤：{ex.Message}");
            }
        }

        // 由 GitHub Copilot 產生 - Upsert PixelToMM
        private void SaveDetectionPrecisionToDatabase()
        {
            if (selectedStation <= 0) throw new Exception("尚未選擇站點，無法儲存 PixelToMM。");
            if (string.IsNullOrEmpty(app.produce_No))
                throw new Exception("無法取得當前料號");

            string parameterName = "PixelToMM";

            using (var db = new MydbDB())
            {
                try
                {
                    int updatedRows = db.@params
                        .Where(p => p.Type == app.produce_No &&
                                    p.Name == parameterName &&
                                    p.Stop == selectedStation)
                        .Set(p => p.Value, currentDetectionPrecision.ToString("F4"))
                        .Set(p => p.ChineseName, $"站點{selectedStation}檢測精度")
                        .Update();

                    if (updatedRows == 0)
                    {
                        db.@params
                          .Value(p => p.Type, app.produce_No)
                          .Value(p => p.Name, parameterName)
                          .Value(p => p.Value, currentDetectionPrecision.ToString("F4"))
                          .Value(p => p.Stop, selectedStation)
                          .Value(p => p.ChineseName, $"站點{selectedStation}檢測精度")
                          .Insert();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"寫入 PixelToMM 失敗: {ex.Message}");
                }
            }
        }

        // 由 GitHub Copilot 產生 - 更新線長 (供內部呼叫)
        private void InternalUpdateLineLength(double pixels)
        {
            currentLineLength = pixels;
            lblLineLength.Text = pixels.ToString("F4");
            CalculateDetectionPrecision();
        }

        // 保留舊介面可能外部呼叫 (若無可刪)
        public void UpdateLineLength(double pixels) => InternalUpdateLineLength(pixels);

        // 由 GitHub Copilot 產生 - 計算 PixelToMM
        private void CalculateDetectionPrecision()
        {
            if (currentLineLength > 0 && currentOD > 0)
            {
                currentDetectionPrecision = currentOD / currentLineLength;
                lblDetectionPrecision.Text = currentDetectionPrecision.ToString("F4");
                btnApplyValues.Enabled = true;
                UpdateStatus($"計算完成 PixelToMM={currentDetectionPrecision:F4}");
            }
            else
            {
                currentDetectionPrecision = 0;
                lblDetectionPrecision.Text = "0.0000";
                btnApplyValues.Enabled = false;
            }
        }

        // 由 GitHub Copilot 產生 - 重置數值
        private void ResetValues()
        {
            currentLineLength = 0;
            currentDetectionPrecision = 0;
            lblLineLength.Text = "0.0000";
            lblDetectionPrecision.Text = "0.0000";
            btnApplyValues.Enabled = false;
        }

        // 由 GitHub Copilot 產生 - 更新狀態
        private void UpdateStatus(string msg)
        {
            lblStatus.Text = msg;
            lblStatus.ForeColor = System.Drawing.Color.DarkBlue;
        }

        // 由 GitHub Copilot 產生 - 套用
        private void BtnApplyValues_Click(object sender, EventArgs e)
        {
            if (selectedStation <= 0)
            {
                MessageBox.Show("請先選擇站點", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (currentDetectionPrecision <= 0)
            {
                MessageBox.Show("請先完成線段量測", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                SaveDetectionPrecisionToDatabase();
                MessageBox.Show(
                    $"儲存成功\n站點: {selectedStation}\nPixelToMM: {currentDetectionPrecision:F4}",
                    "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                btnApplyValues.Text = "✅ 已套用";
                btnApplyValues.BackColor = System.Drawing.Color.LightGreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 由 GitHub Copilot 產生 - 關閉
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // ========================= 線段繪製功能 (整合) =========================

        // 由 GitHub Copilot 產生 - MouseDown
        private void ImagePreview_MouseDown(object sender, MouseEventArgs e)
        {
            if (imagePreview.Image == null || e.Button != MouseButtons.Left) return;

            isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            var rel = ScreenToImageRelative(e.Location);
            if (lineCompleted)
            {
                // 是否點在線段上
                if (IsNearPointRelative(rel, lineStartRelative, 0.02f))
                {
                    isDraggingStart = true;
                }
                else if (IsNearPointRelative(rel, lineEndRelative, 0.02f))
                {
                    isDraggingEnd = true;
                }
                else if (IsOnLineRelative(rel, lineStartRelative, lineEndRelative, 0.02f))
                {
                    isDraggingLine = true;
                    dragOffsetRelative = new System.Drawing.PointF(
                        rel.X - lineStartRelative.X,
                        rel.Y - lineStartRelative.Y
                    );
                }
                else
                {
                    StartNewLine(e.Location);
                }
            }
            else
            {
                StartNewLine(e.Location);
            }
        }

        // 由 GitHub Copilot 產生 - MouseMove
        private void ImagePreview_MouseMove(object sender, MouseEventArgs e)
        {
            currentMousePoint = e.Location;
            if (imagePreview.Image == null) return;

            var rel = ScreenToImageRelative(e.Location);
            bool shiftNow = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            if (isDrawingLine)
            {
                if (isShiftPressed || shiftNow)
                {
                    lineEndRelative = GetConstrainedLineEndPointRelative(lineStartRelative, rel);
                }
                else
                {
                    lineEndRelative = ClampPointToImageBoundsRelative(rel);
                }
                imagePreview.Invalidate();
                UpdateLineLengthFromRelative();
            }
            else if (isDraggingStart)
            {
                if (shiftNow)
                    lineStartRelative = GetConstrainedLineStartPointRelative(lineEndRelative, rel);
                else
                    lineStartRelative = ClampPointToImageBoundsRelative(rel);
                imagePreview.Invalidate();
                UpdateLineLengthFromRelative();
            }
            else if (isDraggingEnd)
            {
                if (shiftNow)
                    lineEndRelative = GetConstrainedLineEndPointRelative(lineStartRelative, rel);
                else
                    lineEndRelative = ClampPointToImageBoundsRelative(rel);
                imagePreview.Invalidate();
                UpdateLineLengthFromRelative();
            }
            else if (isDraggingLine)
            {
                var vec = new System.Drawing.PointF(
                    lineEndRelative.X - lineStartRelative.X,
                    lineEndRelative.Y - lineStartRelative.Y);
                lineStartRelative = new System.Drawing.PointF(rel.X - dragOffsetRelative.X, rel.Y - dragOffsetRelative.Y);
                lineEndRelative = new System.Drawing.PointF(lineStartRelative.X + vec.X, lineStartRelative.Y + vec.Y);
                ClampLineToBounds();
                imagePreview.Invalidate();
                UpdateLineLengthFromRelative();
            }
            else
            {
                UpdateCursor(rel);
            }
        }

        // 由 GitHub Copilot 產生 - MouseUp
        private void ImagePreview_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;

            if (isDrawingLine)
            {
                isDrawingLine = false;
                lineCompleted = true;
                UpdateLineLengthFromRelative();
            }
            isDraggingStart = false;
            isDraggingEnd = false;
            isDraggingLine = false;
            imagePreview.Cursor = Cursors.Default;
        }

        // 由 GitHub Copilot 產生 - Paint
        private void ImagePreview_Paint(object sender, PaintEventArgs e)
        {
            if (imagePreview.Image == null) return;
            if (!(isDrawingLine || lineCompleted)) return;
            if (lineStartRelative.IsEmpty) return;

            var startScreen = ImageRelativeToScreen(lineStartRelative);
            System.Drawing.Point endScreen;

            if (isDrawingLine)
            {
                var curRel = ScreenToImageRelative(currentMousePoint);
                if (isShiftPressed || (Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                    curRel = GetConstrainedLineEndPointRelative(lineStartRelative, curRel);
                endScreen = ImageRelativeToScreen(curRel);
            }
            else
            {
                endScreen = ImageRelativeToScreen(lineEndRelative);
            }

            using (var pen = new Pen(System.Drawing.Color.Red, 3))
            {
                e.Graphics.DrawLine(pen, startScreen, endScreen);
            }

            if (lineCompleted)
            {
                using (var brush = new SolidBrush(System.Drawing.Color.Yellow))
                using (var brd = new Pen(System.Drawing.Color.Black, 1))
                {
                    DrawHandle(e.Graphics, brush, brd, startScreen);
                    DrawHandle(e.Graphics, brush, brd, endScreen);
                }
            }
        }

        // 由 GitHub Copilot 產生 - 畫端點
        private void DrawHandle(Graphics g, SolidBrush b, Pen p, System.Drawing.Point pt)
        {
            int r = 4;
            g.FillEllipse(b, pt.X - r, pt.Y - r, r * 2, r * 2);
            g.DrawEllipse(p, pt.X - r, pt.Y - r, r * 2, r * 2);
        }

        // 由 GitHub Copilot 產生 - 開始新線
        private void StartNewLine(System.Drawing.Point screenPoint)
        {
            lineStartRelative = ScreenToImageRelative(screenPoint);
            lineEndRelative = lineStartRelative;
            isDrawingLine = true;
            lineCompleted = false;
            isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            imagePreview.Cursor = Cursors.Cross;
        }

        // 由 GitHub Copilot 產生 - 更新游標
        private void UpdateCursor(System.Drawing.PointF rel)
        {
            if (lineCompleted)
            {
                if (IsNearPointRelative(rel, lineStartRelative, 0.02f) ||
                    IsNearPointRelative(rel, lineEndRelative, 0.02f))
                    imagePreview.Cursor = Cursors.Cross;
                else if (IsOnLineRelative(rel, lineStartRelative, lineEndRelative, 0.02f))
                    imagePreview.Cursor = Cursors.SizeAll;
                else
                    imagePreview.Cursor = Cursors.Default;
            }
            else
            {
                imagePreview.Cursor = Cursors.Default;
            }
        }

        // 由 GitHub Copilot 產生 - 刪除線
        private void DeleteCurrentLine()
        {
            if (!lineCompleted && !isDrawingLine) return;
            ClearLine();
            imagePreview.Invalidate();
            UpdateStatus("線段已刪除");
        }

        // 由 GitHub Copilot 產生 - 清空線
        private void ClearLine()
        {
            isDrawingLine = false;
            lineCompleted = false;
            isDraggingStart = false;
            isDraggingEnd = false;
            isDraggingLine = false;
            lineStartRelative = System.Drawing.PointF.Empty;
            lineEndRelative = System.Drawing.PointF.Empty;
            InternalUpdateLineLength(0);
            imagePreview.Cursor = Cursors.Default;
        }

        // 由 GitHub Copilot 產生 - 鍵盤事件
        private void PixelCalibrationForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                DeleteCurrentLine();
            }
        }

        // 由 GitHub Copilot 產生 - 計算線長
        private void UpdateLineLengthFromRelative()
        {
            if (lineStartRelative.IsEmpty || lineEndRelative.IsEmpty) return;
            var a = ImageRelativeToAbsolute(lineStartRelative);
            var b = ImageRelativeToAbsolute(lineEndRelative);
            double len = Math.Sqrt(Math.Pow(b.X - a.X, 2) + Math.Pow(b.Y - a.Y, 2));
            InternalUpdateLineLength(len);
        }

        // 由 GitHub Copilot 產生 - 相對座標轉螢幕
        private System.Drawing.Point ImageRelativeToScreen(System.Drawing.PointF rel)
        {
            if (imagePreview.Image == null) return System.Drawing.Point.Empty;
            var rect = GetDisplayedImageRect();
            int x = (int)(rel.X * rect.Width + rect.X);
            int y = (int)(rel.Y * rect.Height + rect.Y);
            return new System.Drawing.Point(x, y);
        }

        // 由 GitHub Copilot 產生 - 螢幕座標轉相對(0~1)
        private System.Drawing.PointF ScreenToImageRelative(System.Drawing.Point pt)
        {
            if (imagePreview.Image == null) return System.Drawing.PointF.Empty;
            var rect = GetDisplayedImageRect();
            if (rect.Width <= 0 || rect.Height <= 0) return System.Drawing.PointF.Empty;

            float rx = (float)(pt.X - rect.X) / rect.Width;
            float ry = (float)(pt.Y - rect.Y) / rect.Height;
            rx = Math.Max(0, Math.Min(1, rx));
            ry = Math.Max(0, Math.Min(1, ry));
            return new System.Drawing.PointF(rx, ry);
        }

        // 由 GitHub Copilot 產生 - 相對轉原圖像素
        private System.Drawing.Point ImageRelativeToAbsolute(System.Drawing.PointF rel)
        {
            if (imagePreview.Image == null) return System.Drawing.Point.Empty;
            int ax = (int)(rel.X * imagePreview.Image.Width);
            int ay = (int)(rel.Y * imagePreview.Image.Height);
            return new System.Drawing.Point(ax, ay);
        }

        // 由 GitHub Copilot 產生 - 取得實際顯示矩形 (SizeMode.Zoom)
        private System.Drawing.Rectangle GetDisplayedImageRect()
        {
            if (imagePreview.Image == null) return System.Drawing.Rectangle.Empty;

            var img = imagePreview.Image;
            double imgW = img.Width;
            double imgH = img.Height;
            double boxW = imagePreview.ClientSize.Width;
            double boxH = imagePreview.ClientSize.Height;

            double ratio = Math.Min(boxW / imgW, boxH / imgH);
            int drawW = (int)(imgW * ratio);
            int drawH = (int)(imgH * ratio);
            int offsetX = (int)((boxW - drawW) / 2);
            int offsetY = (int)((boxH - drawH) / 2);
            return new System.Drawing.Rectangle(offsetX + imagePreview.ClientRectangle.X,
                                                offsetY + imagePreview.ClientRectangle.Y,
                                                drawW, drawH);
        }

        // 由 GitHub Copilot 產生 - 約束線終點（水平或垂直）
        private System.Drawing.PointF GetConstrainedLineEndPointRelative(System.Drawing.PointF start, System.Drawing.PointF end)
        {
            float dx = Math.Abs(end.X - start.X);
            float dy = Math.Abs(end.Y - start.Y);
            if (dx >= dy) return new System.Drawing.PointF(end.X, start.Y);
            return new System.Drawing.PointF(start.X, end.Y);
        }

        // 由 GitHub Copilot 產生 - 約束線起點（拖動另一端）
        private System.Drawing.PointF GetConstrainedLineStartPointRelative(System.Drawing.PointF other, System.Drawing.PointF moving)
        {
            float dx = Math.Abs(moving.X - other.X);
            float dy = Math.Abs(moving.Y - other.Y);
            if (dx >= dy) return new System.Drawing.PointF(moving.X, other.Y);
            return new System.Drawing.PointF(other.X, moving.Y);
        }

        // 由 GitHub Copilot 產生 - 點在線段上判斷
        private bool IsOnLineRelative(System.Drawing.PointF p, System.Drawing.PointF a, System.Drawing.PointF b, float threshold)
        {
            float len = (float)Math.Sqrt(Math.Pow(b.X - a.X, 2) + Math.Pow(b.Y - a.Y, 2));
            if (len < 0.0001f) return false;
            float dist = Math.Abs((b.Y - a.Y) * p.X - (b.X - a.X) * p.Y + b.X * a.Y - b.Y * a.X) / len;
            if (dist > threshold) return false;
            float dot = ((p.X - a.X) * (b.X - a.X) + (p.Y - a.Y) * (b.Y - a.Y)) / (len * len);
            return dot >= 0 && dot <= 1;
        }

        // 由 GitHub Copilot 產生 - 點接近判斷
        private bool IsNearPointRelative(System.Drawing.PointF p1, System.Drawing.PointF p2, float threshold)
        {
            float d = (float)Math.Sqrt(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2));
            return d <= threshold;
        }

        // 由 GitHub Copilot 產生 - 線段限制在範圍
        private void ClampLineToBounds()
        {
            lineStartRelative = ClampPointToImageBoundsRelative(lineStartRelative);
            lineEndRelative = ClampPointToImageBoundsRelative(lineEndRelative);
        }

        // 由 GitHub Copilot 產生 - 限制點
        private System.Drawing.PointF ClampPointToImageBoundsRelative(System.Drawing.PointF p)
        {
            return new System.Drawing.PointF(
                Math.Max(0f, Math.Min(1f, p.X)),
                Math.Max(0f, Math.Min(1f, p.Y))
            );
        }

        // ========================= 線段繪製功能 結束 =========================

        // 由 GitHub Copilot 產生 - 套用存檔
        private void BtnApplyValues_Click(object sender, EventArgs e, bool dummy = false) { } // (留空避免 IntelliSense 混淆)

        // 由 GitHub Copilot 產生 - 真正事件已在上方定義 (保留結構)

        // 由 GitHub Copilot 產生 - 關閉釋放資源
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            calibrationImage?.Dispose();
            if (imagePreview.Image != null)
            {
                var tmp = imagePreview.Image;
                imagePreview.Image = null;
                tmp.Dispose();
            }
            base.OnFormClosing(e);
        }
    }
}