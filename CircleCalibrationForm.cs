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
    public partial class CircleCalibrationForm : Form
    {
        private List<Mat> calibrationImages = new List<Mat>();
        private int currentImageIndex = 0;
        private OpenCvSharp.Point circleCenter = new OpenCvSharp.Point(1000, 1000);
        private int circleRadius = 500;
        private bool isDragging = false;
        private bool isResizing = false;
        private OpenCvSharp.Point lastMousePos;

        // UI 控制項
        private PictureBox imageDisplay;
        private TrackBar radiusTrackBar;
        private NumericUpDown centerXInput, centerYInput, radiusInput;
        private Label statusLabel;
        private ListBox imageListBox;
        private Button btnLoadImages, btnApplyToAll, btnSaveSettings;
        private CheckBox chkShowAllCircles;

        // 新增：料號和站點輸入控制項
        private TextBox txtProductType;
        private ComboBox cmbStation;
        private ComboBox cmbCircleType; // 內圓/外圓/倒角選擇
        private Button btnLoadFromDB;
        private Label lblCurrentSettings;

        private bool hasValidParameters = false; // 新增：標記是否有有效參數
        private bool imageDisplayHasFocus = false; // 圖片區域焦點
        private bool imageListBoxHasFocus = false; // 圖片列表焦點

        private bool isImageEnlarged = false; // 是否處於放大模式
        private EnlargedImageForm enlargedImageForm; // 放大圖片的視窗
        private PictureBox enlargedPictureBox; // 放大圖片的控制項

        private bool hasSavedToDb = false;
        public bool HasSavedToDb => hasSavedToDb;

        #region 放大圖片功能

        // 雙擊事件處理
        private void ImageDisplay_DoubleClick(object sender, EventArgs e)
        {
            if (hasValidParameters && calibrationImages.Count > 0 && !isImageEnlarged)
            {
                ShowEnlargedImage();
            }
        }

        // 顯示放大圖片
        private void ShowEnlargedImage()
        {
            if (calibrationImages.Count == 0) return;

            isImageEnlarged = true;

            // 建立放大視窗
            enlargedImageForm = new EnlargedImageForm(this);
            enlargedPictureBox = enlargedImageForm.ImagePictureBox;

            // 設定放大圖片
            var currentImage = calibrationImages[currentImageIndex];
            enlargedPictureBox.Image = currentImage.ToBitmap();

            // 顯示放大視窗
            enlargedImageForm.Show(this);
            enlargedImageForm.Focus();

            // 更新狀態
            UpdateStatusWithKeyboardHint();
        }

        // 關閉放大圖片
        public void CloseEnlargedImage()
        {
            isImageEnlarged = false;
            enlargedImageForm = null;
            enlargedPictureBox = null;

            // 重新將焦點設到原始圖片
            imageDisplay.Focus();
            UpdateStatusWithKeyboardHint();
        }

        // 在放大圖片上繪製圓圈
        public void DrawCircleOnEnlargedImage(PaintEventArgs e)
        {
            if (calibrationImages.Count == 0 || enlargedPictureBox == null) return;

            // 使用藍色表示放大模式
            Color circleColor = Color.DodgerBlue;
            int penWidth = 2;

            using (var pen = new Pen(circleColor, penWidth))
            {
                var displayRect = GetEnlargedImageDisplayRect();
                if (displayRect.Width > 0 && displayRect.Height > 0)
                {
                    var currentImage = calibrationImages[currentImageIndex];
                    var scaleX = (float)displayRect.Width / currentImage.Width;
                    var scaleY = (float)displayRect.Height / currentImage.Height;

                    var displayCenterX = displayRect.X + circleCenter.X * scaleX;
                    var displayCenterY = displayRect.Y + circleCenter.Y * scaleY;
                    var displayRadius = circleRadius * Math.Min(scaleX, scaleY);

                    // 繪製圓
                    e.Graphics.DrawEllipse(pen,
                        (float)(displayCenterX - displayRadius),
                        (float)(displayCenterY - displayRadius),
                        (float)(displayRadius * 2),
                        (float)(displayRadius * 2));

                    // 繪製圓心十字（放大版）
                    int crossSize = 20;
                    e.Graphics.DrawLine(pen,
                        (float)(displayCenterX - crossSize), (float)displayCenterY,
                        (float)(displayCenterX + crossSize), (float)displayCenterY);
                    e.Graphics.DrawLine(pen,
                        (float)displayCenterX, (float)(displayCenterY - crossSize),
                        (float)displayCenterX, (float)(displayCenterY + crossSize));
                }
            }

            // 顯示操作提示
            using (var brush = new SolidBrush(Color.FromArgb(200, Color.DarkBlue)))
            using (var textBrush = new SolidBrush(Color.White))
            using (var font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold))
            {
                string hint = "放大模式 - ↑↓←→ 移動圓心，ESC 關閉";
                var textSize = e.Graphics.MeasureString(hint, font);
                var rect = new RectangleF(10, 10, textSize.Width + 20, textSize.Height + 10);

                e.Graphics.FillRectangle(brush, rect);
                e.Graphics.DrawString(hint, font, textBrush, 20, 15);
            }
        }

        // 取得放大圖片的顯示區域
        private Rectangle GetEnlargedImageDisplayRect()
        {
            if (enlargedPictureBox?.Image == null) return Rectangle.Empty;

            var imageSize = enlargedPictureBox.Image.Size;
            var displaySize = enlargedPictureBox.ClientSize;

            float scaleX = (float)displaySize.Width / imageSize.Width;
            float scaleY = (float)displaySize.Height / imageSize.Height;
            float scale = Math.Min(scaleX, scaleY);

            int scaledWidth = (int)(imageSize.Width * scale);
            int scaledHeight = (int)(imageSize.Height * scale);

            int x = (displaySize.Width - scaledWidth) / 2;
            int y = (displaySize.Height - scaledHeight) / 2;

            return new Rectangle(x, y, scaledWidth, scaledHeight);
        }

        // 放大圖片的鍵盤事件
        public void EnlargedForm_KeyDown(KeyEventArgs e)
        {
            if (!hasValidParameters || calibrationImages.Count == 0) return;

            switch (e.KeyCode)
            {
                case Keys.Left:
                    MoveCenterBy(-1, 0);
                    e.Handled = true;
                    break;
                case Keys.Right:
                    MoveCenterBy(1, 0);
                    e.Handled = true;
                    break;
                case Keys.Up:
                    MoveCenterBy(0, -1);
                    e.Handled = true;
                    break;
                case Keys.Down:
                    MoveCenterBy(0, 1);
                    e.Handled = true;
                    break;
                case Keys.Escape:
                    enlargedImageForm?.Close();
                    e.Handled = true;
                    break;
            }
        }

        // 放大圖片的滑鼠事件
        public void EnlargedImage_MouseDown(MouseEventArgs e)
        {
            if (calibrationImages.Count == 0 || enlargedPictureBox == null) return;

            OpenCvSharp.Point imagePoint = EnlargedScreenToImageCoordinates(e.Location);
            double distanceToCenter = Math.Sqrt(Math.Pow(imagePoint.X - circleCenter.X, 2) +
                                               Math.Pow(imagePoint.Y - circleCenter.Y, 2));

            if (Math.Abs(distanceToCenter - circleRadius) < 30) // 放大模式下增加點擊容差
            {
                isResizing = true;
                enlargedImageForm.Cursor = Cursors.SizeAll;
            }
            else if (distanceToCenter < circleRadius)
            {
                isDragging = true;
                enlargedImageForm.Cursor = Cursors.Hand;
            }

            lastMousePos = imagePoint;
        }

        public void EnlargedImage_MouseMove(MouseEventArgs e)
        {
            if (calibrationImages.Count == 0 || enlargedPictureBox == null) return;

            OpenCvSharp.Point imagePoint = EnlargedScreenToImageCoordinates(e.Location);

            if (isDragging)
            {
                int deltaX = imagePoint.X - lastMousePos.X;
                int deltaY = imagePoint.Y - lastMousePos.Y;

                circleCenter = new OpenCvSharp.Point(circleCenter.X + deltaX, circleCenter.Y + deltaY);
                UpdateUIWithParameters();

                // 同時更新兩個視窗
                imageDisplay.Invalidate();
                enlargedPictureBox.Invalidate();
            }
            else if (isResizing)
            {
                double newRadius = Math.Sqrt(Math.Pow(imagePoint.X - circleCenter.X, 2) +
                                            Math.Pow(imagePoint.Y - circleCenter.Y, 2));
                circleRadius = Math.Max(50, Math.Min(1000, (int)newRadius));
                UpdateUIWithParameters();

                // 同時更新兩個視窗
                imageDisplay.Invalidate();
                enlargedPictureBox.Invalidate();
            }
            else
            {
                // 更新游標樣式
                double distanceToCenter = Math.Sqrt(Math.Pow(imagePoint.X - circleCenter.X, 2) +
                                                   Math.Pow(imagePoint.Y - circleCenter.Y, 2));

                if (Math.Abs(distanceToCenter - circleRadius) < 30)
                    enlargedImageForm.Cursor = Cursors.SizeAll;
                else if (distanceToCenter < circleRadius)
                    enlargedImageForm.Cursor = Cursors.Hand;
                else
                    enlargedImageForm.Cursor = Cursors.Default;
            }

            lastMousePos = imagePoint;
        }

        public void EnlargedImage_MouseUp(MouseEventArgs e)
        {
            isDragging = false;
            isResizing = false;
            enlargedImageForm.Cursor = Cursors.Default;

            if (chkShowAllCircles.Checked)
            {
                AnalyzeAllImages();
            }
        }

        // 放大圖片的座標轉換
        private OpenCvSharp.Point EnlargedScreenToImageCoordinates(System.Drawing.Point screenPoint)
        {
            if (calibrationImages.Count == 0 || enlargedPictureBox == null)
                return new OpenCvSharp.Point(0, 0);

            var displayRect = GetEnlargedImageDisplayRect();
            if (displayRect.Width <= 0 || displayRect.Height <= 0)
                return new OpenCvSharp.Point(0, 0);

            Mat currentImage = calibrationImages[currentImageIndex];

            float scaleX = (float)currentImage.Width / displayRect.Width;
            float scaleY = (float)currentImage.Height / displayRect.Height;

            int relativeX = screenPoint.X - displayRect.X;
            int relativeY = screenPoint.Y - displayRect.Y;

            return new OpenCvSharp.Point(
                Math.Max(0, Math.Min(currentImage.Width - 1, (int)(relativeX * scaleX))),
                Math.Max(0, Math.Min(currentImage.Height - 1, (int)(relativeY * scaleY)))
            );
        }

        #endregion

        public CircleCalibrationForm()
        {
            InitializeComponent();
            SetupUI();
            this.FormClosing += CircleCalibrationForm_FormClosing;

        }

        private void InitializeComponent()
        {
            // 基本表單設定
            this.SuspendLayout();
            this.AutoScaleDimensions = new SizeF(6F, 12F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1118, 640);
            this.Text = "圓心校正工具";
            this.StartPosition = FormStartPosition.CenterParent;
            this.ResumeLayout(false);
        }
        // 由 GitHub Copilot 產生 - 移除關閉確認對話框，直接允許關閉
        private void CircleCalibrationForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 直接允許關閉，不顯示確認對話框
        }
        private void SetupUI()
        {
            // 由 GitHub Copilot 產生 - 設定表單預設字體
            this.Font = new Font("Microsoft JhengHei", 10F);
            this.KeyPreview = true;
            // 圖像顯示區域
            imageDisplay = new PictureBox
            {
                Size = new System.Drawing.Size(800, 600),
                Location = new System.Drawing.Point(10, 10),
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Gray,
                TabStop = true // 允許接收焦點
            };
            imageDisplay.MouseDown += ImageDisplay_MouseDown;
            imageDisplay.MouseMove += ImageDisplay_MouseMove;
            imageDisplay.MouseUp += ImageDisplay_MouseUp;
            imageDisplay.Paint += ImageDisplay_Paint;
            imageDisplay.Click += ImageDisplay_Click; // 新增：點擊事件
            imageDisplay.Enter += ImageDisplay_Enter; // 新增：獲得焦點事件
            imageDisplay.Leave += ImageDisplay_Leave; // 新增：失去焦點事件
            imageDisplay.DoubleClick += ImageDisplay_DoubleClick; // 🔧 新增雙擊事件
            this.Controls.Add(imageDisplay);

            // 控制面板
            SetupControlPanel();

            // 圖像列表
            SetupImageList();
        }
        #region 圖片鍵盤操控     
        // 新增：更新狀態顯示（不帶鍵盤提示）
        private void UpdateStatusWithoutKeyboardHint()
        {
            string statusText = $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}";

            if (calibrationImages.Count > 0)
            {
                statusText += $"\n目前圖片: {currentImageIndex + 1}/{calibrationImages.Count}";
            }

            statusLabel.Text = statusText;
        }
        #endregion

        private void SetupControlPanel()
        {
            var panel = new Panel
            {
                Location = new System.Drawing.Point(820, 10),
                Size = new System.Drawing.Size(300, 750),
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            };

            int yPos = 10;

            // === 料號和站點設定區域 ===
            panel.Controls.Add(new Label
            {
                Text = "=== 參數確認 ===",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(140, 22),
                Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold)
            });
            yPos += 28;

            // 料號輸入
            panel.Controls.Add(new Label { Text = "料號:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(55, 22) });
            txtProductType = new TextBox
            {
                Location = new System.Drawing.Point(70, yPos),
                Size = new System.Drawing.Size(120, 24),
                Text = app.produce_No ?? ""
            };
            txtProductType.TextChanged += TxtProductType_TextChanged; // 新增事件
            panel.Controls.Add(txtProductType);
            yPos += 32;

            // 站點選擇
            panel.Controls.Add(new Label { Text = "站點:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(55, 22) });
            cmbStation = new ComboBox
            {
                Location = new System.Drawing.Point(70, yPos),
                Size = new System.Drawing.Size(90, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            for (int i = 1; i <= 4; i++)
            {
                cmbStation.Items.Add($"站點 {i}");
            }
            cmbStation.SelectedIndex = 0;
            cmbStation.SelectedIndexChanged += CmbStation_SelectedIndexChanged; // 新增事件
            panel.Controls.Add(cmbStation);
            yPos += 32;

            // 圓類型選擇
            panel.Controls.Add(new Label { Text = "圓類型:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(60, 22) });
            cmbCircleType = new ComboBox
            {
                Location = new System.Drawing.Point(70, yPos),
                Size = new System.Drawing.Size(90, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbCircleType.Items.AddRange(new string[] { "外圓", "內圓", "倒角圓" });
            cmbCircleType.SelectedIndex = 0;
            cmbCircleType.SelectedIndexChanged += CmbCircleType_SelectedIndexChanged; // 新增事件
            panel.Controls.Add(cmbCircleType);
            yPos += 42;

            // 參數檢查按鈕
            var btnCheckParams = new Button
            {
                Text = "檢查參數是否存在",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(150, 32),
                BackColor = Color.Orange
            };
            btnCheckParams.Click += BtnCheckParams_Click;
            panel.Controls.Add(btnCheckParams);
            yPos += 42;

            // 當前設定顯示
            lblCurrentSettings = new Label
            {
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(280, 80),
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                Text = "請先檢查參數是否存在",
                TextAlign = ContentAlignment.TopLeft,
                BackColor = Color.LightGray
            };
            panel.Controls.Add(lblCurrentSettings);
            yPos += 90;

            // 分隔線
            panel.Controls.Add(new Label
            {
                Text = "=== 圓心校正 ===",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(140, 22),
                Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold)
            });
            yPos += 28;

            // 載入圖片按鈕（初始禁用）
            btnLoadImages = new Button
            {
                Text = "載入校正圖片",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(130, 35),
                Enabled = false // 初始禁用
            };
            btnLoadImages.Click += BtnLoadImages_Click;
            panel.Controls.Add(btnLoadImages);
            yPos += 45;

            // 圓心座標輸入（初始禁用）
            panel.Controls.Add(new Label { Text = "圓心X:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(60, 22) });
            centerXInput = new NumericUpDown
            {
                Location = new System.Drawing.Point(75, yPos),
                Size = new System.Drawing.Size(85, 24),
                Maximum = 3000,
                Minimum = 0,
                Value = circleCenter.X,
                Enabled = false // 初始禁用
            };
            centerXInput.ValueChanged += CenterInput_ValueChanged;
            panel.Controls.Add(centerXInput);
            yPos += 32;

            panel.Controls.Add(new Label { Text = "圓心Y:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(60, 22) });
            centerYInput = new NumericUpDown
            {
                Location = new System.Drawing.Point(75, yPos),
                Size = new System.Drawing.Size(85, 24),
                Maximum = 3000,
                Minimum = 0,
                Value = circleCenter.Y,
                Enabled = false // 初始禁用
            };
            centerYInput.ValueChanged += CenterInput_ValueChanged;
            panel.Controls.Add(centerYInput);
            yPos += 32;

            // 半徑輸入（初始禁用）
            panel.Controls.Add(new Label { Text = "半徑:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(55, 22) });
            radiusInput = new NumericUpDown
            {
                Location = new System.Drawing.Point(75, yPos),
                Size = new System.Drawing.Size(85, 24),
                Maximum = 1000,
                Minimum = 50,
                Value = circleRadius,
                Enabled = false // 初始禁用
            };
            radiusInput.ValueChanged += RadiusInput_ValueChanged;
            panel.Controls.Add(radiusInput);
            yPos += 32;

            // 半徑調整滑桿（初始禁用）
            panel.Controls.Add(new Label { Text = "半徑微調:", Location = new System.Drawing.Point(10, yPos), Size = new System.Drawing.Size(75, 22) });
            yPos += 26;
            radiusTrackBar = new TrackBar
            {
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(200, 45),
                Minimum = 50,
                Maximum = 1000,
                Value = circleRadius,
                TickFrequency = 50,
                Enabled = false // 初始禁用
            };
            radiusTrackBar.ValueChanged += RadiusTrackBar_ValueChanged;
            panel.Controls.Add(radiusTrackBar);
            yPos += 55;

            // 狀態顯示
            statusLabel = new Label
            {
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(280, 65),
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                Text = "請先檢查參數是否存在",
                TextAlign = ContentAlignment.TopLeft
            };
            panel.Controls.Add(statusLabel);
            yPos += 75;

            // 顯示選項（初始禁用）
            chkShowAllCircles = new CheckBox
            {
                Text = "顯示所有圖片適配度",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(170, 22),
                Enabled = false, // 初始禁用
                Visible = false
            };
            chkShowAllCircles.CheckedChanged += ChkShowAllCircles_CheckedChanged;
            panel.Controls.Add(chkShowAllCircles);
            yPos += 32;

            // 套用到所有圖片（初始禁用）
            btnApplyToAll = new Button
            {
                Text = "分析所有圖片",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(130, 35),
                BackColor = Color.LightBlue,
                Enabled = false, // 初始禁用
                Visible = false
            };
            btnApplyToAll.Click += BtnApplyToAll_Click;
            panel.Controls.Add(btnApplyToAll);
            yPos += 45;

            // 儲存設定（初始禁用）
            btnSaveSettings = new Button
            {
                Text = "套用推薦值",
                Location = new System.Drawing.Point(10, yPos),
                Size = new System.Drawing.Size(150, 35),
                BackColor = Color.LightGreen,
                Enabled = false // 初始禁用
            };
            btnSaveSettings.Click += BtnSaveSettings_Click;
            panel.Controls.Add(btnSaveSettings);

            this.Controls.Add(panel);
        }
        private void SetupImageList()
        {
            imageListBox = new ListBox
            {
                Location = new System.Drawing.Point(1130, 10),
                Size = new System.Drawing.Size(350, 600),
                TabStop = true // 允許接收焦點
            };
            imageListBox.SelectedIndexChanged += ImageListBox_SelectedIndexChanged;
            imageListBox.Enter += ImageListBox_Enter; // 新增：獲得焦點事件
            imageListBox.Leave += ImageListBox_Leave; // 新增：失去焦點事件
            imageListBox.Click += ImageListBox_Click; // 新增：點擊事件
            this.Controls.Add(imageListBox);
        }
        #region 圖片區域焦點事件
        private void ImageDisplay_Click(object sender, EventArgs e)
        {
            if (hasValidParameters && calibrationImages.Count > 0)
            {
                imageDisplay.Focus();
            }
        }

        private void ImageDisplay_Enter(object sender, EventArgs e)
        {
            imageDisplayHasFocus = true;
            imageListBoxHasFocus = false;
            imageDisplay.Invalidate();
            UpdateStatusWithKeyboardHint();
        }

        private void ImageDisplay_Leave(object sender, EventArgs e)
        {
            imageDisplayHasFocus = false;
            imageDisplay.Invalidate();
            UpdateStatusWithoutKeyboardHint();
        }
        #endregion

        #region 圖片列表焦點事件
        private void ImageListBox_Click(object sender, EventArgs e)
        {
            imageListBox.Focus();
        }

        private void ImageListBox_Enter(object sender, EventArgs e)
        {
            imageListBoxHasFocus = true;
            imageDisplayHasFocus = false;
            UpdateStatusWithListHint();
        }

        private void ImageListBox_Leave(object sender, EventArgs e)
        {
            imageListBoxHasFocus = false;
            UpdateStatusWithoutKeyboardHint();
        }
        #endregion
        // 使用 ProcessCmdKey 統一處理方向鍵
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 圖片區域有焦點 - 移動圓心
            if (imageDisplayHasFocus && hasValidParameters && calibrationImages.Count > 0)
            {
                switch (keyData)
                {
                    case Keys.Left:
                        MoveCenterBy(-1, 0);
                        return true;
                    case Keys.Right:
                        MoveCenterBy(1, 0);
                        return true;
                    case Keys.Up:
                        MoveCenterBy(0, -1);
                        return true;
                    case Keys.Down:
                        MoveCenterBy(0, 1);
                        return true;
                }
            }
            // 圖片列表有焦點 - 選擇圖片（使用預設行為）
            else if (imageListBoxHasFocus && calibrationImages.Count > 0)
            {
                // 讓 ListBox 處理方向鍵，不攔截
                return base.ProcessCmdKey(ref msg, keyData);
            }

            // 其他情況不處理方向鍵
            if (keyData == Keys.Left || keyData == Keys.Right ||
                keyData == Keys.Up || keyData == Keys.Down)
            {
                return true; // 攔截但不處理
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        // 移動圓心的方法
        // 修改現有的 MoveCenterBy 方法
        private void MoveCenterBy(int deltaX, int deltaY)
        {
            if (calibrationImages.Count == 0) return;

            var currentImage = calibrationImages[currentImageIndex];

            // 計算新的圓心位置，確保不超出邊界
            int newX = Math.Max(circleRadius, Math.Min(currentImage.Width - circleRadius, circleCenter.X + deltaX));
            int newY = Math.Max(circleRadius, Math.Min(currentImage.Height - circleRadius, circleCenter.Y + deltaY));

            circleCenter = new OpenCvSharp.Point(newX, newY);

            // 更新UI控制項
            UpdateUIWithParameters();

            // 重繪原始圖片
            imageDisplay.Invalidate();

            // 🔧 如果放大視窗開啟，也要重繪放大圖片
            if (isImageEnlarged && enlargedPictureBox != null)
            {
                enlargedPictureBox.Invalidate();
            }

            // 更新狀態顯示
            UpdateStatusWithKeyboardHint();

            // 如果啟用了全圖分析，重新計算適配度
            if (chkShowAllCircles.Checked)
            {
                AnalyzeAllImages();
            }
        }

        // 狀態顯示方法
        private void UpdateStatusWithKeyboardHint()
        {
            string statusText = $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}";

            if (calibrationImages.Count > 0)
            {
                statusText += $"\n目前圖片: {currentImageIndex + 1}/{calibrationImages.Count}";
            }

            if (isImageEnlarged)
            {
                statusText += "\n放大模式 - 使用 ↑↓←→ 移動圓心，ESC 關閉";
            }
            else
            {
                statusText += "\n圖片焦點 - 使用 ↑↓←→ 移動圓心，雙擊放大";
            }

            statusLabel.Text = statusText;
        }

        private void UpdateStatusWithListHint()
        {
            string statusText = $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}";

            if (calibrationImages.Count > 0)
            {
                statusText += $"\n目前圖片: {currentImageIndex + 1}/{calibrationImages.Count}";
            }

            statusText += "\n列表焦點 - 使用 ↑↓ 選擇圖片";

            statusLabel.Text = statusText;
        }


        // 新增：檢查參數是否存在
        private void BtnCheckParams_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtProductType.Text.Trim()))
            {
                MessageBox.Show("請先輸入料號", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CheckParametersExistence();
        }

        private void CheckParametersExistence()
        {
            try
            {
                string productType = txtProductType.Text.Trim();
                int station = cmbStation.SelectedIndex + 1;
                string circleType = GetCircleTypePrefix();

                using (var db = new MydbDB())
                {
                    // 檢查必要參數是否存在
                    var centerXParam = db.@params.FirstOrDefault(p =>
                        p.Type == productType &&
                        p.Name == $"known_{circleType}_center_x" &&
                        p.Stop == station);

                    var centerYParam = db.@params.FirstOrDefault(p =>
                        p.Type == productType &&
                        p.Name == $"known_{circleType}_center_y" &&
                        p.Stop == station);

                    var radiusParam = db.@params.FirstOrDefault(p =>
                        p.Type == productType &&
                        p.Name == $"known_{circleType}_radius" &&
                        p.Stop == station);

                    if (centerXParam != null && centerYParam != null && radiusParam != null)
                    {
                        // 參數存在，載入並啟用功能
                        circleCenter = new OpenCvSharp.Point(
                            int.Parse(centerXParam.Value),
                            int.Parse(centerYParam.Value)
                        );
                        circleRadius = int.Parse(radiusParam.Value);

                        hasValidParameters = true;
                        EnableCalibrationControls(true);

                        // 🔑 關鍵：延遲更新UI，確保控制項已啟用
                        this.BeginInvoke(new Action(() => {
                            UpdateUIWithParameters();
                            UpdateCurrentSettingsDisplay();
                            imageDisplay.Invalidate();
                        }));

                        lblCurrentSettings.BackColor = Color.LightGreen;
                        lblCurrentSettings.Text = $"✅ 參數已載入\n料號: {productType}\n站點: {station}, 類型: {cmbCircleType.Text}\n" +
                                                $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}";

                        statusLabel.Text = "參數已載入，可以開始校正圓心";

                        MessageBox.Show("參數載入成功！現在可以進行圓心校正。", "成功",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // 參數不存在
                        hasValidParameters = false;
                        EnableCalibrationControls(false);

                        // 由 GitHub Copilot 產生 - 站點3/4的倒角圓不需要設定
                        bool isChamferOnStation3Or4 = (station == 3 || station == 4) && cmbCircleType.SelectedIndex == 2;

                        if (isChamferOnStation3Or4)
                        {
                            // 站點3/4的倒角圓不需要設定
                            lblCurrentSettings.BackColor = Color.LightGray;
                            lblCurrentSettings.Text = $"目前不須設定此參數\n料號: {productType}\n站點: {station}, 類型: {cmbCircleType.Text}";

                            statusLabel.Text = "目前不須設定此參數";

                            MessageBox.Show("目前不須設定此參數", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            lblCurrentSettings.BackColor = Color.LightCoral;
                            lblCurrentSettings.Text = $"❌ 參數不存在\n料號: {productType}\n站點: {station}, 類型: {cmbCircleType.Text}\n" +
                                                    "請先使用料號設定功能新增基礎參數";

                            statusLabel.Text = "參數不存在，無法進行校正";

                            MessageBox.Show($"找不到料號 '{productType}' 站點 {station} 的 {cmbCircleType.Text} 參數！\n\n" +
                                          "請先使用以下步驟新增參數：\n" +
                                          "1. 到主選單的「檢測參數設定 -> 位置參數」\n" +
                                          "2. 複製類似料號的參數到已新增區\n" +
                                          "3. 右下角點擊儲存當前頁面\n" +
                                          "4. 回到此工具進行精確校正",
                                          "參數不存在", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                hasValidParameters = false;
                EnableCalibrationControls(false);

                lblCurrentSettings.BackColor = Color.LightCoral;
                lblCurrentSettings.Text = "❌ 檢查參數時發生錯誤";

                MessageBox.Show($"檢查參數時發生錯誤: {ex.Message}", "錯誤",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 啟用/禁用校正控制項
        private void EnableCalibrationControls(bool enabled)
        {
            btnLoadImages.Enabled = enabled;
            centerXInput.Enabled = enabled;
            centerYInput.Enabled = enabled;
            radiusInput.Enabled = enabled;
            radiusTrackBar.Enabled = enabled;
            chkShowAllCircles.Enabled = enabled;
            btnApplyToAll.Enabled = enabled;
            btnSaveSettings.Enabled = enabled;

            // 如果禁用，也禁用圖像顯示的互動功能
            if (!enabled)
            {
                imageDisplay.Enabled = false;
            }
            else
            {
                imageDisplay.Enabled = true;
            }
        }

        // 當料號、站點、圓類型改變時重新檢查
        private void TxtProductType_TextChanged(object sender, EventArgs e)
        {
            ResetParameterStatus();
        }

        private void CmbStation_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetParameterStatus();
        }

        private void CmbCircleType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetParameterStatusKeepImages();
        }
        private void ResetParameterStatusKeepImages()
        {
            hasValidParameters = false;
            EnableCalibrationControls(false);

            lblCurrentSettings.BackColor = Color.LightGray;
            lblCurrentSettings.Text = "參數狀態未確認\n請點擊「檢查參數是否存在」";

            statusLabel.Text = "請先檢查參數是否存在";

            // ✅ 不清空圖像，只重設選擇
            if (imageListBox.Items.Count > 0)
            {
                imageListBox.SelectedIndex = 0;
            }
        }
        private void ResetParameterStatus()
        {
            hasValidParameters = false;
            EnableCalibrationControls(false);

            lblCurrentSettings.BackColor = Color.LightGray;
            lblCurrentSettings.Text = "參數狀態未確認\n請點擊「檢查參數是否存在」";

            statusLabel.Text = "請先檢查參數是否存在";

            // 清空圖像
            imageDisplay.Image = null;
            calibrationImages.Clear();
            imageListBox.Items.Clear();
        }

        // 根據選擇的圓類型返回對應的前綴
        private string GetCircleTypePrefix()
        {
            switch (cmbCircleType.SelectedIndex)
            {
                case 0: return "outer";     // 外圓
                case 1: return "inner";     // 內圓  
                case 2: return "chamfer";   // 倒角圓
                default: return "outer";
            }
        }
        // 更新當前設定顯示
        private void UpdateCurrentSettingsDisplay()
        {
            if (!hasValidParameters)
            {
                lblCurrentSettings.BackColor = Color.LightGray;
                lblCurrentSettings.Text = "參數狀態未確認\n請點擊「檢查參數是否存在」";
                return;
            }

            string productType = txtProductType.Text.Trim();
            int station = cmbStation.SelectedIndex + 1;
            string circleType = cmbCircleType.Text;

            lblCurrentSettings.BackColor = Color.LightGreen;
            lblCurrentSettings.Text = $"✅ 參數已確認\n" +
                                    $"料號: {productType}\n" +
                                    $"站點: {station}, 類型: {circleType}\n" +
                                    $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}";
        }


        // 事件處理器
        // 由 GitHub Copilot 產生 - 修改載入圖片方法，使用多選檔案對話框
        private void BtnLoadImages_Click(object sender, EventArgs e)
        {
            if (!hasValidParameters)
            {
                MessageBox.Show("請先檢查並載入參數！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 由 GitHub Copilot 產生 - 使用多選檔案對話框讓使用者選擇圖片
            string defaultPath = Path.GetFullPath(@".\image");
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "選擇校正用圖片（可多選）";
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp|所有檔案|*.*";
                openFileDialog.Multiselect = true;
                if (Directory.Exists(defaultPath))
                {
                    openFileDialog.InitialDirectory = defaultPath;
                }

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    LoadCalibrationImages(openFileDialog.FileNames);
                }
            }
        }

        // 由 GitHub Copilot 產生 - 修改為接受檔案路徑陣列
        private void LoadCalibrationImages(string[] imageFiles)
        {
            try
            {
                calibrationImages.Clear();
                imageListBox.Items.Clear();

                // 由 GitHub Copilot 產生 - 直接載入使用者選擇的檔案（最多 30 張）
                foreach (string file in imageFiles.Take(30))
                {
                    try
                    {
                        Mat image = Cv2.ImRead(file);
                        if (!image.Empty())
                        {
                            calibrationImages.Add(image);
                            imageListBox.Items.Add($"圖片 {calibrationImages.Count}: {Path.GetFileName(file)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"載入 {file} 失敗: {ex.Message}");
                    }
                }

                if (calibrationImages.Count > 0)
                {
                    currentImageIndex = 0;
                    imageListBox.SelectedIndex = 0;
                    UpdateImageDisplay();

                    // 如果已經有從資料庫載入的參數，保留它們
                    if (!hasValidParameters)
                    {
                        Mat firstImage = calibrationImages[0];
                        circleCenter = new OpenCvSharp.Point(firstImage.Width / 2, firstImage.Height / 2);
                        circleRadius = Math.Min(firstImage.Width, firstImage.Height) / 4;
                    }

                    UpdateUIWithParameters();
                    UpdateStatusWithoutKeyboardHint();

                    // 自動設定焦點到圖片區域
                    imageDisplay.Focus();
                }
                else
                {
                    statusLabel.Text = "未找到有效圖片";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入圖片時發生錯誤: {ex.Message}");
            }
        }
        // 修改儲存方法
        // 修改儲存方法，確保只能更新現有參數
        private void BtnSaveSettings_Click(object sender, EventArgs e)
        {
            if (!hasValidParameters)
            {
                MessageBox.Show("沒有有效的參數可以更新！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string productType = txtProductType.Text.Trim();
                int station = cmbStation.SelectedIndex + 1;
                string circleTypePrefix = GetCircleTypePrefix();
                string circleTypeName = cmbCircleType.Text;

                using (var db = new MydbDB())
                {
                    // 只更新現有參數，不新增
                    var updateCount = 0;

                    updateCount += db.@params.Where(p => p.Type == productType &&
                                                   p.Name == $"known_{circleTypePrefix}_center_x" &&
                                                   p.Stop == station)
                                            .Set(p => p.Value, circleCenter.X.ToString())
                                            .Update();

                    updateCount += db.@params.Where(p => p.Type == productType &&
                                                   p.Name == $"known_{circleTypePrefix}_center_y" &&
                                                   p.Stop == station)
                                            .Set(p => p.Value, circleCenter.Y.ToString())
                                            .Update();

                    updateCount += db.@params.Where(p => p.Type == productType &&
                                                   p.Name == $"known_{circleTypePrefix}_radius" &&
                                                   p.Stop == station)
                                            .Set(p => p.Value, circleRadius.ToString())
                                            .Update();

                    if (updateCount == 3)
                    {
                        // 由 GitHub Copilot 產生
                        // 標記：本次有成功更新到資料庫
                        hasSavedToDb = true;

                        UpdateCurrentSettingsDisplay();

                        MessageBox.Show($"圓心校正參數已成功更新！\n" +
                                      $"料號: {productType}\n" +
                                      $"站點: {station}\n" +
                                      $"類型: {circleTypeName}\n" +
                                      $"圓心: ({circleCenter.X}, {circleCenter.Y})\n" +
                                      $"半徑: {circleRadius}",
                                      "更新成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"更新失敗！預期更新3個參數，實際更新了{updateCount}個",
                                      "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 更新或插入參數的輔助方法

        private void UpdateImageDisplay()
        {
            if (calibrationImages.Count > 0 && currentImageIndex < calibrationImages.Count)
            {
                var currentImage = calibrationImages[currentImageIndex];
                imageDisplay.Image = currentImage.ToBitmap();
                imageDisplay.Invalidate();
            }
        }

        private void UpdateUI()
        {
            if (!hasValidParameters) return;

            UpdateUIWithParameters();
            statusLabel.Text = $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}";
        }
        private void UpdateUIWithParameters()
        {
            // 暫時停用事件處理，避免遞迴觸發
            centerXInput.ValueChanged -= CenterInput_ValueChanged;
            centerYInput.ValueChanged -= CenterInput_ValueChanged;
            radiusInput.ValueChanged -= RadiusInput_ValueChanged;
            radiusTrackBar.ValueChanged -= RadiusTrackBar_ValueChanged;

            try
            {
                // 強制更新所有相關控制項
                centerXInput.Value = circleCenter.X;
                centerYInput.Value = circleCenter.Y;
                radiusInput.Value = circleRadius;
                radiusTrackBar.Value = circleRadius;

                // 強制重繪控制項
                centerXInput.Refresh();
                centerYInput.Refresh();
                radiusInput.Refresh();
                radiusTrackBar.Refresh();
            }
            finally
            {
                // 重新啟用事件處理
                centerXInput.ValueChanged += CenterInput_ValueChanged;
                centerYInput.ValueChanged += CenterInput_ValueChanged;
                radiusInput.ValueChanged += RadiusInput_ValueChanged;
                radiusTrackBar.ValueChanged += RadiusTrackBar_ValueChanged;
            }
        }

        // 各種事件處理器的空實作，避免編譯錯誤
        // 滑鼠操作處理
        private void ImageDisplay_MouseDown(object sender, MouseEventArgs e)
        {
            if (calibrationImages.Count == 0) return;

            OpenCvSharp.Point imagePoint = ScreenToImageCoordinates(e.Location);
            double distanceToCenter = Math.Sqrt(Math.Pow(imagePoint.X - circleCenter.X, 2) +
                                               Math.Pow(imagePoint.Y - circleCenter.Y, 2));

            if (Math.Abs(distanceToCenter - circleRadius) < 20) // 點擊圓周附近
            {
                isResizing = true;
                Cursor = Cursors.SizeAll;
            }
            else if (distanceToCenter < circleRadius) // 點擊圓內
            {
                isDragging = true;
                Cursor = Cursors.Hand;
            }

            lastMousePos = imagePoint;
        }
        // 修改：滑鼠事件也要更新狀態顯示
        private void ImageDisplay_MouseMove(object sender, MouseEventArgs e)
        {
            if (calibrationImages.Count == 0) return;

            OpenCvSharp.Point imagePoint = ScreenToImageCoordinates(e.Location);

            if (isDragging)
            {
                int deltaX = imagePoint.X - lastMousePos.X;
                int deltaY = imagePoint.Y - lastMousePos.Y;

                circleCenter = new OpenCvSharp.Point(circleCenter.X + deltaX, circleCenter.Y + deltaY);
                UpdateUIWithParameters();
                imageDisplay.Invalidate();
            }
            else if (isResizing)
            {
                double newRadius = Math.Sqrt(Math.Pow(imagePoint.X - circleCenter.X, 2) +
                                            Math.Pow(imagePoint.Y - circleCenter.Y, 2));
                circleRadius = Math.Max(50, Math.Min(1000, (int)newRadius));
                UpdateUIWithParameters();
                imageDisplay.Invalidate();
            }
            else
            {
                double distanceToCenter = Math.Sqrt(Math.Pow(imagePoint.X - circleCenter.X, 2) +
                                                   Math.Pow(imagePoint.Y - circleCenter.Y, 2));

                if (Math.Abs(distanceToCenter - circleRadius) < 20)
                    Cursor = Cursors.SizeAll;
                else if (distanceToCenter < circleRadius)
                    Cursor = Cursors.Hand;
                else
                    Cursor = Cursors.Default;
            }

            lastMousePos = imagePoint;

            // 根據焦點狀態更新狀態顯示
            if (imageDisplayHasFocus)
                UpdateStatusWithKeyboardHint();
            else
                UpdateStatusWithoutKeyboardHint();
        }

        private void UpdateStatusDisplay(OpenCvSharp.Point imagePoint)
        {
            try
            {
                string statusText = $"圓心: ({circleCenter.X}, {circleCenter.Y}), 半徑: {circleRadius}\n";
                statusText += $"游標位置: ({imagePoint.X}, {imagePoint.Y})";

                if (calibrationImages.Count > 0)
                {
                    statusText += $"\n目前圖片: {currentImageIndex + 1}/{calibrationImages.Count}";
                }

                statusLabel.Text = statusText;
            }
            catch (Exception ex)
            {
                statusLabel.Text = $"狀態更新錯誤: {ex.Message}";
            }
        }
        private void ImageDisplay_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
            isResizing = false;
            Cursor = Cursors.Default;

            // 如果啟用了全圖分析，重新計算適配度
            if (chkShowAllCircles.Checked)
            {
                AnalyzeAllImages();
            }
        }
        private void ImageDisplay_Paint(object sender, PaintEventArgs e)
        {
            if (calibrationImages.Count > 0)
            {
                // 根據焦點狀態選擇顏色
                Color circleColor = imageDisplayHasFocus ? Color.DodgerBlue : Color.Red;
                int penWidth = imageDisplayHasFocus ? 1 : 1;

                using (var pen = new Pen(circleColor, penWidth))
                {
                    var displayRect = GetImageDisplayRect();
                    if (displayRect.Width > 0 && displayRect.Height > 0)
                    {
                        var scaleX = (float)displayRect.Width / calibrationImages[currentImageIndex].Width;
                        var scaleY = (float)displayRect.Height / calibrationImages[currentImageIndex].Height;

                        var displayCenterX = displayRect.X + circleCenter.X * scaleX;
                        var displayCenterY = displayRect.Y + circleCenter.Y * scaleY;
                        var displayRadius = circleRadius * Math.Min(scaleX, scaleY);

                        // 繪製圓
                        e.Graphics.DrawEllipse(pen,
                            (float)(displayCenterX - displayRadius),
                            (float)(displayCenterY - displayRadius),
                            (float)(displayRadius * 2),
                            (float)(displayRadius * 2));

                        // 繪製圓心十字
                        int crossSize = imageDisplayHasFocus ? 15 : 10;
                        e.Graphics.DrawLine(pen,
                            (float)(displayCenterX - crossSize), (float)displayCenterY,
                            (float)(displayCenterX + crossSize), (float)displayCenterY);
                        e.Graphics.DrawLine(pen,
                            (float)displayCenterX, (float)(displayCenterY - crossSize),
                            (float)displayCenterX, (float)(displayCenterY + crossSize));
                    }
                }
            }

            // 根據焦點狀態顯示不同提示
            if (imageDisplayHasFocus && hasValidParameters)
            {
                using (var brush = new SolidBrush(Color.FromArgb(200, Color.LightBlue)))
                using (var textBrush = new SolidBrush(Color.Black))
                using (var font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold))
                {
                    string hint = "圖片焦點 - ↑↓←→ 移動圓心";
                    var textSize = e.Graphics.MeasureString(hint, font);
                    var rect = new RectangleF(5, 5, textSize.Width + 10, textSize.Height + 5);

                    e.Graphics.FillRectangle(brush, rect);
                    e.Graphics.DrawString(hint, font, textBrush, 10, 8);
                }
            }
        }
        // 座標轉換
        private OpenCvSharp.Point ScreenToImageCoordinates(System.Drawing.Point screenPoint)
        {
            if (calibrationImages.Count == 0) return new OpenCvSharp.Point(0, 0);

            var displayRect = GetImageDisplayRect();
            if (displayRect.Width <= 0 || displayRect.Height <= 0)
                return new OpenCvSharp.Point(0, 0);

            Mat currentImage = calibrationImages[currentImageIndex];

            // 計算實際的縮放比例
            float scaleX = (float)currentImage.Width / displayRect.Width;
            float scaleY = (float)currentImage.Height / displayRect.Height;

            // 轉換相對於顯示區域的座標
            int relativeX = screenPoint.X - displayRect.X;
            int relativeY = screenPoint.Y - displayRect.Y;

            return new OpenCvSharp.Point(
                Math.Max(0, Math.Min(currentImage.Width - 1, (int)(relativeX * scaleX))),
                Math.Max(0, Math.Min(currentImage.Height - 1, (int)(relativeY * scaleY)))
            );
        }

        private Rectangle GetImageDisplayRect()
        {
            if (imageDisplay.Image == null) return Rectangle.Empty;

            var imageSize = imageDisplay.Image.Size;
            var displaySize = imageDisplay.ClientSize;

            float scaleX = (float)displaySize.Width / imageSize.Width;
            float scaleY = (float)displaySize.Height / imageSize.Height;
            float scale = Math.Min(scaleX, scaleY);

            int scaledWidth = (int)(imageSize.Width * scale);
            int scaledHeight = (int)(imageSize.Height * scale);

            int x = (displaySize.Width - scaledWidth) / 2;
            int y = (displaySize.Height - scaledHeight) / 2;

            return new Rectangle(x, y, scaledWidth, scaledHeight);
        }
        // 在 #region 圓心校正工具 區段中補完以下方法

        private void AnalyzeAllImages()
        {
            var results = new List<CircleAnalysisResult>();

            for (int i = 0; i < calibrationImages.Count; i++)
            {
                var result = AnalyzeSingleImage(calibrationImages[i], circleCenter, circleRadius, i);
                results.Add(result);
            }

            // 更新列表顯示分析結果
            UpdateImageListWithResults(results);

            // 顯示統計資訊
            ShowStatistics(results);
        }

        private CircleAnalysisResult AnalyzeSingleImage(Mat image, OpenCvSharp.Point center, int radius, int imageIndex)
        {
            var result = new CircleAnalysisResult
            {
                ImageIndex = imageIndex,
                CircleCenter = center,
                CircleRadius = radius
            };

            try
            {
                // 檢查圓是否在圖像範圍內
                result.IsInBounds = (center.X - radius >= 0 && center.X + radius < image.Width &&
                                    center.Y - radius >= 0 && center.Y + radius < image.Height);

                if (result.IsInBounds)
                {
                    // 提取圓形ROI
                    Mat mask = new Mat(image.Size(), MatType.CV_8UC1, Scalar.Black);
                    Cv2.Circle(mask, center, radius, Scalar.White, -1);

                    // 計算圓內有效像素比例
                    int totalPixels = (int)(Math.PI * radius * radius);
                    int validPixels = Cv2.CountNonZero(mask);
                    result.ValidPixelRatio = (double)validPixels / totalPixels;

                    // 計算圓內平均亮度
                    Scalar meanValue = Cv2.Mean(image, mask);
                    result.AverageBrightness = (meanValue.Val0 + meanValue.Val1 + meanValue.Val2) / 3.0;

                    // 計算適配度分數 (0-100)
                    result.FitnessScore = CalculateFitnessScore(image, mask, center, radius);

                    mask.Dispose();
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        private double CalculateFitnessScore(Mat image, Mat mask, OpenCvSharp.Point center, int radius)
        {
            // 基於多個因素計算適配度分數
            double score = 100.0;

            // 因素1: 圓是否完全在圖像內 (權重30%)
            if (center.X - radius < 0 || center.X + radius >= image.Width ||
                center.Y - radius < 0 || center.Y + radius >= image.Height)
            {
                score -= 30.0;
            }

            // 因素2: 圓內亮度分布 (權重40%)
            Scalar meanValue = Cv2.Mean(image, mask);
            double avgBrightness = (meanValue.Val0 + meanValue.Val1 + meanValue.Val2) / 3.0;

            if (avgBrightness < 50 || avgBrightness > 200) // 太暗或太亮
            {
                score -= 40.0 * (1.0 - Math.Abs(125 - avgBrightness) / 125.0);
            }

            // 因素3: 圓周邊緣清晰度 (權重30%)
            Mat gray = new Mat();
            Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

            // 在圓周附近檢測邊緣強度
            double edgeStrength = CalculateCircularEdgeStrength(gray, center, radius);
            score -= (1.0 - edgeStrength) * 30.0;

            gray.Dispose();

            return Math.Max(0, Math.Min(100, score));
        }

        private double CalculateCircularEdgeStrength(Mat grayImage, OpenCvSharp.Point center, int radius)
        {
            try
            {
                // 簡化的邊緣強度計算
                Mat edges = new Mat();
                Cv2.Canny(grayImage, edges, 50, 150);

                // 在圓周附近採樣
                int sampleCount = 36; // 每10度採樣一次
                int edgeCount = 0;

                for (int i = 0; i < sampleCount; i++)
                {
                    double angle = (i * 360.0 / sampleCount) * Math.PI / 180.0;
                    int x = (int)(center.X + radius * Math.Cos(angle));
                    int y = (int)(center.Y + radius * Math.Sin(angle));

                    if (x >= 0 && x < edges.Width && y >= 0 && y < edges.Height)
                    {
                        if (edges.At<byte>(y, x) > 0)
                            edgeCount++;
                    }
                }

                edges.Dispose();
                return (double)edgeCount / sampleCount;
            }
            catch
            {
                return 0.5; // 預設值
            }
        }

        private void UpdateImageListWithResults(List<CircleAnalysisResult> results)
        {
            imageListBox.Items.Clear();

            for (int i = 0; i < results.Count; i++)
            {
                var result = results[i];
                string statusText = result.IsInBounds ?
                    $"適配度: {result.FitnessScore:F1}%" :
                    "超出邊界";

                imageListBox.Items.Add($"圖片 {i + 1}: {statusText}");
            }
        }

        private void ShowStatistics(List<CircleAnalysisResult> results)
        {
            var validResults = results.Where(r => r.IsInBounds).ToList();

            if (validResults.Count == 0)
            {
                statusLabel.Text = "沒有有效的分析結果";
                return;
            }

            double avgScore = validResults.Average(r => r.FitnessScore);
            double minScore = validResults.Min(r => r.FitnessScore);
            double maxScore = validResults.Max(r => r.FitnessScore);

            statusLabel.Text = $"平均適配度: {avgScore:F1}% (範圍: {minScore:F1}% - {maxScore:F1}%)";
        }

        private void CenterInput_ValueChanged(object sender, EventArgs e)
        {
            if (calibrationImages.Count > 0)
            {
                var currentImage = calibrationImages[currentImageIndex];
                int newX = Math.Max(0, Math.Min(currentImage.Width - 1, (int)centerXInput.Value));
                int newY = Math.Max(0, Math.Min(currentImage.Height - 1, (int)centerYInput.Value));

                circleCenter = new OpenCvSharp.Point(newX, newY);

                // 同步更新 UI（避免無限遞迴）
                if (centerXInput.Value != newX) centerXInput.Value = newX;
                if (centerYInput.Value != newY) centerYInput.Value = newY;
            }
            else
            {
                circleCenter = new OpenCvSharp.Point((int)centerXInput.Value, (int)centerYInput.Value);
            }

            imageDisplay.Invalidate();
        }

        private void RadiusInput_ValueChanged(object sender, EventArgs e)
        {
            circleRadius = (int)radiusInput.Value;
            radiusTrackBar.Value = circleRadius;
            imageDisplay.Invalidate();
        }

        private void RadiusTrackBar_ValueChanged(object sender, EventArgs e)
        {
            circleRadius = radiusTrackBar.Value;
            radiusInput.Value = circleRadius;
            imageDisplay.Invalidate();
        }

        private void ImageListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (imageListBox.SelectedIndex >= 0 && imageListBox.SelectedIndex < calibrationImages.Count)
            {
                currentImageIndex = imageListBox.SelectedIndex;
                UpdateImageDisplay();
            }
        }

        // ✅ 實作或移除空的事件處理器
        private void ChkShowAllCircles_CheckedChanged(object sender, EventArgs e)
        {
            if (chkShowAllCircles.Checked && calibrationImages.Count > 0)
            {
                AnalyzeAllImages();
            }
        }
        private void BtnApplyToAll_Click(object sender, EventArgs e)
        {
            if (!hasValidParameters)
            {
                MessageBox.Show("請先檢查並載入參數！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (calibrationImages.Count == 0)
            {
                MessageBox.Show("請先載入校正圖片！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AnalyzeAllImages();
        }
    }
    public class EnlargedImageForm : Form
    {
        public PictureBox ImagePictureBox { get; private set; }
        private CircleCalibrationForm parentForm;

        // 新增：縮放相關屬性
        private float zoomFactor = 1.0f;
        private const float MinZoom = 0.1f;
        private const float MaxZoom = 10.0f;
        private const float ZoomStep = 0.1f;
        private System.Drawing.Point lastPanPoint;
        private bool isPanning = false;

        public EnlargedImageForm(CircleCalibrationForm parent)
        {
            parentForm = parent;
            InitializeEnlargedForm();
        }

        private void InitializeEnlargedForm()
        {
            // 由 GitHub Copilot 產生 - 設定放大視窗預設字體
            this.Font = new Font("Microsoft JhengHei", 10F);
            this.Text = "放大圖片 - 使用方向鍵移動圓心";
            this.Size = new System.Drawing.Size(1200, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new System.Drawing.Size(800, 600);

            // 建立放大的 PictureBox
            ImagePictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Black,
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            };

            // 事件處理
            ImagePictureBox.Paint += (s, e) => parentForm.DrawCircleOnEnlargedImage(e);
            ImagePictureBox.MouseDown += (s, e) => parentForm.EnlargedImage_MouseDown(e);
            ImagePictureBox.MouseMove += (s, e) => parentForm.EnlargedImage_MouseMove(e);
            ImagePictureBox.MouseUp += (s, e) => parentForm.EnlargedImage_MouseUp(e);

            // 滾輪縮放事件
            ImagePictureBox.MouseWheel += ImagePictureBox_MouseWheel;

            this.Controls.Add(ImagePictureBox);

            // 處理鍵盤事件
            this.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Escape)
                {
                    this.Close();
                    e.Handled = true;
                }
                else
                {
                    parentForm.EnlargedForm_KeyDown(e);
                }
            };

            // 表單關閉事件
            this.FormClosed += (s, e) => parentForm.CloseEnlargedImage();
        }
        #region 滾輪
        // 滾輪縮放事件處理
        private void ImagePictureBox_MouseWheel(object sender, MouseEventArgs e)
        {
            if (ImagePictureBox.Image == null) return;

            // 計算新的縮放比例
            float delta = e.Delta > 0 ? ZoomStep : -ZoomStep;
            float newZoom = Math.Max(MinZoom, Math.Min(MaxZoom, zoomFactor + delta));

            if (newZoom != zoomFactor)
            {
                // 獲取滑鼠相對於 PictureBox 的位置
                System.Drawing.Point mousePos = ImagePictureBox.PointToClient(Cursor.Position);

                // 計算縮放中心點（以滑鼠位置為中心）
                ZoomAt(newZoom, mousePos);
            }
        }

        // 以指定點為中心進行縮放
        private void ZoomAt(float newZoom, System.Drawing.Point centerPoint)
        {
            if (ImagePictureBox.Image == null) return;

            var oldZoom = zoomFactor;
            zoomFactor = newZoom;

            // 計算新的圖片大小
            var originalSize = ImagePictureBox.Image.Size;
            var newSize = new System.Drawing.Size(
                (int)(originalSize.Width * zoomFactor),
                (int)(originalSize.Height * zoomFactor)
            );

            // 設定新的圖片大小和模式
            ImagePictureBox.SizeMode = PictureBoxSizeMode.Normal;
            ImagePictureBox.Size = newSize;

            // 計算新的位置以保持縮放中心點不變
            var ratioChange = newZoom / oldZoom;
            var newLocation = new System.Drawing.Point(
                (int)(ImagePictureBox.Location.X - (centerPoint.X - ImagePictureBox.Location.X) * (ratioChange - 1)),
                (int)(ImagePictureBox.Location.Y - (centerPoint.Y - ImagePictureBox.Location.Y) * (ratioChange - 1))
            );

            ImagePictureBox.Location = newLocation;

            // 更新標題顯示當前縮放比例
            this.Text = $"放大圖片 - 縮放: {zoomFactor:P0} - 滾輪縮放, 右鍵拖曳平移, 方向鍵移動圓心";

            // 重繪圓圈
            ImagePictureBox.Invalidate();
        }
        #endregion
    }
    public class CircleAnalysisResult
    {
        public int ImageIndex { get; set; }
        public OpenCvSharp.Point CircleCenter { get; set; }
        public int CircleRadius { get; set; }
        public bool IsInBounds { get; set; }
        public double ValidPixelRatio { get; set; }
        public double AverageBrightness { get; set; }
        public double FitnessScore { get; set; }
        public string ErrorMessage { get; set; }
    }
}
