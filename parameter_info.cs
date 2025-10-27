using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LinqToDB;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Point = OpenCvSharp.Point;
using Size = OpenCvSharp.Size;


namespace peilin
{
    public partial class parameter_info : Form
    {
        Mat input = new Mat(new Size(1920, 1200), MatType.CV_8UC3, Scalar.Black);
        string last_tb3 = "";

        // 新增臨時存儲使用者調整後的ROI參數
        private Dictionary<int, (int CenterX, int CenterY, int Radius)> tempROIParams = new Dictionary<int, (int, int, int)>();

        private basler.Camera camera; // 相機引用
        private System.Windows.Forms.Timer updateTimer; // 計時器
        private int currentCameraIndex = -1; // 當前活動相機索引
        private double scaleFactor = 0.3; // 縮放比例，可根據實際需要調整

        private Form enlargedForm = null;
        private PictureBox enlargedPictureBox = null;
        private System.Windows.Forms.Timer enlargedUpdateTimer = null;
        private bool isEnlargedViewActive = false;

        public parameter_info(basler.Camera cameraInstance)
        {
            InitializeComponent();

            // 儲存相機引用
            camera = cameraInstance;

            updateTimer = new System.Windows.Forms.Timer();
            updateTimer.Interval = 100; // 100ms 每次更新
            updateTimer.Tick += UpdateTimer_Tick;

            // 添加表單關閉事件
            //this.FormClosing += Parameter_info_FormClosing;

            // 初始化ROI調整控件
            InitializeROIControls();
        }        // 初始化ROI調整控件
        private void InitializeROIControls()
        {
            // 初始化參數類型選擇
            comboBoxParameterType.SelectedIndex = 0; // 預設選擇 known_inner
            comboBoxParameterType.SelectedIndexChanged += ParameterType_SelectedIndexChanged;
            
            // 限制文本框只能輸入數字
            textBox4.KeyPress += NumericTextBox_KeyPress;
            textBox6.KeyPress += NumericTextBox_KeyPress;
            textBox7.KeyPress += NumericTextBox_KeyPress;

            // 當輸入框內容變更時，更新ROI
            textBox4.TextChanged += ROIParameter_TextChanged;
            textBox6.TextChanged += ROIParameter_TextChanged;
            textBox7.TextChanged += ROIParameter_TextChanged;

            // 重置按鈕事件
            button10.Click += ResetROI_Click;
            
            // 新增儲存按鈕事件
            button11.Click += SaveROI_Click;
            
            // 載入預設參數
            LoadROIParameters();
        }
        
        // 新增儲存ROI參數按鈕事件
        private void SaveROI_Click(object sender, EventArgs e)
        {
            if (currentCameraIndex < 0)
            {
                MessageBox.Show("請先選擇一個相機站點！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 檢查是否有臨時參數可儲存
            if (!tempROIParams.ContainsKey(currentCameraIndex))
            {
                MessageBox.Show("沒有調整過的參數需要儲存！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var tempParam = tempROIParams[currentCameraIndex];
            int stationNumber = currentCameraIndex + 1; // 站點編號 (1-4)
            
            // 確認儲存對話框
            DialogResult result = MessageBox.Show(
                $"確定要儲存站點 {stationNumber} 的ROI參數嗎？\n" +
                $"圓心X: {tempParam.CenterX}\n" +
                $"圓心Y: {tempParam.CenterY}\n" +
                $"半徑: {tempParam.Radius}",
                "確認儲存",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    SaveROIToDatabase(stationNumber, tempParam.CenterX, tempParam.CenterY, tempParam.Radius);
                    
                    // 儲存成功後移除臨時參數
                    tempROIParams.Remove(currentCameraIndex);
                    
                    // 更新顯示
                    UpdateDisplay();
                    
                    lbAdd($"ROI參數已儲存到資料庫 - 站點{stationNumber}", "info", "");
                    MessageBox.Show("ROI參數儲存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    lbAdd("ROI參數儲存失敗", "error", ex.Message);
                    MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }        // 儲存ROI參數到資料庫
        private void SaveROIToDatabase(int stationNumber, int centerX, int centerY, int radius)
        {
            string currentProduce = app.produce_No;
            
            if (string.IsNullOrEmpty(currentProduce))
            {
                throw new Exception("未設定料號，無法儲存參數！");
            }

            string paramPrefix = GetCurrentParameterPrefix();
            string paramTypeName = GetParameterTypeName();

            using (var db = new MydbDB())
            {
                // 修正：使用正確的參數名稱格式 - Name欄位不包含站點編號
                UpdateOrInsertParameter(db, currentProduce, $"{paramPrefix}_center_x", 
                                      centerX.ToString(), stationNumber, $"站{stationNumber}{paramTypeName}圓心X");
                
                UpdateOrInsertParameter(db, currentProduce, $"{paramPrefix}_center_y", 
                                      centerY.ToString(), stationNumber, $"站{stationNumber}{paramTypeName}圓心Y");
                
                UpdateOrInsertParameter(db, currentProduce, $"{paramPrefix}_radius", 
                                      radius.ToString(), stationNumber, $"站{stationNumber}{paramTypeName}半徑");
            }

            // 更新記憶體中的參數 - 使用TypeSetting的格式
            if (app.param != null)
            {
                app.param[$"{paramPrefix}_center_x_{stationNumber}"] = centerX.ToString();
                app.param[$"{paramPrefix}_center_y_{stationNumber}"] = centerY.ToString();
                app.param[$"{paramPrefix}_radius_{stationNumber}"] = radius.ToString();
            }
        }        // 更新或插入參數的輔助方法
        private void UpdateOrInsertParameter(MydbDB db, string productNumber, string paramName, 
                                           string paramValue, int stop, string chineseName)
        {
            // 修正：使用正確的查詢條件 - Name和Stop都要匹配
            var updateResult = db.@params
                .Where(p => p.Type == productNumber && p.Name == paramName && p.Stop == stop)
                .Set(p => p.Value, paramValue)
                .Update();

            // 如果沒有更新任何記錄，說明參數不存在，需要新增
            if (updateResult == 0)
            {
                db.@params
                    .Value(p => p.Type, productNumber)
                    .Value(p => p.Name, paramName)
                    .Value(p => p.Value, paramValue)
                    .Value(p => p.Stop, stop)
                    .Value(p => p.ChineseName, chineseName)
                    .Insert();
            }
        }

        // 取得參數類型的中文名稱
        private string GetParameterTypeName()
        {
            switch (comboBoxParameterType.SelectedIndex)
            {
                case 0: return "內圓"; // known_inner
                case 1: return "外圓"; // known_outer  
                case 2: return "倒角"; // chamfer
                default: return "內圓";
            }
        }

        // 數字輸入框限制
        private void NumericTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) ||
                e.KeyChar == (Char)8 || e.KeyChar == (Char)13)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        // ROI參數變更事件
        private void ROIParameter_TextChanged(object sender, EventArgs e)
        {
            if (currentCameraIndex < 0) return;

            int centerX, centerY, radius;
            if (!int.TryParse(textBox4.Text, out centerX)) return;
            if (!int.TryParse(textBox6.Text, out centerY)) return;
            if (!int.TryParse(textBox7.Text, out radius)) return;

            // 更新臨時參數
            tempROIParams[currentCameraIndex] = (centerX, centerY, radius);

            // 更新顯示
            UpdateDisplay();
        }        // 重置ROI按鈕事件
        private void ResetROI_Click(object sender, EventArgs e)
        {
            if (currentCameraIndex < 0) return;

            // 移除當前相機的臨時參數
            if (tempROIParams.ContainsKey(currentCameraIndex))
            {
                tempROIParams.Remove(currentCameraIndex);
            }

            // 重新讀取資料庫中的參數，使用當前選擇的參數類型
            LoadROIParametersForCamera(currentCameraIndex + 1);

            // 更新顯示
            UpdateDisplay();
        }

        // 啟動相機顯示
        private void StartCamera(int cameraIndex)
        {
            try
            {
                // 檢查相機是否可用
                if (!camera.CheckCamera(cameraIndex))
                {
                    MessageBox.Show($"相機 {cameraIndex + 1} 不可用！");
                    return;
                }

                // 停止當前相機
                StopCamera();
                camera.Stop();

                // 設定新的相機索引
                currentCameraIndex = cameraIndex;

                // 設置相機為連續模式並啟動
                camera.Setting(0); // 連續模式
                camera.Start();

                // 更新按鈕狀態
                UpdateButtonsState(cameraIndex);                // 讀取ROI參數並顯示在輸入框中，使用當前選擇的參數類型
                LoadROIParametersForCamera(cameraIndex + 1);

                // 啟動計時器
                updateTimer.Start();

                // 顯示提示訊息
                lbAdd($"開始顯示相機 {cameraIndex + 1} 的影像", "info", "");
            }
            catch (Exception ex)
            {
                lbAdd("啟動相機顯示失敗", "error", ex.Message);
            }
        }

        // 更新按鈕狀態
        private void UpdateButtonsState(int activeIndex)
        {
            button3.BackColor = (activeIndex == 0) ? Color.LightGreen : SystemColors.Control;
            button7.BackColor = (activeIndex == 1) ? Color.LightGreen : SystemColors.Control;
            button8.BackColor = (activeIndex == 2) ? Color.LightGreen : SystemColors.Control;
            button9.BackColor = (activeIndex == 3) ? Color.LightGreen : SystemColors.Control;
        }

        // 停止相機顯示
        private void StopCamera()
        {
            if (currentCameraIndex >= 0)
            {
                updateTimer.Stop();
                currentCameraIndex = -1;
                UpdateButtonsState(-1);

                // 清空ROI參數輸入框
                textBox4.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";

                // 設置相機為硬體觸發模式
                camera.Setting(1);
            }
        }

        // 計時器事件
        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            if (currentCameraIndex >= 0)
            {
                UpdateDisplay();
            }
        }

        // 更新顯示
        private void UpdateDisplay()
        {
            try
            {
                // 獲取當前相機的影像
                Mat frame = GetLatestFrame(currentCameraIndex);
                if (frame == null || frame.Empty())
                {
                    return;
                }

                // 處理影像 (繪製 ROI)
                Mat processedFrame = DrawROIOnFrame(frame, currentCameraIndex);

                // 縮放影像以適合顯示
                Mat resizedFrame = new Mat();
                Cv2.Resize(processedFrame, resizedFrame, new Size(), scaleFactor, scaleFactor, InterpolationFlags.Linear);

                // 顯示到 PictureBox
                cherngerPictureBox1.Image?.Dispose(); // 釋放舊影像資源
                cherngerPictureBox1.Image = resizedFrame.ToBitmap();

                // 釋放資源
                frame.Dispose();
                processedFrame.Dispose();
            }
            catch (Exception ex)
            {
                lbAdd("更新顯示失敗", "error", ex.Message);
                StopCamera();
            }
        }

        // 獲取最新的相機影像
        private Mat GetLatestFrame(int cameraIndex)
        {
            try
            {
                // 直接從相機獲取一幀
                
                Mat frame = camera.GetCurrentFrame(cameraIndex);

                if (frame != null && !frame.Empty())
                {
                    return frame;
                }
                
                // 如果無法直接從相機獲取，嘗試從隊列獲取
                if (app.Queue_Bitmap1.Count > 0 && cameraIndex == 0)
                {
                    // 現有的從隊列獲取影像的代碼...
                }
                // ...其餘的現有代碼...

                // 如果都失敗，則創建測試影像
                return CreateTestImage(cameraIndex);
            }
            catch (Exception ex)
            {
                lbAdd("獲取影像失敗", "error", ex.Message);
                return null;
            }
        }

        // 創建測試用影像 (如果無法獲取真實相機影像)
        private Mat CreateTestImage(int cameraIndex)
        {
            Mat testImage = new Mat(2048, 2448, MatType.CV_8UC3, new Scalar(240, 240, 240));

            // 在測試影像上畫一些內容
            Cv2.Circle(testImage, new Point(1224, 1024), 500, new Scalar(200, 200, 200), -1);
            Cv2.Circle(testImage, new Point(1224, 1024), 350, new Scalar(150, 150, 150), -1);

            // 添加相機識別文字
            Cv2.PutText(testImage, $"Camera {cameraIndex + 1}", new Point(100, 100),
                HersheyFonts.HersheyComplex, 2, new Scalar(0, 0, 255), 3);

            return testImage;
        }

        // 在影像上繪製 ROI (修改為使用臨時參數)
        private Mat DrawROIOnFrame(Mat inputFrame, int cameraIndex)
        {
            Mat result = inputFrame.Clone();

            try
            {
                int centerX, centerY, radius;
                bool hasROI = false;

                // 優先使用臨時參數
                if (tempROIParams.ContainsKey(cameraIndex))
                {
                    var temp = tempROIParams[cameraIndex];
                    centerX = temp.CenterX;
                    centerY = temp.CenterY;
                    radius = temp.Radius;
                    hasROI = true;

                    // 標示這是臨時參數
                    Cv2.PutText(result, "使用臨時參數 (未儲存)", new Point(50, 100),
                                HersheyFonts.HersheyDuplex, 1, new Scalar(0, 165, 255), 2);
                }
                else
                {
                    // 從數據庫或配置中讀取 ROI 參數
                    hasROI = ReadROIParameters(cameraIndex + 1, out centerX, out centerY, out radius);
                }

                // 如果找到 ROI 參數，在原始影像上繪製 ROI
                if (hasROI)
                {
                    Cv2.Circle(result, new Point(centerX, centerY), radius, new Scalar(0, 255, 0), 3);
                    Cv2.Circle(result, new Point(centerX, centerY), 5, new Scalar(0, 0, 255), 1);

                    // 添加十字線以便對準
                    Cv2.Line(result, new Point(centerX - 50, centerY), new Point(centerX + 50, centerY), new Scalar(0, 0, 255), 2);
                    Cv2.Line(result, new Point(centerX, centerY - 50), new Point(centerX, centerY + 50), new Scalar(0, 0, 255), 2);

                    // 顯示座標信息
                    Cv2.PutText(result, $"X: {centerX}, Y: {centerY}, R: {radius}", new Point(50, 50),
                                HersheyFonts.HersheyDuplex, 1, new Scalar(0, 255, 255), 2);
                }
            }
            catch (Exception ex)
            {
                lbAdd("ROI 繪製失敗", "error", ex.Message);
            }

            return result;
        }

        // 從數據庫或配置中讀取 ROI 參數 (保持原有方法不變)
        private bool ReadROIParameters(int cameraIndex, out int centerX, out int centerY, out int radius)
        {
            centerX = 0;
            centerY = 0;
            radius = 0;

            try
            {
                // 從 app.param 中讀取參數
                string produceNo = app.produce_No; // 當前料號

                if (string.IsNullOrEmpty(produceNo) || app.param == null)
                {
                    return false;
                }

                // 嘗試讀取指定相機和料號的 ROI 參數
                string centerXKey = $"known_inner_center_x_{cameraIndex}";
                string centerYKey = $"known_inner_center_y_{cameraIndex}";
                string radiusKey = $"known_inner_radius_{cameraIndex}"; // 使用外圓最大半徑作為 ROI

                if (app.param.ContainsKey(centerXKey) && app.param.ContainsKey(centerYKey) && app.param.ContainsKey(radiusKey))
                {
                    centerX = int.Parse(app.param[centerXKey]);
                    centerY = int.Parse(app.param[centerYKey]);
                    radius = int.Parse(app.param[radiusKey]);
                    return true;
                }

                // 如果沒有特定參數，使用默認值
                if (app.param.ContainsKey($"outer_maxRadius_{cameraIndex}"))
                {
                    //Console.WriteLine("-----------");
                    radius = int.Parse(app.param[$"outer_maxRadius_{cameraIndex}"]);
                    // 使用影像中心作為圓心
                    centerX = 1224; // 假設影像寬度為 2448
                    centerY = 1024; // 假設影像高度為 2048
                    //return true;
                }
            }
            catch (Exception ex)
            {
                lbAdd("讀取 ROI 參數失敗", "error", ex.Message);
            }            return false;
        }        // 表單關閉事件
        private void Parameter_info_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                // 先停止相機
                StopCamera();
                /*
                // 然後更新數據庫參數
                using (var db = new MydbDB())
                {
                    db.@params
                    .Where(p => p.Name == "continue_NG")
                    .Set(p => p.Value, textBox1.Text)
                    .Update();

                    db.@params
                    .Where(p => p.Name == "continue_NULL")
                    .Set(p => p.Value, textBox2.Text)
                    .Update();

                    db.@params
                    .Where(p => p.Name == "empty_th")
                    .Set(p => p.Value, textBox5.Text)
                    .Update();
                }
                */
            }
            catch (Exception ex)
            {
                Console.WriteLine($"表單關閉時發生錯誤: {ex.Message}");
            }
        }

        // 添加記錄行
        private void lbAdd(string message, string type, string details)
        {
            // 根據您的實際情況實現此方法
            Console.WriteLine($"[{type}] {message} {details}");
        }

        private void db_load()
        {
            dataGridView2.Rows.Clear();

            using (var db = new MydbDB())
            {
                var q =
                    from c in db.@params
                    where c.Type == comboBox5.Text
                    orderby c.Type, c.Stop
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        //dataGridView1.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                        dataGridView2.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                    }
                }
            }
        }
        private void parameter_Load(object sender, EventArgs e)
        {
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Types
                    orderby c.TypeColumn
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        comboBox3.Items.Add(c.TypeColumn);
                        comboBox5.Items.Add(c.TypeColumn);
                    }
                }
            }
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.@params
                    orderby c.Type, c.Stop
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (!comboBox1.Items.Contains(c.Name))
                        {
                            comboBox1.Items.Add(c.Name);
                            comboBox4.Items.Add(c.ChineseName);
                        }
                    }
                }
            }
            db_load();
            comboBox5.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button5.Enabled = false;
            button4.Text = "儲存";

            comboBox3.Enabled = true;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";
                comboBox3.Enabled = false;
                comboBox1.Enabled = false;
                textBox3.Enabled = false;
                comboBox2.Enabled = false;
                /*
                comboBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                comboBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                */
                button4.Enabled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "編輯")
            {
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.@params
                        where c.Type == comboBox3.Text && c.Name == comboBox1.Text && c.Stop == int.Parse(comboBox2.Text)
                        orderby c.Type, c.Stop
                        select c;

                    if (q.Count() > 0)
                    {
                        textBox3.Enabled = true;

                        button1.Enabled = false;
                        button5.Enabled = false;

                        button4.Text = "儲存";
                    }
                    else
                    {
                        MessageBox.Show("該料號不存在!");
                    }
                }
            }
            else //儲存
            {
                if (!comboBox3.Enabled)
                {
                    using (var db = new MydbDB())
                    {
                        db.@params
                        .Where(p => p.Type == comboBox3.Text && p.Name == comboBox1.Text && p.Stop == int.Parse(comboBox2.Text))
                        .Set(p => p.Name, comboBox1.Text)
                        .Set(p => p.Value, textBox3.Text)
                        .Set(p => p.Stop, int.Parse(comboBox2.Text))
                        .Update();
                    }

                    comboBox3.Enabled = false;
                    comboBox1.Enabled = false;
                    textBox3.Enabled = false;
                    comboBox2.Enabled = false;

                    button1.Enabled = true;
                    button5.Enabled = true;
                    button4.Text = "編輯";

                    db_load();
                }
                else
                {
                    if (comboBox3.Text != "")
                    {
                        using (var db = new MydbDB())
                        {
                            var q =
                                from c in db.@params
                                where c.Type == comboBox3.Text && c.Name == comboBox1.Text
                                orderby c.Type, c.Stop
                                select c;

                            if (q.Count() > 0)
                            {
                                MessageBox.Show("該參數已存在!");
                            }
                            else
                            {
                                foreach (var item in comboBox1.Items)
                                {
                                    for (int i = 0; i < 4; i++)
                                    {
                                        db.@params
                                           .Value(p => p.Type, comboBox3.Text)
                                           .Value(p => p.Name, item)
                                           .Value(p => p.Value, "0")
                                           .Value(p => p.Stop, i)
                                           .Insert();
                                    }
                                }

                                comboBox3.Enabled = false;
                                comboBox1.Enabled = false;
                                textBox3.Enabled = false;
                                comboBox2.Enabled = false;
                                button1.Enabled = true;
                                button5.Enabled = true;
                                button4.Text = "編輯";

                                db_load();
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("資料輸入不完整!");
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("確定要刪除料號：" + comboBox3.Text + "的參數資料?", "警告", MessageBoxButtons.OKCancel);

            if (result == DialogResult.OK)
            {
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.@params
                        where c.Type == comboBox3.Text
                        orderby c.Type, c.Stop
                        select c;

                    if (q.Count() > 0)
                    {
                        db.@params
                          .Delete(p => p.Type == comboBox3.Text);

                        comboBox3.Text = "";
                        comboBox1.Text = "";
                        textBox3.Text = "";
                        comboBox2.Text = "";
                        button5.Enabled = false;

                        db_load();
                    }
                    else
                    {
                        MessageBox.Show("該帳號不存在!");
                    }
                }
            }
        }

        private void comboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                e.DrawBackground();
                if (e.Index >= 0)
                {
                    //文字置中
                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;

                    Color textColor = cbx.Enabled ? Color.Black : Color.DimGray;
                    Brush brush = new SolidBrush(textColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;

                    //重繪字串
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) ||
                 e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //及時讀取使用者輸入，反饋在圖片上，以調整參數
            label6.Text = "";
            var ipt_c = input.Clone();
            if (last_tb3 != textBox3.Text)
            {
                last_tb3 = textBox3.Text;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            string path = System.IO.Directory.GetCurrentDirectory();
            //openFileDialog.InitialDirectory = path;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "讀取設定檔";
            openFileDialog.Filter = "所有文件(*.*)|*.*";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog.FileName != null)
            {
                input = Cv2.ImRead(openFileDialog.FileName);
                cherngerPictureBox1.Image = input.ToBitmap();
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //comboBox1.SelectedIndex = comboBox4.SelectedIndex;
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (button4.Text == "儲存") //強制退出編輯狀態
            {
                // 取消編輯狀態
                button4.Text = "編輯";
                textBox3.Enabled = false;
                comboBox3.Enabled = false;
                // ...
                button1.Enabled = true;
                button5.Enabled = true;
            }

            if (e.RowIndex >= 0)
            {
                button1.Enabled = true;
                button5.Enabled = true;
                comboBox3.Enabled = false;
                comboBox1.Enabled = false;
                textBox3.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                comboBox1.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox4.Text = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                comboBox2.Text = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                var cellValue = dataGridView2.Rows[e.RowIndex].Cells[2].Value;

                if (cellValue != null)
                {
                    Console.WriteLine("test00");
                    textBox3.Text = cellValue.ToString();
                    //Console.WriteLine($"成功赋值: {textBox3.Text}");
                }
                //textBox3.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                button4.Enabled = true;
            }

        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) || e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }

        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) || e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }

        }
        /*
        private void parameter_info_FormClosing(object sender, FormClosingEventArgs e)
        {
            using (var db = new MydbDB())
            {
                db.@params
                .Where(p => p.Name == "continue_NG")
                .Set(p => p.Value, textBox1.Text)
                .Update();

                db.@params
                .Where(p => p.Name == "continue_NULL")
                .Set(p => p.Value, textBox2.Text)
                .Update();

                db.@params
                .Where(p => p.Name == "empty_th")
                .Set(p => p.Value, textBox5.Text)
                .Update();
            }
        }
        */
        private void button2_Click(object sender, EventArgs e)
        {
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.@params
                    orderby c.Type, c.Stop
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        //dataGridView1.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                        dataGridView2.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                    }
                }
            }
            comboBox3.Text = "";
            comboBox3.SelectedIndex = -1;
            comboBox2.Text = "";
            comboBox2.SelectedIndex = -1;
            comboBox4.Text = "";
            comboBox4.SelectedIndex = -1;
            textBox3.Text = "";
            button4.Text = "編輯";
            button4.Enabled = false;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.@params
                    where c.Type == (comboBox5.Text)
                    orderby c.Type, c.Stop
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        /*
                        if (c.Group4show == "body_th")
                        {
                            dataGridView1.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                        }
                        else if (c.Group4show == "detect_th")
                        {
                            dataGridView2.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                        }
                        */
                        dataGridView2.Rows.Add(c.Type, c.Name, c.Value, c.Stop, c.ChineseName);
                    }
                }
            }

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) || e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }


        private void cherngerPictureBox1_Click(object sender, EventArgs e)
        {
            /*
            // 確保有圖像可以顯示
            if (cherngerPictureBox1.Image == null) return;

            // 獲取當前顯示的圖像
            using (var currentImage = GetLatestFrame(currentCameraIndex))
            {
                if (currentImage == null || currentImage.Empty()) return;

                // 處理影像 (繪製 ROI)
                using (Mat processedFrame = DrawROIOnFrame(currentImage, currentCameraIndex))
                {
                    // 創建並顯示放大窗體
                    ShowEnlargedImage(processedFrame);
                }
            }
            */

            // 確保有圖像可以顯示
            if (cherngerPictureBox1.Image == null) return;

            // 如果放大窗口已經存在，則關閉它
            if (isEnlargedViewActive && enlargedForm != null && !enlargedForm.IsDisposed)
            {
                enlargedForm.Close();
                isEnlargedViewActive = false;
                return;
            }

            // 建立新的放大窗體
            ShowFullScreenLiveImage();
        }
        private void ShowFullScreenLiveImage()
        {
            // 創建一個新窗體用於顯示放大的圖像
            enlargedForm = new Form();
            enlargedForm.Text = $"全螢幕實時圖像 - 相機 {currentCameraIndex + 1}";

            // 設定為全螢幕模式
            enlargedForm.WindowState = FormWindowState.Maximized;
            enlargedForm.FormBorderStyle = FormBorderStyle.None; // 無邊框
            enlargedForm.TopMost = true; // 置頂顯示
            enlargedForm.Icon = this.Icon;

            // 創建PictureBox填滿整個窗體
            enlargedPictureBox = new PictureBox();
            enlargedPictureBox.Dock = DockStyle.Fill;
            enlargedPictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            enlargedPictureBox.BackColor = Color.Black;

            // 添加顯示相機資訊的標籤
            Label infoLabel = new Label();
            infoLabel.AutoSize = true;
            infoLabel.ForeColor = Color.LightGreen;
            infoLabel.BackColor = Color.Transparent;
            infoLabel.Font = new Font("Arial", 12, FontStyle.Bold);
            infoLabel.Text = $"相機 {currentCameraIndex + 1} 即時畫面 | 按ESC關閉視窗";
            infoLabel.Location = new System.Drawing.Point(10, 10);

            // 創建半透明面板來容納標籤
            Panel overlayPanel = new Panel();
            overlayPanel.BackColor = Color.FromArgb(100, 0, 0, 0);
            overlayPanel.AutoSize = true;
            overlayPanel.Padding = new Padding(5);
            overlayPanel.Controls.Add(infoLabel);
            overlayPanel.Location = new System.Drawing.Point(10, 10);

            // 添加PictureBox和資訊面板到窗體
            enlargedForm.Controls.Add(enlargedPictureBox);
            enlargedForm.Controls.Add(overlayPanel);

            // 確保資訊面板在最上層
            overlayPanel.BringToFront();

            // 按ESC或點擊關閉窗體的處理
            enlargedForm.KeyPreview = true;
            enlargedForm.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Escape)
                    enlargedForm.Close();
            };

            // 點擊任意位置關閉窗體
            enlargedPictureBox.Click += (s, e) => enlargedForm.Close();

            // 處理窗體關閉事件，停止更新計時器
            enlargedForm.FormClosed += (s, e) => {
                if (enlargedUpdateTimer != null)
                {
                    enlargedUpdateTimer.Stop();
                    enlargedUpdateTimer.Dispose();
                    enlargedUpdateTimer = null;
                }
                isEnlargedViewActive = false;
            };

            // 創建和啟動實時更新計時器
            enlargedUpdateTimer = new System.Windows.Forms.Timer();
            enlargedUpdateTimer.Interval = 30; // 更新頻率提高，更流暢
            enlargedUpdateTimer.Tick += EnlargedUpdateTimer_Tick;

            // 顯示窗體
            enlargedForm.Show();

            // 標記放大視圖為活動狀態
            isEnlargedViewActive = true;

            // 立即更新一次
            UpdateEnlargedImage();

            // 啟動定時更新
            enlargedUpdateTimer.Start();
        }


        private void EnlargedUpdateTimer_Tick(object sender, EventArgs e)
        {
            UpdateEnlargedImage();
        }

        private void UpdateEnlargedImage()
        {
            if (enlargedPictureBox == null || !isEnlargedViewActive) return;

            try
            {
                // 獲取當前相機的影像
                using (Mat frame = GetLatestFrame(currentCameraIndex))
                {
                    if (frame == null || frame.Empty()) return;

                    // 處理影像 (繪製 ROI)
                    using (Mat processedFrame = DrawROIOnFrame(frame, currentCameraIndex))
                    {
                        // 釋放舊影像資源
                        enlargedPictureBox.Image?.Dispose();

                        // 設置新圖像
                        enlargedPictureBox.Image = processedFrame.ToBitmap();
                    }
                }
            }
            catch (Exception ex)
            {
                lbAdd("更新放大圖像失敗", "error", ex.Message);
            }
        }
        private void ShowEnlargedImage(Mat image)
        {
            // 創建一個新窗體用於顯示放大的圖像
            Form enlargedForm = new Form();
            enlargedForm.Text = $"放大圖像 - 相機 {currentCameraIndex + 1}";
            enlargedForm.StartPosition = FormStartPosition.CenterScreen;
            enlargedForm.Size = new System.Drawing.Size(1024, 768);
            enlargedForm.FormBorderStyle = FormBorderStyle.Sizable;
            enlargedForm.Icon = this.Icon;

            // 創建PictureBox填滿整個窗體
            PictureBox enlargedPictureBox = new PictureBox();
            enlargedPictureBox.Dock = DockStyle.Fill;
            enlargedPictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            enlargedPictureBox.BackColor = Color.Black;

            // 顯示圖像
            enlargedPictureBox.Image = image.ToBitmap();

            // 按ESC關閉窗體的處理
            enlargedForm.KeyPreview = true;
            enlargedForm.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Escape)
                    enlargedForm.Close();
            };

            // 添加PictureBox到窗體
            enlargedForm.Controls.Add(enlargedPictureBox);

            // 顯示窗體
            enlargedForm.ShowDialog();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            StartCamera(0); // 啟動相機1顯示
        }

        private void button7_Click(object sender, EventArgs e)
        {
            StartCamera(1); // 啟動相機2顯示
        }

        private void button8_Click(object sender, EventArgs e)
        {
            StartCamera(2); // 啟動相機3顯示
        }

        private void button9_Click(object sender, EventArgs e)
        {
            StartCamera(3); // 啟動相機4顯示
        }

        // 由 GitHub Copilot 產生 - 取像按鈕點擊事件
        private void buttonCapture_Click(object sender, EventArgs e)
        {
            if (currentCameraIndex < 0)
            {
                MessageBox.Show("請先選擇站點（站1-站4）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 獲取原始影像（不包含輔助線）
                using (Mat originalFrame = GetLatestFrame(currentCameraIndex))
                {
                    if (originalFrame == null || originalFrame.Empty())
                    {
                        MessageBox.Show("無法獲取相機影像", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 使用 SaveFileDialog 讓使用者選擇儲存位置
                    using (SaveFileDialog saveDialog = new SaveFileDialog())
                    {
                        saveDialog.Filter = "JPEG 圖片|*.jpg|PNG 圖片|*.png|BMP 圖片|*.bmp|所有檔案|*.*";
                        saveDialog.FilterIndex = 1;
                        saveDialog.FileName = $"{DateTime.Now:yyyyMMdd_HHmmss}-{currentCameraIndex + 1}";
                        saveDialog.Title = "儲存取像圖片";

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            // 儲存原始影像
                            Cv2.ImWrite(saveDialog.FileName, originalFrame);
                            MessageBox.Show($"取像成功！\n儲存位置：{saveDialog.FileName}", "成功",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"取像失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 參數類型選擇變更事件
        private void ParameterType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 當使用者變更參數類型時，重新載入當前相機的參數
            if (currentCameraIndex >= 0)
            {
                LoadROIParametersForCamera(currentCameraIndex + 1);
            }
        }
        
        // 取得目前選擇的參數類型前綴
        private string GetCurrentParameterPrefix()
        {
            switch (comboBoxParameterType.SelectedIndex)
            {
                case 0: return "known_inner"; // 內圓
                case 1: return "known_outer"; // 外圓  
                case 2: return "chamfer";     // 倒角
                default: return "known_inner";
            }
        }

        // 載入ROI參數
        private void LoadROIParameters()
        {
            try
            {
                // 這裡假設 station 是站點編號（1~4），需轉成 int
                if (comboBox3.SelectedItem == null) return;
                int stationNumber;
                if (!int.TryParse(comboBox3.SelectedItem.ToString(), out stationNumber))
                {
                    // 若無法轉換，預設為 1
                    stationNumber = 1;
                }
                LoadROIParametersForCamera(stationNumber);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入ROI參數時發生錯誤: {ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadROIParametersForCamera(int stationNumber)
        {
            try
            {
                string currentProduce = app.produce_No;
                if (string.IsNullOrEmpty(currentProduce))
                {
                    lbAdd("未設定料號，使用預設值", "warning", "");
                    // 使用預設值
                    textBox4.Text = "1224";
                    textBox6.Text = "1024";
                    textBox7.Text = "100";
                    return;
                }                string paramPrefix = GetCurrentParameterPrefix();
                
                using (var db = new MydbDB())
                {
                    // 使用正確的查詢方式 - Name欄位不包含站點編號，站點編號在Stop欄位中
                    string centerXParam = $"{paramPrefix}_center_x";
                    string centerYParam = $"{paramPrefix}_center_y";
                    string radiusParam = $"{paramPrefix}_radius";

                    // 從數據庫查詢參數
                    var centerXData = db.@params
                        .Where(p => p.Name == centerXParam && p.Type == currentProduce && p.Stop == stationNumber)
                        .FirstOrDefault();
                    
                    var centerYData = db.@params
                        .Where(p => p.Name == centerYParam && p.Type == currentProduce && p.Stop == stationNumber)
                        .FirstOrDefault();
                    
                    var radiusData = db.@params
                        .Where(p => p.Name == radiusParam && p.Type == currentProduce && p.Stop == stationNumber)
                        .FirstOrDefault();

                    // 更新文本框顯示
                    textBox4.Text = centerXData?.Value ?? "1224"; // 預設中心X
                    textBox6.Text = centerYData?.Value ?? "1024"; // 預設中心Y  
                    textBox7.Text = radiusData?.Value ?? "100";   // 預設半徑
                }
            }            catch (Exception ex)
            {
                lbAdd($"載入站點{stationNumber}的參數失敗", "error", ex.Message);
                // 使用預設值
                textBox4.Text = "1224";
                textBox6.Text = "1024";
                textBox7.Text = "100";
            }
        }

    }
}
