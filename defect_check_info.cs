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


namespace peilin
{
    public partial class defect_check_info : Form
    {
        // 紀錄當前選到的 DataGridView 列 (若是 -1 代表目前未選取任何列)
        private int selectedRowIndex = -1;

        public defect_check_info()
        {
            InitializeComponent();
        }

        /// 表單載入事件
        private void defect_check_info_Load(object sender, EventArgs e)
        {
            // 載入所有料號到 comboBox5
            LoadPartNumbersIntoComboBox5();

            // 預設選第一個 (若有資料)
            if (comboBox5.Items.Count > 0)
            {
                comboBox5.SelectedIndex = 0;
            }

            // 先把資料讀進 DataGridView
            db_load();

            // 預設右側所有欄位鎖住
            SetAllControlsEnabled(false);

            // 預設「編輯」、「刪除」按鈕先關閉，直到選了一筆資料
            button4.Enabled = false;
            button5.Enabled = false;
        }

        /// 將 defect_check 表裡的所有 料號(Type) Distinct 抓出來，放入 comboBox5
        private void LoadPartNumbersIntoComboBox5()
        {
            comboBox5.Items.Clear();
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
                        comboBox5.Items.Add(c.TypeColumn);
                    }
                }
            }
        }
        

        /// <summary>
        /// 依照 comboBox5 選的 料號，載入對應資料到 DataGridView
        /// </summary>
        private void db_load()
        {
            dataGridView1.Rows.Clear();
            selectedRowIndex = -1;

            using (var db = new MydbDB())
            {
                var selectedPart = comboBox5.Text?.Trim();
                var query = db.DefectChecks.AsQueryable();

                if (!string.IsNullOrEmpty(selectedPart))
                {
                    query = query.Where(dc => dc.Type == selectedPart);
                }

                // 排序
                query = query.OrderBy(dc => dc.Type)
                             .ThenBy(dc => dc.Stop)
                             .ThenBy(dc => dc.Name);

                var list = query.ToList();

                foreach (var item in list)
                {
                    string ynText = (item.Yn.HasValue && item.Yn.Value == 1) ? "是" : "否";
                    dataGridView1.Rows.Add(
                        item.Type,
                        item.Stop,
                        item.Name,
                        ynText,
                        item.Threshold,
                        item.ChineseName
                    );
                }
            }

            // 重置右邊顯示
            ClearRightPanel();
            button4.Enabled = false;
            button5.Enabled = false;
        }
        private void parameter_Load(object sender, EventArgs e)
        {
            db_load();

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
                        comboBox5.Items.Add(c.TypeColumn);
                    }
                }
            }
        }

        /// comboBox5 選擇變更時，重新載入 DataGridView
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectChecks
                    where c.Type == (comboBox5.Text)
                    orderby c.Type, c.Yn
                    select c;

                string ynText = "";
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        ynText = (c.Yn.HasValue && c.Yn.Value == 1) ? "是" : "否";
                        dataGridView1.Rows.Add(c.Type, c.Stop, c.Name, ynText, c.Threshold, c.ChineseName);
                    }
                }
            }
        }

        /// 按下 "顯示所有料號" (button2)，代表不篩選料號，重設 comboBox5
        private void button2_Click(object sender, EventArgs e)
        {
            comboBox5.SelectedIndex = -1; // 移除選取
            comboBox5.Text = "";
            db_load();
        }

        /// DataGridView 點擊，帶資料回到右側的欄位
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count)
            {
                selectedRowIndex = e.RowIndex;

                // 料號
                textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value?.ToString() ?? "";
                // 站數
                comboBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? "";
                // 瑕疵名稱
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                // 中文名稱
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value?.ToString() ?? "";
                // 是否檢測 => "是"/"否"
                var ynText = dataGridView1.Rows[e.RowIndex].Cells[3].Value?.ToString() ?? "";
                comboBox1.Text = ynText;
                // 閥值
                textBoxThreshold.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value?.ToString() ?? "";

                SetAllControlsEnabled(false);

                // 選到資料後，可進行「編輯」或「刪除」
                button4.Text = "編輯";
                button4.Enabled = true;
                button5.Enabled = true;
            }
        }

        /// "新增" 按鈕：清空欄位並讓欄位可輸入，最後由 button4 儲存
        private void button1_Click(object sender, EventArgs e)
        {
            // 清空欄位
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.SelectedIndex = -1; // "是否檢測"
            comboBox3.SelectedIndex = -1; // "站數"
            textBoxThreshold.Text = "";

            // 啟用輸入欄位
            SetAllControlsEnabled(true);

            // 改 button4 文字 => "儲存(新增)"
            button4.Text = "儲存(新增)";
            button4.Enabled = true;

            // selectedRowIndex = -1，表示是 "新增" 模式
            selectedRowIndex = -1;
            // 「刪除」暫時關掉
            button5.Enabled = false;
        }

        /// "刪除" 按鈕：刪除當前選中那筆資料
        private void button5_Click(object sender, EventArgs e)
        {
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("請先選取要刪除的資料。");
                return;
            }

            // 取得 DataGridView 這列的資料 (料號, 站數, 瑕疵)
            var typeVal = dataGridView1.Rows[selectedRowIndex].Cells[0].Value?.ToString();
            var stopVal = dataGridView1.Rows[selectedRowIndex].Cells[1].Value?.ToString();
            var nameVal = dataGridView1.Rows[selectedRowIndex].Cells[2].Value?.ToString();
            var ChineseName = dataGridView1.Rows[selectedRowIndex].Cells[5].Value?.ToString();
            int stopInt = 0;
            int.TryParse(stopVal, out stopInt);

            DialogResult result = MessageBox.Show($"確定要刪除：\n料號={typeVal}\n站數={stopVal}\n瑕疵={nameVal} ?", 
                                                  "警告", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                using (var db = new MydbDB())
                {
                    db.DefectChecks
                      .Where(dc => dc.Type == typeVal
                                && dc.Stop == stopInt
                                && dc.Name == nameVal
                                && dc.ChineseName == ChineseName)
                      .Delete();
                }

                MessageBox.Show("刪除成功！");
                db_load();
            }
        }

        /// "編輯" or "儲存" 按鈕：依照文字判斷目前要做編輯還是做儲存
        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "編輯")
            {
                // 解鎖是否檢測、閥值
                textBoxThreshold.Enabled = true;
                comboBox1.Enabled = true;

                // 其他欄位仍保持鎖定
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                comboBox3.Enabled = false;

                // 按鈕改為 "儲存"
                button4.Text = "儲存(編輯)";

                // 同時鎖住 "新增" 和 "刪除"
                button1.Enabled = false;
                button5.Enabled = false;
            }
            else if (button4.Text == "儲存(編輯)")
            {
                // 進行「儲存」
                SaveEdit();
                SaveThresholdAndYn();

                // 儲存後，按鈕回到 "編輯"
                button4.Text = "編輯";

                // 解鎖新增、刪除
                button1.Enabled = true;
                button5.Enabled = true;

                // 全欄位鎖回
                SetAllControlsEnabled(false);
            }
            else if (button4.Text == "儲存(新增)")
            {
                // 執行「新增儲存」
                SaveNew();

                // 結束後，回到「編輯」狀態
                button4.Text = "編輯";
                button1.Enabled = true;
                button5.Enabled = true;

                // 全欄位鎖回
                SetAllControlsEnabled(false);
            }
        }
        private void SaveEdit()
        {
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("請先選取要編輯的資料。");
                return;
            }

            // 舊資料
            var oldType = dataGridView1.Rows[selectedRowIndex].Cells[0].Value?.ToString();
            var oldStop = dataGridView1.Rows[selectedRowIndex].Cells[1].Value?.ToString();
            var oldName = dataGridView1.Rows[selectedRowIndex].Cells[2].Value?.ToString();
            var oldCHName = dataGridView1.Rows[selectedRowIndex].Cells[5].Value?.ToString();

            // 轉型
            int oldStopInt;
            if (!int.TryParse(oldStop, out oldStopInt))
            {
                MessageBox.Show("站數資料無法轉換。");
                return;
            }

            // 新資料 (使用者在右側填的)
            string newType = textBox1.Text.Trim();
            string newStopStr = comboBox3.Text.Trim();
            int newStop;
            if (!int.TryParse(newStopStr, out newStop))
            {
                MessageBox.Show("站數必須是數字。");
                return;
            }
            string newName = textBox2.Text.Trim();
            string newCHName = textBox3.Text.Trim();
            int newYn = (comboBox1.Text == "是") ? 1 : 0;
            string newThreshold = textBoxThreshold.Text.Trim();

            // 檢查是否會跟既有資料撞 (Type, Stop, Name) => 如果你想禁止使用者改成一筆已存在的記錄
            // 可在這裡檢查 DB
            using (var db = new MydbDB())
            {
                // 若 (Type,Stop,Name) 有改，需判斷 DB 中是否已有相同三欄
                if (!newType.Equals(oldType) || newStop != oldStopInt || !newName.Equals(oldName) || !newCHName.Equals(oldCHName))
                {
                    var qDup = db.DefectChecks
                                 .Where(dc => dc.Type == newType
                                           && dc.Stop == newStop
                                           && dc.Name == newName
                                           && dc.ChineseName == newCHName);
                    if (qDup.Any())
                    {
                        MessageBox.Show("已存在相同 [料號, 站數, 瑕疵]，請更改後再儲存。");
                        return;
                    }
                }

                // Update
                db.DefectChecks
                  .Where(dc => dc.Type == oldType
                            && dc.Stop == oldStopInt
                            && dc.Name == oldName)
                  .Set(dc => dc.Type, newType)
                  .Set(dc => dc.Stop, newStop)
                  .Set(dc => dc.Name, newName)
                  .Set(dc => dc.Yn, newYn)
                  .Set(dc => dc.Threshold, newThreshold)
                  .Set(dc => dc.ChineseName, newCHName)
                  .Update();
            }

            MessageBox.Show("編輯儲存成功！");
            db_load();
        }

        /// <summary>
        /// 執行「新增」儲存 (插入 DefectChecks)
        /// </summary>
        private void SaveNew()
        {
            // 直接從右側欄位抓資料
            string newType = textBox1.Text.Trim();
            string newStopStr = comboBox3.Text.Trim();
            string newName = textBox2.Text.Trim();
            string newCHName = textBox3.Text.Trim();
            int newYn = (comboBox1.Text == "是") ? 1 : 0;
            string newThreshold = textBoxThreshold.Text.Trim();

            // 基本檢查
            if (string.IsNullOrWhiteSpace(newType) ||
                string.IsNullOrWhiteSpace(newStopStr) ||
                string.IsNullOrWhiteSpace(newName) ||
                string.IsNullOrWhiteSpace(newCHName))
            {
                MessageBox.Show("請輸入完整的 [料號, 站數, 瑕疵名稱]。");
                return;
            }

            int newStop;
            if (!int.TryParse(newStopStr, out newStop))
            {
                MessageBox.Show("站數必須是數字。");
                return;
            }

            // 檢查資料庫是否已存在
            using (var db = new MydbDB())
            {
                var qDup = db.DefectChecks
                             .Where(dc => dc.Type == newType
                                       && dc.Stop == newStop
                                       && dc.Name == newName
                                       && dc.ChineseName == newCHName);
                if (qDup.Any())
                {
                    MessageBox.Show("已存在相同 [料號, 站數, 瑕疵]。無法新增！");
                    return;
                }

                // Insert
                db.DefectChecks
                  .Value(dc => dc.Type, newType)
                  .Value(dc => dc.Stop, newStop)
                  .Value(dc => dc.Name, newName)
                  .Value(dc => dc.Yn, newYn)
                  .Value(dc => dc.Threshold, newThreshold)
                  .Value(dc => dc.ChineseName, newCHName)
                  .Insert();
            }

            MessageBox.Show("新增成功！");
            db_load();
        }
        /// 儲存(新增 or 編輯) 實際執行
        private void SaveThresholdAndYn()
        {
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("未選取任何資料，無法儲存。");
                return;
            }

            // 取得舊的 key
            var oldType = dataGridView1.Rows[selectedRowIndex].Cells[0].Value?.ToString();
            var oldStop = dataGridView1.Rows[selectedRowIndex].Cells[1].Value?.ToString();
            var oldName = dataGridView1.Rows[selectedRowIndex].Cells[2].Value?.ToString();
            var oldCHName = dataGridView1.Rows[selectedRowIndex].Cells[5].Value?.ToString();

            int stopInt;
            if (!int.TryParse(oldStop, out stopInt))
            {
                MessageBox.Show("站數資料轉換失敗，請檢查資料。");
                return;
            }

            // 新的閥值 => 來自 textBoxThreshold
            string newThreshold = textBoxThreshold.Text.Trim();
            int newYn = (comboBox1.Text == "是") ? 1 : 0;

            // 寫回資料庫
            using (var db = new MydbDB())
            {
                db.DefectChecks
                  .Where(dc => dc.Type == oldType
                            && dc.Stop == stopInt
                            && dc.Name == oldName
                            && dc.ChineseName == oldCHName)
                  .Set(dc => dc.Threshold, newThreshold)
                  .Set(dc => dc.Yn, newYn)
                  .Update();
            }

            // 更新畫面
            MessageBox.Show("閥值儲存成功！");
            db_load();
        }

        /// <summary>
        /// 控制欄位啟用/停用
        /// </summary>
        private void ClearRightPanel()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.Text = "";
            comboBox3.Text = "";
            textBoxThreshold.Text = "";
        }
        private void SetAllControlsEnabled(bool enabled)
        {
            textBox1.Enabled = enabled;
            textBox2.Enabled = enabled;
            textBox3.Enabled = enabled;
            comboBox1.Enabled = enabled;
            comboBox3.Enabled = enabled;
            textBoxThreshold.Enabled = enabled;
        }
        private void comboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            ComboBox cbx = sender as ComboBox;
            e.DrawBackground();

            // 判斷是否滑鼠懸停或選中
            bool isHovered = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // 繪製背景顏色（優先處理滑鼠懸停的藍底效果）
            Color backgroundColor = isHovered ? Color.LightSteelBlue : SystemColors.Window;

            using (Brush backgroundBrush = new SolidBrush(backgroundColor))
            {
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds);
            }

            // 確定文字顏色
            Color textColor = cbx.Enabled ? Color.Black : Color.DimGray;

            // 檢查是否需要繪製文字
            if (e.Index >= 0 && cbx != null)
            {
                // 文字置中
                StringFormat sf = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Alignment = StringAlignment.Center
                };

                // 重繪字串
                using (Brush textBrush = new SolidBrush(textColor))
                {
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, textBrush, e.Bounds, sf);
                }
            }

            // 繪製焦點矩形
            e.DrawFocusRectangle();
        }


        private void textBoxThreshold_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
