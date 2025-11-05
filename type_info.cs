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
    public partial class type_info : Form
    {
        Dictionary<string, string> camera_param = new Dictionary<string, string>();
        Dictionary<string, string> camera_chinesename = new Dictionary<string, string>();
        Dictionary<string, int> defect_type = new Dictionary<string, int>();
        Dictionary<string, int> defect_type_clsid = new Dictionary<string, int>();
        Dictionary<string, int> defect_type_stop = new Dictionary<string, int>();
        Dictionary<string, int> defect_type_PTFE = new Dictionary<string, int>();
        Dictionary<string, int> defect_type_PTFE_clsid = new Dictionary<string, int>();
        Dictionary<string, int> defect_type_PTFE_stop = new Dictionary<string, int>();
        Dictionary<string, string> param = new Dictionary<string, string>();        
        Dictionary<string, string> param_chinesename = new Dictionary<string, string>();
        Dictionary<string, string> param_Group4show = new Dictionary<string, string>();
        Dictionary<string, string> param_PTFE = new Dictionary<string, string>();
        Dictionary<string, string> param_PTFE_chinesename = new Dictionary<string, string>();
        Dictionary<string, string> param_PTFE_Group4show = new Dictionary<string, string>();
        Dictionary<string, int> blow = new Dictionary<string, int>();
        Dictionary<string, string> blow_chinesename = new Dictionary<string, string>();

        public type_info()
        {
            InitializeComponent();
        }
        private void db_load()
        {
            dataGridView1.Rows.Clear();
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
                        if (true)
                        {
                            dataGridView1.Rows.Add(c.TypeColumn, c.material, c.thick, c.PTFEColor, 
                                                    c.ID, c.OD, c.H, 
                                                    c.hasgroove, c.boxorpack, c.hasYZP, c.package);
                        }                        
                        comboBox5.Items.Add(c.TypeColumn);
                    }
                }
            }

            if (textBox1.Text != "")
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value.ToString() == textBox1.Text)
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
                        break;
                    }
                }
            }
        }
        private void parameter_Load(object sender, EventArgs e)
        {
            /*
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Cameras
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (!camera_param.ContainsKey(c.Name))
                        {
                            camera_param.Add(c.Name, c.Value);
                            camera_chinesename.Add(c.Name, c.ChineseName);
                        }
                    }
                }
            }
            */
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectChecks
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (c.Type == "10211111410T")
                        {
                            if (!defect_type.ContainsKey(c.Name))
                            {
                                defect_type.Add(c.Name, (int)c.Yn);
                                //defect_type_clsid.Add(c.Name, (int)c.ID);
                                //defect_type_stop.Add(c.Name, (int)c.Stop);
                            }
                        }
                        else if (c.Type == "10214062510T")
                        {
                            if (!defect_type_PTFE.ContainsKey(c.Name))
                            {
                                defect_type_PTFE.Add(c.Name, (int)c.Yn);
                               // defect_type_PTFE_clsid.Add(c.Name, (int)c.ClsId);
                               // defect_type_PTFE_stop.Add(c.Name, (int)c.Stop);
                            }
                        }
                    }
                }

            }

            using (var db = new MydbDB())
            {
                var q =
                    from c in db.@params
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (c.Type == "10211111410T")
                        {
                            if (!param.ContainsKey(c.Name + c.Stop))
                            {
                                param.Add(c.Name + c.Stop, c.Value);
                                if (!param_chinesename.ContainsKey(c.Name + c.Stop))
                                    param_chinesename.Add(c.Name + c.Stop, c.ChineseName);
                               /* if (!param_Group4show.ContainsKey(c.Name + c.Stop))
                                    param_Group4show.Add(c.Name + c.Stop, c.Group4show);*/
                            }
                        }
                        else if (c.Type == "10214062510T")
                        {
                            if (!param_PTFE.ContainsKey(c.Name + c.Stop))
                            {
                                param_PTFE.Add(c.Name + c.Stop, c.Value);
                                if (!param_PTFE_chinesename.ContainsKey(c.Name + c.Stop))
                                    param_PTFE_chinesename.Add(c.Name + c.Stop, c.ChineseName);
                                /*if (!param_PTFE_Group4show.ContainsKey(c.Name + c.Stop))
                                    param_PTFE_Group4show.Add(c.Name + c.Stop, c.Group4show);*/
                            }
                        }
                    }
                }
            }

            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Blows
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (!blow.ContainsKey(c.Name + c.Stop))
                        {
                            blow.Add(c.Name + c.Stop, c.Time);
                            blow_chinesename.Add(c.Name + c.Stop, c.ChineseName);
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
            button4.Text = "儲存(新增)";

            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            comboBox6.Enabled = true;
            comboBox4.Enabled = true;
            comboBox3.Enabled = true;
            comboBox2.Enabled = true;
            comboBox1.Enabled = true;

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            comboBox6.Text = "";
            comboBox4.Text = "";
            comboBox3.Text = "";
            comboBox2.Text = "";
            comboBox1.Text = "";
            comboBox6.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox1.SelectedIndex = -1;

            button4.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            db_load();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
               comboBox6.Enabled = false;
                comboBox4.Enabled = false;
                comboBox3.Enabled = false;
                comboBox2.Enabled = false;
                comboBox1.Enabled = false;

                textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value?.ToString() ?? "";
                comboBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? "";
                textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value?.ToString() ?? "";
                comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value?.ToString() ?? "";
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value?.ToString() ?? "";
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value?.ToString() ?? "";
                textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value?.ToString() ?? "";
                comboBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value?.ToString() ?? "";
                comboBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value?.ToString() ?? "";
                comboBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value?.ToString() ?? ""; //invisable

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
                        from c in db.Types
                        where c.TypeColumn == textBox1.Text
                        select c;

                    if (q.Any())
                    {
                        textBox1.Enabled = true;
                        textBox2.Enabled = true;
                        textBox3.Enabled = true;
                        textBox4.Enabled = true;
                        textBox5.Enabled = true;
                        textBox6.Enabled = true;
                        comboBox6.Enabled = true;
                        comboBox4.Enabled = true;
                        comboBox3.Enabled = true;
                        comboBox2.Enabled = true;
                        comboBox1.Enabled = true;

                        button1.Enabled = false;
                        button5.Enabled = false;

                        button4.Text = "儲存(編輯)";
                    }
                    else
                    {
                        MessageBox.Show("該料號不存在!");
                    }
                }
            }
            //邏輯待更改
            else if (button4.Text == "儲存(編輯)")
            {
                // 使用者編輯完後，按下「儲存(編輯)」
                // 1. 檢查是否資料輸入完整
                if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                    string.IsNullOrWhiteSpace(textBox2.Text) ||
                    string.IsNullOrWhiteSpace(textBox3.Text) ||
                    string.IsNullOrWhiteSpace(textBox4.Text))
                {
                    MessageBox.Show("請填寫完整資料!");
                    return;
                }

                using (var db = new MydbDB())
                {
                    // 2. 假設「舊料號」就是 dataGridView1帶進來的
                    string oldType = dataGridView1.CurrentRow?.Cells[0].Value?.ToString();

                    if (string.IsNullOrEmpty(oldType))
                    {
                        MessageBox.Show("無法取得舊料號，請重新選擇!");
                        return;
                    }

                    // 3. 若使用者改了 textBox1(料號)，要檢查是否與其他料號衝突
                    string newType = textBox1.Text.Trim();
                    if (!newType.Equals(oldType, StringComparison.OrdinalIgnoreCase)) //是否與當前選擇料號相等 (到底有沒有更改)
                    {
                        // 檢查 DB 是否已存在 newType
                        var checkExists = from t in db.Types
                                          where t.TypeColumn == newType
                                          select t;
                        if (checkExists.Any())
                        {
                            // 已存在相同料號 => 不允許
                            MessageBox.Show("已有相同料號，請更改其他料號!");
                            return;
                        }
                    }

                    // 4. Update
                    double valID, valOD, valH, valthick;
                    if (!double.TryParse(textBox2.Text, out valID) ||
                        !double.TryParse(textBox3.Text, out valOD) ||
                        !double.TryParse(textBox4.Text, out valH) ||
                         !double.TryParse(textBox5.Text, out valthick))
                    {
                        MessageBox.Show("ID / OD / H / thick 必須是數字!");
                        return;
                    }

                    db.Types
                      .Where(p => p.TypeColumn == oldType)
                      .Set(p => p.TypeColumn, newType)
                      .Set(p => p.ID, valID)
                      .Set(p => p.OD, valOD)
                      .Set(p => p.H, valH)
                      // 油溝(PTFEType) 直接存 comboBox2.Text
                      .Set(p => p.hasgroove, comboBox2.Text)
                      // PTFE顏色 => comboBox1
                      .Set(p => p.PTFEColor, comboBox1.Text)
                      .Set(p => p.material, comboBox3.Text)
                      .Set(p => p.thick, valthick)
                      .Set(p => p.boxorpack, comboBox4.Text)
                      .Set(p => p.hasYZP, comboBox6.Text)
                      .Set(p => p.package, textBox6.Text)
                      .Update();
                }

                // 編輯完成，回到「編輯」按鈕狀態
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                comboBox6.Enabled = false;
                comboBox4.Enabled = false;
                comboBox3.Enabled = false;
                comboBox2.Enabled = false;
                comboBox1.Enabled = false;
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";

                db_load(); // 重新載入 DataGridView
            }
            //邏輯待更改
            else if (button4.Text == "儲存(新增)")
            {
                // 使用者按下「新增」後 => 填好資料 => 按「儲存(新增)」
                // 1. 檢查資料欄位
                if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                    string.IsNullOrWhiteSpace(textBox2.Text) ||
                    string.IsNullOrWhiteSpace(textBox3.Text) ||
                    string.IsNullOrWhiteSpace(textBox4.Text) ||
                    string.IsNullOrWhiteSpace(textBox5.Text) ||
                    string.IsNullOrWhiteSpace(textBox6.Text))
                {
                    MessageBox.Show("請填寫完整資料!");
                    return;
                }

                double valID, valOD, valH, valthick;
                if (!double.TryParse(textBox2.Text, out valID) ||
                    !double.TryParse(textBox3.Text, out valOD) ||
                    !double.TryParse(textBox4.Text, out valH) ||
                    !double.TryParse(textBox5.Text, out valthick))
                {
                    MessageBox.Show("ID / OD / H 必須是數字!");
                    return;
                }

                string newType = textBox1.Text.Trim();

                using (var db = new MydbDB())
                {
                    // 2. 檢查是否已存在相同料號
                    var checkExists = from t in db.Types
                                      where t.TypeColumn == newType
                                      select t;
                    if (checkExists.Any())
                    {
                        MessageBox.Show("該料號已存在，無法新增!");
                        return;
                    }

                    // 3. Insert 到 db.Types
                    db.Types
                      .Value(p => p.TypeColumn, newType)
                      .Value(p => p.material, comboBox3.Text)
                      .Value(p => p.thick, valthick)
                      .Value(p => p.ID, valID)
                      .Value(p => p.OD, valOD)
                      .Value(p => p.H, valH)
                      .Value(p => p.hasgroove, comboBox2.Text)
                      .Value(p => p.PTFEColor, comboBox1.Text)
                      .Value(p => p.boxorpack, comboBox4.Text)
                      .Value(p => p.hasYZP, comboBox6.Text)
                      .Value(p => p.package, textBox6.Text)
                      .Insert();

                    // 如果還有其他相機參數 / DefectChecks / @params / blows 要一併 Insert
                    // 可以在這裡進行
                }

                // 新增完成，回到「編輯」按鈕狀態
                textBox1.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                comboBox6.Enabled = false;
                comboBox4.Enabled = false;
                comboBox3.Enabled = false;
                comboBox2.Enabled = false;
                comboBox1.Enabled = false;
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";

                db_load(); // 重新載入 DataGridView
            }
        }
        

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("確定要刪除料號：" + textBox1.Text + "的資料?", "警告", MessageBoxButtons.OKCancel);

            if (result == DialogResult.OK)
            {
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.Types
                        where c.TypeColumn == textBox1.Text
                        orderby c.TypeColumn
                        select c;

                    if (q.Count() > 0)
                    {
                        db.Types
                          .Delete(p => p.TypeColumn == textBox1.Text);

                        #region 相機參數刪除
                        var q2 = from c in db.Cameras
                                 where c.Type == textBox1.Text
                                 select c;

                        if (q2.Count() > 0)
                        {
                            db.Cameras
                              .Delete(p => p.Type == textBox1.Text);
                        }
                        #endregion

                        #region DefectChecks參數刪除
                        var q3 = from c in db.DefectChecks
                                 where c.Type == textBox1.Text
                                 select c;

                        if (q3.Count() > 0)
                        {
                            db.DefectChecks
                              .Delete(p => p.Type == textBox1.Text);
                        }
                        #endregion

                        #region params參數刪除
                        var q4 = from c in db.@params
                                 where c.Type == textBox1.Text
                                 select c;

                        if (q4.Count() > 0)
                        {
                            db.@params
                              .Delete(p => p.Type == textBox1.Text);
                        }
                        #endregion

                        #region blows參數刪除
                        var q5 = from c in db.Blows
                                 where c.Type == textBox1.Text
                                 select c;

                        if (q5.Count() > 0)
                        {
                            db.Blows
                              .Delete(p => p.Type == textBox1.Text);
                        }
                        #endregion

                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox3.Text = "";
                        textBox4.Text = "";
                        textBox5.Text = "";
                        textBox6.Text = "";
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

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) ||
                (e.KeyChar >= (Char)65 && e.KeyChar <= (Char)90) ||
                (e.KeyChar >= (Char)97 && e.KeyChar <= (Char)122) ||
                 e.KeyChar == (Char)13 || e.KeyChar == (Char)8 ||
                 e.KeyChar == (Char)32 || e.KeyChar == (Char)46)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void comboBox2_DrawItem(object sender, DrawItemEventArgs e)
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

                    Brush brush = new SolidBrush(cbx.ForeColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;

                    //重繪字串
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }
        private void comboBox6_DrawItem(object sender, DrawItemEventArgs e)
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

                    Brush brush = new SolidBrush(cbx.ForeColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;

                    //重繪字串
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Types
                    where c.TypeColumn == (comboBox5.Text)
                    orderby c.TypeColumn
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dataGridView1.Rows.Add(c.TypeColumn, c.material, c.thick, c.PTFEColor,
                                                c.ID, c.OD, c.H,
                                                c.hasgroove == "groove" ? "有" : "無",
                                                c.boxorpack,
                                                c.hasYZP,
                                                c.package);  // 加上這個缺少的欄位
                    }
                }
            }
        }
    }
}
