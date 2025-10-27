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
    public partial class defect_type_info : Form
    {
        string origin_name = "";
        public defect_type_info()
        {
            InitializeComponent();
        }
        private void db_load()
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectTypes
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dataGridView1.Rows.Add(c.Name);
                    }
                }
            }

            if (textBox3.Text != "")
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value.ToString() == textBox3.Text)
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
                        break;
                    }
                }
            }
        }
        private void parameter_Load(object sender, EventArgs e)
        {
            db_load();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button5.Enabled = false;

            textBox3.Enabled = true;
            textBox3.Text = "";
            origin_name = "";

            button4.Text = "儲存(新增)";
            button4.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            db_load();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectTypes
                    where c.Name.Contains(textBox4.Text)
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dataGridView1.Rows.Add(c.Name);
                    }
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex>=0)
            {
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";
                textBox3.Enabled = false;
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            }            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "編輯")
            {
                using (var db = new MydbDB())
                {
                    string defectName = textBox3.Text.Trim();
                    var q = from c in db.DefectTypes
                            where c.Name == defectName
                            select c;

                    if (q.Any())
                    {
                        // 找得到 => 進入「編輯模式」
                        origin_name = defectName; // 記錄舊名稱
                        textBox3.Enabled = true;

                        button1.Enabled = false; // 關閉新增
                        button5.Enabled = false; // 關閉刪除
                        button4.Text = "儲存(編輯)";
                    }
                    else
                    {
                        MessageBox.Show("該瑕疵種類不存在，無法編輯！");
                    }
                }
            }
            else if (button4.Text == "儲存(編輯)")
            {
                // 正在編輯一筆舊名稱 = origin_name
                if (string.IsNullOrWhiteSpace(textBox3.Text))
                {
                    MessageBox.Show("請輸入瑕疵名稱！");
                    return;
                }
                string newName = textBox3.Text.Trim();
                if (string.IsNullOrWhiteSpace(origin_name))
                {
                    MessageBox.Show("舊名稱不明，請重新選擇要編輯的瑕疵種類！");
                    return;
                }

                using (var db = new MydbDB())
                {
                    // 若使用者改了名稱 => 檢查有沒有重複
                    if (!newName.Equals(origin_name, StringComparison.OrdinalIgnoreCase))
                    {
                        var qDup = from c in db.DefectTypes
                                   where c.Name == newName
                                   select c;
                        if (qDup.Any())
                        {
                            MessageBox.Show("該瑕疵種類已存在，請換其他名稱！");
                            return;
                        }
                    }

                    // 進行 Update
                    db.DefectTypes
                      .Where(p => p.Name == origin_name)
                      .Set(p => p.Name, newName)
                      .Update();

                    // 同步更新 DefectChecks
                    db.DefectChecks
                      .Where(p => p.Name == origin_name)
                      .Set(p => p.Name, newName)
                      .Update();
                }

                // 編輯完成，重置狀態
                origin_name = "";
                textBox3.Enabled = false;
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";

                db_load();
            }
            else if (button4.Text == "儲存(新增)")
            {
                // 使用者想新增一筆新的瑕疵種類
                if (string.IsNullOrWhiteSpace(textBox3.Text))
                {
                    MessageBox.Show("請輸入瑕疵名稱！");
                    return;
                }
                string newName = textBox3.Text.Trim();

                using (var db = new MydbDB())
                {
                    // 檢查有無重複
                    var q = from c in db.DefectTypes
                            where c.Name == newName
                            select c;
                    if (q.Any())
                    {
                        MessageBox.Show("該瑕疵種類已存在，無法新增！");
                        return;
                    }

                    // Insert 到 DefectTypes
                    db.DefectTypes
                       .Value(p => p.Name, newName)
                       .Insert();

                    // 再給所有料號 Types 建立 DefectChecks (和原程式類似)
                    var qTypes = from t in db.Types
                                 select t;
                    foreach (var t in qTypes)
                    {
                        db.DefectChecks
                           .Value(p => p.Type, t.TypeColumn)
                           .Value(p => p.Name, newName)
                           .Value(p => p.Yn, 0)
                           .Insert();
                    }
                }

                // 新增完成，重置狀態
                textBox3.Enabled = false;
                button1.Enabled = true;
                button5.Enabled = true;
                button4.Text = "編輯";

                db_load();
            }
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button3.Focus();
                button3_Click(sender, e);
                textBox4.Focus();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("確定要刪除瑕疵種類：" + textBox3.Text + "?", "警告", MessageBoxButtons.OKCancel);

            if (result == DialogResult.OK)
            {
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.DefectTypes
                        where c.Name == textBox3.Text
                        select c;

                    if (q.Count() > 0)
                    {
                        db.DefectTypes
                          .Delete(p => p.Name == textBox3.Text);
                        
                        db.DefectChecks
                          .Delete(p => p.Name == textBox3.Text);

                        textBox3.Text = "";
                        button5.Enabled = false;

                        db_load();
                    }
                    else
                    {
                        MessageBox.Show("該瑕疵種類不存在!");
                    }
                }
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
