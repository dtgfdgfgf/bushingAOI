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
    public partial class user_info : Form
    {

        public user_info()
        {
            InitializeComponent();
        }
        private void db_load()
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Users
                    orderby c.UserName
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (c.UserName != "engineer")
                        {
                            if (c.Level==0)
                            {
                                dataGridView1.Rows.Add(c.UserName, c.Password, c.Level,"工程師");
                            }
                            else if (c.Level == 1)
                            {
                                dataGridView1.Rows.Add(c.UserName, c.Password, c.Level,"管理者");
                            }
                            else if (c.Level == 2)
                            {
                                dataGridView1.Rows.Add(c.UserName, c.Password, c.Level,"作業員");
                            }
                        }
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
            db_load();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button5.Enabled = false;
            button4.Text = "儲存";

            button4.Enabled = true;
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            comboBox2.Enabled = true;
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
                    from c in db.Users
                    where c.UserName.Contains(textBox4.Text)
                    orderby c.UserName
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (c.UserName != "engineer")
                        {
                            if (c.Level == 1)
                            {
                                dataGridView1.Rows.Add(c.UserName, c.Password, "管理者");
                            }
                            else
                            {
                                dataGridView1.Rows.Add(c.UserName, c.Password, "作業員");
                            }
                        }
                    }
                }
            }
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
                comboBox2.Enabled = false;
                textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox2.SelectedIndex = (int)dataGridView1.Rows[e.RowIndex].Cells[2].Value - 1;
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
                        from c in db.Users
                        where c.UserName == textBox1.Text
                        orderby c.UserName
                        select c;

                    if (q.Count() > 0)
                    {
                        textBox2.Enabled = true;
                        comboBox2.Enabled = true;

                        button1.Enabled = false;
                        button5.Enabled = false;

                        button4.Text = "儲存";
                    }
                    else
                    {
                        MessageBox.Show("該使用者不存在!");
                    }
                }
            }
            else
            {
                if (!textBox1.Enabled)
                {
                    using (var db = new MydbDB())
                    {
                        db.Users
                        .Where(p => p.UserName == textBox1.Text)
                        .Set(p => p.Password, textBox2.Text)
                        .Set(p => p.Level, comboBox2.SelectedIndex + 1)
                        .Update();
                    }

                    textBox1.Enabled = false;
                    textBox2.Enabled = false;
                    comboBox2.Enabled = false;

                    button1.Enabled = true;
                    button5.Enabled = true;
                    button4.Text = "編輯";

                    db_load();
                }
                else
                {
                    if (textBox1.Text != "" && textBox2.Text != "" && comboBox2.Text != "")
                    {

                        using (var db = new MydbDB())
                        {
                            var q =
                                from c in db.Users
                                where c.UserName == textBox1.Text
                                orderby c.UserName
                                select c;

                            if (q.Count() > 0)
                            {
                                MessageBox.Show("該使用者已存在!");
                            }
                            else
                            {
                                db.Users
                                  .Value(p => p.UserName, textBox1.Text)
                                  .Value(p => p.Password, textBox2.Text)
                                  .Value(p => p.Level, comboBox2.SelectedIndex + 1)
                                  .Insert();

                                textBox1.Enabled = false;
                                textBox2.Enabled = false;
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
            DialogResult result = MessageBox.Show("確定要刪除使用者：" + textBox1.Text + "的資料?", "警告", MessageBoxButtons.OKCancel);

            if (result == DialogResult.OK)
            {

                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.Users
                        where c.UserName == textBox1.Text
                        orderby c.UserName
                        select c;

                    if (q.Count() > 0)
                    {
                        db.Users
                          .Delete(p => p.UserName == textBox1.Text);

                        textBox1.Text = "";
                        textBox2.Text = "";
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

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) ||
                (e.KeyChar >= (Char)65 && e.KeyChar <= (Char)90) ||
                (e.KeyChar >= (Char)97 && e.KeyChar <= (Char)122) ||
                 e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
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

                    Brush brush = new SolidBrush(cbx.ForeColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;

                    //重繪字串
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
