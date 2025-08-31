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
    public partial class blow_info : Form
    {
        Mat input = new Mat(new Size(1920, 1200), MatType.CV_8UC3, Scalar.Black);
        public blow_info()
        {
            InitializeComponent();
        }
        private void db_load()
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Blows
                    where c.Type == comboBox5.Text
                    orderby c.Type, c.Stop
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dataGridView1.Rows.Add(c.Type, c.Name, c.Time, c.Stop, c.ChineseName);
                    }
                }
            }

            if (comboBox3.Text != "")
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value.ToString() == comboBox3.Text &&
                        dataGridView1.Rows[i].Cells[1].Value.ToString() == comboBox1.Text &&
                        dataGridView1.Rows[i].Cells[3].Value.ToString() == comboBox2.Text)
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
                        break;
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
                    from c in db.Blows
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

        private void button2_Click(object sender, EventArgs e)
        {
            db_load();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
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
                comboBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                comboBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                comboBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                comboBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
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
                        from c in db.Blows
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
            else
            {
                if (!comboBox3.Enabled)
                {
                    using (var db = new MydbDB())
                    {
                        db.Blows
                        .Where(p => p.Type == comboBox3.Text && p.Name == comboBox1.Text && p.Stop == int.Parse(comboBox2.Text))
                        .Set(p => p.Name, comboBox1.Text)
                        .Set(p => p.Time, int.Parse(textBox3.Text))
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
                                from c in db.Blows
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
                                        db.Blows
                                           .Value(p => p.Type, comboBox3.Text)
                                           .Value(p => p.Name, item)
                                           .Value(p => p.Time, 0)
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
                        from c in db.Blows
                        where c.Type == comboBox3.Text
                        orderby c.Type, c.Stop
                        select c;

                    if (q.Count() > 0)
                    {
                        db.Blows
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

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = comboBox4.SelectedIndex;
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Blows
                    where c.Type==(comboBox5.Text)
                    orderby c.Type, c.Stop
                    select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dataGridView1.Rows.Add(c.Type, c.Name, c.Time, c.Stop, c.ChineseName);
                    }
                }
            }
        }
    }
}
