using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace peilin
{
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.DialogResult = DialogResult.OK;
            Close();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.Focus();                
                button1_Click(sender, e);
                textBox1.Focus();
            }
        }
        public string TextBoxMsg
        {
            set
            {
                textBox1.Text = value;
            }
            get
            {
                if (button1.DialogResult == DialogResult.OK)
                {
                    return textBox1.Text;
                }
                else
                {
                    return "無";
                }
            }
        }
    }
}
