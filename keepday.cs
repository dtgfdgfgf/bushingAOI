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
    public partial class keepday : Form
    {
        public keepday()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            using (var db = new MydbDB())
            {                
                db.Parameters.Where(p => p.Name == "KeepDay").Set(p => p.Value, textBox6.Text).Update();
            }
            app.paramUpdate = true;
            Close();
        }

        private void user_param_Load(object sender, EventArgs e)
        {
            textBox6.Text = app.param["KeepDay"];
        }
    }
}
