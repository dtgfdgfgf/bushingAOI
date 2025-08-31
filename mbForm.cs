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
    public partial class mbForm : Form
    {
        public mbForm()
        {
            InitializeComponent();
        }
        public mbForm(string caption,string content) 
        {
            InitializeComponent();

            this.Text = caption;
            this.label1.Text = content;
            this.Height = 50;
            this.Width = label1.Width + 24;            
        }
    }
}
