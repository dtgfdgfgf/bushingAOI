using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PLC;

namespace peilin
{
    public partial class alert : Form
    {
        Panel stop_panel = new Panel();
        Label stop_text = new Label();
        Button stop_button = new Button();

        public alert()
        {
            InitializeComponent();
        }

        private void alert_Load(object sender, EventArgs e)
        {
            this.TopMost = true;

            stop_panel.BackColor = Color.Yellow;
            stop_panel.Enabled = true;
            stop_panel.Visible = false;
            stop_panel.Size = new System.Drawing.Size(500, 150);
            stop_panel.Location = new System.Drawing.Point(this.Width / 2 - stop_panel.Width / 2, this.Height / 2 - stop_panel.Height / 2);
            stop_panel.Name = "stop_panel";
            stop_panel.TabIndex = 322;
            this.Controls.Add(stop_panel);
            stop_panel.BringToFront();

            stop_text.Visible = true;
            stop_text.Font = new System.Drawing.Font("微軟正黑體", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            stop_text.ForeColor = System.Drawing.Color.Black;
            stop_text.Size = new System.Drawing.Size(500, 50);
            stop_text.Location = new System.Drawing.Point(stop_panel.Width / 2 - stop_text.Width / 2, 10);
            stop_text.TextAlign = ContentAlignment.TopCenter;
            stop_text.Name = "stop_text";
            stop_text.TabIndex = 323;
            stop_text.Text = "急停中，請先排除錯誤";
            stop_panel.Controls.Add(stop_text);
            stop_text.BringToFront();

            stop_button.BackColor = System.Drawing.SystemColors.Control;
            stop_button.Visible = true;
            stop_button.Font = new System.Drawing.Font("微軟正黑體", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            stop_button.ForeColor = System.Drawing.Color.Black;
            stop_button.Size = new System.Drawing.Size(200, 50);
            stop_button.Location = new System.Drawing.Point(stop_text.Width / 2 - stop_button.Width / 2, 80);
            stop_button.Name = "stop_button";
            stop_button.TabIndex = 324;
            stop_button.Text = "解除鎖定";
            stop_button.Click += new System.EventHandler(this.stop_button_Click);
            stop_panel.Controls.Add(stop_button);
            stop_button.BringToFront();

            app.status = false;
        }
        private void stop_button_Click(object sender, EventArgs e)
        {
            Form1.PLC_SetM(25, true);            //解除警報
            timer1.Enabled = false;
            Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!Form1.PLC_CheckX(0))
            {
                BeginInvoke(new Action(() => stop_text.Text = "完成急停錯誤排除後，按下解除鎖定"));
                BeginInvoke(new Action(() => stop_button.Enabled = false));
                stop_panel.Location = new System.Drawing.Point(this.Width / 2 - stop_panel.Width / 2, this.Height / 2 - stop_panel.Height / 2);
                BeginInvoke(new Action(() => stop_panel.Visible = true));
            }
            else if (Form1.PLC_CheckM(27)) //培霖是23
            {
                BeginInvoke(new Action(() => stop_text.Text = "門已開啟，請關上門後重新 啟動"));
                BeginInvoke(new Action(() => stop_button.Enabled = false));
                stop_panel.Location = new System.Drawing.Point(this.Width / 2 - stop_panel.Width / 2, this.Height / 2 - stop_panel.Height / 2);
                BeginInvoke(new Action(() => stop_panel.Visible = true));
            }
            else if (Form1.PLC_CheckM(21))
            {
                BeginInvoke(new Action(() => stop_panel.Visible = true));
                BeginInvoke(new Action(() => stop_text.Text = GetErrorMessage()));
                BeginInvoke(new Action(() => stop_button.Enabled = true));
                //Console.WriteLine("[TICK] 應該可以按解除鎖定");
            }

        }
        private string GetErrorMessage()
        {
            if (Form1.PLC_CheckM(800)) return "緊急停止 - 請檢查急停按鈕";
            if (Form1.PLC_CheckM(801)) return "固定急停 - 請檢查固定急停開關";
            if (Form1.PLC_CheckM(802)) return "輸送帶急停1 - 請檢查輸送帶";
            if (Form1.PLC_CheckM(804)) return "門禁急停 - 請檢查安全門是否關閉";
            if (Form1.PLC_CheckM(805)) return "滿料急停 - 請檢查裝箱是否滿料";
            if (Form1.PLC_CheckM(806)) return "汽缸未歸位 - 請檢查汽缸/紙箱位置";
            if (Form1.PLC_CheckM(807)) return "卡料急停 - 請檢查圓震料道通行";
            if (Form1.PLC_CheckM(808)) return "儲料斗急停 - 請檢查儲料斗是否沒料";
            if (Form1.PLC_CheckM(809)) return "沒箱子急停 - 請檢查紙箱位置";
            if (Form1.PLC_CheckM(810)) return "NG/NULL滿料急停 - 請檢查NG//NULL裝箱是否已滿";

            return "急停中，請先排除錯誤"; // 預設訊息
        }
    }
}
