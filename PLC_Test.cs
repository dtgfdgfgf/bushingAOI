using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Concurrent;
using System.IO.Ports;
using System.Threading;
using Serilog;

namespace PLC
{
    public partial class PLC_Test : Form
    {
        public PLC_Test()
        {
            InitializeComponent();
        }

        private void PLC_Test_Load(object sender, EventArgs e)
        {
        }

        private void Point_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton B = sender as RadioButton;
            comboBox1.Items.Clear();
            comboBox1.Text = "";
            switch (B.Text)
            {
                case "X":
                {
                    for (int i = 0; i < 18; i++)
                    {
                        for (int j = 0; j < 8; j++)
                        {
                            comboBox1.Items.Add("X" + (i * 10 + j));
                        }
                    }
                }
                    break;
                case "Y":
                {
                    for (int i = 0; i < 18; i++)
                    {
                        for (int j = 0; j < 8; j++)
                        {
                            comboBox1.Items.Add("Y" + (i * 10 + j));
                        }
                    }
                }
                    break;
                case "M":
                {
                    for (int i = 0; i < 9512; i++)
                    {
                        comboBox1.Items.Add("M" + i);
                    }
                }
                    break;
                case "T":
                {
                    for (int i = 0; i < 512; i++)
                    {
                        comboBox1.Items.Add("T" + i);
                    }
                }
                    break;
                case "C":
                {
                    for (int i = 0; i < 256; i++)
                    {
                        comboBox1.Items.Add("C" + i);
                    }
                }
                    break;
                case "S":
                {
                    for (int i = 0; i < 4096; i++)
                    {
                        comboBox1.Items.Add("S" + i);
                    }
                }
                    break;
                case "EC1X":
                {
                    for (int i = 0; i < 8; i++)
                    {
                        comboBox1.Items.Add("EC1X" + i);
                    }
                }
                    break;
                case "EC1Y":
                {
                    for (int i = 0; i < 8; i++)
                    {
                        comboBox1.Items.Add("EC1Y" + i);
                    }
                }
                    break;
                case "EC2X":
                {
                    for (int i = 0; i < 8; i++)
                    {
                        comboBox1.Items.Add("EC2X" + i);
                    }
                }
                    break;
                case "EC2Y":
                {
                    for (int i = 0; i < 8; i++)
                    {
                        comboBox1.Items.Add("EC2Y" + i);
                    }
                }
                    break;
                case "EC3X":
                {
                    for (int i = 0; i < 8; i++)
                    {
                        comboBox1.Items.Add("EC3X" + i);
                    }
                }
                    break;
                case "EC3Y":
                {
                    for (int i = 0; i < 8; i++)
                    {
                        comboBox1.Items.Add("EC3Y" + i);
                    }
                }
                    break;
            }
        }

        private void Value_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton B = sender as RadioButton;
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            switch (B.Text)
            {
                case "D":
                {
                    for (int i = 0; i < 9512; i++)
                    {
                        comboBox2.Items.Add("D" + i);
                    }
                }
                    break;
                case "T":
                {
                    for (int i = 0; i < 512; i++)
                    {
                        comboBox2.Items.Add("T" + i);
                    }
                }
                    break;
                case "C_16":
                {
                    for (int i = 0; i < 200; i++)
                    {
                        comboBox2.Items.Add("C" + i);
                    }
                }
                    break;
                case "C_32":
                {
                    for (int i = 200; i < 256; i++)
                    {
                        comboBox2.Items.Add("C" + i);
                    }
                }
                    break;
                case "R":
                {
                    for (int i = 0; i < 10000; i++)
                    {
                        comboBox2.Items.Add("R" + i);
                    }
                }
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && !radioButton0.Checked)
            {
                foreach (var item in groupBox1.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];

                            try
                            {
                                int p = 0, n = 0;
                                p = int.Parse(s);
                                if (p >= 6)
                                {
                                    n = int.Parse(comboBox1.Text.Substring(4, comboBox1.Text.Length - 4));
                                    if (control.Text.Substring(0, 4) != comboBox1.Text.Substring(0, 4))
                                    {
                                        MessageBox.Show("編號錯誤");
                                        break;
                                    }
                                }
                                else
                                {
                                    n = int.Parse(comboBox1.Text.Substring(1, comboBox1.Text.Length - 1));
                                    if (control.Text.Substring(0, 1) != comboBox1.Text.Substring(0, 1))
                                    {
                                        MessageBox.Show("編號錯誤");
                                        break;
                                    }
                                }

                                PLC_ModBus.SetPoint(1, (BoolUnit)p, n, true, false, null);

                                break;
                            }
                            catch
                            {
                                MessageBox.Show("編號錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && !radioButton0.Checked)
            {
                foreach (var item in groupBox1.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];
                            try
                            {
                                int p = 0, n = 0;
                                p = int.Parse(s);
                                if (p >= 6)
                                {
                                    n = int.Parse(comboBox1.Text.Substring(4, comboBox1.Text.Length - 4));
                                    if (control.Text.Substring(0, 4) != comboBox1.Text.Substring(0, 4))
                                    {
                                        MessageBox.Show("編號錯誤");
                                        break;
                                    }
                                }
                                else
                                {
                                    n = int.Parse(comboBox1.Text.Substring(1, comboBox1.Text.Length - 1));
                                    if (control.Text.Substring(0, 1) != comboBox1.Text.Substring(0, 1))
                                    {
                                        MessageBox.Show("編號錯誤");
                                        break;
                                    }
                                }

                                PLC_ModBus.SetPoint(1, (BoolUnit)p, n, false, false, null);

                                break;
                            }
                            catch
                            {
                                MessageBox.Show("編號錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                foreach (var item in groupBox1.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];
                            try
                            {
                                int p = 0, n = 0;
                                p = int.Parse(s);
                                if (p >= 6)
                                {
                                    n = int.Parse(comboBox1.Text.Substring(4, comboBox1.Text.Length - 4));
                                    if (control.Text.Substring(0, 4) != comboBox1.Text.Substring(0, 4))
                                    {
                                        MessageBox.Show("編號錯誤");
                                        break;
                                    }
                                }
                                else
                                {
                                    n = int.Parse(comboBox1.Text.Substring(1, comboBox1.Text.Length - 1));
                                    if (control.Text.Substring(0, 1) != comboBox1.Text.Substring(0, 1))
                                    {
                                        MessageBox.Show("編號錯誤");
                                        break;
                                    }
                                }

                                ManualResetEvent Handle = new ManualResetEvent(false);
                                PLC_ModBus.CheckPoint(1, (BoolUnit)p, n, 1, true, Handle);
                                if (Handle.WaitOne(500))
                                {
                                    label2.Text = comboBox1.Text;
                                    switch ((BoolUnit)p)
                                    {
                                        case BoolUnit.X:
                                        {
                                            label4.Text = PLC_Value.Point_X[n / 10][n % 10].ToString();
                                        }
                                            break;
                                        case BoolUnit.Y:
                                        {
                                            label4.Text = PLC_Value.Point_Y[n / 10][n % 10].ToString();
                                        }
                                            break;
                                        case BoolUnit.M:
                                        {
                                            label4.Text = PLC_Value.Point_M[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.S:
                                        {
                                            label4.Text = PLC_Value.Point_S[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.T:
                                        {
                                            label4.Text = PLC_Value.Point_T[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.C:
                                        {
                                            label4.Text = PLC_Value.Point_C[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.EC1X:
                                        {
                                            label4.Text = PLC_Value.Point_EC1X[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.EC1Y:
                                        {
                                            label4.Text = PLC_Value.Point_EC1Y[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.EC2X:
                                        {
                                            label4.Text = PLC_Value.Point_EC2X[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.EC2Y:
                                        {
                                            label4.Text = PLC_Value.Point_EC2Y[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.EC3X:
                                        {
                                            label4.Text = PLC_Value.Point_EC3X[n].ToString();
                                        }
                                            break;
                                        case BoolUnit.EC3Y:
                                        {
                                            label4.Text = PLC_Value.Point_EC3Y[n].ToString();
                                        }
                                            break;
                                    }
                                }
                                else
                                    MessageBox.Show("PLC回應延遲");


                                break;
                            }
                            catch
                            {
                                MessageBox.Show("編號錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && textBox1.Text != "")
            {
                foreach (var item in groupBox2.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];

                            try
                            {
                                int p = 0, n = 0, v = 0;
                                v = int.Parse(textBox1.Text);
                                p = int.Parse(s) - 20;
                                n = int.Parse(comboBox2.Text.Substring(1, comboBox2.Text.Length - 1));
                                if (p == 3)
                                {
                                    if (v > int.MaxValue || v < int.MinValue)
                                        throw new Exception();
                                }
                                else
                                {
                                    if (v > Int16.MaxValue || v < Int16.MinValue)
                                        throw new Exception();
                                }

                                if (control.Text.Substring(0, 1) != comboBox2.Text.Substring(0, 1))
                                {
                                    MessageBox.Show("編號或數值錯誤");
                                    break;
                                }

                                PLC_ModBus.SendValue(1, (ValueUnit)p, n, v, false, null);

                                break;
                            }
                            catch
                            {
                                MessageBox.Show("編號或數值錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && textBox1.Text != "")
            {
                foreach (var item in groupBox2.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];

                            try
                            {
                                int p = 0, n = 0, v = 0;
                                v = int.Parse(textBox1.Text);
                                p = int.Parse(s) - 20;
                                n = int.Parse(comboBox2.Text.Substring(1, comboBox2.Text.Length - 1));
                                if (v > int.MaxValue || v < int.MinValue)
                                    throw new Exception();
                                if (control.Text.Substring(0, 1) != comboBox2.Text.Substring(0, 1))
                                {
                                    MessageBox.Show("編號或數值錯誤");
                                    break;
                                }

                                PLC_ModBus.SendValue_32bit(1, (ValueUnit)p, n, v, false, null);
                                break;
                            }
                            catch
                            {
                                MessageBox.Show("編號或數值錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
            {
                if (radioButton23.Checked)
                    label6.Text = "數值(32bit)";
                else
                    label6.Text = "數值";
                foreach (var item in groupBox2.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];

                            try
                            {
                                int p = 0, n = 0;
                                p = int.Parse(s) - 20;
                                n = int.Parse(comboBox2.Text.Substring(1, comboBox2.Text.Length - 1));
                                if (control.Text.Substring(0, 1) != comboBox2.Text.Substring(0, 1))
                                {
                                    Log.Warning("編號或數值錯誤");
                                    break;
                                }

                                ManualResetEvent Handle = new ManualResetEvent(false);
                                PLC_ModBus.CheckValue(1, (ValueUnit)p, n, 1, false, Handle);
                                if (Handle.WaitOne(500))
                                {
                                    label7.Text = comboBox2.Text;
                                    switch ((ValueUnit)p)
                                    {
                                        case ValueUnit.D:
                                        {
                                            label5.Text = PLC_Value.Value_D[n].ToString();
                                        }
                                            break;
                                        case ValueUnit.R:
                                        {
                                            label5.Text = PLC_Value.Value_R[n].ToString();
                                        }
                                            break;
                                        case ValueUnit.T:
                                        {
                                            label5.Text = PLC_Value.Value_T[n].ToString();
                                        }
                                            break;
                                        case ValueUnit.C_16bit:
                                        case ValueUnit.C_32bit:
                                        {
                                            label5.Text = PLC_Value.Value_C[n].ToString();
                                        }
                                            break;
                                    }
                                }
                                else
                                    Log.Warning("PLC回應延遲");


                                break;
                            }
                            catch
                            {
                                Log.Warning("編號或數值錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
            {
                label6.Text = "數值(32bit)";
                foreach (var item in groupBox2.Controls)
                {
                    if (item.GetType().ToString() == "System.Windows.Forms.RadioButton")
                    {
                        var control = item as RadioButton;
                        if (control.Checked)
                        {
                            string s = control.Name.Split(new string[] { "radioButton" }, 3,
                                StringSplitOptions.RemoveEmptyEntries)[0];

                            try
                            {
                                int p = 0, n = 0;
                                p = int.Parse(s) - 20;
                                n = int.Parse(comboBox2.Text.Substring(1, comboBox2.Text.Length - 1));
                                if (control.Text.Substring(0, 1) != comboBox2.Text.Substring(0, 1))
                                {
                                    MessageBox.Show("編號或數值錯誤");
                                    break;
                                }

                                ManualResetEvent Handle = new ManualResetEvent(false);
                                if ((ValueUnit)p == ValueUnit.C_32bit)
                                    PLC_ModBus.CheckValue(1, (ValueUnit)p, n, 1, false, Handle);
                                else
                                    PLC_ModBus.CheckValue(1, (ValueUnit)p, n, 2, false, Handle);

                                if (Handle.WaitOne(500))
                                {
                                    label7.Text = comboBox2.Text;
                                    switch ((ValueUnit)p)
                                    {
                                        case ValueUnit.D:
                                        {
                                            int Value = 0;
                                            if (PLC_Value.Value_D[n] < 0)
                                            {
                                                Value = 65536 + PLC_Value.Value_D[n] + PLC_Value.Value_D[n + 1] * 65536;
                                            }
                                            else
                                            {
                                                Value = PLC_Value.Value_D[n] + PLC_Value.Value_D[n + 1] * 65536;
                                            }

                                            label5.Text = Value.ToString();
                                        }
                                            break;
                                        case ValueUnit.R:
                                        {
                                            int Value = 0;
                                            if (PLC_Value.Value_R[n] < 0)
                                            {
                                                Value = 65536 + PLC_Value.Value_R[n] + PLC_Value.Value_R[n + 1] * 65536;
                                            }
                                            else
                                            {
                                                Value = PLC_Value.Value_R[n] + PLC_Value.Value_R[n + 1] * 65536;
                                            }

                                            label5.Text = Value.ToString();
                                        }
                                            break;
                                        case ValueUnit.T:
                                        {
                                            int Value = 0;
                                            if (PLC_Value.Value_T[n] < 0)
                                            {
                                                Value = 65536 + PLC_Value.Value_T[n] + PLC_Value.Value_T[n + 1] * 65536;
                                            }
                                            else
                                            {
                                                Value = PLC_Value.Value_T[n] + PLC_Value.Value_T[n + 1] * 65536;
                                            }

                                            label5.Text = Value.ToString();
                                        }
                                            break;
                                        //case ValueUnit.C_16bit:
                                        //    {
                                        //        int Value = 0;
                                        //        if (PLC_Value.Value_C2[n] < 0)
                                        //        {
                                        //            Value = 65536 + PLC_Value.Value_C[n] + PLC_Value.Value_C[n + 1] * 65536;
                                        //        }
                                        //        else
                                        //        {
                                        //            Value = PLC_Value.Value_C[n] + PLC_Value.Value_C[n + 1] * 65536;
                                        //        }
                                        //        label5.Text = Value.ToString();
                                        //    }
                                        //    break;
                                        case ValueUnit.C_16bit:
                                        case ValueUnit.C_32bit:
                                        {
                                            label5.Text = PLC_Value.Value_C[n].ToString();
                                        }
                                            break;
                                    }
                                }
                                else
                                    Log.Warning("PLC回應延遲");


                                break;
                            }
                            catch
                            {
                                Log.Warning("編號或數值錯誤");
                                break;
                            }
                        }
                    }
                }
            }
        }
    }

    public static class PLC_ModBus
    {
        #region 宣告

        public static ManualResetEvent Action_Handle;
        public static string COM = "COM4";
        public static Protocol Protocol_Type = Protocol.ASCII;
        private static SerialPort com;
        private static ConcurrentQueue<CommentData> SendData = new ConcurrentQueue<CommentData>();
        private static ConcurrentQueue<CommentData> SendData_Emergency = new ConcurrentQueue<CommentData>();
        private static EventWaitHandle Handle = new EventWaitHandle(false, EventResetMode.ManualReset);
        private static byte[] Comment = new byte[1024];
        private static byte[] Temp_Comment = new byte[0];
        private static ValueUnit Check_ValueUnit = 0;
        private static BoolUnit Check_BoolUnit = 0;
        private static int Check_ValueUnit_Stop = 0;
        private static int Check_ValueUnit_Start = 0;
        private static int Check_ValueUnit_Lenth = 0;
        private static int Check_BoolUnit_Stop = 0;
        private static int Check_BoolUnit_Start = 0;
        private static int Check_BoolUnit_Lenth = 0;
        private static bool PLC_Run = false;
        public static bool PLC_Reboot = false;
        private static bool PLC_SendComment = false;
        private static bool PLC_ShowData = false;
        private static Task T = new Task(new Action(() => SendComment()), TaskCreationOptions.LongRunning);
        private static readonly object Locker = new object();
        public static AutoResetEvent event1 = new AutoResetEvent(false);
        public static AutoResetEvent event2 =  new AutoResetEvent(false);
        public static AutoResetEvent event3 =  new AutoResetEvent(false);

        #endregion

        #region 列舉

        public enum Contect_M
        {
            /// <summary>
            /// 轉盤
            /// </summary>
            //Motor = 0,
            //Vibrat_Plate = 1,
            //Red_light = 2,
            //Green_light = 3,
            //BlowStop1 = 4,
            //BlowStop2 = 5,
            //BlowStop3 = 6,
            //BlowStop4 = 7,
            //BlowStop5 = 10,
            //BlowStop6 = 11
        }

        public enum Contect_D
        {
        }

        public enum Contect_X
        {
        }

        public enum Contect_Y
        {
        }

        public enum Contect_S
        {
        }

        public enum Contect_T
        {
        }

        public enum Contect_C
        {
        }

        public enum Contect_R
        {
        }

        #endregion

        #region 讀寫

        #region 寫入指令

        //寫值進入暫存器
        public static void SendValue(int Stop, ValueUnit unit, int Number, int Value, bool Emer,
            ManualResetEvent Action_Handle)
        {
            CommentData comment = new CommentData();
            if (Protocol_Type == Protocol.ASCII)
            {
                string Start_Str = "3A ";
                string End_Str = " 0D 0A";
                string Action_Str = "30 36 ";
                string Stop_Str = "";
                string Unit_Str = "";
                string Value_Str = "";
                string LRC_Str = "";
                string Lenth_Str = "";
                string ByteLenth_Str = "";

                #region 站號處理

                if (Stop < 1)
                    throw new Exception("站號錯誤(Stop>0)");
                int StopN1 = 30 + (Stop / 16);
                int StopN2 = 30 + (Stop % 16);
                if (StopN1 >= 40)
                    StopN1++;
                if (StopN2 >= 40)
                    StopN2++;
                Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

                #endregion

                #region 暫存器字串處理

                int UnitN1 = 30;
                int UnitN2 = 30;
                int UnitN3 = 30;
                int UnitN4 = 30;
                switch (unit)
                {
                    case ValueUnit.D:
                    {
                        if (Number < 0 || Number > 9511)
                            throw new Exception("站存器編號錯誤(有效編號D0~D9511)");
                        UnitN1 = 30;
                        UnitN2 = 30;
                        UnitN3 = 30;
                        UnitN4 = 30;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 >= 40)
                            UnitN4++;
                        if (UnitN4 > 46)
                        {
                            UnitN4 -= 17;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 >= 40)
                            UnitN3++;
                        if (UnitN3 > 46)
                        {
                            UnitN3 -= 17;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 >= 40)
                            UnitN2++;
                        if (UnitN2 > 46)
                        {
                            UnitN2 -= 17;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 >= 40)
                            UnitN1++;
                        if (UnitN1 > 46)
                        {
                            UnitN1 = 46;
                        }
                    }
                        break;
                    case ValueUnit.T:
                    {
                        if (Number < 0 || Number > 511)
                            throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                        UnitN1 = 32;
                        UnitN2 = 36;
                        UnitN3 = 30;
                        UnitN4 = 30;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 >= 40)
                            UnitN4++;
                        if (UnitN4 > 46)
                        {
                            UnitN4 -= 17;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 >= 40)
                            UnitN3++;
                        if (UnitN3 > 46)
                        {
                            UnitN3 -= 17;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 >= 40)
                            UnitN2++;
                        if (UnitN2 > 46)
                        {
                            UnitN2 -= 17;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 >= 40)
                            UnitN1++;
                        if (UnitN1 > 46)
                        {
                            UnitN1 = 46;
                        }
                    }
                        break;
                    case ValueUnit.C_16bit:
                    {
                        if (Number < 0 || Number > 199)
                            throw new Exception("站存器編號錯誤(有效編號C0~C199)");
                        UnitN1 = 32;
                        UnitN2 = 38;
                        UnitN3 = 30;
                        UnitN4 = 30;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 >= 40)
                            UnitN4++;
                        if (UnitN4 > 46)
                        {
                            UnitN4 -= 17;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 >= 40)
                            UnitN3++;
                        if (UnitN3 > 46)
                        {
                            UnitN3 -= 17;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 >= 40)
                            UnitN2++;
                        if (UnitN2 > 46)
                        {
                            UnitN2 -= 17;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 >= 40)
                            UnitN1++;
                        if (UnitN1 > 46)
                        {
                            UnitN1 = 46;
                        }
                    }
                        break;
                    case ValueUnit.C_32bit:
                    {
                        if (Number < 200 || Number > 255)
                            throw new Exception("站存器編號錯誤(有效編號C200~C255)");
                        Number -= 200;
                        Number = Number * 2;
                        UnitN1 = 32;
                        UnitN2 = 39;
                        UnitN3 = 30;
                        UnitN4 = 30;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 >= 40)
                            UnitN4++;
                        if (UnitN4 > 46)
                        {
                            UnitN4 -= 17;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 >= 40)
                            UnitN3++;
                        if (UnitN3 > 46)
                        {
                            UnitN3 -= 17;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 >= 40)
                            UnitN2++;
                        if (UnitN2 > 46)
                        {
                            UnitN2 -= 17;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 >= 40)
                            UnitN1++;
                        if (UnitN1 > 46)
                        {
                            UnitN1 = 46;
                        }
                    }
                        break;
                    case ValueUnit.R:
                    {
                        if (Number < 0 || Number > 9999)
                            throw new Exception("站存器編號錯誤(有效編號R0~R9999)");
                        UnitN1 = 32;
                        UnitN2 = 40;
                        UnitN3 = 30;
                        UnitN4 = 30;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 >= 40)
                            UnitN4++;
                        if (UnitN4 > 46)
                        {
                            UnitN4 -= 17;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 >= 40)
                            UnitN3++;
                        if (UnitN3 > 46)
                        {
                            UnitN3 -= 17;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 >= 40)
                            UnitN2++;
                        if (UnitN2 > 46)
                        {
                            UnitN2 -= 17;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 >= 40)
                            UnitN1++;
                        if (UnitN1 > 46)
                        {
                            UnitN1 = 46;
                        }
                    }
                        break;
                }

                Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " +
                           UnitN4.ToString() + " ";

                #endregion

                #region 資料處理

                switch (unit)
                {
                    case ValueUnit.D:
                    case ValueUnit.T:
                    case ValueUnit.R:
                    case ValueUnit.C_16bit:
                    {
                        if (Value > Int16.MaxValue || Value < Int16.MinValue)
                            throw new Exception("數值過大(-32768<=Value<=32767)");

                        if (Value >= 0)
                        {
                            int N1 = 30;
                            int N2 = 30;
                            int N3 = 30;
                            int N4 = 30;
                            int Value_N1 = Value / 16 / 16 / 16 % 16;
                            int Value_N2 = Value / 16 / 16 % 16;
                            int Value_N3 = Value / 16 % 16;
                            int Value_N4 = Value % 16;
                            N4 += Value_N4;
                            if (N4 >= 40)
                                N4++;
                            if (N4 > 46)
                            {
                                N4 -= 17;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 >= 40)
                                N3++;
                            if (N3 > 46)
                            {
                                N3 -= 17;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 >= 40)
                                N2++;
                            if (N2 > 46)
                            {
                                N2 -= 17;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 >= 40)
                                N1++;
                            if (N1 > 46)
                            {
                                N1 = 46;
                            }

                            Value_Str = N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " + N4.ToString();
                        }
                        else
                        {
                            int N1 = 30;
                            int N2 = 30;
                            int N3 = 30;
                            int N4 = 30;
                            int Value_N1 = 15 - (-Value / 16 / 16 / 16 % 16);
                            int Value_N2 = 15 - (-Value / 16 / 16 % 16);
                            int Value_N3 = 15 - (-Value / 16 % 16);
                            int Value_N4 = 16 - (-Value % 16);
                            N4 += Value_N4;
                            if (N4 >= 40)
                                N4++;
                            if (N4 > 46)
                            {
                                N4 -= 17;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 >= 40)
                                N3++;
                            if (N3 > 46)
                            {
                                N3 -= 17;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 >= 40)
                                N2++;
                            if (N2 > 46)
                            {
                                N2 -= 17;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 >= 40)
                                N1++;
                            if (N1 > 46)
                            {
                                N1 = 46;
                            }

                            Value_Str = N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " + N4.ToString();
                        }
                    }
                        break;
                    case ValueUnit.C_32bit:
                    {
                        if (Value <= int.MaxValue || Value >= int.MinValue)
                        {
                            Action_Str = "31 30 ";

                            #region 組數處理

                            int Len = 2;
                            int LenN1 = 30 + (Len / 16 / 16 / 16);
                            int LenN2 = 30 + (Len / 16 / 16);
                            int LenN3 = 30 + (Len / 16);
                            int LenN4 = 30 + (Len % 16);
                            if (LenN1 >= 40)
                                LenN1++;
                            if (LenN2 >= 40)
                                LenN2++;
                            if (LenN3 >= 40)
                                LenN3++;
                            if (LenN4 >= 40)
                                LenN4++;
                            Lenth_Str = LenN1.ToString() + " " + LenN2.ToString() + " " + LenN3.ToString() + " " +
                                        LenN4.ToString() + " ";

                            #endregion

                            #region 組數處理

                            int BLen = Len * 2;
                            int BLenN1 = 30 + (BLen / 16);
                            int BLenN2 = 30 + (BLen % 16);
                            if (BLenN1 >= 40)
                                BLenN1++;
                            if (BLenN2 >= 40)
                                BLenN2++;
                            ByteLenth_Str = BLenN1.ToString() + " " + BLenN2.ToString() + " ";

                            #endregion

                            if (Value >= 0)
                            {
                                int N1 = 30;
                                int N2 = 30;
                                int N3 = 30;
                                int N4 = 30;
                                int N5 = 30;
                                int N6 = 30;
                                int N7 = 30;
                                int N8 = 30;
                                int Value_N1 = Value / 16 / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N2 = Value / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N3 = Value / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N4 = Value / 16 / 16 / 16 / 16 % 16;
                                int Value_N5 = Value / 16 / 16 / 16 % 16;
                                int Value_N6 = Value / 16 / 16 % 16;
                                int Value_N7 = Value / 16 % 16;
                                int Value_N8 = Value % 16;
                                N8 += Value_N8;
                                if (N8 >= 40)
                                    N8++;
                                if (N8 > 46)
                                {
                                    N8 -= 17;
                                    N7++;
                                }

                                N7 += Value_N7;
                                if (N7 >= 40)
                                    N7++;
                                if (N7 > 46)
                                {
                                    N7 -= 17;
                                    N6++;
                                }

                                N6 += Value_N6;
                                if (N6 >= 40)
                                    N6++;
                                if (N6 > 46)
                                {
                                    N6 -= 17;
                                    N5++;
                                }

                                N5 += Value_N5;
                                if (N5 >= 40)
                                    N5++;
                                if (N5 > 46)
                                {
                                    N5 -= 17;
                                    N4++;
                                }

                                N4 += Value_N4;
                                if (N4 >= 40)
                                    N4++;
                                if (N4 > 46)
                                {
                                    N4 -= 17;
                                    N3++;
                                }

                                N3 += Value_N3;
                                if (N3 >= 40)
                                    N3++;
                                if (N3 > 46)
                                {
                                    N3 -= 17;
                                    N2++;
                                }

                                N2 += Value_N2;
                                if (N2 >= 40)
                                    N2++;
                                if (N2 > 46)
                                {
                                    N2 -= 17;
                                    N1++;
                                }

                                N1 += Value_N1;
                                if (N1 >= 40)
                                    N1++;
                                if (N1 > 46)
                                {
                                    N1 = 46;
                                }

                                Value_Str = N5.ToString() + " " + N6.ToString() + " " + N7.ToString() + " " +
                                            N8.ToString() + " " + N1.ToString() + " " + N2.ToString() + " " +
                                            N3.ToString() + " " + N4.ToString();
                            }
                            else
                            {
                                int N = int.MinValue;
                                Value = Value - N;
                                int N1 = 38;
                                int N2 = 30;
                                int N3 = 30;
                                int N4 = 30;
                                int N5 = 30;
                                int N6 = 30;
                                int N7 = 30;
                                int N8 = 30;
                                int Value_N1 = Value / 16 / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N2 = Value / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N3 = Value / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N4 = Value / 16 / 16 / 16 / 16 % 16;
                                int Value_N5 = Value / 16 / 16 / 16 % 16;
                                int Value_N6 = Value / 16 / 16 % 16;
                                int Value_N7 = Value / 16 % 16;
                                int Value_N8 = Value % 16;
                                N8 += Value_N8;
                                if (N8 >= 40)
                                    N8++;
                                if (N8 > 46)
                                {
                                    N8 -= 17;
                                    N7++;
                                }

                                N7 += Value_N7;
                                if (N7 >= 40)
                                    N7++;
                                if (N7 > 46)
                                {
                                    N7 -= 17;
                                    N6++;
                                }

                                N6 += Value_N6;
                                if (N6 >= 40)
                                    N6++;
                                if (N6 > 46)
                                {
                                    N6 -= 17;
                                    N5++;
                                }

                                N5 += Value_N5;
                                if (N5 >= 40)
                                    N5++;
                                if (N5 > 46)
                                {
                                    N5 -= 17;
                                    N4++;
                                }

                                N4 += Value_N4;
                                if (N4 >= 40)
                                    N4++;
                                if (N4 > 46)
                                {
                                    N4 -= 17;
                                    N3++;
                                }

                                N3 += Value_N3;
                                if (N3 >= 40)
                                    N3++;
                                if (N3 > 46)
                                {
                                    N3 -= 17;
                                    N2++;
                                }

                                N2 += Value_N2;
                                if (N2 >= 40)
                                    N2++;
                                if (N2 > 46)
                                {
                                    N2 -= 17;
                                    N1++;
                                }

                                N1 += Value_N1;
                                if (N1 >= 40)
                                    N1++;
                                if (N1 > 46)
                                {
                                    N1 = 46;
                                }

                                Value_Str = N5.ToString() + " " + N6.ToString() + " " + N7.ToString() + " " +
                                            N8.ToString() + " " + N1.ToString() + " " + N2.ToString() + " " +
                                            N3.ToString() + " " + N4.ToString();
                            }
                        }
                        else
                        {
                            throw new Exception("數值過大(-2147483648<=Value<=2147483647)");
                        }
                    }
                        break;
                }

                #endregion

                string Data = Stop_Str + Action_Str + Unit_Str + Lenth_Str + ByteLenth_Str + Value_Str;
                //Console.WriteLine(Data);
                LRC_Str = LRC2(Data);
                LRC_Str = LRC_Str.ToUpper();
                LRC_Str = convert16(LRC_Str).TrimEnd();
                string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
                string hexValuesReplace = Comment.Replace(" ", "");
                byte[] buffer = StringToByteArray(hexValuesReplace);
                comment.Comment = buffer;
                comment.valueUnit = unit;
                comment.stop = Stop;
            }
            else if (Protocol_Type == Protocol.RTU)
            {
                string Action_Str = "06 ";
                string Stop_Str = "";
                string Unit_Str = "";
                string Value_Str = "";
                string Lenth_Str = "";
                string ByteLenth_Str = "";

                #region 站號處理

                if (Stop < 1)
                    throw new Exception("站號錯誤(Stop>0)");
                int StopN1 = (Stop / 16);
                int StopN2 = (Stop % 16);
                Stop_Str = StopN1.ToString("X") + StopN2.ToString("X") + " ";

                #endregion

                #region 暫存器字串處理

                int UnitN1 = 0;
                int UnitN2 = 0;
                int UnitN3 = 0;
                int UnitN4 = 0;
                switch (unit)
                {
                    case ValueUnit.D:
                    {
                        if (Number < 0 || Number > 9511)
                            throw new Exception("站存器編號錯誤(有效編號D0~D9511)");
                        UnitN1 = 0;
                        UnitN2 = 0;
                        UnitN3 = 0;
                        UnitN4 = 0;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 > 15)
                        {
                            UnitN4 -= 16;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 > 15)
                        {
                            UnitN3 -= 16;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 > 15)
                        {
                            UnitN2 -= 16;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 > 15)
                        {
                            UnitN1 = 15;
                        }
                    }
                        break;
                    case ValueUnit.T:
                    {
                        if (Number < 0 || Number > 511)
                            throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                        UnitN1 = 2;
                        UnitN2 = 6;
                        UnitN3 = 0;
                        UnitN4 = 0;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 > 15)
                        {
                            UnitN4 -= 16;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 > 15)
                        {
                            UnitN3 -= 16;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 > 15)
                        {
                            UnitN2 -= 16;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 > 15)
                        {
                            UnitN1 = 15;
                        }
                    }
                        break;
                    case ValueUnit.C_16bit:
                    {
                        if (Number < 0 || Number > 199)
                            throw new Exception("站存器編號錯誤(有效編號C0~C199)");
                        UnitN1 = 2;
                        UnitN2 = 8;
                        UnitN3 = 0;
                        UnitN4 = 0;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 > 15)
                        {
                            UnitN4 -= 16;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 > 15)
                        {
                            UnitN3 -= 16;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 > 15)
                        {
                            UnitN2 -= 16;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 > 15)
                        {
                            UnitN1 = 15;
                        }
                    }
                        break;
                    case ValueUnit.C_32bit:
                    {
                        if (Number < 200 || Number > 255)
                            throw new Exception("站存器編號錯誤(有效編號C200~C255)");
                        Number -= 200;
                        Number = Number * 2;
                        UnitN1 = 2;
                        UnitN2 = 9;
                        UnitN3 = 0;
                        UnitN4 = 0;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 > 15)
                        {
                            UnitN4 -= 16;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 > 15)
                        {
                            UnitN3 -= 16;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 > 15)
                        {
                            UnitN2 -= 16;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 > 15)
                        {
                            UnitN1 = 15;
                        }
                    }
                        break;
                    case ValueUnit.R:
                    {
                        if (Number < 0 || Number > 9999)
                            throw new Exception("站存器編號錯誤(有效編號R0~R9999)");
                        UnitN1 = 2;
                        UnitN2 = 10;
                        UnitN3 = 0;
                        UnitN4 = 0;
                        int Number_N1 = Number / 16 / 16 / 16 % 16;
                        int Number_N2 = Number / 16 / 16 % 16;
                        int Number_N3 = Number / 16 % 16;
                        int Number_N4 = Number % 16;
                        UnitN4 += Number_N4;
                        if (UnitN4 > 15)
                        {
                            UnitN4 -= 16;
                            UnitN3++;
                        }

                        UnitN3 += Number_N3;
                        if (UnitN3 > 15)
                        {
                            UnitN3 -= 16;
                            UnitN2++;
                        }

                        UnitN2 += Number_N2;
                        if (UnitN2 > 15)
                        {
                            UnitN2 -= 16;
                            UnitN1++;
                        }

                        UnitN1 += Number_N1;
                        if (UnitN1 > 15)
                        {
                            UnitN1 = 15;
                        }
                    }
                        break;
                }

                Unit_Str = UnitN1.ToString("X") + UnitN2.ToString("X") + " " + UnitN3.ToString("X") +
                           UnitN4.ToString("X") + " ";

                #endregion

                #region 資料處理

                switch (unit)
                {
                    case ValueUnit.D:
                    case ValueUnit.T:
                    case ValueUnit.R:
                    case ValueUnit.C_16bit:
                    {
                        if (Value > Int16.MaxValue || Value < Int16.MinValue)
                            throw new Exception("數值過大(-32768<=Value<=32767)");

                        if (Value >= 0)
                        {
                            int N1 = 0;
                            int N2 = 0;
                            int N3 = 0;
                            int N4 = 0;
                            int Value_N1 = Value / 16 / 16 / 16 % 16;
                            int Value_N2 = Value / 16 / 16 % 16;
                            int Value_N3 = Value / 16 % 16;
                            int Value_N4 = Value % 16;
                            N4 += Value_N4;
                            if (N4 > 15)
                            {
                                N4 -= 16;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 > 15)
                            {
                                N3 -= 16;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 > 15)
                            {
                                N2 -= 16;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 > 15)
                            {
                                N1 = 15;
                            }

                            Value_Str = N1.ToString("X") + N2.ToString("X") + " " + N3.ToString("X") + N4.ToString("X");
                        }
                        else
                        {
                            int N1 = 0;
                            int N2 = 0;
                            int N3 = 0;
                            int N4 = 0;
                            int Value_N1 = 15 - (-Value / 16 / 16 / 16 % 16);
                            int Value_N2 = 15 - (-Value / 16 / 16 % 16);
                            int Value_N3 = 15 - (-Value / 16 % 16);
                            int Value_N4 = 16 - (-Value % 16);
                            N4 += Value_N4;
                            if (N4 > 15)
                            {
                                N4 -= 16;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 > 15)
                            {
                                N3 -= 16;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 > 15)
                            {
                                N2 -= 16;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 > 15)
                            {
                                N1 = 15;
                            }

                            Value_Str = N1.ToString("X") + N2.ToString("X") + " " + N3.ToString("X") + N4.ToString("X");
                        }
                    }
                        break;
                    case ValueUnit.C_32bit:
                    {
                        if (Value <= int.MaxValue || Value >= int.MinValue)
                        {
                            Action_Str = "10 ";

                            #region 組數處理

                            int Len = 2;
                            int LenN1 = (Len / 16 / 16 / 16);
                            int LenN2 = (Len / 16 / 16);
                            int LenN3 = (Len / 16);
                            int LenN4 = (Len % 16);
                            Lenth_Str = LenN1.ToString("X") + LenN2.ToString("X") + " " + LenN3.ToString("X") +
                                        LenN4.ToString("X") + " ";

                            #endregion

                            #region 組數處理

                            int BLen = Len * 2;
                            int BLenN1 = (BLen / 16);
                            int BLenN2 = (BLen % 16);
                            ByteLenth_Str = BLenN1.ToString("X") + BLenN2.ToString("X") + " ";

                            #endregion

                            if (Value >= 0)
                            {
                                int N1 = 0;
                                int N2 = 0;
                                int N3 = 0;
                                int N4 = 0;
                                int N5 = 0;
                                int N6 = 0;
                                int N7 = 0;
                                int N8 = 0;
                                int Value_N1 = Value / 16 / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N2 = Value / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N3 = Value / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N4 = Value / 16 / 16 / 16 / 16 % 16;
                                int Value_N5 = Value / 16 / 16 / 16 % 16;
                                int Value_N6 = Value / 16 / 16 % 16;
                                int Value_N7 = Value / 16 % 16;
                                int Value_N8 = Value % 16;
                                N8 += Value_N8;
                                if (N8 > 15)
                                {
                                    N8 -= 16;
                                    N7++;
                                }

                                N7 += Value_N7;
                                if (N7 > 15)
                                {
                                    N7 -= 16;
                                    N6++;
                                }

                                N6 += Value_N6;
                                if (N6 > 15)
                                {
                                    N6 -= 16;
                                    N5++;
                                }

                                N5 += Value_N5;
                                if (N5 > 15)
                                {
                                    N5 -= 16;
                                    N4++;
                                }

                                N4 += Value_N4;
                                if (N4 > 15)
                                {
                                    N4 -= 16;
                                    N3++;
                                }

                                N3 += Value_N3;
                                if (N3 > 15)
                                {
                                    N3 -= 16;
                                    N2++;
                                }

                                N2 += Value_N2;
                                if (N2 > 15)
                                {
                                    N2 -= 16;
                                    N1++;
                                }

                                N1 += Value_N1;
                                if (N1 > 15)
                                {
                                    N1 = 15;
                                }

                                Value_Str = N5.ToString("X") + N6.ToString("X") + " " + N7.ToString("X") +
                                            N8.ToString("X") + " " + N1.ToString("X") + N2.ToString("X") + " " +
                                            N3.ToString("X") + N4.ToString("X");
                            }
                            else
                            {
                                int N = int.MinValue;
                                Value = Value - N;
                                int N1 = 8;
                                int N2 = 0;
                                int N3 = 0;
                                int N4 = 0;
                                int N5 = 0;
                                int N6 = 0;
                                int N7 = 0;
                                int N8 = 0;
                                int Value_N1 = Value / 16 / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N2 = Value / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N3 = Value / 16 / 16 / 16 / 16 / 16 % 16;
                                int Value_N4 = Value / 16 / 16 / 16 / 16 % 16;
                                int Value_N5 = Value / 16 / 16 / 16 % 16;
                                int Value_N6 = Value / 16 / 16 % 16;
                                int Value_N7 = Value / 16 % 16;
                                int Value_N8 = Value % 16;
                                N8 += Value_N8;
                                if (N8 > 15)
                                {
                                    N8 -= 16;
                                    N7++;
                                }

                                N7 += Value_N7;
                                if (N7 > 15)
                                {
                                    N7 -= 16;
                                    N6++;
                                }

                                N6 += Value_N6;
                                if (N6 > 15)
                                {
                                    N6 -= 16;
                                    N5++;
                                }

                                N5 += Value_N5;
                                if (N5 > 15)
                                {
                                    N5 -= 16;
                                    N4++;
                                }

                                N4 += Value_N4;
                                if (N4 > 15)
                                {
                                    N4 -= 16;
                                    N3++;
                                }

                                N3 += Value_N3;
                                if (N3 > 15)
                                {
                                    N3 -= 16;
                                    N2++;
                                }

                                N2 += Value_N2;
                                if (N2 > 15)
                                {
                                    N2 -= 16;
                                    N1++;
                                }

                                N1 += Value_N1;
                                if (N1 > 15)
                                {
                                    N1 = 15;
                                }

                                Value_Str = N5.ToString("X") + N6.ToString("X") + " " + N7.ToString("X") +
                                            N8.ToString("X") + " " + N1.ToString("X") + N2.ToString("X") + " " +
                                            N3.ToString("X") + N4.ToString();
                            }
                        }
                        else
                        {
                            throw new Exception("數值過大(-2147483648<=Value<=2147483647)");
                        }
                    }
                        break;
                }

                #endregion

                string Data = Stop_Str + Action_Str + Unit_Str + Lenth_Str + ByteLenth_Str + Value_Str;

                Console.WriteLine(Data);
                UInt16 _CRC = CRC(Data);
                string hexValue = _CRC.ToString("X2");
                string loCRC, hiCRC;
                if (hexValue.Length == 3)
                {
                    loCRC = hexValue.Substring(1, 2); //crc low byte   
                    hiCRC = hexValue.Substring(0, 1); //crc high byte  
                }
                else
                {
                    loCRC = hexValue.Substring(2, 2); //crc low byte   
                    hiCRC = hexValue.Substring(0, 2); //crc high byte  
                }
                //Console.WriteLine(loCRC);
                //Console.WriteLine(hiCRC);

                string Comment = Data + " " + loCRC + " " + hiCRC;
                string hexValuesReplace = Comment.Replace(" ", "");
                byte[] buffer = StringToByteArray(hexValuesReplace);

                comment.Comment = buffer;
                comment.valueUnit = unit;
                comment.stop = Stop;
            }

            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (Action_Handle != null)
                {
                    SendData.Enqueue(comment);
                }
                else
                {
                    if (!SendData.Contains(comment))
                        SendData.Enqueue(comment);
                }
            }
            //Console.WriteLine(Comment);
        }

        //連續寫入點位
        public static void SetPoint(int Stop, BoolUnit unit, int Number, List<bool> list, bool Emer,
            ManualResetEvent Action_Handle)
        {
            string Start_Str = "3A ";
            string End_Str = " 0D 0A";
            string Action_Str = "30 46 ";
            string Stop_Str = "";
            string Unit_Str = "";
            string Value_Str = "";
            string LRC_Str = "";
            string Lenth_Str = "";
            string ByteLenth_Str = "";
            List<string> slist = new List<string>();

            #region 站號處理

            if (Stop < 1)
                throw new Exception("站號錯誤(Stop>0)");
            int StopN1 = 30 + (Stop / 16);
            int StopN2 = 30 + (Stop % 16);
            if (StopN1 >= 40)
                StopN1++;
            if (StopN2 >= 40)
                StopN2++;
            Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

            #endregion

            #region 暫存器字串處理

            int UnitN1 = 30;
            int UnitN2 = 30;
            int UnitN3 = 30;
            int UnitN4 = 30;
            switch (unit)
            {
                case BoolUnit.Y:
                {
                    if (Number < 0 || Number > 177)
                        throw new Exception("站存器編號錯誤(有效編號Y0~Y177)");
                    if (Number % 10 > 7)
                        throw new Exception("站存器編號錯誤(有效編號尾數Y0~Y7)");
                    Number = (Number / 100 % 10 * 8 * 8) + (Number / 10 % 10 * 8) + Number % 10;
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.M:
                {
                    if (Number < 0 || (Number > 8191 && Number < 9000) || Number > 9511)
                        throw new Exception("站存器編號錯誤(有效編號M0~M8191、M9000~M9511)");
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.S:
                {
                    if (Number < 0 || Number > 4095)
                        throw new Exception("站存器編號錯誤(有效編號S0~S4095)");
                    UnitN1 = 32;
                    UnitN2 = 36;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.T:
                {
                    if (Number < 0 || Number > 511)
                        throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                    UnitN1 = 30;
                    UnitN2 = 31;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.C:
                {
                    if (Number < 0 || Number > 255)
                        throw new Exception("站存器編號錯誤(有效編號C0~C255)");
                    UnitN1 = 30;
                    UnitN2 = 33;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC1X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC1X0~EC1X7)");
                    Number += 9260;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC1Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC1Y0~EC1Y7)");
                    Number += 9270;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC2X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC2X0~EC2X7)");
                    Number += 9280;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC2Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC2Y0~EC2Y7)");
                    Number += 9290;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC3X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC3X0~EC3X7)");
                    Number += 9300;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC3Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC3Y0~EC3Y7)");
                    Number += 9310;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.X:
                {
                    throw new Exception("X點位不能做寫入動作)");
                }
            }

            Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " + UnitN4.ToString() +
                       " ";

            #endregion

            #region 組數處理

            int Len = list.Count;
            int LenN1 = 30 + (Len / 16 / 16 / 16);
            int LenN2 = 30 + (Len / 16 / 16);
            int LenN3 = 30 + (Len / 16);
            int LenN4 = 30 + (Len % 16);
            if (LenN1 >= 40)
                LenN1++;
            if (LenN2 >= 40)
                LenN2++;
            if (LenN3 >= 40)
                LenN3++;
            if (LenN4 >= 40)
                LenN4++;
            Lenth_Str = LenN1.ToString() + " " + LenN2.ToString() + " " + LenN3.ToString() + " " + LenN4.ToString() +
                        " ";

            #endregion

            #region 組數處理

            string s = "";
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i])
                    s = 1 + s;
                else
                    s = 0 + s;
                if (i % 8 == 7 || i == list.Count - 1)
                {
                    slist.Add(s);
                    s = "";
                }
            }

            int BLen = slist.Count;
            int BLenN1 = 30 + (BLen / 16);
            int BLenN2 = 30 + (BLen % 16);
            if (BLenN1 >= 40)
                BLenN1++;
            if (BLenN2 >= 40)
                BLenN2++;
            ByteLenth_Str = BLenN1.ToString() + " " + BLenN2.ToString() + " ";

            #endregion

            #region 資料處理

            for (int i = 0; i < slist.Count; i++)
            {
                int Value = Convert.ToInt32(slist[i], 2);
                int N1 = 30;
                int N2 = 30;
                int Value_N1 = Value / 16 % 16;
                int Value_N2 = Value % 16;
                N2 += Value_N2;
                if (N2 >= 40)
                    N2++;
                if (N2 > 46)
                {
                    N2 -= 17;
                    N1++;
                }

                N1 += Value_N1;
                if (N1 >= 40)
                    N1++;
                if (N1 > 46)
                {
                    N1 = 46;
                }

                Value_Str += N1.ToString() + " " + N2.ToString();
                if (i != slist.Count - 1)
                    Value_Str += " ";
            }

            #endregion


            string Data = Stop_Str + Action_Str + Unit_Str + Lenth_Str + ByteLenth_Str + Value_Str;

            LRC_Str = LRC2(Data);
            LRC_Str = LRC_Str.ToUpper();
            LRC_Str = convert16(LRC_Str).TrimEnd();
            string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
            Console.WriteLine(Comment);
            string hexValuesReplace = Comment.Replace(" ", "");
            byte[] buffer = StringToByteArray(hexValuesReplace);
            CommentData comment = new CommentData();
            comment.Comment = buffer;
            comment.boolUnit = unit;
            comment.stop = Stop;
            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (Action_Handle != null)
                {
                    SendData.Enqueue(comment);
                }
                else
                {
                    if (!SendData.Contains(comment))
                        SendData.Enqueue(comment);
                }
            }
        }

        //連續寫入點位
        public static void SetPoint(int Stop, BoolUnit unit, int Number, string S, bool Emer,
            ManualResetEvent Action_Handle)
        {
            string Start_Str = "3A ";
            string End_Str = " 0D 0A";
            string Action_Str = "30 46 ";
            string Stop_Str = "";
            string Unit_Str = "";
            string Value_Str = "";
            string LRC_Str = "";
            string Lenth_Str = "";
            string ByteLenth_Str = "";
            List<string> slist = new List<string>();

            #region 站號處理

            if (Stop < 1)
                throw new Exception("站號錯誤(Stop>0)");
            int StopN1 = 30 + (Stop / 16);
            int StopN2 = 30 + (Stop % 16);
            if (StopN1 >= 40)
                StopN1++;
            if (StopN2 >= 40)
                StopN2++;
            Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

            #endregion

            #region 暫存器字串處理

            int UnitN1 = 30;
            int UnitN2 = 30;
            int UnitN3 = 30;
            int UnitN4 = 30;
            switch (unit)
            {
                case BoolUnit.Y:
                {
                    if (Number < 0 || Number > 177)
                        throw new Exception("站存器編號錯誤(有效編號Y0~Y177)");
                    if (Number % 10 > 7)
                        throw new Exception("站存器編號錯誤(有效編號尾數Y0~Y7)");
                    Number = (Number / 100 % 10 * 8 * 8) + (Number / 10 % 10 * 8) + Number % 10;
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.M:
                {
                    if (Number < 0 || (Number > 8191 && Number < 9000) || Number > 9511)
                        throw new Exception("站存器編號錯誤(有效編號M0~M8191、M9000~M9511)");
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.S:
                {
                    if (Number < 0 || Number > 4095)
                        throw new Exception("站存器編號錯誤(有效編號S0~S4095)");
                    UnitN1 = 32;
                    UnitN2 = 36;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.T:
                {
                    if (Number < 0 || Number > 511)
                        throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                    UnitN1 = 30;
                    UnitN2 = 31;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.C:
                {
                    if (Number < 0 || Number > 255)
                        throw new Exception("站存器編號錯誤(有效編號C0~C255)");
                    UnitN1 = 30;
                    UnitN2 = 33;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC1X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC1X0~EC1X7)");
                    Number += 9260;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC1Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC1Y0~EC1Y7)");
                    Number += 9270;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC2X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC2X0~EC2X7)");
                    Number += 9280;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC2Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC2Y0~EC2Y7)");
                    Number += 9290;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC3X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC3X0~EC3X7)");
                    Number += 9300;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC3Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC3Y0~EC3Y7)");
                    Number += 9310;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.X:
                {
                    throw new Exception("X點位不能做寫入動作)");
                }
            }

            Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " + UnitN4.ToString() +
                       " ";

            #endregion

            #region 組數處理

            int Len = S.Length;
            int LenN1 = 30 + (Len / 16 / 16 / 16);
            int LenN2 = 30 + (Len / 16 / 16);
            int LenN3 = 30 + (Len / 16);
            int LenN4 = 30 + (Len % 16);
            if (LenN1 >= 40)
                LenN1++;
            if (LenN2 >= 40)
                LenN2++;
            if (LenN3 >= 40)
                LenN3++;
            if (LenN4 >= 40)
                LenN4++;
            Lenth_Str = LenN1.ToString() + " " + LenN2.ToString() + " " + LenN3.ToString() + " " + LenN4.ToString() +
                        " ";

            #endregion

            #region 組數處理

            if (S.Split(new char[] { '1', '0' }, StringSplitOptions.RemoveEmptyEntries).Count() > 0)
                throw new Exception("訊號字串錯誤(只能1、0): " + S);
            string s = "";
            for (int i = 0; i < S.Length; i++)
            {
                if (S.Substring(i, 1) == "1")
                    s = 1 + s;
                else
                    s = 0 + s;
                if (i % 8 == 7 || i == S.Length - 1)
                {
                    slist.Add(s);
                    s = "";
                }
            }

            int BLen = slist.Count;
            int BLenN1 = 30 + (BLen / 16);
            int BLenN2 = 30 + (BLen % 16);
            if (BLenN1 >= 40)
                BLenN1++;
            if (BLenN2 >= 40)
                BLenN2++;
            ByteLenth_Str = BLenN1.ToString() + " " + BLenN2.ToString() + " ";

            #endregion

            #region 資料處理

            for (int i = 0; i < slist.Count; i++)
            {
                int Value = Convert.ToInt32(slist[i], 2);
                int N1 = 30;
                int N2 = 30;
                int Value_N1 = Value / 16 % 16;
                int Value_N2 = Value % 16;
                N2 += Value_N2;
                if (N2 >= 40)
                    N2++;
                if (N2 > 46)
                {
                    N2 -= 17;
                    N1++;
                }

                N1 += Value_N1;
                if (N1 >= 40)
                    N1++;
                if (N1 > 46)
                {
                    N1 = 46;
                }

                Value_Str += N1.ToString() + " " + N2.ToString();
                if (i != slist.Count - 1)
                    Value_Str += " ";
            }

            #endregion


            string Data = Stop_Str + Action_Str + Unit_Str + Lenth_Str + ByteLenth_Str + Value_Str;

            LRC_Str = LRC2(Data);
            LRC_Str = LRC_Str.ToUpper();
            LRC_Str = convert16(LRC_Str).TrimEnd();
            string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
            Console.WriteLine(Comment);
            string hexValuesReplace = Comment.Replace(" ", "");
            byte[] buffer = StringToByteArray(hexValuesReplace);
            CommentData comment = new CommentData();
            comment.Comment = buffer;
            comment.boolUnit = unit;
            comment.stop = Stop;
            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (Action_Handle != null)
                {
                    SendData.Enqueue(comment);
                }
                else
                {
                    if (!SendData.Contains(comment))
                        SendData.Enqueue(comment);
                }
            }
        }

        //連續寫入暫存器
        public static void SendValue(int Stop, ValueUnit unit, int Number, List<int> list, bool Emer,
            ManualResetEvent Action_Handle)
        {
            string Start_Str = "3A ";
            string End_Str = " 0D 0A";
            string Action_Str = "31 30 ";
            string Stop_Str = "";
            string Unit_Str = "";
            string Lenth_Str = "";
            string ByteLenth_Str = "";
            string Value_Str = "";
            string LRC_Str = "";
            int Len = list.Count;
            switch (unit)
            {
                case ValueUnit.D:
                case ValueUnit.T:
                case ValueUnit.R:
                case ValueUnit.C_16bit:
                    Len = list.Count;
                    break;
                case ValueUnit.C_32bit:
                    Len = list.Count * 2;
                    break;
            }

            #region 站號處理

            if (Stop < 1)
                throw new Exception("站號錯誤(Stop>0)");
            int StopN1 = 30 + (Stop / 16);
            int StopN2 = 30 + (Stop % 16);
            if (StopN1 >= 40)
                StopN1++;
            if (StopN2 >= 40)
                StopN2++;
            Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

            #endregion

            #region 暫存器字串處理

            int UnitN1 = 30;
            int UnitN2 = 30;
            int UnitN3 = 30;
            int UnitN4 = 30;
            switch (unit)
            {
                case ValueUnit.D:
                {
                    if (Number < 0 || Number > 9511)
                        throw new Exception("站存器編號錯誤(有效編號D0~D9511)");
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.T:
                {
                    if (Number < 0 || Number > 511)
                        throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                    UnitN1 = 32;
                    UnitN2 = 36;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.C_16bit:
                {
                    if (Number < 0 || Number > 199)
                        throw new Exception("站存器編號錯誤(有效編號C0~C199)");
                    UnitN1 = 32;
                    UnitN2 = 38;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.C_32bit:
                {
                    if (Number < 200 || Number > 255)
                        throw new Exception("站存器編號錯誤(有效編號C200~C255)");
                    Number -= 200;
                    Number = Number * 2;
                    UnitN1 = 32;
                    UnitN2 = 39;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.R:
                {
                    if (Number < 0 || Number > 9999)
                        throw new Exception("站存器編號錯誤(有效編號R0~R9999)");
                    UnitN1 = 32;
                    UnitN2 = 40;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
            }

            Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " + UnitN4.ToString() +
                       " ";

            #endregion

            #region 組數處理

            int LenN1 = 30 + (Len / 16 / 16 / 16);
            int LenN2 = 30 + (Len / 16 / 16);
            int LenN3 = 30 + (Len / 16);
            int LenN4 = 30 + (Len % 16);
            if (LenN1 >= 40)
                LenN1++;
            if (LenN2 >= 40)
                LenN2++;
            if (LenN3 >= 40)
                LenN3++;
            if (LenN4 >= 40)
                LenN4++;
            Lenth_Str = LenN1.ToString() + " " + LenN2.ToString() + " " + LenN3.ToString() + " " + LenN4.ToString() +
                        " ";

            #endregion

            #region 組數處理

            int BLen = Len * 2;
            int BLenN1 = 30 + (BLen / 16);
            int BLenN2 = 30 + (BLen % 16);
            if (BLenN1 >= 40)
                BLenN1++;
            if (BLenN2 >= 40)
                BLenN2++;
            ByteLenth_Str = BLenN1.ToString() + " " + BLenN2.ToString() + " ";

            #endregion

            #region 資料處理

            for (int i = 0; i < list.Count; i++)
            {
                int Value = list[i];
                switch (unit)
                {
                    case ValueUnit.D:
                    case ValueUnit.T:
                    case ValueUnit.R:
                    case ValueUnit.C_16bit:
                    {
                        if (Value < Int16.MinValue || Value > Int16.MaxValue)
                            throw new Exception("數值錯誤(-32768<=Value<=32767): " + Value);

                        if (Value >= 0)
                        {
                            int N1 = 30;
                            int N2 = 30;
                            int N3 = 30;
                            int N4 = 30;
                            int Value_N1 = Value / 16 / 16 / 16 % 16;
                            int Value_N2 = Value / 16 / 16 % 16;
                            int Value_N3 = Value / 16 % 16;
                            int Value_N4 = Value % 16;
                            N4 += Value_N4;
                            if (N4 >= 40)
                                N4++;
                            if (N4 > 46)
                            {
                                N4 -= 17;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 >= 40)
                                N3++;
                            if (N3 > 46)
                            {
                                N3 -= 17;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 >= 40)
                                N2++;
                            if (N2 > 46)
                            {
                                N2 -= 17;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 >= 40)
                                N1++;
                            if (N1 > 46)
                            {
                                N1 = 46;
                            }

                            Value_Str += N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " +
                                         N4.ToString();
                        }
                        else
                        {
                            int N1 = 30;
                            int N2 = 30;
                            int N3 = 30;
                            int N4 = 30;
                            int Value_N1 = 15 - (-Value / 16 / 16 / 16 % 16);
                            int Value_N2 = 15 - (-Value / 16 / 16 % 16);
                            int Value_N3 = 15 - (-Value / 16 % 16);
                            int Value_N4 = 16 - (-Value % 16);
                            N4 += Value_N4;
                            if (N4 >= 40)
                                N4++;
                            if (N4 > 46)
                            {
                                N4 -= 17;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 >= 40)
                                N3++;
                            if (N3 > 46)
                            {
                                N3 -= 17;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 >= 40)
                                N2++;
                            if (N2 > 46)
                            {
                                N2 -= 17;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 >= 40)
                                N1++;
                            if (N1 > 46)
                            {
                                N1 = 46;
                            }

                            Value_Str += N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " +
                                         N4.ToString();
                        }
                    }
                        break;
                    case ValueUnit.C_32bit:
                    {
                        if (Value < int.MinValue || Value > int.MaxValue)
                            throw new Exception("數值錯誤(-2147483648<=Value<=2147483647): " + Value);
                        if (Value >= 0)
                        {
                            int N1 = 30;
                            int N2 = 30;
                            int N3 = 30;
                            int N4 = 30;
                            int N5 = 30;
                            int N6 = 30;
                            int N7 = 30;
                            int N8 = 30;
                            int Value_N1 = Value / 16 / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                            int Value_N2 = Value / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                            int Value_N3 = Value / 16 / 16 / 16 / 16 / 16 % 16;
                            int Value_N4 = Value / 16 / 16 / 16 / 16 % 16;
                            int Value_N5 = Value / 16 / 16 / 16 % 16;
                            int Value_N6 = Value / 16 / 16 % 16;
                            int Value_N7 = Value / 16 % 16;
                            int Value_N8 = Value % 16;
                            N8 += Value_N8;
                            if (N8 >= 40)
                                N8++;
                            if (N8 > 46)
                            {
                                N8 -= 17;
                                N7++;
                            }

                            N7 += Value_N7;
                            if (N7 >= 40)
                                N7++;
                            if (N7 > 46)
                            {
                                N7 -= 17;
                                N6++;
                            }

                            N6 += Value_N6;
                            if (N6 >= 40)
                                N6++;
                            if (N6 > 46)
                            {
                                N6 -= 17;
                                N5++;
                            }

                            N5 += Value_N5;
                            if (N5 >= 40)
                                N5++;
                            if (N5 > 46)
                            {
                                N5 -= 17;
                                N4++;
                            }

                            N4 += Value_N4;
                            if (N4 >= 40)
                                N4++;
                            if (N4 > 46)
                            {
                                N4 -= 17;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 >= 40)
                                N3++;
                            if (N3 > 46)
                            {
                                N3 -= 17;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 >= 40)
                                N2++;
                            if (N2 > 46)
                            {
                                N2 -= 17;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 >= 40)
                                N1++;
                            if (N1 > 46)
                            {
                                N1 = 46;
                            }

                            Value_Str += N5.ToString() + " " + N6.ToString() + " " + N7.ToString() + " " +
                                         N8.ToString() + " " + N1.ToString() + " " + N2.ToString() + " " +
                                         N3.ToString() + " " + N4.ToString();
                        }
                        else
                        {
                            int N = int.MinValue;
                            Value = Value - N;
                            int N1 = 38;
                            int N2 = 30;
                            int N3 = 30;
                            int N4 = 30;
                            int N5 = 30;
                            int N6 = 30;
                            int N7 = 30;
                            int N8 = 30;
                            int Value_N1 = Value / 16 / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                            int Value_N2 = Value / 16 / 16 / 16 / 16 / 16 / 16 % 16;
                            int Value_N3 = Value / 16 / 16 / 16 / 16 / 16 % 16;
                            int Value_N4 = Value / 16 / 16 / 16 / 16 % 16;
                            int Value_N5 = Value / 16 / 16 / 16 % 16;
                            int Value_N6 = Value / 16 / 16 % 16;
                            int Value_N7 = Value / 16 % 16;
                            int Value_N8 = Value % 16;
                            N8 += Value_N8;
                            if (N8 >= 40)
                                N8++;
                            if (N8 > 46)
                            {
                                N8 -= 17;
                                N7++;
                            }

                            N7 += Value_N7;
                            if (N7 >= 40)
                                N7++;
                            if (N7 > 46)
                            {
                                N7 -= 17;
                                N6++;
                            }

                            N6 += Value_N6;
                            if (N6 >= 40)
                                N6++;
                            if (N6 > 46)
                            {
                                N6 -= 17;
                                N5++;
                            }

                            N5 += Value_N5;
                            if (N5 >= 40)
                                N5++;
                            if (N5 > 46)
                            {
                                N5 -= 17;
                                N4++;
                            }

                            N4 += Value_N4;
                            if (N4 >= 40)
                                N4++;
                            if (N4 > 46)
                            {
                                N4 -= 17;
                                N3++;
                            }

                            N3 += Value_N3;
                            if (N3 >= 40)
                                N3++;
                            if (N3 > 46)
                            {
                                N3 -= 17;
                                N2++;
                            }

                            N2 += Value_N2;
                            if (N2 >= 40)
                                N2++;
                            if (N2 > 46)
                            {
                                N2 -= 17;
                                N1++;
                            }

                            N1 += Value_N1;
                            if (N1 >= 40)
                                N1++;
                            if (N1 > 46)
                            {
                                N1 = 46;
                            }

                            Value_Str += N5.ToString() + " " + N6.ToString() + " " + N7.ToString() + " " +
                                         N8.ToString() + " " + N1.ToString() + " " + N2.ToString() + " " +
                                         N3.ToString() + " " + N4.ToString();
                        }
                    }
                        break;
                }

                if (i < list.Count - 1)
                    Value_Str += " ";
            }

            #endregion

            string Data = Stop_Str + Action_Str + Unit_Str + Lenth_Str + ByteLenth_Str + Value_Str;

            LRC_Str = LRC2(Data);
            LRC_Str = LRC_Str.ToUpper();
            LRC_Str = convert16(LRC_Str).TrimEnd();
            string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
            string hexValuesReplace = Comment.Replace(" ", "");
            byte[] buffer = StringToByteArray(hexValuesReplace);
            CommentData comment = new CommentData();
            comment.Comment = buffer;
            comment.valueUnit = unit;
            comment.stop = Stop;
            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (Action_Handle != null)
                {
                    SendData.Enqueue(comment);
                }
                else
                {
                    if (!SendData.Contains(comment))
                        SendData.Enqueue(comment);
                }
            }
        }

        //寫狀態進入點位
        public static void SetPoint(int Stop, BoolUnit unit, int Number, bool Value, bool Emer,
            ManualResetEvent Action_Handle)
        {
            string Start_Str = "3A ";
            string End_Str = " 0D 0A";
            string Action_Str = "30 35 ";
            string Stop_Str = "";
            string Unit_Str = "";
            string Value_Str = "";
            string LRC_Str = "";

            #region 站號處理

            if (Stop < 1)
                throw new Exception("站號錯誤(Stop>0)");
            int StopN1 = 30 + (Stop / 16);
            int StopN2 = 30 + (Stop % 16);
            if (StopN1 >= 40)
                StopN1++;
            if (StopN2 >= 40)
                StopN2++;
            Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

            #endregion

            #region 暫存器字串處理

            int UnitN1 = 30;
            int UnitN2 = 30;
            int UnitN3 = 30;
            int UnitN4 = 30;
            switch (unit)
            {
                case BoolUnit.Y:
                {
                    if (Number < 0 || Number > 177)
                        throw new Exception("站存器編號錯誤(有效編號Y0~Y177)");
                    if (Number % 10 > 7)
                        throw new Exception("站存器編號錯誤(有效編號尾數Y0~Y7)");
                    Number = (Number / 100 % 10 * 8 * 8) + (Number / 10 % 10 * 8) + Number % 10;
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.M:
                {
                    if (Number < 0 || (Number > 8191 && Number < 9000) || Number > 9511)
                        throw new Exception("站存器編號錯誤(有效編號M0~M8191、M9000~M9511)");
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.S:
                {
                    if (Number < 0 || Number > 4095)
                        throw new Exception("站存器編號錯誤(有效編號S0~S4095)");
                    UnitN1 = 32;
                    UnitN2 = 36;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.T:
                {
                    if (Number < 0 || Number > 511)
                        throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                    UnitN1 = 30;
                    UnitN2 = 31;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.C:
                {
                    if (Number < 0 || Number > 255)
                        throw new Exception("站存器編號錯誤(有效編號C0~C255)");
                    UnitN1 = 30;
                    UnitN2 = 33;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC1X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC1X0~EC1X7)");
                    Number += 9260;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC1Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC1Y0~EC1Y7)");
                    Number += 9270;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC2X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC2X0~EC2X7)");
                    Number += 9280;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC2Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC2Y0~EC2Y7)");
                    Number += 9290;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC3X:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC3X0~EC3X7)");
                    Number += 9300;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.EC3Y:
                {
                    if (Number < 0 || Number > 7)
                        throw new Exception("站存器編號錯誤(有效編號EC3Y0~EC3Y7)");
                    Number += 9310;
                    if (Number >= 9000)
                        Number -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = Number / 16 / 16 / 16 % 16;
                    int Number_N2 = Number / 16 / 16 % 16;
                    int Number_N3 = Number / 16 % 16;
                    int Number_N4 = Number % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.X:
                {
                    throw new Exception("X點位不能做寫入動作)");
                }
            }

            Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " + UnitN4.ToString() +
                       " ";

            #endregion

            #region 資料處理

            int N1 = 30;
            int N2 = 30;
            int N3 = 30;
            int N4 = 30;
            if (Value)
            {
                N1 = 46;
                N2 = 46;
                N3 = 30;
                N4 = 30;
            }
            else
            {
                N1 = 30;
                N2 = 30;
                N3 = 30;
                N4 = 30;
            }

            Value_Str = N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " + N4.ToString();

            #endregion

            string Data = Stop_Str + Action_Str + Unit_Str + Value_Str;

            LRC_Str = LRC2(Data);
            LRC_Str = LRC_Str.ToUpper();
            LRC_Str = convert16(LRC_Str).TrimEnd();
            string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
            //Console.WriteLine(Comment);
            string hexValuesReplace = Comment.Replace(" ", "");
            byte[] buffer = StringToByteArray(hexValuesReplace);
            CommentData comment = new CommentData();
            comment.Comment = buffer;
            comment.boolUnit = unit;
            comment.stop = Stop;
            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (Action_Handle != null)
                {
                    SendData.Enqueue(comment);
                }
                else
                {
                    if (!SendData.Contains(comment))
                        SendData.Enqueue(comment);
                }
            }
        }


        //直接寫入指令
        public static string SendComment_Test(string STR1)
        {
            try
            {
                Handle.Reset();
                string Start_Str = "3A ";
                string End_Str = " 0D 0A";
                string STR2 = LRC2(STR1);
                STR2 = STR2.ToUpper();
                STR2 = convert16(STR2).TrimEnd();
                STR2 = STR1 + " " + STR2;
                STR2 = Start_Str + STR2 + End_Str;
                string hexValuesReplace = STR2.Replace(" ", "");
                byte[] buffer = StringToByteArray(hexValuesReplace);
                CommentData comment = new CommentData();
                comment.Comment = buffer;
                if (!SendData.Contains(comment))
                    SendData.Enqueue(comment);
                return STR2;
            }
            catch
            {
                return "";
            }
        }

        #endregion

        #region 讀取指令

        //取暫存器值
        public static void CheckValue(int Stop, ValueUnit unit, int StartNumber, int lenth, bool Emer,
            ManualResetEvent Action_Handle)
        {
            string Start_Str = "3A ";
            string End_Str = " 0D 0A";
            string Action_Str = "30 33 ";
            string Stop_Str = "";
            string Unit_Str = "";
            string Value_Str = "";
            string LRC_Str = "";

            #region 站號處理

            if (Stop < 1)
                throw new Exception("站號錯誤(Stop>0)");
            int StopN1 = 30 + (Stop / 16);
            int StopN2 = 30 + (Stop % 16);
            if (StopN1 >= 40)
                StopN1++;
            if (StopN2 >= 40)
                StopN2++;
            Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

            #endregion

            #region 暫存器字串處理

            int UnitN1 = 30;
            int UnitN2 = 30;
            int UnitN3 = 30;
            int UnitN4 = 30;
            switch (unit)
            {
                case ValueUnit.D:
                {
                    if (StartNumber < 0 || StartNumber > 9511)
                        throw new Exception("站存器編號錯誤(有效編號D0~D9511)");
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.T:
                {
                    if (StartNumber < 0 || StartNumber > 511)
                        throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                    UnitN1 = 32;
                    UnitN2 = 36;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.C_16bit:
                {
                    if (StartNumber < 0 || StartNumber > 199)
                        throw new Exception("站存器編號錯誤(有效編號C0~C199)");
                    UnitN1 = 32;
                    UnitN2 = 38;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.C_32bit:
                {
                    if (StartNumber < 200 || StartNumber > 255)
                        throw new Exception("站存器編號錯誤(有效編號C200~C255)");
                    StartNumber -= 200;
                    StartNumber = StartNumber * 2;
                    lenth = lenth * 2;
                    UnitN1 = 32;
                    UnitN2 = 39;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case ValueUnit.R:
                {
                    if (StartNumber < 0 || StartNumber > 9999)
                        throw new Exception("站存器編號錯誤(有效編號R0~R9999)");
                    UnitN1 = 32;
                    UnitN2 = 40;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
            }

            Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " + UnitN4.ToString() +
                       " ";

            #endregion

            #region 資料處理

            int N1 = 30;
            int N2 = 30;
            int N3 = 30;
            int N4 = 30;
            int Value_N1 = lenth / 16 / 16 / 16 % 16;
            int Value_N2 = lenth / 16 / 16 % 16;
            int Value_N3 = lenth / 16 % 16;
            int Value_N4 = lenth % 16;
            N4 += Value_N4;
            if (N4 >= 40)
                N4++;
            if (N4 > 46)
            {
                N4 -= 17;
                N3++;
            }

            N3 += Value_N3;
            if (N3 >= 40)
                N3++;
            if (N3 > 46)
            {
                N3 -= 17;
                N2++;
            }

            N2 += Value_N2;
            if (N2 >= 40)
                N2++;
            if (N2 > 46)
            {
                N2 -= 17;
                N1++;
            }

            N1 += Value_N1;
            if (N1 >= 40)
                N1++;
            if (N1 > 46)
            {
                N1 = 46;
            }

            Value_Str = N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " + N4.ToString();

            #endregion

            string Data = Stop_Str + Action_Str + Unit_Str + Value_Str;

            LRC_Str = LRC2(Data);
            LRC_Str = LRC_Str.ToUpper();
            LRC_Str = convert16(LRC_Str).TrimEnd();
            string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
            string hexValuesReplace = Comment.Replace(" ", "");
            byte[] buffer = StringToByteArray(hexValuesReplace);
            CommentData comment = new CommentData();
            comment.Comment = buffer;
            comment.valueUnit = unit;
            comment.stop = Stop;
            comment.Lenth = lenth;
            comment.Start = StartNumber;
            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (!SendData.Contains(comment))
                    SendData.Enqueue(comment);
            }

            //Console.WriteLine(Comment);
        }

        //取點位狀態
        public static void CheckPoint(int Stop, BoolUnit unit, int StartNumber, int lenth, bool Emer,
            ManualResetEvent Action_Handle)
        {
            string Start_Str = "3A ";
            string End_Str = " 0D 0A";
            string Action_Str = "30 31 ";
            if (unit == BoolUnit.X)
                Action_Str = "30 32 ";
            string Stop_Str = "";
            string Unit_Str = "";
            string Value_Str = "";
            string LRC_Str = "";

            #region 站號處理

            if (Stop < 1)
                throw new Exception("站號錯誤(Stop>0)");
            int StopN1 = 30 + (Stop / 16);
            int StopN2 = 30 + (Stop % 16);
            if (StopN1 >= 40)
                StopN1++;
            if (StopN2 >= 40)
                StopN2++;
            Stop_Str = StopN1.ToString() + " " + StopN2.ToString() + " ";

            #endregion

            #region 暫存器字串處理

            int UnitN1 = 30;
            int UnitN2 = 30;
            int UnitN3 = 30;
            int UnitN4 = 30;
            switch (unit)
            {
                case BoolUnit.X:
                {
                    if (StartNumber < 0 || StartNumber > 177)
                        throw new Exception("站存器編號錯誤(有效編號X0~X177)");
                    if (StartNumber % 10 > 7)
                        throw new Exception("站存器編號錯誤(有效編號尾數X0~X7)");
                    int N = (StartNumber / 100 % 10 * 8 * 8) + (StartNumber / 10 % 10 * 8) + StartNumber % 10;
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = N / 16 / 16 / 16 % 16;
                    int Number_N2 = N / 16 / 16 % 16;
                    int Number_N3 = N / 16 % 16;
                    int Number_N4 = N % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.Y:
                {
                    if (StartNumber < 0 || StartNumber > 177)
                        throw new Exception("站存器編號錯誤(有效編號Y0~Y177)");
                    if (StartNumber % 10 > 7)
                        throw new Exception("站存器編號錯誤(有效編號尾數Y0~Y7)");
                    int N = (StartNumber / 100 % 10 * 8 * 8) + (StartNumber / 10 % 10 * 8) + StartNumber % 10;
                    UnitN1 = 30;
                    UnitN2 = 30;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = N / 16 / 16 / 16 % 16;
                    int Number_N2 = N / 16 / 16 % 16;
                    int Number_N3 = N / 16 % 16;
                    int Number_N4 = N % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.M:
                {
                    if (StartNumber < 0 || (StartNumber > 8191 && StartNumber < 9000) || StartNumber > 9511)
                        throw new Exception("站存器編號錯誤(有效編號M0~M8191、M9000~M9511)");
                    int N = StartNumber;
                    if (N >= 9000)
                        N -= 808;
                    UnitN1 = 30;
                    UnitN2 = 34;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = N / 16 / 16 / 16 % 16;
                    int Number_N2 = N / 16 / 16 % 16;
                    int Number_N3 = N / 16 % 16;
                    int Number_N4 = N % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.S:
                {
                    if (StartNumber < 0 || StartNumber > 4095)
                        throw new Exception("站存器編號錯誤(有效編號S0~S4095)");
                    UnitN1 = 32;
                    UnitN2 = 36;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.T:
                {
                    if (StartNumber < 0 || StartNumber > 511)
                        throw new Exception("站存器編號錯誤(有效編號T0~T511)");
                    UnitN1 = 30;
                    UnitN2 = 31;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
                case BoolUnit.C:
                {
                    if (StartNumber < 0 || StartNumber > 255)
                        throw new Exception("站存器編號錯誤(有效編號C0~C255)");
                    UnitN1 = 30;
                    UnitN2 = 33;
                    UnitN3 = 30;
                    UnitN4 = 30;
                    int Number_N1 = StartNumber / 16 / 16 / 16 % 16;
                    int Number_N2 = StartNumber / 16 / 16 % 16;
                    int Number_N3 = StartNumber / 16 % 16;
                    int Number_N4 = StartNumber % 16;
                    UnitN4 += Number_N4;
                    if (UnitN4 >= 40)
                        UnitN4++;
                    if (UnitN4 > 46)
                    {
                        UnitN4 -= 17;
                        UnitN3++;
                    }

                    UnitN3 += Number_N3;
                    if (UnitN3 >= 40)
                        UnitN3++;
                    if (UnitN3 > 46)
                    {
                        UnitN3 -= 17;
                        UnitN2++;
                    }

                    UnitN2 += Number_N2;
                    if (UnitN2 >= 40)
                        UnitN2++;
                    if (UnitN2 > 46)
                    {
                        UnitN2 -= 17;
                        UnitN1++;
                    }

                    UnitN1 += Number_N1;
                    if (UnitN1 >= 40)
                        UnitN1++;
                    if (UnitN1 > 46)
                    {
                        UnitN1 = 46;
                    }
                }
                    break;
            }

            Unit_Str = UnitN1.ToString() + " " + UnitN2.ToString() + " " + UnitN3.ToString() + " " + UnitN4.ToString() +
                       " ";

            #endregion

            #region 資料處理

            int N1 = 30;
            int N2 = 30;
            int N3 = 30;
            int N4 = 30;
            int Value_N1 = lenth / 16 / 16 / 16 % 16;
            int Value_N2 = lenth / 16 / 16 % 16;
            int Value_N3 = lenth / 16 % 16;
            int Value_N4 = lenth % 16;
            N4 += Value_N4;
            if (N4 >= 40)
                N4++;
            if (N4 > 46)
            {
                N4 -= 17;
                N3++;
            }

            N3 += Value_N3;
            if (N3 >= 40)
                N3++;
            if (N3 > 46)
            {
                N3 -= 17;
                N2++;
            }

            N2 += Value_N2;
            if (N2 >= 40)
                N2++;
            if (N2 > 46)
            {
                N2 -= 17;
                N1++;
            }

            N1 += Value_N1;
            if (N1 >= 40)
                N1++;
            if (N1 > 46)
            {
                N1 = 46;
            }

            Value_Str = N1.ToString() + " " + N2.ToString() + " " + N3.ToString() + " " + N4.ToString();

            #endregion

            string Data = Stop_Str + Action_Str + Unit_Str + Value_Str;

            LRC_Str = LRC2(Data);
            LRC_Str = LRC_Str.ToUpper();
            LRC_Str = convert16(LRC_Str).TrimEnd();
            string Comment = Start_Str + Data + " " + LRC_Str + End_Str;
            string hexValuesReplace = Comment.Replace(" ", "");
            byte[] buffer = StringToByteArray(hexValuesReplace);
            CommentData comment = new CommentData();
            comment.Comment = buffer;
            comment.boolUnit = unit;
            comment.stop = Stop;
            comment.Lenth = lenth;
            comment.Start = StartNumber;
            if (Action_Handle != null)
                comment.Action_Handle = Action_Handle;
            if (Emer)
                SendData_Emergency.Enqueue(comment);
            else
            {
                if (!SendData.Contains(comment))
                    SendData.Enqueue(comment);
            }
            //Serilog.Log.Information($"PLC DataQueue = {SendData.Count}");
            //Console.WriteLine(Comment);
        }

        #endregion

        //get the queue current count
        public static int GetQueueCount()
        {
            return SendData.Count;
        }

        //指令匯集緩衝區
        public async static void SendComment()
        {
            while (PLC_Run)
            {
                if (PLC_Reboot)
                {
                    event2.WaitOne();
                    Log.Warning("wait2");//wait for reconnection
                    event3.Set();//send signal to outside loop
                }
                    
                if ((SendData.Count > 0 || SendData_Emergency.Count > 0) && com.BytesToRead == 0 )
                {
                    if (SendData_Emergency.Count > 0)
                    {
                        if (SendData_Emergency.TryDequeue(out var Data))
                        {
                            if (Protocol_Type == Protocol.ASCII)
                            {
                                if (Data.Comment[0] == 58 && Data.Comment[Data.Comment.Length - 2] == 13 &&
                                    Data.Comment[Data.Comment.Length - 1] == 10)
                                {
                                    Array.Resize(ref Comment, Data.Comment.Length);
                                    Comment = Data.Comment;
                                    if (Data.Comment[3] == 48 && Data.Comment[4] == 49)
                                    {
                                        Check_BoolUnit_Stop = Data.stop;
                                        Check_BoolUnit = Data.boolUnit;
                                        Check_BoolUnit_Start = Data.Start;
                                        Check_BoolUnit_Lenth = Data.Lenth;
                                    }
                                    else if (Data.Comment[3] == 48 && Data.Comment[4] == 50)
                                    {
                                        Check_BoolUnit_Stop = Data.stop;
                                        Check_BoolUnit = Data.boolUnit;
                                        Check_BoolUnit_Start = Data.Start;
                                        Check_BoolUnit_Lenth = Data.Lenth;
                                    }
                                    else if (Data.Comment[3] == 48 && Data.Comment[4] == 51)
                                    {
                                        Check_ValueUnit_Stop = Data.stop;
                                        Check_ValueUnit = Data.valueUnit;
                                        Check_ValueUnit_Start = Data.Start;
                                        Check_ValueUnit_Lenth = Data.Lenth;
                                    }

                                    int time = 0;
                                    if (Data.Action_Handle != null)
                                        Action_Handle = Data.Action_Handle;
                                    while (PLC_SendComment && time < 3)
                                    {
                                        try
                                        {
                                            Handle.Reset();
                                            Temp_Comment = new byte[0];
                                            com.Write(Data.Comment, 0, Data.Comment.Length);
                                            //await com.BaseStream.WriteAsync(Data.Comment ,0, Data.Comment.Length);
                                        }
                                        catch
                                        {
                                            Log.Error("PLC寫入異常");
                                            Temp_Comment = new byte[0];
                                            break;
                                        }

                                        if (Handle.WaitOne(200))
                                        {
                                            Temp_Comment = new byte[0];
                                            break;
                                        }

                                        time++;
                                    }

                                    if (time >= 3)
                                    {
                                        Log.Error("PLC指令沒回應");
                                        // var comport_status = com.IsOpen;
                                        // if (comport_status)
                                        //     Log.Information("當前PLC COMPORT 為 on ");
                                        // else
                                        //     Log.Information("當前PLC COMPORT 為 off  ");

                                        if (Data.Action_Handle != null)
                                            Data.Action_Handle.Set();
                                        Action_Handle = null;
                                    }
                                    else
                                    {
                                        if (!(Data.Comment[3] == 48 && Data.Comment[4] == 49) &&
                                            !(Data.Comment[3] == 48 && Data.Comment[4] == 50) &&
                                            !(Data.Comment[3] == 48 && Data.Comment[4] == 51))
                                        {
                                            if (Data.Action_Handle != null)
                                                Data.Action_Handle.Set();
                                            Action_Handle = null;
                                        }
                                    }
                                }
                            }
                            else if (Protocol_Type == Protocol.RTU)
                            {
                                Array.Resize(ref Comment, Data.Comment.Length);
                                Comment = Data.Comment;
                                if (Data.Comment[1] == 1)
                                {
                                    Check_BoolUnit_Stop = Data.stop;
                                    Check_BoolUnit = Data.boolUnit;
                                    Check_BoolUnit_Start = Data.Start;
                                    Check_BoolUnit_Lenth = Data.Lenth;
                                }
                                else if (Data.Comment[1] == 2)
                                {
                                    Check_BoolUnit_Stop = Data.stop;
                                    Check_BoolUnit = Data.boolUnit;
                                    Check_BoolUnit_Start = Data.Start;
                                    Check_BoolUnit_Lenth = Data.Lenth;
                                }
                                else if (Data.Comment[1] == 3)
                                {
                                    Check_ValueUnit_Stop = Data.stop;
                                    Check_ValueUnit = Data.valueUnit;
                                    Check_ValueUnit_Start = Data.Start;
                                    Check_ValueUnit_Lenth = Data.Lenth;
                                }

                                if (Data.Action_Handle != null)
                                    Action_Handle = Data.Action_Handle;
                                int time = 0;
                                while (PLC_SendComment && time < 3)
                                {
                                    try
                                    {
                                        Handle.Reset();
                                        Temp_Comment = new byte[0];
                                        com.Write(Data.Comment, 0, Data.Comment.Length);
                                        //await com.BaseStream.WriteAsync(Data.Comment ,0, Data.Comment.Length);
                                        
                                    }
                                    catch
                                    {
                                        Console.WriteLine("PLC寫入異常");
                                        Temp_Comment = new byte[0];
                                        break;
                                    }

                                    if (Handle.WaitOne(200))
                                    {
                                        Temp_Comment = new byte[0];
                                        break;
                                    }

                                    time++;
                                }

                                if (time >= 3)
                                {
                                    Log.Error("PLC指令沒回應");
                                    if (Data.Action_Handle != null)
                                        Data.Action_Handle.Set();
                                    Action_Handle = null;
                                }
                                else
                                {
                                    if (!(Data.Comment[1] == 1) && !(Data.Comment[1] == 2) &&
                                        !(Data.Comment[1] == 3))
                                    {
                                        if (Data.Action_Handle != null)
                                            Data.Action_Handle.Set();
                                        Action_Handle = null;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (SendData.TryDequeue(out var Data))
                        {
                            if (Protocol_Type == Protocol.ASCII)
                            {
                                if (Data.Comment[0] == 58 && Data.Comment[Data.Comment.Length - 2] == 13 &&
                                    Data.Comment[Data.Comment.Length - 1] == 10)
                                {
                                    Array.Resize(ref Comment, Data.Comment.Length);
                                    Comment = Data.Comment;
                                    if (Data.Comment[3] == 48 && Data.Comment[4] == 49)
                                    {
                                        Check_BoolUnit_Stop = Data.stop;
                                        Check_BoolUnit = Data.boolUnit;
                                        Check_BoolUnit_Start = Data.Start;
                                        Check_BoolUnit_Lenth = Data.Lenth;
                                    }
                                    else if (Data.Comment[3] == 48 && Data.Comment[4] == 50)
                                    {
                                        Check_BoolUnit_Stop = Data.stop;
                                        Check_BoolUnit = Data.boolUnit;
                                        Check_BoolUnit_Start = Data.Start;
                                        Check_BoolUnit_Lenth = Data.Lenth;
                                    }
                                    else if (Data.Comment[3] == 48 && Data.Comment[4] == 51)
                                    {
                                        Check_ValueUnit_Stop = Data.stop;
                                        Check_ValueUnit = Data.valueUnit;
                                        Check_ValueUnit_Start = Data.Start;
                                        Check_ValueUnit_Lenth = Data.Lenth;
                                    }

                                    if (Data.Action_Handle != null)
                                        Action_Handle = Data.Action_Handle;
                                    int time = 0;
                                    while (PLC_SendComment && time < 3)
                                    {
                                        try
                                        {
                                            Handle.Reset();
                                            Temp_Comment = Array.Empty<byte>();
                                            //await com.BaseStream.WriteAsync(Data.Comment ,0, Data.Comment.Length);
                                            com.Write(Data.Comment, 0, Data.Comment.Length);
                                            
                                        }
                                        catch (Exception e1)
                                        {
                                            Log.Error("PLC寫入異常" + e1.ToString());
                                            Temp_Comment = Array.Empty<byte>();
                                            break;
                                        }

                                        if (Handle.WaitOne(200))
                                        {
                                            Temp_Comment = Array.Empty<byte>();
                                            break;
                                        }

                                        time++;
                                    }

                                    if (time >= 3)
                                    {
                                        Log.Error("PLC指令沒回應");
                                        // var comport_status = com.IsOpen;
                                        // if (comport_status)
                                        //     Log.Information("當前PLC COMPORT 為 on ");
                                        // else
                                        //     Log.Information("當前PLC COMPORT 為 off  ");
                                        if (Data.Action_Handle != null)
                                            Data.Action_Handle.Set();
                                        Action_Handle = null;
                                    }
                                    else
                                    {
                                        if (!(Data.Comment[3] == 48 && Data.Comment[4] == 49) &&
                                            !(Data.Comment[3] == 48 && Data.Comment[4] == 50) &&
                                            !(Data.Comment[3] == 48 && Data.Comment[4] == 51))
                                        {
                                            if (Data.Action_Handle != null)
                                                Data.Action_Handle.Set();
                                            Action_Handle = null;
                                        }
                                    }
                                }
                            }
                            else if (Protocol_Type == Protocol.RTU)
                            {
                                Array.Resize(ref Comment, Data.Comment.Length);
                                Comment = Data.Comment;
                                if (Data.Comment[1] == 1)
                                {
                                    Check_BoolUnit_Stop = Data.stop;
                                    Check_BoolUnit = Data.boolUnit;
                                    Check_BoolUnit_Start = Data.Start;
                                    Check_BoolUnit_Lenth = Data.Lenth;
                                }
                                else if (Data.Comment[1] == 2)
                                {
                                    Check_BoolUnit_Stop = Data.stop;
                                    Check_BoolUnit = Data.boolUnit;
                                    Check_BoolUnit_Start = Data.Start;
                                    Check_BoolUnit_Lenth = Data.Lenth;
                                }
                                else if (Data.Comment[1] == 3)
                                {
                                    Check_ValueUnit_Stop = Data.stop;
                                    Check_ValueUnit = Data.valueUnit;
                                    Check_ValueUnit_Start = Data.Start;
                                    Check_ValueUnit_Lenth = Data.Lenth;
                                }

                                if (Data.Action_Handle != null)
                                    Action_Handle = Data.Action_Handle;
                                int time = 0;
                                while (PLC_SendComment && time < 3)
                                {
                                    try
                                    {
                                        Handle.Reset();
                                        Temp_Comment = new byte[0];
                                        //await com.BaseStream.WriteAsync(Data.Comment ,0, Data.Comment.Length);
                                        com.Write(Data.Comment, 0, Data.Comment.Length);
                                    }
                                    catch
                                    {
                                        Console.WriteLine("PLC寫入異常");
                                        Temp_Comment = new byte[0];
                                        break;
                                    }

                                    if (Handle.WaitOne(200))
                                    {
                                        Temp_Comment = new byte[0];
                                        break;
                                    }

                                    time++;
                                }

                                if (time >= 3)
                                {
                                    Log.Error("PLC指令沒回應");
                                    var comport_status = com.IsOpen;
                                    // if (comport_status)
                                    //     Log.Information("當前PLC COMPORT 為 on ");
                                    // else
                                    //     Log.Information("當前PLC COMPORT 為 off  ");
                                    if (Data.Action_Handle != null)
                                        Data.Action_Handle.Set();
                                    Action_Handle = null;
                                }
                                else
                                {
                                    if (!(Data.Comment[1] == 1) && !(Data.Comment[1] == 2) &&
                                        !(Data.Comment[1] == 3))
                                    {
                                        if (Data.Action_Handle != null)
                                            Data.Action_Handle.Set();
                                        Action_Handle = null;
                                    }
                                }
                            }
                        }
                    }
                }


                event1.Set();
                Thread.Sleep(5);
            }
        }

        public static void Reconnect()
        {
            if (PLC_Run)
            {
                
                event1.WaitOne();
                Log.Warning("wait1");
                Log.Warning("PLC重新連線(手動)");

                PLC_Reboot = true;
                // Attempt to reconnect after a short delay
                Thread.Sleep(1000);
                com.Close();
                com.Encoding = System.Text.Encoding.GetEncoding(28591);
                com.Open();
                
                PLC_Reboot = false;
                event2.Set();
                Handle.Reset();
                PLC_Run = true;
                PLC_SendComment = true;
                Log.Warning("reconnect success");
                Thread.Sleep(1000);
            }
        }

        private static void ErrorReceivedHandler(object sender, SerialErrorReceivedEventArgs e)
        {
            Log.Error("PLCError:"+e.ToString());
            // if (PLC_Run)
            // {
            //     Log.Error("PLC發生錯誤 嘗試重新連線");
            //     // Attempt to reconnect after a short delay
            //     Thread.Sleep(1000);
            //     com.Close();
            //     com.Open();
            // }
        }
        private static void DataReceived(Byte[] Buffer,int length)
        {
            Array.Resize(ref Buffer, length);
            string Back = "";
            Back += String.Format("{0}{1}", BitConverter.ToString(Buffer), Environment.NewLine);
            string BackValue = Back.Replace("-", " ");
            
            if (Protocol_Type == Protocol.ASCII)
            {
                if (Buffer[0] == 58)
                {
                    if (length < 2 || Buffer[length - 2] != 13 || Buffer[length - 1] != 10)
                    {
                        int Temp_Lenth = Temp_Comment.Length;
                        Array.Resize(ref Temp_Comment, length + Temp_Lenth);
                        for (int i = 0; i < length; i++)
                        {
                            Temp_Comment[Temp_Lenth + i] = Buffer[i];
                        }

                        Buffer = Temp_Comment;
                        length = Buffer.Length;
                    }
                }
                else
                {
                    int Temp_Lenth = Temp_Comment.Length;
                    Array.Resize(ref Temp_Comment, length + Temp_Lenth);
                    for (int i = 0; i < length; i++)
                    {
                        Temp_Comment[Temp_Lenth + i] = Buffer[i];
                    }

                    Buffer = Temp_Comment;
                    length = Buffer.Length;
                }

                if (length >= 3 && Buffer[0] == 58 && Buffer[length - 2] == 13 && Buffer[length - 1] == 10)
                {
                    if (Buffer[3] != 48 || Buffer[4] != 49)
                    {
                        if (Buffer[3] == 48 && Buffer[4] == 50)
                        {
                            int DataCount = 0;

                            #region 確認資料組數

                            int N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[5] })) - 30;
                            if (N1 >= 10)
                                N1--;
                            int N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[6] })) - 30;
                            if (N2 >= 10)
                                N2--;
                            DataCount = N1 * 16 + N2;

                            #endregion

                            #region 填入資料

                            for (int i = 0; i < DataCount; i++)
                            {
                                int value_N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[7 + 2 * i] })) - 30;
                                if (value_N1 >= 10)
                                    value_N1--;
                                int value_N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[8 + 2 * i] })) - 30;
                                if (value_N2 >= 10)
                                    value_N2--;
                                switch (Check_BoolUnit)
                                {
                                    case BoolUnit.X:
                                    {
                                        int Value = value_N1 * 16 + value_N2;
                                        string s = Convert.ToString(Value, 2);
                                        int n = s.Length;
                                        if (i == DataCount - 1)
                                        {
                                            for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                            {
                                                s = "0" + s;
                                            }
                                        }
                                        else
                                        {
                                            for (int j = n; j < 8; j++)
                                            {
                                                s = "0" + s;
                                            }
                                        }

                                        char[] Bit = s.Reverse().ToArray();
                                        for (int k = 0; k < Bit.Length; k++)
                                        {
                                            int A = Check_BoolUnit_Start / 10, B = Check_BoolUnit_Start % 10;
                                            //B = B + k + i * 8;
                                            //A = A + (B / 8);
                                            //A = (A / 8) * 10 + A % 8;
                                            //B = B % 8;
                                            if (Bit[k] == '0')
                                            {
                                                PLC_Value.Point_X[A][B] = false;
                                            }
                                            else if (Bit[k] == '1')
                                            {
                                                PLC_Value.Point_X[A][B] = true;
                                            }
                                        }
                                    }
                                        break;
                                }
                            }

                            #endregion
                        }
                        else if (Buffer[3] == 48 && Buffer[4] == 51)
                        {
                            int DataCount = 0;

                            #region 確認資料組數

                            int N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[5] })) - 30;
                            if (N1 >= 10)
                                N1--;
                            int N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[6] })) - 30;
                            if (N2 >= 10)
                                N2--;
                            DataCount = N1 * 16 + N2;

                            #endregion

                            #region 填入資料

                            for (int i = 0; i < DataCount / 2; i++)
                            {
                                int value_N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[7 + 4 * i] })) - 30;
                                if (value_N1 >= 10)
                                    value_N1--;
                                int value_N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[8 + 4 * i] })) - 30;
                                if (value_N2 >= 10)
                                    value_N2--;
                                int value_N3 = int.Parse(BitConverter.ToString(new byte[] { Buffer[9 + 4 * i] })) - 30;
                                if (value_N3 >= 10)
                                    value_N3--;
                                int value_N4 = int.Parse(BitConverter.ToString(new byte[] { Buffer[10 + 4 * i] })) - 30;
                                if (value_N4 >= 10)
                                    value_N4--;
                                switch (Check_ValueUnit)
                                {
                                    case ValueUnit.D:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_D[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[D" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_D[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.R:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_R[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[R" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_R[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.T:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_T[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[T" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_T[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.C_16bit:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_Ctemp[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[C" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_C3[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.C_32bit:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_Ctemp[Check_ValueUnit_Start + i + 200] =
                                            value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //if(i/2==0)
                                        //Console.WriteLine("[C" + (Check_ValueUnit_Start + i/2) + "上位]: " + PLC_Value.Value_C3[Check_ValueUnit_Start + i] + " ");
                                        //else
                                        //Console.WriteLine("[C" + (Check_ValueUnit_Start + i/2) + "下位]: " + PLC_Value.Value_C3[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                }
                            }

                            #endregion
                        }
                    }
                    else
                    {
                        int DataCount = 0;

                        #region 確認資料組數

                        int N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[5] })) - 30;
                        if (N1 >= 10)
                            N1--;
                        int N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[6] })) - 30;
                        if (N2 >= 10)
                            N2--;
                        DataCount = N1 * 16 + N2;

                        #endregion

                        #region 填入資料

                        for (int i = 0; i < DataCount; i++)
                        {
                            int value_N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[7 + 2 * i] })) - 30;
                            if (value_N1 >= 10)
                                value_N1--;
                            int value_N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[8 + 2 * i] })) - 30;
                            if (value_N2 >= 10)
                                value_N2--;
                            switch (Check_BoolUnit)
                            {
                                case BoolUnit.Y:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        int A = Check_BoolUnit_Start / 10, B = Check_BoolUnit_Start % 10;
                                        //B = B + k + i * 8;
                                        //A = A + (B / 8);
                                        //A = (A / 8) * 10 + A % 8;
                                        //B = B % 8;
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_Y[A][B] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_Y[A][B] = true;
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.M:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        int M = Check_BoolUnit_Start + i * 8 + k;
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_M[M] = false;
                                            if (M >= 9260 && M <= 9267)
                                                PLC_Value.Point_EC1X[M - 9260] = false;
                                            if (M >= 9270 && M <= 9277)
                                                PLC_Value.Point_EC1Y[M - 9270] = false;
                                            if (M >= 9280 && M <= 9287)
                                                PLC_Value.Point_EC2X[M - 9280] = false;
                                            if (M >= 9290 && M <= 9297)
                                                PLC_Value.Point_EC2Y[M - 9290] = false;
                                            if (M >= 9300 && M <= 9307)
                                                PLC_Value.Point_EC3X[M - 9300] = false;
                                            if (M >= 9310 && M <= 9317)
                                                PLC_Value.Point_EC3Y[M - 9310] = false;
                                            //Console.WriteLine("M" + (Check_BoolUnit_Start + i * 8 + k) + " :" + "False");
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_M[M] = true;
                                            if (M >= 9260 && M <= 9267)
                                                PLC_Value.Point_EC1X[M - 9260] = true;
                                            if (M >= 9270 && M <= 9277)
                                                PLC_Value.Point_EC1Y[M - 9270] = true;
                                            if (M >= 9280 && M <= 9287)
                                                PLC_Value.Point_EC2X[M - 9280] = true;
                                            if (M >= 9290 && M <= 9297)
                                                PLC_Value.Point_EC2Y[M - 9290] = true;
                                            if (M >= 9300 && M <= 9307)
                                                PLC_Value.Point_EC3X[M - 9300] = true;
                                            if (M >= 9310 && M <= 9317)
                                                PLC_Value.Point_EC3Y[M - 9310] = true;
                                            //Console.WriteLine("M" + (Check_BoolUnit_Start + i * 8 + k) + " :" + "True");
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.S:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_S[Check_BoolUnit_Start + i * 8 + k] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_S[Check_BoolUnit_Start + i * 8 + k] = true;
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.T:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_T[Check_BoolUnit_Start + i * 8 + k] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_T[Check_BoolUnit_Start + i * 8 + k] = true;
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.C:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_C[Check_BoolUnit_Start + i * 8 + k] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_C[Check_BoolUnit_Start + i * 8 + k] = true;
                                        }
                                    }
                                }
                                    break;
                            }
                        }

                        #endregion
                    }

                    Temp_Comment = new byte[0];
                    if (Action_Handle != null)
                        Action_Handle.Set();
                    Handle.Set();
                }
            }
        }

        //回傳指令接收函式
        private static void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            //SerialPort sp = plcconnector.getSerialPort();

            Byte[] Buffer = new Byte[1024];
            Int32 length = com.Read(Buffer, 0, Buffer.Length);
            Array.Resize(ref Buffer, length);
            string Back = "";
            Back += String.Format("{0}{1}", BitConverter.ToString(Buffer), Environment.NewLine);
            string BackValue = Back.Replace("-", " ");
            // if (PLC_ShowData)
            // {
            //     Console.WriteLine("[Back_value]: " + BackValue + " ");
            //     Console.WriteLine("==========================");
            // }//58 13 10
            //Console.WriteLine("+++"+ Buffer[0]+" "+ Buffer[length-2]+" "+ Buffer[length-1]);
            if (Protocol_Type == Protocol.ASCII)
            {
                if (Buffer[0] == 58)
                {
                    if (length < 2 || Buffer[length - 2] != 13 || Buffer[length - 1] != 10)
                    {
                        int Temp_Lenth = Temp_Comment.Length;
                        Array.Resize(ref Temp_Comment, length + Temp_Lenth);
                        for (int i = 0; i < length; i++)
                        {
                            Temp_Comment[Temp_Lenth + i] = Buffer[i];
                        }

                        Buffer = Temp_Comment;
                        length = Buffer.Length;
                    }
                }
                else
                {
                    int Temp_Lenth = Temp_Comment.Length;
                    Array.Resize(ref Temp_Comment, length + Temp_Lenth);
                    for (int i = 0; i < length; i++)
                    {
                        Temp_Comment[Temp_Lenth + i] = Buffer[i];
                    }

                    Buffer = Temp_Comment;
                    length = Buffer.Length;
                }

                if (length >= 3 && Buffer[0] == 58 && Buffer[length - 2] == 13 && Buffer[length - 1] == 10)
                {
                    if (Buffer[3] != 48 || Buffer[4] != 49)
                    {
                        if (Buffer[3] == 48 && Buffer[4] == 50)
                        {
                            int DataCount = 0;

                            #region 確認資料組數

                            int N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[5] })) - 30;
                            if (N1 >= 10)
                                N1--;
                            int N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[6] })) - 30;
                            if (N2 >= 10)
                                N2--;
                            DataCount = N1 * 16 + N2;

                            #endregion

                            #region 填入資料

                            for (int i = 0; i < DataCount; i++)
                            {
                                int value_N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[7 + 2 * i] })) - 30;
                                if (value_N1 >= 10)
                                    value_N1--;
                                int value_N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[8 + 2 * i] })) - 30;
                                if (value_N2 >= 10)
                                    value_N2--;
                                switch (Check_BoolUnit)
                                {
                                    case BoolUnit.X:
                                    {
                                        int Value = value_N1 * 16 + value_N2;
                                        string s = Convert.ToString(Value, 2);
                                        int n = s.Length;
                                        if (i == DataCount - 1)
                                        {
                                            for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                            {
                                                s = "0" + s;
                                            }
                                        }
                                        else
                                        {
                                            for (int j = n; j < 8; j++)
                                            {
                                                s = "0" + s;
                                            }
                                        }

                                        char[] Bit = s.Reverse().ToArray();
                                        for (int k = 0; k < Bit.Length; k++)
                                        {
                                            int A = Check_BoolUnit_Start / 10, B = Check_BoolUnit_Start % 10;
                                            //B = B + k + i * 8;
                                            //A = A + (B / 8);
                                            //A = (A / 8) * 10 + A % 8;
                                            //B = B % 8;
                                            if (Bit[k] == '0')
                                            {
                                                PLC_Value.Point_X[A][B] = false;
                                            }
                                            else if (Bit[k] == '1')
                                            {
                                                PLC_Value.Point_X[A][B] = true;
                                            }
                                        }
                                    }
                                        break;
                                }
                            }

                            #endregion
                        }
                        else if (Buffer[3] == 48 && Buffer[4] == 51)
                        {
                            int DataCount = 0;

                            #region 確認資料組數

                            int N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[5] })) - 30;
                            if (N1 >= 10)
                                N1--;
                            int N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[6] })) - 30;
                            if (N2 >= 10)
                                N2--;
                            DataCount = N1 * 16 + N2;

                            #endregion

                            #region 填入資料

                            for (int i = 0; i < DataCount / 2; i++)
                            {
                                int value_N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[7 + 4 * i] })) - 30;
                                if (value_N1 >= 10)
                                    value_N1--;
                                int value_N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[8 + 4 * i] })) - 30;
                                if (value_N2 >= 10)
                                    value_N2--;
                                int value_N3 = int.Parse(BitConverter.ToString(new byte[] { Buffer[9 + 4 * i] })) - 30;
                                if (value_N3 >= 10)
                                    value_N3--;
                                int value_N4 = int.Parse(BitConverter.ToString(new byte[] { Buffer[10 + 4 * i] })) - 30;
                                if (value_N4 >= 10)
                                    value_N4--;
                                switch (Check_ValueUnit)
                                {
                                    case ValueUnit.D:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_D[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[D" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_D[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.R:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_R[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[R" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_R[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.T:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_T[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[T" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_T[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.C_16bit:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_Ctemp[Check_ValueUnit_Start + i] = value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //Console.WriteLine("[C" + (Check_ValueUnit_Start + i) + "]: " + PLC_Value.Value_C3[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                    case ValueUnit.C_32bit:
                                    {
                                        if (value_N1 > 7)
                                        {
                                            value_N1 = -(15 - value_N1);
                                            value_N2 = -(15 - value_N2);
                                            value_N3 = -(15 - value_N3);
                                            value_N4 = -(16 - value_N4);
                                        }

                                        PLC_Value.Value_Ctemp[Check_ValueUnit_Start + i + 200] =
                                            value_N1 * 16 * 16 * 16 +
                                            value_N2 * 16 * 16 + value_N3 * 16 + value_N4;
                                        //if(i/2==0)
                                        //Console.WriteLine("[C" + (Check_ValueUnit_Start + i/2) + "上位]: " + PLC_Value.Value_C3[Check_ValueUnit_Start + i] + " ");
                                        //else
                                        //Console.WriteLine("[C" + (Check_ValueUnit_Start + i/2) + "下位]: " + PLC_Value.Value_C3[Check_ValueUnit_Start + i] + " ");
                                    }
                                        break;
                                }
                            }

                            #endregion
                        }
                    }
                    else
                    {
                        int DataCount = 0;

                        #region 確認資料組數

                        int N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[5] })) - 30;
                        if (N1 >= 10)
                            N1--;
                        int N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[6] })) - 30;
                        if (N2 >= 10)
                            N2--;
                        DataCount = N1 * 16 + N2;

                        #endregion

                        #region 填入資料

                        for (int i = 0; i < DataCount; i++)
                        {
                            int value_N1 = int.Parse(BitConverter.ToString(new byte[] { Buffer[7 + 2 * i] })) - 30;
                            if (value_N1 >= 10)
                                value_N1--;
                            int value_N2 = int.Parse(BitConverter.ToString(new byte[] { Buffer[8 + 2 * i] })) - 30;
                            if (value_N2 >= 10)
                                value_N2--;
                            switch (Check_BoolUnit)
                            {
                                case BoolUnit.Y:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        int A = Check_BoolUnit_Start / 10, B = Check_BoolUnit_Start % 10;
                                        //B = B + k + i * 8;
                                        //A = A + (B / 8);
                                        //A = (A / 8) * 10 + A % 8;
                                        //B = B % 8;
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_Y[A][B] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_Y[A][B] = true;
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.M:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        int M = Check_BoolUnit_Start + i * 8 + k;
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_M[M] = false;
                                            if (M >= 9260 && M <= 9267)
                                                PLC_Value.Point_EC1X[M - 9260] = false;
                                            if (M >= 9270 && M <= 9277)
                                                PLC_Value.Point_EC1Y[M - 9270] = false;
                                            if (M >= 9280 && M <= 9287)
                                                PLC_Value.Point_EC2X[M - 9280] = false;
                                            if (M >= 9290 && M <= 9297)
                                                PLC_Value.Point_EC2Y[M - 9290] = false;
                                            if (M >= 9300 && M <= 9307)
                                                PLC_Value.Point_EC3X[M - 9300] = false;
                                            if (M >= 9310 && M <= 9317)
                                                PLC_Value.Point_EC3Y[M - 9310] = false;
                                            //Console.WriteLine("M" + (Check_BoolUnit_Start + i * 8 + k) + " :" + "False");
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_M[M] = true;
                                            if (M >= 9260 && M <= 9267)
                                                PLC_Value.Point_EC1X[M - 9260] = true;
                                            if (M >= 9270 && M <= 9277)
                                                PLC_Value.Point_EC1Y[M - 9270] = true;
                                            if (M >= 9280 && M <= 9287)
                                                PLC_Value.Point_EC2X[M - 9280] = true;
                                            if (M >= 9290 && M <= 9297)
                                                PLC_Value.Point_EC2Y[M - 9290] = true;
                                            if (M >= 9300 && M <= 9307)
                                                PLC_Value.Point_EC3X[M - 9300] = true;
                                            if (M >= 9310 && M <= 9317)
                                                PLC_Value.Point_EC3Y[M - 9310] = true;
                                            //Console.WriteLine("M" + (Check_BoolUnit_Start + i * 8 + k) + " :" + "True");
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.S:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_S[Check_BoolUnit_Start + i * 8 + k] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_S[Check_BoolUnit_Start + i * 8 + k] = true;
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.T:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_T[Check_BoolUnit_Start + i * 8 + k] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_T[Check_BoolUnit_Start + i * 8 + k] = true;
                                        }
                                    }
                                }
                                    break;
                                case BoolUnit.C:
                                {
                                    int Value = value_N1 * 16 + value_N2;
                                    string s = Convert.ToString(Value, 2);
                                    int n = s.Length;
                                    if (i == DataCount - 1)
                                    {
                                        for (int j = n; j < Check_BoolUnit_Lenth - i * 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }
                                    else
                                    {
                                        for (int j = n; j < 8; j++)
                                        {
                                            s = "0" + s;
                                        }
                                    }

                                    char[] Bit = s.Reverse().ToArray();
                                    for (int k = 0; k < Bit.Length; k++)
                                    {
                                        if (Bit[k] == '0')
                                        {
                                            PLC_Value.Point_C[Check_BoolUnit_Start + i * 8 + k] = false;
                                        }
                                        else if (Bit[k] == '1')
                                        {
                                            PLC_Value.Point_C[Check_BoolUnit_Start + i * 8 + k] = true;
                                        }
                                    }
                                }
                                    break;
                            }
                        }

                        #endregion
                    }

                    Temp_Comment = new byte[0];
                    if (Action_Handle != null)
                        Action_Handle.Set();
                    Handle.Set();
                }
            }
            else if (Protocol_Type == Protocol.RTU)
            {
            }
            Array.Clear(Buffer, 0, Buffer.Length);
        }

        #endregion

        #region 開啟/關閉連接

        //開啟
        public static void PLC_On()
        {
            try
            {
                PLC_Value.Reset();
                com = new SerialPort(COM, 115200, Parity.None, 8, StopBits.One);
                com.Encoding = System.Text.Encoding.GetEncoding(28591);
                com.WriteTimeout = 1000;// 設置WriteTimeout屬性（以毫秒為單位）
                com.Open();
                int blockLimit = 1024;
                byte[] buffer = new byte[blockLimit];
                
                Handle.Reset();
                
                // SerialPort DataReceived code from this blog:
                // Original DataReceived event has some IO Exception issue, so we replace with BaseStream method.
                // https://sparxeng.com/blog/software/must-use-net-system-io-ports-serialport
                Action kickoffRead = null;
                kickoffRead = delegate {
                    com.BaseStream.BeginRead(buffer, 0, buffer.Length, delegate (IAsyncResult ar) {
                        try {
                            int actualLength = com.BaseStream.EndRead(ar);
                            byte[] received = new byte[actualLength];
                            Buffer.BlockCopy(buffer, 0, received, 0, actualLength);
                            DataReceived(received,actualLength);
                            Array.Clear(buffer,0,blockLimit);
                        }
                        catch (Exception exc) {
                            Log.Error("Exception at PLC reading:"+exc.ToString());
                        }
                        kickoffRead();
                    }, null);
                };
                kickoffRead();
                PLC_Run = true;
                PLC_SendComment = true;
                T.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "PLC連接錯誤");
            }
        }

        //關閉
        public static void PLC_Close()
        {
            PLC_Run = false;
            PLC_SendComment = false;
            //com.DataReceived -= new SerialDataReceivedEventHandler(DataReceivedHandler);
            com.Close();
        }

        #endregion

        #region 方法

        public static byte[] StringToByteArray(string hex) //函數:將字串轉Byte
        {
            return Enumerable.Range(0, hex.Length)
                .Where(x => x % 2 == 0)
                .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                .ToArray();
        }

        private static string LRC2(string n) //函數:LRC驗證 寫入數值
        {
            char[] kk; //存放 HexToChar 的值
            string[] kka; //存放kk相加的值
            ushort[] num; //存放轉換kka型態的值
            byte sumk = 0; //存放最後得到的LRC值
            string yui2 = "";
            yui2 = n;


            string hexValues = yui2;
            string[] strArray = hexValues.Split(' '); //分割字串，找出分割點字元"|" ，並將結果存入陣列
            kk = new char[strArray.Length];
            kka = new string[(strArray.Length) / 2];
            num = new ushort[(strArray.Length) / 2];

            //for (int i = 0; i < strArray.Length; i++)        //透過迴圈將陣列值取出 也可用foreach
            //{
            //    Console.WriteLine(strArray[i].ToString());
            //}
            //Console.ReadLine();

            for (int i = 0; i < strArray.Length; i++) //透過迴圈將陣列值轉換Char
            {
                int value = Convert.ToInt32(strArray[i], 16);
                kk[i] = (char)value;
            }

            for (int i = 0, j = 0; i < strArray.Length; i = i + 2, j++) //透過迴圈將兩個一組字串合併,接著相加出總和
            {
                kka[j] = kk[i].ToString() + kk[i + 1].ToString();
                num[j] = (ushort)Convert.ToInt32(kka[j], 16);
                sumk = (byte)(sumk + num[j]);
            }

            sumk = (byte)(~sumk);
            sumk += 1;
            string sumk16 = Convert.ToString(sumk, 16);

            sumk16 = sumk16.PadLeft(2, '0');
            return sumk16;
        }

        private static UInt16 CRC(string S) //RTU CRC驗證
        {
            int len = (S.Length + 1) / 3;
            byte[] buf = new byte[len];
            for (int i = 0; i < len; i++)
            {
                int N = Convert.ToInt32(S.Substring(i * 3, 2), 16);
                buf[i] = (byte)N;
            }

            UInt16 crc = 0xFFFF;

            for (int pos = 0; pos < len; pos++)
            {
                crc ^= (UInt16)buf[pos]; // 取出第一個byte與crc XOR

                for (int i = 8; i != 0; i--)
                {
                    // 巡檢每個bit  
                    if ((crc & 0x0001) != 0)
                    {
                        // 如果 LSB = 1   
                        crc >>= 1; // 右移1bit 並且 XOR 0xA001  
                        crc ^= 0xA001;
                    }
                    else // 如果 LSB != 1  
                        crc >>= 1; // 右移1bit
                }
            }

            return crc;
        }

        private static string convert16(string instr) //函數:char to hex
        {
            char[] wdchars = instr.ToCharArray();
            string outstr = ""; //輸出字串
            foreach (char wdchar in wdchars)
                outstr = (outstr + Convert.ToString(wdchar, 16) + " ").ToUpper();

            return outstr;
        }

        //開啟印出傳送指令及回傳指令功能
        public static void Open_ShowData()
        {
            PLC_ShowData = true;
        }

        //關閉印出傳送指令及回傳指令功能
        public static void Close_ShowData()
        {
            PLC_ShowData = false;
        }

        //開啟傳送指令
        public static void Open_SendComment()
        {
            PLC_SendComment = true;
        }

        //關閉傳送指令
        public static void Close_SendComment()
        {
            PLC_SendComment = false;
        }

        //32位元暫存器寫入
        public static void SendValue_32bit(int Stop, ValueUnit unit, int Number, int Value, bool Emer,
            ManualResetEvent Action_Handle)
        {
            int _startN1 = 0, _startN2 = 0;
            List<int> _list1 = new List<int>();
            if (Value >= 0)
            {
                if (Value >= 65536)
                {
                    _startN2 = Value / 65536;
                    int _n = Value % 65536;
                    if (_n > 32767)
                    {
                        _startN1 = -65536 + _n;
                    }
                    else
                        _startN1 = _n;
                }
                else
                {
                    if (Value > 32767)
                    {
                        _startN1 = -65536 + Value;
                    }
                    else
                        _startN1 = Value;
                }
            }
            else
            {
                _startN2 = ((Value + 1) / 65536) - 1;
                int _n = -Value % 65536;
                if (_n <= 32768)
                {
                    _startN1 = -_n;
                }
                else
                {
                    _startN1 = 65536 - _n;
                }
            }

            //Console.WriteLine("N1: "+ _startN1+" N2: "+ _startN2);
            _list1.Add(_startN1);
            _list1.Add(_startN2);
            if (unit == ValueUnit.C_32bit)
                SendValue(Stop, (ValueUnit)unit, Number, Value, Emer, Action_Handle);
            else
                SendValue(Stop, (ValueUnit)unit, Number, _list1, Emer, Action_Handle);
        }

        //32位元暫存器連續寫入
        public static void SendValue_32bit(int Stop, ValueUnit unit, int Number, List<int> Value, bool Emer,
            ManualResetEvent Action_Handle)
        {
            List<int> _list1 = new List<int>();
            for (int i = 0; i < Value.Count; i++)
            {
                int _startN1 = 0, _startN2 = 0;
                if (Value[i] >= 0)
                {
                    if (Value[i] >= 65536)
                    {
                        _startN2 = Value[i] / 65536;
                        int _n = Value[i] % 65536;
                        if (_n > 32767)
                        {
                            _startN1 = -65536 + _n;
                        }
                        else
                            _startN1 = _n;
                    }
                    else
                    {
                        if (Value[i] > 32767)
                        {
                            _startN1 = -65536 + Value[i];
                        }
                        else
                            _startN1 = Value[i];
                    }
                }
                else
                {
                    _startN2 = ((Value[i] + 1) / 65536) - 1;
                    int _n = -Value[i] % 65536;
                    if (_n <= 32768)
                    {
                        _startN1 = -_n;
                    }
                    else
                    {
                        _startN1 = 65536 - _n;
                    }
                }


                _list1.Add(_startN1);
                _list1.Add(_startN2);
            }

            SendValue(Stop, ValueUnit.D, Number, _list1, Emer, Action_Handle);
        }

        public static int GetValue_32bit(ValueUnit_32bt Unit, int Number)
        {
            int Value = 0;
            switch (Unit)
            {
                case ValueUnit_32bt.D:
                {
                    int value1 = PLC_Value.Value_D[Number];
                    int value2 = PLC_Value.Value_D[Number + 1];
                    if (value2 >= 0)
                    {
                        if (value1 < 0)
                        {
                            Value = value2 * 65536 + value1 + 65536;
                        }
                        else
                        {
                            Value = value2 * 65536 + value1;
                        }
                    }
                    else
                    {
                        if (value1 < 0)
                        {
                            Value = (value2 + 1) * 65536 + value1;
                        }
                        else
                        {
                            Value = value2 * 65536 + value1;
                        }
                    }
                }
                    break;
                case ValueUnit_32bt.R:
                {
                    int value1 = PLC_Value.Value_R[Number];
                    int value2 = PLC_Value.Value_R[Number + 1];
                    if (value2 >= 0)
                    {
                        if (value1 < 0)
                        {
                            Value = value2 * 65536 + value1 + 65536;
                        }
                        else
                        {
                            Value = value2 * 65536 + value1;
                        }
                    }
                    else
                    {
                        if (value1 < 0)
                        {
                            Value = (value2 + 1) * 65536 + value1;
                        }
                        else
                        {
                            Value = value2 * 65536 + value1;
                        }
                    }
                }
                    break;
                case ValueUnit_32bt.C_32bit:
                {
                    Value = int.Parse(PLC_Value.Value_C[Number].ToString());
                }
                    break;
            }

            return Value;
        }

        #endregion
    }

    #region 點位容器

    public class PLC_Value
    {
        public class int_C
        {
            public int this[int index]
            {
                get
                {
                    if (index < 200)
                        return Value_Ctemp[index];
                    else
                    {
                        if (Value_Ctemp[(index - 200) * 2 + 200] < 0)
                        {
                            return 65536 + Value_Ctemp[(index - 200) * 2 + 200] +
                                   Value_Ctemp[(index - 200) * 2 + 201] * 65536;
                        }
                        else
                        {
                            return Value_Ctemp[(index - 200) * 2 + 200] + Value_Ctemp[(index - 200) * 2 + 201] * 65536;
                        }
                    }
                }
            }
        }

        public static int[] Value_D;
        public static int[] Value_T;
        public static int_C Value_C;
        public static int[] Value_Ctemp;
        public static int[] Value_R;
        public static List<bool[]> Point_X;
        public static List<bool[]> Point_Y;
        public static bool[] Point_M;
        public static bool[] Point_S;
        public static bool[] Point_T;
        public static bool[] Point_C;
        public static bool[] Point_EC1Y;
        public static bool[] Point_EC1X;
        public static bool[] Point_EC2Y;
        public static bool[] Point_EC2X;
        public static bool[] Point_EC3Y;
        public static bool[] Point_EC3X;
        public static bool[] Point_Total;
        public static List<SavePoint_Data> SaveData_X = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_Y = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_D = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_T = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_C = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_R = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_M = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_S = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_EC1Y = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_EC2Y = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_EC3Y = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_EC1X = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_EC2X = new List<SavePoint_Data>();
        public static List<SavePoint_Data> SaveData_EC3X = new List<SavePoint_Data>();


        public static void Reset()
        {
            #region 暫存器

            Value_D = new int[9512];
            Value_T = new int[512];
            Value_Ctemp = new int[312];
            Value_R = new int[10000];
            Value_C = new int_C();

            #endregion

            #region 點位

            Point_X = new List<bool[]>();
            Point_Y = new List<bool[]>();
            for (int i = 0; i < 18; i++)
            {
                bool[] X = new bool[8] { false, false, false, false, false, false, false, false };
                Point_X.Add(X);
            }

            for (int i = 0; i < 18; i++)
            {
                bool[] Y = new bool[8] { false, false, false, false, false, false, false, false };
                Point_Y.Add(Y);
            }

            Point_M = new bool[9512];
            Point_S = new bool[4096];
            Point_T = new bool[512];
            Point_C = new bool[256];
            Point_EC1X = new bool[8];
            Point_EC1Y = new bool[8];
            Point_EC2X = new bool[8];
            Point_EC2Y = new bool[8];
            Point_EC3X = new bool[8];
            Point_EC3Y = new bool[8];
            Point_Total = new bool[13824];

            #endregion
        }

        public static void Save_Reset()
        {
            SaveData_X = new List<SavePoint_Data>();
            SaveData_Y = new List<SavePoint_Data>();
            SaveData_D = new List<SavePoint_Data>();
            SaveData_T = new List<SavePoint_Data>();
            SaveData_C = new List<SavePoint_Data>();
            SaveData_R = new List<SavePoint_Data>();
            SaveData_M = new List<SavePoint_Data>();
            SaveData_S = new List<SavePoint_Data>();
            SaveData_EC1Y = new List<SavePoint_Data>();
            SaveData_EC2Y = new List<SavePoint_Data>();
            SaveData_EC3Y = new List<SavePoint_Data>();
            SaveData_EC1X = new List<SavePoint_Data>();
            SaveData_EC2X = new List<SavePoint_Data>();
            SaveData_EC3X = new List<SavePoint_Data>();
        }
    }

    #endregion

    #region enum

    public enum ValueUnit
    {
        D = 0,
        T = 1,
        C_16bit = 2,
        C_32bit = 3,
        R = 4,
    }

    public enum BoolUnit
    {
        X = 0,
        Y = 1,
        M = 2,
        S = 3,
        T = 4,
        C = 5,
        EC1X = 6,
        EC2X = 7,
        EC3X = 8,
        EC1Y = 9,
        EC2Y = 10,
        EC3Y = 11,
        All = 12,
    }

    public enum ValueUnit_32bt
    {
        D = 0,
        C_32bit = 1,
        R = 2,
    }

    public enum Protocol
    {
        ASCII = 0,
        RTU = 1,
    }

    #endregion

    #region 下指令用存取結構

    public struct CommentData
    {
        public int stop;
        public ValueUnit valueUnit;
        public BoolUnit boolUnit;
        public int Start;
        public int Lenth;
        public byte[] Comment;
        public ManualResetEvent Action_Handle;
    }

    #endregion

    public struct SavePoint_Data
    {
        public int Number;
        public string Data;
    }
}