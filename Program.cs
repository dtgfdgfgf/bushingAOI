using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using PLC;
using OnnxTest;
using AnomalyTensorRT;

using Basler.Pylon;
using basler;  // 你的 Camera 類別命名空間
using OpenCvSharp;

namespace peilin
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            //OnnxTester.TestOnnxModel();
            //.SingleTest1();
            //Application.Run(new defect_check_info());
            //Application.Run(new defect_type_info());
            //Application.Run(new type_info());
            //Application.Run(new parameter_info());
            //onnx_Test.onnxTest();
            //testPerPixel.test_PerPixel();
            //testAOI.test_AOI();
            //testAOI2.test_AOI2();
            //Application.Run(new PLC_Test());  
            //testroi.test_roi();
            //parameter_info parameter_info = new parameter_info();
        }
    }
}
