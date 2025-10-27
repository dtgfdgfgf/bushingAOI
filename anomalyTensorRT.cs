using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;
using OpenCvSharp;
using System.Diagnostics;
using System.Threading;
using OpenCvSharp.Internal.Vectors;

namespace AnomalyTensorRT
{
    public class TensorRT
    {
        [DllImport("AD_TRT_dll1.dll", EntryPoint = "single_inference", CallingConvention = CallingConvention.Cdecl)]
        private static extern float Inference1(IntPtr imgPtr, int num, int mode, out IntPtr dstPtr);
        
        [DllImport("AD_TRT_dll1.dll", EntryPoint = "create_model", CallingConvention = CallingConvention.Cdecl)]
        private static extern void CreateModel1(string modelPath, string metaPath, bool effAd);

        [DllImport("AD_TRT_dll1.dll", EntryPoint = "close", CallingConvention = CallingConvention.Cdecl)]
        private static extern void DestroyModel1();
        [DllImport("AD_TRT_dll2.dll", EntryPoint = "single_inference", CallingConvention = CallingConvention.Cdecl)]
        private static extern float Inference2(IntPtr imgPtr, int num, int mode, out IntPtr dstPtr);

        [DllImport("AD_TRT_dll2.dll", EntryPoint = "create_model", CallingConvention = CallingConvention.Cdecl)]
        private static extern void CreateModel2(string modelPath, string metaPath, bool effAd);

        [DllImport("AD_TRT_dll2.dll", EntryPoint = "close", CallingConvention = CallingConvention.Cdecl)]
        private static extern void DestroyModel2();
        [DllImport("AD_TRT_dll3.dll", EntryPoint = "single_inference", CallingConvention = CallingConvention.Cdecl)]
        private static extern float Inference3(IntPtr imgPtr, int num, int mode, out IntPtr dstPtr);

        [DllImport("AD_TRT_dll3.dll", EntryPoint = "create_model", CallingConvention = CallingConvention.Cdecl)]
        private static extern void CreateModel3(string modelPath, string metaPath, bool effAd);

        [DllImport("AD_TRT_dll3.dll", EntryPoint = "close", CallingConvention = CallingConvention.Cdecl)]
        private static extern void DestroyModel3();
        [DllImport("AD_TRT_dll4.dll", EntryPoint = "single_inference", CallingConvention = CallingConvention.Cdecl)]
        private static extern float Inference4(IntPtr imgPtr, int num, int mode, out IntPtr dstPtr);

        [DllImport("AD_TRT_dll4.dll", EntryPoint = "create_model", CallingConvention = CallingConvention.Cdecl)]
        private static extern void CreateModel4(string modelPath, string metaPath, bool effAd);

        [DllImport("AD_TRT_dll4.dll", EntryPoint = "close", CallingConvention = CallingConvention.Cdecl)]
        private static extern void DestroyModel4();
        /*
        static void Main(string[] args)
        {
            SingleTest1();

            //SingleTest2();
            //BatchTest();
            //BatchTest2();
            //BatchTest2();

        }
        */
        [HandleProcessCorruptedStateExceptions]
        public static (Mat, float) SingleTest1(Mat img, string enginename, string meta_dataname)
        {
            Console.WriteLine("Test started...");

            //var img = Cv2.ImRead(imgPath);
            Task.Factory.StartNew(() => CreateModel1(enginename, meta_dataname, false));
            Thread.Sleep(1000);
            
            var timer = new Stopwatch();
            timer.Start();
   
            float score = Inference1(img.CvPtr, 0, 1, out var dstPtr);

            Mat dst = new Mat(dstPtr);           
            timer.Stop();

            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");
            DestroyModel1();
            return (dst, score);        
        }
        public static (Mat, float) SingleTest2(Mat img, string enginename, string meta_dataname)
        {
            Console.WriteLine("Test started...");

            //var img = Cv2.ImRead(imgPath);
            Task.Factory.StartNew(() => CreateModel2(enginename, meta_dataname, false));
            Thread.Sleep(1000);

            var timer = new Stopwatch();
            timer.Start();

            float score = Inference2(img.CvPtr, 0, 1, out var dstPtr);

            Mat dst = new Mat(dstPtr);
            timer.Stop();

            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");
            DestroyModel2();
            return (dst, score);
        }
        public static (Mat, float) SingleTest3(Mat img, string enginename, string meta_dataname)
        {
            Console.WriteLine("Test started...");

            //var img = Cv2.ImRead(imgPath);
            Task.Factory.StartNew(() => CreateModel3(enginename, meta_dataname, false));
            Thread.Sleep(1000);

            var timer = new Stopwatch();
            timer.Start();

            float score = Inference3(img.CvPtr, 0, 1, out var dstPtr);

            Mat dst = new Mat(dstPtr);
            timer.Stop();

            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");
            DestroyModel3();
            return (dst, score);
        }
        public static (Mat, float) SingleTest4(Mat img, string enginename, string meta_dataname)
        {
            Console.WriteLine("Test started...");

            //var img = Cv2.ImRead(imgPath);
            Task.Factory.StartNew(() => CreateModel4(enginename, meta_dataname, false));
            Thread.Sleep(1000);

            var timer = new Stopwatch();
            timer.Start();

            float score = Inference4(img.CvPtr, 0, 1, out var dstPtr);

            Mat dst = new Mat(dstPtr);
            timer.Stop();

            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");
            DestroyModel4();
            return (dst, score);
        }
        #region unused
        /*
        public static (Mat ,float) SingleTest1(string imgPath, string enginename, string meta_dataname)
        {
            
            //Console.WriteLine("test");
            //var img = Cv2.ImRead(@"C:\Users\User\Desktop\datasets\030.png");
            //var enginename = @"C:\Users\User\Desktop\peilin - 複製\models\dirty1.engine";
            //var meta_dataname = @"C:\Users\User\Desktop\peilin - 複製\models\dirty1.json";
            //Task.Factory.StartNew(() => CreateModel(enginename, meta_dataname, false));
            
            //Task.Factory.StartNew(() => CreateModel(modelPath, metaPath, false));
            Console.WriteLine("Test started...");
            var img = Cv2.ImRead(imgPath);
            Task.Factory.StartNew(() => CreateModel(enginename, meta_dataname, false));

            Thread.Sleep(1000);
            var timer = new Stopwatch();
            timer.Start();
            float score = Inference(img.CvPtr, 0, 1, out var dstPtr);
            timer.Stop();
            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");

            Mat dst = new Mat(dstPtr);
            float threshold = 0.5f;
            string thresholdText = $"Threshold: {threshold:F2}";
            string scoreText = $"Score: {score:F2}";
            //int fontFace = HersheyFonts.HersheySimplex;
            double fontScale = 3;
            Scalar textColor = new Scalar(0, 255, 0); // 绿色
            int thickness = 3;
            int baseline = 0;

            // 计算文字的位置
            var thresholdSize = Cv2.GetTextSize(thresholdText, HersheyFonts.HersheySimplex, fontScale, thickness, out baseline);
            var scoreSize = Cv2.GetTextSize(scoreText, HersheyFonts.HersheySimplex, fontScale, thickness, out baseline);

            Point thresholdPoint = new Point(50, 100); // Threshold 的起点
            Point scorePoint = new Point(50, 100 + thresholdSize.Height + 5); // Score 在 Threshold 下方

            // 在图像上绘制文字
            Cv2.PutText(dst, thresholdText, thresholdPoint, HersheyFonts.HersheySimplex, fontScale, textColor, thickness);
            Cv2.PutText(dst, scoreText, scorePoint, HersheyFonts.HersheySimplex, fontScale, textColor, thickness);

            Console.WriteLine("score= " + score);
            Cv2.NamedWindow("dst", WindowFlags.KeepRatio);
            Cv2.ImShow("dst", dst);
            Cv2.WaitKey();
            DestroyModel();
            return (dst, score);
        }
        public static void SingleTest2()
        {
            var img = Cv2.ImRead("8_25.jpg");
            Task.Factory.StartNew(() => CreateModel("model2.engine", "metadata2.json", false));
            Thread.Sleep(2000);
            var timer = new Stopwatch();
            timer.Start();
            float score = Inference(img.CvPtr, 0, 0, out var dstPtr);
            timer.Stop();
            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");
            Mat dst = new Mat(dstPtr);
            Console.WriteLine("score= " + score);
            Cv2.ImShow("dst", dst);
            Cv2.WaitKey();
            var img2 = Cv2.ImRead("8_25.jpg");
            float score2 = Inference(img2.CvPtr, 0, 1, out var dstPtr2);
            Mat dst2 = new Mat(dstPtr2);
            Cv2.ImShow("dst", dst2);
            Cv2.WaitKey();
            DestroyModel();
        }
        public static void MultiTest()
        {
            var img = Cv2.ImRead("15_91.png");
            img = img.CvtColor(ColorConversionCodes.BGR2RGB);
            Task.Factory.StartNew(() => CreateModel("model.engine", "metadata.json", true));
            Thread.Sleep(2000);

            var timer = new Stopwatch();
            timer.Start();
            for (int i = 0; i < 20; i++)
            {
                float score = Inference(img.CvPtr, 0, 0, out var dstPtr);
                Mat dst = new Mat(dstPtr);
                Console.WriteLine(score);
            }

            timer.Stop();
            Console.WriteLine("Elapsed Time:" + timer.ElapsedMilliseconds + "ms");
            Console.ReadLine();
            DestroyModel();
        }
        public static void BatchTest()
        {
            var img = Cv2.ImRead("15_91.png");
            var img2 = Cv2.ImRead("49_162.png");
            Task.Factory.StartNew(() => CreateModel2("model1.engine", "metadata1.json", false, 0.5f));
            Thread.Sleep(5000);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Test1(img, img2);
            Console.ReadLine();

            DestroyModel();

        }

        private static void Test1(Mat img, Mat img2)
        {
            var timer = new Stopwatch();
            timer.Start();
            var ImgList = new List<Mat>();
            int count = 1000;
            for (int i = 0; i < 1024; i++)
            {
                if (i % 2 == 0)
                {
                    ImgList.Add(img.Clone());
                }
                else
                {
                    ImgList.Add(img2.Clone());
                }
            }

            var imgsPointer = ImgList.Select(x => x.CvPtr).ToArray();
            var vecfloats = new VectorOfFloat();
            var vecMats = new VectorOfMat();

            BatchInference(imgsPointer, imgsPointer.Length, 0, 0, vecMats.CvPtr, vecfloats.CvPtr);
            var dsts = vecMats.ToArray();
            var scores = vecMats.ToArray();

            timer.Stop();
            Console.WriteLine("Elapsed Time:" + (double)timer.ElapsedMilliseconds / (double)count + "ms");
            Console.WriteLine($"FPS: {1000.0 / (double)(timer.ElapsedMilliseconds / (double)count)}");
            //GC.Collect();
        }

        public static void BatchTest2()
        {
            var img = Cv2.ImRead("15_91.png");
            img = img.CvtColor(ColorConversionCodes.BGR2RGB);
            Task.Factory.StartNew(() => CreateModel2("model2.engine", "metadata2.json", true, 0.5f));
            Thread.Sleep(2000);
            var ImgList = new List<Mat>();
            int count = 512;
            for (int i = 0; i < count; i++)
            {
                ImgList.Add(img.Clone());
            }
            var timer = new Stopwatch();
            timer.Start();
            var imgsPointer = ImgList.Select(x => x.CvPtr).ToArray();
            var vecfloats = new VectorOfFloat();
            var vecMats = new VectorOfMat();
            BatchInference(imgsPointer, imgsPointer.Length, 0, 0, vecMats.CvPtr, vecfloats.CvPtr);
            var dsts = vecMats.ToArray();
            var scores = vecMats.ToArray();
            timer.Stop();
            Console.WriteLine("Elapsed Time:" + (double)timer.ElapsedMilliseconds / (double)count + "ms");
            Console.WriteLine($"FPS: {1000.0 / (double)(timer.ElapsedMilliseconds / (double)count)}");
            Console.ReadLine();

            timer.Stop();
            DestroyModel();
        }
        */
        #endregion  
    }
}
