using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenCvSharp;

namespace peilin
{
    public class testAOI2
    {
        public static void test_AOI2()
        {
            string folderPath = @"C:\Users\User\Desktop\peilin2\bin\x64\Release\image\2025-02\0210\origin";
            //string folderPath = @"C:\Users\User\Desktop\peilin2\bin\x64\Release\test\";
            string storePath = @"C:\Users\User\Desktop\peilin2\bin\x64\Release\testdeform\";
            // 獲取資料夾內所有 PNG 文件
            string[] pngFiles = Directory.GetFiles(folderPath, "*.jpg", SearchOption.TopDirectoryOnly);
            //string imagePath = @"C:\Workspace\anomalib\datasets\MVTec\bush\ori\bush2in\PTFE刮痕\041.png";
            foreach (string file in pngFiles)
            {
                //Console.WriteLine($"正在處理文件: {file}");
                Mat src = Cv2.ImRead(file);
                //Mat src = Cv2.ImRead(imagePath, ImreadModes.Color);
                if (src.Empty())
                {
                    Console.WriteLine("读图失败.");
                    return;
                }
                Mat gray = new Mat();
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

                Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3), new OpenCvSharp.Point(-1, -1));
                Mat ringThresh = new Mat();
                Cv2.Threshold(gray, ringThresh, 200, 255, ThresholdTypes.Binary);
                //Cv2.Erode(ringThresh, ringThresh, kernel);
                /*
                Cv2.NamedWindow("ringThresh", WindowFlags.KeepRatio);
                Cv2.ImShow("ringThresh", ringThresh);
                Cv2.WaitKey();
                */
                var pro = new Mat();
                Point[][] contours;
                HierarchyIndex[] h;
                Cv2.Erode(ringThresh, pro, kernel);
                Cv2.Erode(pro, pro, kernel);
                Cv2.FindContours(pro, out contours, out h, RetrievalModes.List, ContourApproximationModes.ApproxSimple);
                //Cv2.NamedWindow("FindContours", WindowFlags.KeepRatio);
                //Cv2.ImShow("FindContours", pro);
                //Cv2.WaitKey();

                Point[][] temp = new Point[1][];
                var m = new Mat(pro.Size(), MatType.CV_8UC1, Scalar.Black);
                foreach (var item in contours)
                {
                    temp[0] = item;
                    var area = Cv2.ContourArea(item);
                    var rect = Cv2.BoundingRect(item);

                    if (rect.Width < 1650 && rect.Height < 1650 && rect.Width > 1000 && rect.Height > 1000 && area > 5000)
                    {
                        Cv2.DrawContours(m, temp, 0, Scalar.White, -1);
                    }
                }
                //Cv2.NamedWindow("m", WindowFlags.KeepRatio);
                //Cv2.ImShow("m", m);
                Cv2.ImWrite(storePath + Path.GetFileName(file), m);
                Cv2.WaitKey();
            }
        }
    }
}
