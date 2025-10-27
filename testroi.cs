using System;
using System.IO;
using OpenCvSharp;

namespace peilin
{
    class testroi
    {
        public static void test_roi()
        {
            string folderPath = @"C:\Workspace\anomalib\datasets\MVTec\bushbgr\bush2in\test\good";
            string storePath = @"C:\Users\User\Desktop\peilin2\bin\x64\Release\testroi\";

            // 抓取 .jpg
            string[] jpgFiles = Directory.GetFiles(folderPath, "*.png", SearchOption.TopDirectoryOnly);
            Directory.CreateDirectory(storePath);

            foreach (var filePath in jpgFiles)
            {
                // 1. 讀取 (如果你要最終保留彩色，就先用 Color)
                Mat src = Cv2.ImRead(filePath, ImreadModes.Color);

                // 2. 轉灰階後二值化
                Mat gray = new Mat();
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

                Mat mask = new Mat();
                Cv2.Threshold(gray, mask, 200, 255, ThresholdTypes.Binary);

                // 3. 用兩次 Erode + 兩次 Dilate 來嘗試斷開細連結
                Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));

                // 連續做兩次 Erode
                for (int i = 0; i < 2; i++)
                {
                    Cv2.Erode(mask, mask, kernel);
                }
                // 連續做兩次 Dilate
                for (int i = 0; i < 2; i++)
                {
                    Cv2.Dilate(mask, mask, kernel);
                }

                // 4. 找出所有外部輪廓
                Cv2.FindContours(
                    mask,
                    out Point[][] contours,
                    out HierarchyIndex[] hierarchy,
                    RetrievalModes.External,
                    ContourApproximationModes.ApproxSimple
                );

                // 5. 根據面積篩選：大於某個閾值 → 視為要去除的大區
                double areaThreshold = 1500;
                Mat removeMask = new Mat(mask.Size(), MatType.CV_8UC1, Scalar.Black);

                foreach (var contour in contours)
                {
                    double area = Cv2.ContourArea(contour);
                    if (area > areaThreshold)
                    {
                        // 把這些大塊輪廓填到 removeMask 裡
                        Cv2.DrawContours(
                            removeMask,
                            new[] { contour },
                            -1,
                            Scalar.White,
                            -1
                        );
                    }
                }

                // 6. 將原圖上 removeMask 白色的區域塗黑
                src.SetTo(new Scalar(0, 0, 0), removeMask);

                // 7. 輸出結果
                string fileName = Path.GetFileName(filePath);
                string outPath = Path.Combine(storePath, fileName);
                Cv2.ImWrite(outPath, src);

                Console.WriteLine($"Done: {fileName}");
            }

            Console.WriteLine("All done!");
        }
    }
}
