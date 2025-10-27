using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace peilin
{
    public class testPerPixel
    {
        public static void test_PerPixel()
        {
            // 加载测试图片
            string imagePath = @"C:\Users\User\Desktop\peilin2\bin\x64\Release\025.png"; // 替换为你的图片路径
            Mat inputImage = Cv2.ImRead(imagePath);

            if (inputImage.Empty())
            {
                Console.WriteLine("无法加载图片！");
                return;
            }

            // 创建实例并调用 ROI 提取方法
            testPerPixel instance = new testPerPixel();
            Mat roiImage = instance.DetectAndExtractROI(inputImage, 1);

            if (roiImage == null)
            {
                Console.WriteLine("未检测到有效的内外圆！");
                return;
            }

            // 使用 DetectAndExtractROI 检测到的外圈进行计算
            int outerRadiusPixels = instance.GetDetectedOuterRadius();

            if (outerRadiusPixels > 0)
            {
                double actualOuterDiameterMm = 40.03; // 已知外圈实际直径（毫米）
                double mmPerPixel = actualOuterDiameterMm / (2 * outerRadiusPixels);

                Console.WriteLine($"每像素对应的毫米值: {mmPerPixel:F4} mm/pixel");
            }
            else
            {
                Console.WriteLine("外圈未成功检测，无法计算每像素毫米值。");
            }
        }

        private int detectedOuterRadius = -1; // 存储检测到的外圈半径

        private int GetDetectedOuterRadius()
        {
            return detectedOuterRadius;
        }

        private Mat DetectAndExtractROI(Mat inputImage, int stop)
        {
            // 从 app.param 中读取霍夫圆参数
            int outerMinRadius = 550;
            int outerMaxRadius = 570;
            int innerMinRadius = 450;
            int innerMaxRadius = 500;
            int outerP1 = 120;
            int outerP2 = 20;
            int innerP1 = 120;
            int innerP2 = 20;
            int outerMinDist = 50;
            int innerMinDist = 50;

            int centerTolerance = 50; // 圆心容许偏差
            bool matched = false;
            Mat roi_blurred = null;

            Mat gray = inputImage.CvtColor(ColorConversionCodes.BGR2GRAY);
            Mat blurred = gray.GaussianBlur(new Size(5, 5), 1);

            // 霍夫圆检测：外圈
            var outerCircles = Cv2.HoughCircles(
                blurred,
                HoughModes.Gradient,
                dp: 1,
                minDist: outerMinDist,
                param1: outerP1,
                param2: outerP2,
                minRadius: outerMinRadius,
                maxRadius: outerMaxRadius);

            // 霍夫圆检测：内圈
            var innerCircles = Cv2.HoughCircles(
                blurred,
                HoughModes.Gradient,
                dp: 1,
                minDist: innerMinDist,
                param1: innerP1,
                param2: innerP2,
                minRadius: innerMinRadius,
                maxRadius: innerMaxRadius);

            if (outerCircles.Length > 0 && innerCircles.Length > 0)
            {
                foreach (var outer in outerCircles)
                {
                    foreach (var inner in innerCircles)
                    {
                        // 计算圆心距离
                        double centerDistance = Math.Sqrt(Math.Pow(outer.Center.X - inner.Center.X, 2) + Math.Pow(outer.Center.Y - inner.Center.Y, 2));

                        if (centerDistance <= centerTolerance)
                        {
                            // 更新检测到的外圈半径
                            detectedOuterRadius = (int)outer.Radius;

                            // 建mask
                            Mat mask = new Mat(inputImage.Size(), MatType.CV_8UC1, Scalar.Black);
                            Cv2.Circle(mask, (Point)outer.Center, (int)outer.Radius, Scalar.White, -1);
                            Cv2.Circle(mask, (Point)inner.Center, (int)inner.Radius, Scalar.Black, -1);

                            Mat roi_full = new Mat();
                            Cv2.BitwiseAnd(inputImage, inputImage, roi_full, mask);

                            // 转灰阶 + 模糊 => roi_blurred
                            Mat roi_gray = new Mat();
                            Cv2.CvtColor(roi_full, roi_gray, ColorConversionCodes.BGR2GRAY);
                            roi_blurred = new Mat();
                            Cv2.GaussianBlur(roi_gray, roi_blurred, new Size(5, 5), 0);

                            Cv2.CvtColor(roi_blurred, roi_blurred, ColorConversionCodes.GRAY2BGR);

                            matched = true;
                            break;
                        }
                    }
                    if (matched) break;
                }
            }
            if (!matched)
            {
                // If not matched => 直接用原图 or return null
                return inputImage; // or return null
            }

            return roi_blurred;
        }
    }
}
