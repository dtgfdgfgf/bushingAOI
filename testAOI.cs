using System;
using System.IO;
using OpenCvSharp;

namespace peilin
{
    public class testAOI
    {
        const double PIXEL_TO_MM = 0.0266;  // 像素→毫米 (示例)
        const double MAX_GAP_MM = 1.5;      // 若开口>1.5mm => NG

        public static void test_AOI()
        {
            // 1) 读图
            string folderPath = @"C:\Workspace\anomalib\datasets\MVTec\bush\ori\bush2out\金屬面刮痕";
            //string imagePath = @"C:\Workspace\anomalib\datasets\MVTec\bush\ori\bush2out\金屬面刮痕\005.png";
            string storePath = @"C:\Users\User\Desktop\peilin2\bin\x64\Release\test1\";
            string[] pngFiles = Directory.GetFiles(folderPath, "*.png", SearchOption.TopDirectoryOnly);

            foreach (string file in pngFiles)
            {
                //Mat src = Cv2.ImRead(imagePath, ImreadModes.Color);
                Mat src = Cv2.ImRead(file);
                if (src.Empty())
                {
                    Console.WriteLine("读图失败.");
                    return;
                }

                // 2) 可先做一次简单阈值(可选),
                //    也可等DetectAndExtractROI后再做, 视你的流程而定
                Mat gray = new Mat();
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

                Mat ringThresh = new Mat();
                Cv2.Threshold(gray, ringThresh, 210, 255, ThresholdTypes.Binary);
                // 这里 200 仅是demo, 可换Otsu或别的

                //Cv2.NamedWindow("ringThresh", WindowFlags.KeepRatio);
                //Cv2.ImShow("ringThresh", ringThresh);
                //Cv2.WaitKey(300);

                // 3) 调用 "DetectAndExtractROI" 用霍夫圆找外圈+内圈 => 匹配 => 生成 roi_blurred
                var (matched, innerCircle, roi_blurred) = DetectAndExtractROI(src);
                if (!matched || innerCircle == null)
                {
                    Console.WriteLine("外圈内圈未成功匹配, 无法继续开口检测.");
                    return;
                }
                Point2f center = innerCircle.Value.Center;
                double radius = innerCircle.Value.Radius;

                // 4) 对 "ringThresh" (或roi_blurred做阈值后) 执行极坐标扫描
                double angleStep = 0.5;
                int nSteps = (int)(360.0 / angleStep);

                bool[] isHole = new bool[nSteps];

                // 用 min(R, 10) 之类做一点margin? 依需求
                radius = radius - 10; 

                // 可视化: 在 src 上画点(绿=环, 红=hole)
                for (int i = 0; i < nSteps; i++)
                {
                    double angleDeg = i * angleStep;
                    double rad = angleDeg * Math.PI / 180.0;

                    double rx = center.X + radius * Math.Cos(rad);
                    double ry = center.Y + radius * Math.Sin(rad);

                    int px = (int)Math.Round(rx);
                    int py = (int)Math.Round(ry);

                    if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                    {
                        // hole
                        isHole[i] = true;
                        Cv2.Circle(src, new Point(px, py), 2, Scalar.Red, -1);
                    }
                    else
                    {
                        byte val = ringThresh.Get<byte>(py, px);
                        if (val < 127)
                        {
                            // ring
                            isHole[i] = false;
                            Cv2.Circle(src, new Point(px, py), 1, Scalar.Green, -1);
                        }
                        else
                        {
                            // hole
                            isHole[i] = true;
                            Cv2.Circle(src, new Point(px, py), 2, Scalar.Red, -1);
                        }
                    }
                }

                // 5) 找最大连续 hole(双倍循环+用 idx 计)
                double maxGapAngleDeg = 0;
                int startIdx = -1;
                for (int idx = 0; idx < nSteps * 2; idx++)
                {
                    int realIdx = idx % nSteps;
                    if (isHole[realIdx])
                    {
                        if (startIdx < 0) startIdx = idx;
                    }
                    else
                    {
                        if (startIdx >= 0)
                        {
                            int length = idx - startIdx;
                            double gapDeg = length * angleStep;
                            if (gapDeg > maxGapAngleDeg) maxGapAngleDeg = gapDeg;
                            startIdx = -1;
                        }
                    }
                }
                if (startIdx >= 0)
                {
                    int length = (nSteps * 2) - startIdx;
                    double gapDeg = length * angleStep;
                    if (gapDeg > maxGapAngleDeg) maxGapAngleDeg = gapDeg;
                }

                // 6) gapArcPx= radius*(gapAngle(弧度)), => mm
                double gapArcPx = radius * (maxGapAngleDeg * Math.PI / 180.0);
                double gapArcMm = gapArcPx * PIXEL_TO_MM;

                Console.WriteLine($"最大開口角度={maxGapAngleDeg:F2} deg => 弧長(px)={gapArcPx:F2} => {gapArcMm:F2} mm");

                bool isNG = gapArcMm >= MAX_GAP_MM;
                if (isNG) Console.WriteLine("=> 開口過大 => NG");
                else Console.WriteLine("=> 開口 OK");

                Cv2.PutText(
                            src,                                // 图像
                            $"maxDeg: {maxGapAngleDeg:F2} deg  arcLen(px): {gapArcPx:F2}  arcLen(mm): {gapArcMm:F2}",  // 显示的第一行文字
                            new Point(30, 60),                    // 文字起始位置
                            HersheyFonts.HersheySimplex,       // 字体
                            2,                               // 字体大小
                            Scalar.White,                      // 文字颜色 (白色)
                            3,                                 // 文字线条粗细
                            LineTypes.AntiAlias                // 抗锯齿
                        );

                Cv2.ImWrite(storePath + Path.GetFileName(file), src);
                //Cv2.ImWrite(storePath + Path.GetFileName(imagePath), src);
                //Cv2.NamedWindow("src", WindowFlags.KeepRatio);
                //Cv2.ImShow("src", src);
                //Cv2.WaitKey();
            }
        }


        /// <summary>
        /// 用霍夫圆先找外圈(outer)与内圈(inner),
        /// 若圆心距离<= centerTolerance => 视为匹配,
        /// 然后建一个mask(外圈=white, 内圈=black), bitwiseAnd,
        /// 产出 "roi_blurred" 并返回.
        /// 
        /// 同时返回 "outerCircle" 以便后续极坐标扫描.
        /// 
        /// 你也可改成仅找外圈. 
        /// </summary>
        private static (bool matched, CircleSegment? outerCircle, Mat roi_blurred)
            DetectAndExtractROI(Mat inputImage)
        {
            // 你给的参数(可再调):
            int outerMinRadius = 610;
            int outerMaxRadius = 620;
            int innerMinRadius = 390;
            int innerMaxRadius = 400;
            int outerP1 = 120;
            int outerP2 = 20;
            int innerP1 = 120;
            int innerP2 = 20;
            int outerMinDist = 50;
            int innerMinDist = 50;
            int centerTolerance = 50;

            bool matched = false;
            Mat roi_blurred = null;
            CircleSegment? bestInner = null;

            // 1) 灰阶+模糊
            Mat gray = inputImage.CvtColor(ColorConversionCodes.BGR2GRAY);
            Mat blurred = gray.GaussianBlur(new Size(5, 5), 1);

            // 2) 霍夫圆检测：外圈
            var outerCircles = Cv2.HoughCircles(
                blurred, HoughModes.Gradient,
                dp: 1, minDist: outerMinDist,
                param1: outerP1, param2: outerP2,
                minRadius: outerMinRadius, maxRadius: outerMaxRadius);

            // 3) 霍夫圆检测：内圈
            var innerCircles = Cv2.HoughCircles(
                blurred, HoughModes.Gradient,
                dp: 1, minDist: innerMinDist,
                param1: innerP1, param2: innerP2,
                minRadius: innerMinRadius, maxRadius: innerMaxRadius);

            if (outerCircles.Length > 0 && innerCircles.Length > 0)
            {
                foreach (var outer in outerCircles)
                {
                    foreach (var inner in innerCircles)
                    {
                        double centerDistance = Math.Sqrt(
                            Math.Pow(outer.Center.X - inner.Center.X, 2) +
                            Math.Pow(outer.Center.Y - inner.Center.Y, 2));
                        if (centerDistance <= centerTolerance)
                        {
                            // 视为匹配
                            bestInner = inner;
                            // 建 mask
                            Mat mask = new Mat(inputImage.Size(), MatType.CV_8UC1, Scalar.Black);
                            // 外圈=white
                            Cv2.Circle(mask, (Point)outer.Center, (int)outer.Radius, Scalar.White, -1);
                            // 内圈=black
                            Cv2.Circle(mask, (Point)inner.Center, (int)inner.Radius, Scalar.Black, -1);

                            Mat roi_full = new Mat();
                            Cv2.BitwiseAnd(inputImage, inputImage, roi_full, mask);

                            // 转灰阶+模糊
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
                // 失败 => 返回 false
                return (false, null, inputImage);
            }
            // 成功 => 返回 outerCircle & roi_blurred
            return (true, bestInner, roi_blurred);
        }
    }
}
