using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using OpenCvSharp;

namespace OnnxTest
{
    public class OnnxTester
    {
        public static void TestOnnxModel()
        {
            string modelPath = @"C:\\Users\\Chernger\\Desktop\\peilin - 複製\\models\\scratch2.onnx"; // 替換為你的 ONNX 模型路徑
            string imagePath = @"C:\\Users\\Chernger\\Desktop\\peilin - 複製\\testImg\\002.png"; // 替換為你要測試的影像路徑
            string metadataPath = @"C:\\Users\\Chernger\\Desktop\\peilin - 複製\\models\\scratch2.json"; // 假設 JSON 與 ONNX 檔案同名

            try
            {
                // 檢查模型、影像與元數據是否存在
                if (!File.Exists(modelPath)) throw new FileNotFoundException($"模型檔案未找到: {modelPath}");
                if (!File.Exists(imagePath)) throw new FileNotFoundException($"影像檔案未找到: {imagePath}");
                if (!File.Exists(metadataPath)) throw new FileNotFoundException($"元數據檔案未找到: {metadataPath}");

                // 加載元數據
                var metadata = JsonSerializer.Deserialize<Metadata>(File.ReadAllText(metadataPath));
                if (metadata == null) throw new InvalidOperationException("無法加載元數據！");

                // 加載 ONNX 模型
                var session = new InferenceSession(modelPath);
                Console.WriteLine("模型成功載入！");

                // 載入影像並進行預處理
                Mat image = Cv2.ImRead(imagePath);
                if (image.Empty()) throw new InvalidOperationException("無法讀取影像！");

                Console.WriteLine("成功讀取影像！");
                var preprocessedImage = PreprocessImage(image, metadata.InputSize);

                // 準備推理輸入
                var inputTensor = NormalizeImage(preprocessedImage, metadata.InputSize);
                var inputName = session.InputMetadata.Keys.First();
                var inputs = new[] { NamedOnnxValue.CreateFromTensor(inputName, inputTensor) };

                // 執行推理
                var results = session.Run(inputs);

                var output0 = results.First().AsTensor<float>().ToArray();
                var prediction = Array.IndexOf(output0, output0.Max());
                //Console.WriteLine($"異常分數: {prediction}");

                Console.WriteLine("推理執行成功！");

                // 解析結果
                foreach (var output in results)
                {
                    Console.WriteLine($"名稱: {output.Name}");
                    if (output.Value is Tensor<float> tensor)
                    {
                        var data = tensor.ToArray();
                        Console.WriteLine($"資料大小: {data.Length}");

                        // 顯示熱圖
                        DisplayBlendedHeatmap(image, data, tensor.Dimensions[2], tensor.Dimensions[3]);

                        // 計算異常分數
                        CalculateAnomalyScores(data, tensor.Dimensions[2], tensor.Dimensions[3], metadata.PixelThreshold);
                    }
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"測試過程中出錯: {ex.Message}");
            }
        }

        private static void DisplayBlendedHeatmap(Mat originalImage, float[] heatmapArray, int width, int height)
        {
            var heatmap = new Mat(height, width, MatType.CV_32F, heatmapArray);
            Cv2.Normalize(heatmap, heatmap, 0, 255, NormTypes.MinMax);
            var heatmapUint8 = new Mat();
            heatmap.ConvertTo(heatmapUint8, MatType.CV_8U);
            Cv2.ApplyColorMap(heatmapUint8, heatmapUint8, ColormapTypes.Jet);

            Mat resizedHeatmap = new Mat();
            Cv2.Resize(heatmapUint8, resizedHeatmap, new Size(originalImage.Width, originalImage.Height));

            Mat resizedOriginal = new Mat();
            Cv2.Resize(originalImage, resizedOriginal, new Size(originalImage.Width, originalImage.Height));
            if (resizedOriginal.Channels() == 1)
            {
                Cv2.CvtColor(resizedOriginal, resizedOriginal, ColorConversionCodes.GRAY2BGR);
            }

            Mat blended = new Mat();
            Cv2.AddWeighted(resizedOriginal, 0.85, resizedHeatmap, 0.15, 0, blended);

            Cv2.ImShow("Blended Heatmap", blended);
            Cv2.ImWrite(@"C:\\Users\\Chernger\\Desktop\\peilin - 複製\\testImg\\031.png", blended);
            Cv2.WaitKey(0);
        }

        private static void CalculateAnomalyScores(float[] heatmapArray, int width, int height, float threshold)
        {
            var totalPixels = width * height;

            var heatmap = new Mat(height, width, MatType.CV_32F, heatmapArray);
            Cv2.Normalize(heatmap, heatmap, 0, 255, NormTypes.MinMax);
            var heatmapUint8 = new Mat();
            heatmap.ConvertTo(heatmapUint8, MatType.CV_8U);

            var maxScore = heatmapArray.Max();
            var avgScore = heatmapArray.Average();
            var anomalyPixelCount = heatmapArray.Count(value => value > threshold);
            var anomalyPixelRatio = (float)anomalyPixelCount / totalPixels;

            Console.WriteLine($"異常分數計算結果：");
            Console.WriteLine($"最大異常分數: {maxScore}");
            Console.WriteLine($"平均異常分數: {avgScore}");
            Console.WriteLine($"超過閥值 ({threshold}) 的像素比例: {anomalyPixelRatio * 100:F2}%");
        }

        private static Mat PreprocessImage(Mat img, int size)
        {
            var resized = new Mat();
            Cv2.Resize(img, resized, new Size(size, size));
            if (resized.Channels() == 1)
            {
                Cv2.CvtColor(resized, resized, ColorConversionCodes.GRAY2BGR);
            }
            return resized;
        }

        private static DenseTensor<float> NormalizeImage(Mat img, int size)
        {
            var mean = new[] { 0.485f, 0.456f, 0.406f };
            var std = new[] { 0.229f, 0.224f, 0.225f };
            var tensor = new DenseTensor<float>(new[] { 1, 3, size, size });

            for (int c = 0; c < 3; c++)
            {
                for (int h = 0; h < size; h++)
                {
                    for (int w = 0; w < size; w++)
                    {
                        var pixel = img.At<Vec3b>(h, w);
                        tensor[0, c, h, w] = (pixel[c] / 255.0f - mean[c]) / std[c];
                    }
                }
            }
            return tensor;
        }

        private class Metadata
        {
            public int InputSize { get; set; } = 432;
            public float PixelThreshold { get; set; } = 0.5f;
        }
    }
}
