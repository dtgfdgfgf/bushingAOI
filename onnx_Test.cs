using System;
using System.IO;
using System.Linq;
using System.Drawing; // 若要用 Bitmap 讀圖
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using System.Collections.Generic;

namespace peilin
{
    public class onnx_Test
    {
        public static void onnxTest()
        {
            // (1) 寫死路徑: ONNX 模型、圖片資料夾
            string onnxPath = @"C:\Users\User\Desktop\01.onnx";
            string folderPath = @"C:\Workspace\anomalib\datasets\MVTec\bush\crop\b1_256_2\dirty";

            // (2) 建立 InferenceSession (載入 ONNX)
            var session = new InferenceSession(onnxPath);
            Console.WriteLine("ONNX 模型載入成功!");

            // (3) 初始化結果計數字典
            Dictionary<int, int> resultCount = new Dictionary<int, int>();

            // (4) 掃描資料夾內的所有 PNG 圖片
            var imageFiles = Directory.GetFiles(folderPath, "*.png");
            Console.WriteLine($"發現 {imageFiles.Length} 張圖片!");

            foreach (var imagePath in imageFiles)
            {
                Console.WriteLine($"正在推論圖片: {Path.GetFileName(imagePath)}");

                // (5) 讀取影像並轉成 Tensor
                var inputTensor = PreprocessImage(imagePath, 256, 256);

                // (6) 準備 NamedOnnxValue
                string inputName = session.InputMetadata.Keys.First();
                var inputs = new[] { NamedOnnxValue.CreateFromTensor(inputName, inputTensor) };

                // (7) 進行推論
                var results = session.Run(inputs);

                // (8) 取輸出並找到最大值索引
                var resultTensor = results.First().AsTensor<float>();
                float[] scores = resultTensor.ToArray();
                int predictedIndex = Array.IndexOf(scores, scores.Max());

                // (9) 更新結果計數
                if (resultCount.ContainsKey(predictedIndex))
                {
                    resultCount[predictedIndex]++;
                }
                else
                {
                    resultCount[predictedIndex] = 1;
                }

                Console.WriteLine($"圖片: {Path.GetFileName(imagePath)}, 索引: {predictedIndex}, 分數: {scores[predictedIndex]:F4}");
            }

            // (10) 統計並輸出結果
            Console.WriteLine("\n推論結果統計:");
            foreach (var kvp in resultCount.OrderBy(k => k.Key))
            {
                Console.WriteLine($"索引: {kvp.Key}, 出現次數: {kvp.Value}");
            }

            Console.WriteLine("推論測試完畢!");
            Console.ReadLine();
        }

        /// <summary>
        /// 簡單的影像前處理: 讀取 Bitmap, Resize 到 (newWidth,newHeight), 
        /// 然後做 1x3xH xW => float Tensor (沒做mean-std, 要視模型而定).
        /// </summary>
        private static DenseTensor<float> PreprocessImage(string imagePath, int newWidth, int newHeight)
        {
            // 讀取影像
            var bmp = new Bitmap(imagePath);
            // Resize
            var resized = new Bitmap(bmp, newWidth, newHeight);

            // 建立 Tensor: 假設 batch=1, channels=3
            var tensor = new DenseTensor<float>(new[] { 1, 3, newHeight, newWidth });

            for (int y = 0; y < newHeight; y++)
            {
                for (int x = 0; x < newWidth; x++)
                {
                    Color c = resized.GetPixel(x, y);
                    tensor[0, 0, y, x] = c.B; // channel 0
                    tensor[0, 1, y, x] = c.G; // channel 1
                    tensor[0, 2, y, x] = c.R; // channel 2
                }
            }

            return tensor;
        }
    }
}
