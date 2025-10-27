using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using OpenCvSharp;
using System.Diagnostics;
using System.Threading;

namespace peilin 
{
    #region 資料模型類別
    public class LoadModelResponse
    {
        public string message { get; set; }
        public string error { get; set; }
    }

    public class DetectionResult
    {
        public List<int> box { get; set; }
        public int class_id { get; set; }
        public string class_name { get; set; }
        public double score { get; set; }
    }

    public class DetectionResponse
    {
        public List<DetectionResult> detections { get; set; }
        public string error { get; set; }
    }
    #endregion

    public class YoloDetection
    {
        // 添加預熱相關的私有變數
        private readonly Dictionary<string, DateTime> _lastDetectionTime = new Dictionary<string, DateTime>();
        private readonly Dictionary<string, CancellationTokenSource> _warmupTokens = new Dictionary<string, CancellationTokenSource>();
        private readonly object _warmupLock = new object();

        // 將靜態方法轉換為實例方法
        public async Task<bool> IsServerAvailable(string serverBaseUrl)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await client.GetAsync(serverBaseUrl);
                    return response.IsSuccessStatusCode;
                }
                catch (HttpRequestException)
                {
                    return false;
                }
            }
        }

        public bool StartPythonServer(string batFilePath, string modelPath, string port = "5000")
        {
            try
            {
                if (!File.Exists(batFilePath))
                {
                    Debug.WriteLine($"BAT 檔案不存在: {batFilePath}");
                    return false;
                }

                ProcessStartInfo startInfo = new ProcessStartInfo(batFilePath);
                startInfo.CreateNoWindow = true;
                startInfo.Arguments = $"{modelPath} {port}";
                startInfo.UseShellExecute = true;
                startInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;

                Process.Start(startInfo);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"啟動伺服器 BAT 檔案失敗: {ex.Message}");
                return false;
            }
        }

        public async Task<LoadModelResponse> LoadYoloModel(string modelPath, string serverBaseUrl)
        {
            LoadModelResponse loadModelResponse = new LoadModelResponse() { message = null, error = null };

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string loadModelUrl = $"{serverBaseUrl}/load_model";
                    var requestData = new { model_path = modelPath };
                    string jsonData = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await client.PostAsync(loadModelUrl, content);
                    response.EnsureSuccessStatusCode();

                    string responseBody = await response.Content.ReadAsStringAsync();
                    loadModelResponse = JsonConvert.DeserializeObject<LoadModelResponse>(responseBody);
                }
                catch (HttpRequestException ex)
                {
                    loadModelResponse.error = $"載入模型 HTTP 請求錯誤: {ex.Message}";
                }
                catch (JsonException ex)
                {
                    loadModelResponse.error = $"載入模型 JSON 解析錯誤: {ex.Message}";
                }
                catch (Exception ex)
                {
                    loadModelResponse.error = $"載入模型發生錯誤: {ex.Message}";
                }
            }
            return loadModelResponse;
        }

        public async Task<DetectionResponse> PerformObjectDetection(Mat image, string serverUrl, bool isRealDetection = true)
        {
            // 如果是真實檢測，更新最後檢測時間
            if (isRealDetection)
            {
                lock (_warmupLock)
                {
                    // 從 serverUrl 提取基礎 URL
                    string baseUrl = serverUrl.Replace("/detect", "");
                    _lastDetectionTime[baseUrl] = DateTime.Now;
                }
            }

            DetectionResponse detectionResponse = new DetectionResponse() { detections = null, error = null };

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    byte[] imageBytes = image.ToBytes(".jpg");
                    string base64Image = Convert.ToBase64String(imageBytes);
                    var requestData = new { image = base64Image };
                    string jsonData = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await client.PostAsync(serverUrl, content);
                    response.EnsureSuccessStatusCode();

                    string responseBody = await response.Content.ReadAsStringAsync();
                    detectionResponse = JsonConvert.DeserializeObject<DetectionResponse>(responseBody);
                }
                catch (HttpRequestException ex)
                {
                    detectionResponse.error = $"HTTP 請求錯誤: {ex.Message}";
                }
                catch (JsonException ex)
                {
                    detectionResponse.error = $"JSON 解析錯誤: {ex.Message}";
                }
                catch (Exception ex)
                {
                    detectionResponse.error = $"發生錯誤: {ex.Message}";
                }
            }
            return detectionResponse;
        }

        public async Task<DetectionResponse> PerformSplitObjectDetection(Mat image, string serverUrl,
            Size imgSize, Size subSize, int step, float confThreshold, float nmsThreshold)
        {
            DetectionResponse detectionResponse = new DetectionResponse { detections = null, error = null };

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    byte[] imageBytes = image.ToBytes(".png");
                    string base64Image = Convert.ToBase64String(imageBytes);

                    var requestData = new
                    {
                        image = base64Image,
                        img_size = new int[] { imgSize.Width, imgSize.Height },
                        sub_size = new int[] { subSize.Width, subSize.Height },
                        step = step,
                        conf_threshold = confThreshold,
                        nms_threshold = nmsThreshold
                    };

                    string jsonData = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await client.PostAsync(serverUrl, content);
                    response.EnsureSuccessStatusCode();

                    string responseBody = await response.Content.ReadAsStringAsync();
                    detectionResponse = JsonConvert.DeserializeObject<DetectionResponse>(responseBody);
                }
                catch (HttpRequestException ex)
                {
                    detectionResponse.error = $"HTTP 請求錯誤: {ex.Message}";
                }
                catch (JsonException ex)
                {
                    detectionResponse.error = $"JSON 解析錯誤: {ex.Message}";
                }
                catch (Exception ex)
                {
                    detectionResponse.error = $"發生錯誤: {ex.Message}";
                }
            }

            return detectionResponse;
        }

        // 新增繪製結果的方法
        public Mat DrawDetectionResults(Mat image, DetectionResponse detectionResponse, float scoreThreshold = 0.5f)
        {
            if (detectionResponse.detections != null && detectionResponse.detections.Count > 0)
            {
                Scalar[] colors = { Scalar.Red, Scalar.Blue, Scalar.Green, Scalar.Yellow, Scalar.Orange,
                                    Scalar.Pink, Scalar.Violet, Scalar.LightGray, Scalar.White, Scalar.Black };

                foreach (var detection in detectionResponse.detections)
                {
                    if (detection.score > scoreThreshold)
                    {
                        // 計算瑕疵面積
                        int width = detection.box[2] - detection.box[0];
                        int height = detection.box[3] - detection.box[1];
                        int area = width * height;

                        // 計算長寬比 (取較大值除以較小值)
                        float aspectRatio = Math.Max(width, height) / (float)Math.Min(width, height);

                        // 如果面積小於閾值，跳過此瑕疵（不顯示，視為OK）
                        if (aspectRatio <= 2.5f && area < 400)
                        {
                            continue;
                        }

                        Cv2.Rectangle(image,
                                    new Point(detection.box[0], detection.box[1]),
                                    new Point(detection.box[2], detection.box[3]),
                                    colors[detection.class_id % colors.Length], 2);

                        Cv2.PutText(image,
                                    $"{detection.class_name} {detection.score:F2}",
                                    new Point(detection.box[0], detection.box[1] - 10),
                                    HersheyFonts.HersheySimplex, 0.5,
                                    colors[detection.class_id % colors.Length], 1, LineTypes.AntiAlias);

                        // 顯示長寬 (寬度x高度)
                        //int width = detection.box[2] - detection.box[0];
                        //int height = detection.box[3] - detection.box[1];
                        string sizeText = $"{width}x{height}";
                        Cv2.PutText(image,
                                    sizeText,
                                    new Point(detection.box[0], detection.box[3] + 20), // 框下方顯示
                                    HersheyFonts.HersheySimplex, 0.5,
                                    colors[detection.class_id % colors.Length], 1, LineTypes.AntiAlias);
                    }
                }
            }
            return image;
        }

        // 新增整合啟動服務器的方法
        public async Task<bool> ServerOn(string serverBaseUrl, string batFilePath, string modelPath, IProgress<string> progress = null)
        {
            progress?.Report("檢查伺服器連線...");

            // 從 serverBaseUrl 解析端口號
            string port = "5000"; // 預設端口
            if (serverBaseUrl.Contains(":"))
            {
                string[] parts = serverBaseUrl.Split(':');
                if (parts.Length > 2)
                {
                    // 格式: http://localhost:5001
                    string portPart = parts[2];
                    if (portPart.Contains("/"))
                    {
                        portPart = portPart.Split('/')[0];
                    }
                    port = portPart;
                }
            }
            //
            //
            //
            //riteLine(port);

            bool serverAvailable = await IsServerAvailable(serverBaseUrl);
            if (!serverAvailable)
            {
                progress?.Report("伺服器未連線，嘗試啟動伺服器...");
                bool serverStarted = StartPythonServer(batFilePath, modelPath, port);

                if (serverStarted)
                {
                    progress?.Report("伺服器啟動中，等待連線建立...");
                    await Task.Delay(3000);

                    serverAvailable = await IsServerAvailable(serverBaseUrl);
                    if (!serverAvailable)
                    {
                        progress?.Report("伺服器啟動失敗，請檢查設定。");
                        return false;
                    }
                    else
                    {
                        progress?.Report("伺服器已啟動成功。");
                        return true;
                    }
                }
                else
                {
                    progress?.Report("無法啟動伺服器，請檢查 BAT 檔案是否存在或是否有執行權限。");
                    return false;
                }
            }
            else
            {
                progress?.Report("伺服器已連線。");
                return true;
            }
        }
        // 開始持續預熱
        public void StartContinuousWarmup(string serverUrl, int intervalSeconds = 5)
        {
            lock (_warmupLock)
            {
                // 如果已經在預熱，先停止
                if (_warmupTokens.ContainsKey(serverUrl))
                {
                    _warmupTokens[serverUrl].Cancel();
                    _warmupTokens.Remove(serverUrl);
                }

                // 初始化記錄時間
                _lastDetectionTime[serverUrl] = DateTime.Now;

                // 建立新的取消令牌
                var cancellationToken = new CancellationTokenSource();
                _warmupTokens[serverUrl] = cancellationToken;

                // 啟動背景預熱任務
                Task.Run(async () => await ContinuousWarmupLoop(serverUrl, intervalSeconds, cancellationToken.Token));
            }
        }
        // 持續預熱循環
        private async Task ContinuousWarmupLoop(string serverUrl, int intervalSeconds, CancellationToken cancellationToken)
        {
            Console.WriteLine($"[WarmUp] 開始對 {serverUrl} 進行持續預熱 (間隔: {intervalSeconds}秒)");

            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    DateTime now = DateTime.Now;
                    DateTime lastDetection = _lastDetectionTime.ContainsKey(serverUrl) ? _lastDetectionTime[serverUrl] : DateTime.MinValue;

                    // 如果超過指定時間沒有真實檢測，執行預熱
                    if ((now - lastDetection).TotalSeconds >= intervalSeconds)
                    {
                        // 執行預熱
                        using (var dummyMat = new Mat(new Size(2448, 2048), MatType.CV_8UC3, Scalar.Black))
                        {
                            try
                            {
                                await PerformObjectDetection(dummyMat, $"{serverUrl}/detect", false);
                                //Console.WriteLine($"[WarmUp] 預熱成功: {serverUrl}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[WarmUp] 預熱失敗: {serverUrl}, 錯誤: {ex.Message}");
                            }
                        }
                    }

                    // 等待一段時間再檢查
                    await Task.Delay(1000, cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[WarmUp] 預熱循環發生錯誤: {ex.Message}");
                    await Task.Delay(5000, cancellationToken);
                }
            }

            Console.WriteLine($"[WarmUp] 停止對 {serverUrl} 的持續預熱");
        }
        // 停止所有預熱
        public void StopAllWarmup()
        {
            lock (_warmupLock)
            {
                foreach (var token in _warmupTokens.Values)
                {
                    token.Cancel();
                }
                _warmupTokens.Clear();
                _lastDetectionTime.Clear();
            }
        }

    }
}
