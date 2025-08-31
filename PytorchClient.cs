using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;
using OpenCvSharp;
using System.Diagnostics;

namespace CherngerUI
{
    public class ServerChecker
    {
        private class ServerInfo
        {
            public string Name { get; set; }
            public string BatFilePath { get; set; }
            public string ServerBaseUrl { get; set; }
        }
        public static async Task CheckAndStartServers()
        {
            var servers = new ServerInfo[]
            {
            new ServerInfo { Name = "front", BatFilePath = "start_server.bat", ServerBaseUrl = "http://localhost:5001" },
            new ServerInfo { Name = "back", BatFilePath = "start_server2.bat", ServerBaseUrl = "http://localhost:5002" }
            };

            foreach (var server in servers)
            {
                await ProcessServer(server);
            }
        }

        private static async Task ProcessServer(ServerInfo server)
        {
            Console.WriteLine($"檢查伺服器({server.Name})是否啟動...");
            bool serverAvailable = await PytorchClient.IsServerAvailable(server.ServerBaseUrl);

            if (!serverAvailable)
            {
                Console.WriteLine($"伺服器({server.Name})未連線，嘗試啟動伺服器...");
                bool serverStarted = PytorchClient.StartPythonServer(server.BatFilePath);

                if (serverStarted)
                {
                    Console.WriteLine($"伺服器({server.Name})啟動中，等待 5 秒...");
                    await Task.Delay(5000);

                    serverAvailable = await PytorchClient.IsServerAvailable(server.ServerBaseUrl);

                    if (!serverAvailable)
                    {
                        Console.WriteLine($"伺服器({server.Name})啟動失敗，請檢查 {server.BatFilePath} 檔案和 Python 環境設定。");
                        Console.WriteLine("\n結束...");
                        Environment.Exit(1); // 使用 Environment.Exit 以更乾淨地終止
                    }
                    else
                    {
                        Console.WriteLine($"伺服器({server.Name})已啟動成功。");
                    }
                }
                else
                {
                    Console.WriteLine($"無法啟動伺服器({server.Name})，請檢查 {server.BatFilePath} 檔案是否存在或是否有執行權限。");
                    Console.WriteLine("\n結束...");
                    Environment.Exit(1); // 使用 Environment.Exit 以更乾淨地終止
                }
            }
            else
            {
                Console.WriteLine($"伺服器({server.Name})已連線。");
            }
        }
    }

    public class LoadModelResponse // 定義 LoadModel API 的回應類別
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

    public class DetectionResponse // 定義 Detect API 回應類別
    {
        public List<DetectionResult> detections { get; set; }
        public string error { get; set; }
    }

    public static class PytorchClient
    {
        public static async Task Test() //測試連線用
        {
            string serverBaseUrl = "http://localhost:5002"; // 伺服器基礎 URL
            string detectApiUrl = $"{serverBaseUrl}/detect";
            string loadModelApiUrl = $"{serverBaseUrl}/load_model";
            string imagePath =
                "C:\\Users\\User\\Desktop\\TJ-003(B)_241211_crop_4\\TJ-003(B)_241211_crop_4\\test\\images\\24112543-7_4.jpg"; //"john-mayer.png";
            string batFilePath = "start_server.bat"; // BAT 檔案路徑 (假設 BAT 檔案與 C# 執行檔在同一個目錄下)

            // 檢查伺服器是否已啟動
            bool serverAvailable = await IsServerAvailable(serverBaseUrl);
            if (!serverAvailable)
            {
                Console.WriteLine("伺服器未連線，嘗試啟動伺服器...");
                bool serverStarted = StartPythonServer(batFilePath); // 啟動 Python 伺服器

                if (serverStarted)
                {
                    Console.WriteLine("伺服器啟動中，等待 3 秒...");
                    await Task.Delay(3000); // 等待 5 秒讓伺服器啟動

                    serverAvailable = await IsServerAvailable(serverBaseUrl); // 再次檢查伺服器是否啟動
                    if (!serverAvailable)
                    {
                        Console.WriteLine("伺服器啟動失敗，請檢查 start_server.bat 檔案和 Python 環境設定。");
                        Console.WriteLine("\n結束...");
                        return; // 伺服器啟動失敗，直接結束程式
                    }
                    else
                    {
                        Console.WriteLine("伺服器已啟動成功。");
                    }
                }
                else
                {
                    Console.WriteLine("無法啟動伺服器，請檢查 start_server.bat 檔案是否存在或是否有執行權限。");
                    Console.WriteLine("\n結束...");
                    //return; // 無法執行 BAT 檔案，直接結束程式
                }
            }
            else
            {
                Console.WriteLine("伺服器已連線。");
            }

            // Console.WriteLine("是否要更換 YOLOv8 模型？ (y/n)");
            // string changeModelInput = Console.ReadLine().ToLower();
            //
            // if (changeModelInput == "y")
            // {
            //     Console.WriteLine("請輸入新的模型路徑 (例如 yolov8m.pt 或自訂模型路徑):");
            //     string newModelPath = Console.ReadLine();
            //     LoadModelResponse
            //         loadModelResponse = await LoadYoloModel(newModelPath, serverBaseUrl); // 呼叫 LoadYoloModel 函數
            //
            //     if (loadModelResponse.error != null)
            //     {
            //         Console.WriteLine($"載入模型失敗: {loadModelResponse.error}");
            //         Console.WriteLine("將使用預設模型進行物件偵測。");
            //     }
            //     else
            //     {
            //         Console.WriteLine($"模型載入成功: {loadModelResponse.message}");
            //     }
            // }

            Mat image = Cv2.ImRead(imagePath);
            if (image.Empty())
            {
                Console.WriteLine($"無法讀取圖片: {imagePath}");
                return;
            }

            DetectionResponse
                detectionResponse = await PerformObjectDetection(image, $"{serverBaseUrl}/detect"); // 注意 URL 要完整
            if (detectionResponse.error != null)
            {
                Console.WriteLine($"伺服器偵測錯誤: {detectionResponse.error}");
            }
            else if (detectionResponse.detections != null && detectionResponse.detections.Count > 0)
            {
                Console.WriteLine("物件偵測結果:");
                Scalar[] colors =
                {
                    Scalar.Red, Scalar.Blue, Scalar.Green, Scalar.Yellow, Scalar.Orange, Scalar.Pink, Scalar.Violet,
                    Scalar.LightGray, Scalar.White, Scalar.Black
                };

                foreach (var detection in detectionResponse.detections)
                {
                    Console.WriteLine(
                        $"- 類別: {detection.class_name} (ID: {detection.class_id}), 分數: {detection.score:F4}, Box: [{string.Join(", ", detection.box)}]");
                    if (detection.score > 0.5)
                    {
                        Cv2.Rectangle(image, new Point(detection.box[0], detection.box[1]),
                            new Point(detection.box[2], detection.box[3]), colors[detection.class_id % colors.Length],
                            2);
                        Cv2.PutText(image, $"{detection.class_name} {detection.score:F2}",
                            new Point(detection.box[0], detection.box[1] - 10), HersheyFonts.HersheySimplex, 0.5,
                            colors[detection.class_id % colors.Length], 1, LineTypes.AntiAlias);
                    }
                }

                Cv2.ImShow("物件偵測結果", image);
                Cv2.WaitKey(0);
                Cv2.DestroyAllWindows();
            }
            else
            {
                Console.WriteLine("沒有偵測到物件。");
            }

            Console.WriteLine("\n按任意鍵結束...");
            //Console.ReadKey();
        }

        // 檢查伺服器是否可連線的函數
        public static async Task<bool> IsServerAvailable(string serverBaseUrl)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await client.GetAsync(serverBaseUrl); // 發送 GET 請求到伺服器根路徑
                    return response.IsSuccessStatusCode; // 如果狀態碼是 2xx，則視為伺服器可用
                }
                catch (HttpRequestException)
                {
                    return false; // 如果發生連線錯誤 (例如伺服器未運行)，則視為不可用
                }
            }
        }


        // 執行 BAT 檔案啟動 Python 伺服器的函數
        public static bool StartPythonServer(string batFilePath)
        {
            try
            {
                if (!File.Exists(batFilePath))
                {
                    Console.WriteLine($"BAT 檔案不存在: {batFilePath}");
                    return false;
                }

                ProcessStartInfo startInfo = new ProcessStartInfo(batFilePath);
                startInfo.CreateNoWindow = true; // 不要創建新視窗 (背景執行) 可以改為 false 來顯示視窗以便查看伺服器輸出
                startInfo.UseShellExecute = true; // 需要 UseShellExecute=true 才能執行 bat 檔案，預設是 false
                startInfo.WorkingDirectory =
                    AppDomain.CurrentDomain.BaseDirectory; // 設定工作目錄為 C# 程式執行目錄，確保 BAT 檔案能正確找到 server.py 或相關資源
                                                           // 設定視窗樣式為隱藏
                //startInfo.WindowStyle = ProcessWindowStyle.Hidden; //隱藏視窗
                Process.Start(startInfo);
                return true; // BAT 檔案啟動成功
            }
            catch (Exception ex)
            {
                Console.WriteLine($"啟動伺服器 BAT 檔案失敗: {ex.Message}");
                return false; // BAT 檔案啟動失敗
            }
        }
        public static async Task<bool> ShutdownServerAsync(string serverBaseUrl)
        {
            return await SendAdminRequestAsync(serverBaseUrl + "/shutdown");
        }

        public static async Task<bool> RestartServerAsync(string serverBaseUrl)
        {
            return await SendAdminRequestAsync(serverBaseUrl + "/restart");
        }

        private static async Task<bool> SendAdminRequestAsync(string serverBaseUrl)
        {
            bool serverClosed = false;
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string url = serverBaseUrl;
                    var response = await client.PostAsync(url, null);  // POST 請求沒有 body，所以傳 null

                    if (response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        dynamic jsonResponse = JsonConvert.DeserializeObject(responseBody);

                        // 檢查伺服器是否回傳成功訊息
                        if (jsonResponse != null && jsonResponse.message != null)
                        {
                            Console.WriteLine(jsonResponse.message);
                            return true;
                        }
                        else
                        {
                            Console.WriteLine("伺服器成功執行指令，但回應訊息有問題");
                            return true;
                        }


                    }
                    else
                    {
                        Console.WriteLine($"請求失敗：{response.StatusCode}");
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"發生錯誤：{ex.Message}");
                    return false;  // 處理其他錯誤
                }
                finally
                {
                    Console.WriteLine("伺服器可能已經關閉。");
                    serverClosed = true;
                }
            }
        }
        // 新增呼叫 /load_model API 的函數
        public static async Task<LoadModelResponse> LoadYoloModel(string modelPath, string serverBaseUrl)
        {
            LoadModelResponse loadModelResponse = new LoadModelResponse() { message = null, error = null };

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string loadModelUrl = $"{serverBaseUrl}/load_model"; // 完整的 /load_model API URL
                    var requestData = new { model_path = modelPath }; // JSON 請求資料
                    string jsonData = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await client.PostAsync(loadModelUrl, content);
                    response.EnsureSuccessStatusCode(); // 確保請求成功

                    string responseBody = await response.Content.ReadAsStringAsync();
                    loadModelResponse =
                        JsonConvert.DeserializeObject<LoadModelResponse>(responseBody); // 嘗試反序列化為 LoadModelResponse
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"載入模型 HTTP 請求錯誤: {ex.Message}");
                    loadModelResponse.error = $"載入模型 HTTP 請求錯誤: {ex.Message}";
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"載入模型 JSON 解析錯誤: {ex.Message}");
                    loadModelResponse.error = $"載入模型 JSON 解析錯誤: {ex.Message}";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"載入模型發生錯誤: {ex.Message}");
                    loadModelResponse.error = $"載入模型發生錯誤: {ex.Message}";
                }
                finally
                {
                    Console.WriteLine($"載入模型 {modelPath} 結束");
                }
            }

            return loadModelResponse;
        }
        //Split YOLO Detection by YOLO-Server
        public static async Task<DetectionResponse> PerformSplitObjectDetection(Mat image, string serverUrl,
            Size imgSize, Size subSize, int step, float confThreshold, float nmsThreshold)
        {
            DetectionResponse detectionResponse = new DetectionResponse { detections = null, error = null };

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // 1. 將 Mat 圖片轉換為 JPG byte 陣列
                    byte[] imageBytes = image.ToBytes(".png");

                    // 2. 將 byte 陣列轉換為 Base64 字串
                    string base64Image = Convert.ToBase64String(imageBytes);

                    // 3. 建構 JSON 請求內容，包含切割參數
                    var requestData = new
                    {
                        image = base64Image,
                        img_size = new int[] { imgSize.Width, imgSize.Height }, // 轉換為 int[]
                        sub_size = new int[] { subSize.Width, subSize.Height }, // 轉換為 int[]
                        step = step,
                        conf_threshold = confThreshold,
                        nms_threshold = nmsThreshold
                    };

                    string jsonData = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    // 4. 發送 POST 請求到 /split_detect API
                    HttpResponseMessage response = await client.PostAsync(serverUrl, content);
                    response.EnsureSuccessStatusCode();

                    // 5. 讀取並解析 JSON 回應
                    string responseBody = await response.Content.ReadAsStringAsync();
                    detectionResponse = JsonConvert.DeserializeObject<DetectionResponse>(responseBody);
                    Console.WriteLine($"收到檢測結果:{detectionResponse}");
                    Console.WriteLine($"檢測物件數量:{detectionResponse.detections.Count}");
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"HTTP 請求錯誤: {ex.Message}");
                    detectionResponse.error = $"HTTP 請求錯誤: {ex.Message}";
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"JSON 解析錯誤: {ex.Message}");
                    detectionResponse.error = $"JSON 解析錯誤: {ex.Message}";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"發生錯誤: {ex.Message}");
                    detectionResponse.error = $"發生錯誤: {ex.Message}";
                }
            }

            return detectionResponse;
        }

        // 將推理過程包裝成 OpenCV 風格的函數 (接收 Mat 物件)
        public static async Task<DetectionResponse> PerformObjectDetection(Mat image, string serverUrl)
        {
            DetectionResponse detectionResponse = new DetectionResponse()
            { detections = null, error = null }; // 初始化返回物件

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // 1. 使用 OpenCVSharp 將 Mat 圖片轉換為 JPG byte 陣列
                    byte[] imageBytes = image.ToBytes(".png"); // 可以選擇其他格式，例如 ".png"

                    // 2. 將 byte 陣列轉換為 Base64 字串
                    string base64Image = Convert.ToBase64String(imageBytes);

                    // 3. 建構 JSON 請求內容
                    var requestData = new { image = base64Image };
                    string jsonData = JsonConvert.SerializeObject(requestData);

                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    // 4. 發送 POST 請求到 YOLOv8 伺服器
                    HttpResponseMessage response = await client.PostAsync(serverUrl, content);
                    response.EnsureSuccessStatusCode();

                    // 5. 讀取並解析 JSON 回應
                    string responseBody = await response.Content.ReadAsStringAsync();


                    detectionResponse = JsonConvert.DeserializeObject<DetectionResponse>(responseBody);
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"HTTP 請求錯誤: {ex.Message}");
                    detectionResponse.error = $"HTTP 請求錯誤: {ex.Message}"; // 記錄錯誤訊息
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"JSON 解析錯誤: {ex.Message}");
                    detectionResponse.error = $"JSON 解析錯誤: {ex.Message}"; // 記錄錯誤訊息
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"發生錯誤: {ex.Message}");
                    detectionResponse.error = $"發生錯誤: {ex.Message}"; // 記錄錯誤訊息
                }
            }

            return detectionResponse; // 返回 DetectionResponse 物件
        }
    }
}