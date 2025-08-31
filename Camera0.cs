using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Basler.Pylon;             // ★ Basler Pylon SDK
using OpenCvSharp;             // ★ OpenCvSharp
using OpenCvSharp.Extensions;  // ★ OpenCvSharp 將 Bitmap 等轉成 Mat 用
using PLC;                     // 你本來的 PLC 命名空間 (若有需要)
using peilin;
using System.Collections.Generic;

namespace basler
{
    public class Camera
    {
        // Basler Camera 陣列 (模擬多相機)
        private static Basler.Pylon.Camera[] m_BaslerCameras = new Basler.Pylon.Camera[app.MaxCameraCount];

        // 是否已啟動抓圖 (對應原本 GrabOver 的概念)
        private static bool[] isGrabOver = new bool[app.MaxCameraCount];

        // 紀錄實際偵測到的相機數
        private static int CameraCount = 0;

        // 為了與舊程式相容，保留一個「哪一家廠商」的旗標 (不一定要用)
        private static byte SelectCameraCompany = 1; // 0: StCamera, 1: Basler (預設改成 Basler)

        // 與 form1 溝通，用來呼叫 form1.Receiver(...)
        private static Form1 form1;
        #region 監控

        // 在 Camera 類內部的頂部添加
        private static readonly object _logLock = new object();
        private static string _logFilePath = $"CameraLog_{DateTime.Now:yyyyMMdd}.txt";
        private static Dictionary<int, EventHandler<ImageGrabbedEventArgs>> _imageGrabbedHandlers =
            new Dictionary<int, EventHandler<ImageGrabbedEventArgs>>();
        private static Dictionary<int, EventHandler> _connectionLostHandlers =
            new Dictionary<int, EventHandler>();
        private static Dictionary<int, DateTime> _lastGrabTimes = new Dictionary<int, DateTime>();
        private static Dictionary<int, int> _grabCounters = new Dictionary<int, int>();
        private static System.Windows.Forms.Timer _monitorTimer;
        private static bool _monitoringActive = false;

        private static readonly Dictionary<int, bool> _processingImage = new Dictionary<int, bool>();
        private static readonly Dictionary<int, double> _minAllowedInterval = new Dictionary<int, double>();
        private static readonly Dictionary<int, Stopwatch> _cameraStopwatches = new Dictionary<int, Stopwatch>();
        private static readonly Dictionary<int, int> _touchedCounters = new Dictionary<int, int>();
        DateTime st = DateTime.Now;

        // 日誌記錄方法
        private static void LogMessage(string message, bool isWarning = false)
        {
            try
            {
                lock (_logLock)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    string logEntry = $"[{timestamp}] {(isWarning ? "[警告]" : "[資訊]")} {message}";

                    // 同時輸出到控制台和文件
                    Console.WriteLine(logEntry);
                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);

                    // 如果是警告，同時保存到專用警告日誌
                    if (isWarning)
                    {
                        File.AppendAllText($"CameraWarning_{DateTime.Now:yyyyMMdd}.txt",
                            logEntry + Environment.NewLine);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"日誌記錄失敗: {ex.Message}");
            }
        }

        // 初始化監控計時器
        private void InitializeMonitoring()
        {
            if (!_monitoringActive)
            {
                _monitorTimer = new System.Windows.Forms.Timer();
                _monitorTimer.Interval = 60000; // 每分鐘檢查一次
                _monitorTimer.Tick += (sender, e) => MonitorCameraBehavior();
                _monitorTimer.Start();
                _monitoringActive = true;
                LogMessage("相機監控系統已啟動");
            }
        }

        // 監控相機行為
        private void MonitorCameraBehavior()
        {
            try
            {
                // 檢查相機狀態和事件計數
                for (int i = 0; i < app.MaxCameraCount; i++)
                {
                    if (m_BaslerCameras[i] != null && m_BaslerCameras[i].IsOpen)
                    {
                        // 特別關注第二站相機
                        if (i == 1)
                        {
                            // 檢查觸發模式
                            bool isTriggerOn = m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].GetValue() == PLCamera.TriggerMode.On;
                            string triggerSource = m_BaslerCameras[i].Parameters[PLCamera.TriggerSource].GetValue();

                            // 如果在硬體觸發模式下但長時間沒有影像，可能表示有問題
                            if (isTriggerOn && triggerSource == PLCamera.TriggerSource.Line3)
                            {
                                if (_lastGrabTimes.ContainsKey(i))
                                {
                                    TimeSpan timeSinceLastGrab = DateTime.Now - _lastGrabTimes[i];
                                    if (timeSinceLastGrab.TotalMinutes > 5)
                                    {
                                        LogMessage($"警告：第二站相機已在硬體觸發模式下 {timeSinceLastGrab.TotalMinutes:F1} 分鐘未接收影像", true);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"監控檢查時發生錯誤: {ex.Message}", true);
            }
        }


        #endregion
        /// <summary>
        /// 開啟所有 Basler 相機，並完成初始化
        /// </summary>
        /// <param name="form">主要用來呼叫 form1.Receiver 的參考</param>
        /// <returns>若未偵測到任一相機，回傳 false</returns>
        public bool Open(Form1 form)
        {
            form1 = form;

            // 取得系統中可用的 Basler 相機清單
            var allDevices = CameraFinder.Enumerate();
            CameraCount = allDevices.Count;

            if (CameraCount == 0)
            {
                // 完全沒有偵測到相機
                return false;
            }

            // 最高只開啟 app.MaxCameraCount 台
            if (CameraCount > app.MaxCameraCount)
                CameraCount = app.MaxCameraCount;

            app.SoftTriggerMode = false;
            //Setting(1);

            // 開始連結相機
            AddrLink();
            Setting(1);

            return true;
        }

        /// <summary>
        /// 與舊程式介面相容：叫出「相機設定對話框」
        /// Basler SDK 不提供像 StCamera 那樣的通用設定對話框，可以自行實作或忽略。
        /// 這裡給一個空殼。
        /// </summary>
        public void ShowSettingDlg(int CameraID)
        {
            // Basler 並沒有對應的「內建設定視窗」如 StCamera.ShowSettingDlg
            // 可自行撰寫或調用外部工具。
            Console.WriteLine("Basler does not provide a built-in Setting Dialog. (Stub)");
        }

        /// <summary>
        /// 檢查指定 CameraID 是否可用
        /// </summary>
        public bool CheckCamera(int CameraID)
        {
            if (CameraID < 0 || CameraID >= app.MaxCameraCount) return false;
            return (m_BaslerCameras[CameraID] != null);
        }

        /// <summary>
        /// 模擬原本的 Setting(Mode) 來切換取像模式：
        /// Mode 0: 連續 (FreeRun) 
        /// Mode 1: 硬體觸發 
        /// Mode 2: 軟體觸發 
        /// </summary>
        /// <param name="Mode"></param>
        public void Setting(byte Mode)
        {
            #region 舊程式不含監控
            /*
            for (int i = 0; i < app.MaxCameraCount; i++)
            {
                // 檢查相機是否存在
                if (m_BaslerCameras[i] == null) continue;

                // 先停止抓圖
                StopSingle(i);

                // 依 Mode 來設定觸發模式
                // Basler 的參數名請查 Pylon C# API
                // 下方各項參數可能需視相機型號調整
                if (Mode == 0)
                {
                    // FreeRun (連續)
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSelector].SetValue(PLCamera.TriggerSelector.FrameStart);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.Off);
                    m_BaslerCameras[i].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
                    app.SoftTriggerMode = true;
                }
                else if (Mode == 1)
                {
                    // 硬體觸發
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSelector].SetValue(PLCamera.TriggerSelector.FrameStart);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSource].SetValue(PLCamera.TriggerSource.Line3);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.On);
                    m_BaslerCameras[i].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
                    app.SoftTriggerMode = false;

                }
                else if (Mode == 2)
                {
                    // 軟體觸發
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSelector].SetValue(PLCamera.TriggerSelector.FrameStart);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSource].SetValue(PLCamera.TriggerSource.Software);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.On);
                    m_BaslerCameras[i].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
                    app.SoftTriggerMode = true;
                }
            }
            */
            #endregion
            // 記錄設定調用
            LogMessage($"相機觸發模式變更請求: Mode={Mode} ({(Mode == 0 ? "連續" : Mode == 1 ? "硬體觸發" : "軟體觸發")})");

            // 記錄呼叫堆疊信息
            try
            {
                var stackTrace = new System.Diagnostics.StackTrace(true);
                string caller = "Unknown";

                if (stackTrace.FrameCount > 1)
                {
                    var frame = stackTrace.GetFrame(1);
                    caller = $"{frame.GetMethod().DeclaringType}.{frame.GetMethod().Name}";

                    // 記錄更多調用者信息
                    var callerInfo = $"調用來源: {caller}, 行號: {frame.GetFileLineNumber()}";
                    LogMessage(callerInfo);
                }
            }
            catch { }

            for (int i = 0; i < app.MaxCameraCount; i++)
            {
                // 檢查相機是否存在
                if (m_BaslerCameras[i] == null) continue;

                // 先停止抓圖
                StopSingle(i);

                // 記錄每個相機的模式變更前狀態
                bool wasGrabbing = false;
                string previousMode = "Unknown";

                try
                {
                    wasGrabbing = m_BaslerCameras[i].StreamGrabber.IsGrabbing;
                    previousMode = m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].GetValue() == PLCamera.TriggerMode.On ?
                                   "觸發模式" : "連續模式";
                }
                catch { }

                // 特別關注第二站相機
                if (i == 1)
                {
                    LogMessage($"第二站相機模式變更: 從 {previousMode} 到 Mode={Mode}, 之前是否在抓取: {wasGrabbing}");
                }

                // 依 Mode 來設定觸發模式
                // Basler 的參數名請查 Pylon C# API
                // 下方各項參數可能需視相機型號調整
                if (Mode == 0)
                {
                    // FreeRun (連續)
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSelector].SetValue(PLCamera.TriggerSelector.FrameStart);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.Off);
                    m_BaslerCameras[i].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
                    app.SoftTriggerMode = true;
                }
                else if (Mode == 1)
                {
                    // 硬體觸發
                    // 1. 首先設定觸發器相關參數
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSelector].SetValue(PLCamera.TriggerSelector.FrameStart);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSource].SetValue(PLCamera.TriggerSource.Line3);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.On);

                    // 2. 再設定線路相關參數（包括去抖動時間）
                    m_BaslerCameras[i].Parameters[PLCamera.LineSelector].SetValue(PLCamera.LineSelector.Line3);
                    m_BaslerCameras[i].Parameters[PLCamera.LineDebouncerTime].SetValue(10.0); 

                    // 3. 最後設定擷取模式
                    m_BaslerCameras[i].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);

                    app.SoftTriggerMode = false;

                }
                else if (Mode == 2)
                {
                    // 軟體觸發
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSelector].SetValue(PLCamera.TriggerSelector.FrameStart);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerSource].SetValue(PLCamera.TriggerSource.Software);
                    m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.On);
                    m_BaslerCameras[i].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
                    app.SoftTriggerMode = true;
                }

                // 在第二站相機上確認設置是否成功
                if (i == 1)
                {
                    try
                    {
                        bool isTriggerOn = m_BaslerCameras[i].Parameters[PLCamera.TriggerMode].GetValue() == PLCamera.TriggerMode.On;
                        string triggerSource = m_BaslerCameras[i].Parameters[PLCamera.TriggerSource].GetValue();
                        string newMode = isTriggerOn ? $"觸發模式({triggerSource})" : "連續模式";

                        LogMessage($"第二站相機模式設置結果: {newMode}");

                        // 檢查模式是否與預期一致
                        bool isExpectedMode = (Mode == 0 && !isTriggerOn) ||
                                             (Mode == 1 && isTriggerOn && triggerSource == PLCamera.TriggerSource.Line3) ||
                                             (Mode == 2 && isTriggerOn && triggerSource == PLCamera.TriggerSource.Software);

                        if (!isExpectedMode)
                        {
                            LogMessage($"警告：第二站相機模式設置可能有問題，預期Mode={Mode}，實際為 {newMode}", true);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"檢查第二站相機模式時出錯: {ex.Message}", true);
                    }
                }
            }

            LogMessage($"相機觸發模式變更完成: Mode={Mode}");
        }

        /// <summary>
        /// 範例白平衡設定，Basler 部分型號支援 BalanceRatio。
        /// Mode=0 自動, Mode=1 手動
        /// </summary>
        public void SetWhiteBalance(int Mode, int i, ushort GainR, ushort GainGr, ushort GainGb, ushort GainB)
        {
            if (m_BaslerCameras[i] == null) return;

            if (Mode == 0)
            {
                // 自動
                m_BaslerCameras[i].Parameters[PLCamera.BalanceWhiteAuto].TrySetValue(PLCamera.BalanceWhiteAuto.Continuous);
            }
            else
            {
                // 手動
                m_BaslerCameras[i].Parameters[PLCamera.BalanceWhiteAuto].TrySetValue(PLCamera.BalanceWhiteAuto.Off);

                // Basler 常用三通道 (Red, Green, Blue)，也有些機種會拆成 Gr/Gb。
                // 下方程式碼僅做示範，實際值需依產品手冊測試、換算。
                // Red
                m_BaslerCameras[i].Parameters[PLCamera.BalanceRatioSelector].SetValue(PLCamera.BalanceRatioSelector.Red);
                m_BaslerCameras[i].Parameters[PLCamera.BalanceRatio].SetValue(GainR / 128.0);

                // Green (一部分機種有 Gr/Gb 可分別設定)
                m_BaslerCameras[i].Parameters[PLCamera.BalanceRatioSelector].SetValue(PLCamera.BalanceRatioSelector.Green);
                m_BaslerCameras[i].Parameters[PLCamera.BalanceRatio].SetValue(GainGr / 128.0);

                // Blue
                m_BaslerCameras[i].Parameters[PLCamera.BalanceRatioSelector].SetValue(PLCamera.BalanceRatioSelector.Blue);
                m_BaslerCameras[i].Parameters[PLCamera.BalanceRatio].SetValue(GainB / 128.0);
            }
        }

        public void GetWhite()
        {
            // 範例：讀取當前白平衡是否自動
            // 依需求顯示或修改
            if (m_BaslerCameras[0] != null)
            {
                var val = m_BaslerCameras[0].Parameters[PLCamera.BalanceWhiteAuto].GetValue();
                //Console.WriteLine(val);
            }
        }

        /// <summary>
        /// 與舊程式介面一致，設定曝光時脈(只是一個命名示意)。Basler 的曝光值單位通常為微秒
        /// </summary>
        public void ClockChange(int ID, int ExposureClock)
        {
            if (m_BaslerCameras[ID] == null) return;
            // 依相機需求調整，Basler 通常是用 microseconds
            m_BaslerCameras[ID].Parameters[PLCamera.ExposureTime].SetValue((double)ExposureClock);
        }

        /// <summary>
        /// Shutter Mode 設定 (Basler 不一定有相同參數)，先保留空殼或自行實作
        /// </summary>
        public void SensorShutterModeSet(int i, uint mode)
        {
            if (m_BaslerCameras[i] == null) return;
            // 依產品手冊與 Pylon SDK 的參數可自行擴充
            Console.WriteLine("Basler shutter mode is not implemented. (Stub)");
        }

        /// <summary>
        /// 鏡像翻轉設定 (Basler 可使用 ReverseX, ReverseY 兩個參數)
        /// Mode = 0: 不翻轉; 1: X翻轉; 2: Y翻轉; 3: XY都翻轉等... 視需求自行定義
        /// </summary>
        public void MirrorSet(int i, byte Mode)
        {
            if (m_BaslerCameras[i] == null) return;

            switch (Mode)
            {
                case 0: // 不翻轉
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseX].SetValue(false);
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseY].SetValue(false);
                    break;
                case 1: // X翻轉
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseX].SetValue(true);
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseY].SetValue(false);
                    break;
                case 2: // Y翻轉
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseX].SetValue(false);
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseY].SetValue(true);
                    break;
                case 3: // XY翻轉
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseX].SetValue(true);
                    m_BaslerCameras[i].Parameters[PLCamera.ReverseY].SetValue(true);
                    break;
            }
        }

        /// <summary>
        /// 設定類比增益 (Baseline)
        /// </summary>
        public void GainChange(int ID, int Gain)
        {
            if (m_BaslerCameras[ID] == null) return;
            m_BaslerCameras[ID].Parameters[PLCamera.Gain].SetValue((double)Gain);
        }

        /// <summary>
        /// 開始所有相機的擷取（連續模式或硬體模式皆可在 Setting 時先行設定）
        /// </summary>
        public void Start()
        {
            for (int i = 0; i < app.MaxCameraCount; i++)
            {
                //StartSingle(i);
                try
                {
                    // 如果相機已在運行，先停止它
                    if (m_BaslerCameras[i] != null &&
                        m_BaslerCameras[i].StreamGrabber.IsGrabbing)
                    {
                        StopSingle(i);
                        // 短暫延遲以確保資源已釋放
                        System.Threading.Thread.Sleep(50);
                    }

                    StartSingle(i);
                }
                catch (Exception ex)
                {
                    // 記錄或處理錯誤，避免整個操作因單個相機失敗而中斷
                    Console.WriteLine($"啟動相機 {i} 失敗: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 設定曝光時間
        /// </summary>
        public void SetExposureClock(int CameraID, uint Value)
        {
            if (m_BaslerCameras[CameraID] != null)
            {
                // Basler 是以微秒為單位
                m_BaslerCameras[CameraID].Parameters[PLCamera.ExposureTime].SetValue(Value);
            }
        }

        /// <summary>
        /// 設定 Gain
        /// </summary>
        public void SetGain(int CameraID, ushort Value)
        {
            if (m_BaslerCameras[CameraID] != null)
            {
                m_BaslerCameras[CameraID].Parameters[PLCamera.Gain].SetValue((double)Value);
            }
        }

        /// <summary>
        /// 模擬數位增益。Basler 未必有「DigitalGain」參數，依型號而定。
        /// 這裡僅示範和 SetGain 相同處理。
        /// </summary>
        public void SetDigitalGain(int CameraID, ushort Value)
        {
            if (m_BaslerCameras[CameraID] != null)
            {
                // 如果型號支援 digital gain，可換成實際參數名稱
                // 暫時一樣呼叫 Gain
                m_BaslerCameras[CameraID].Parameters[PLCamera.Gain].SetValue((double)Value);
            }
        }

        /// <summary>
        /// 停止所有相機擷取
        /// </summary>
        public void Stop()
        {
            for (int i = 0; i < app.MaxCameraCount; i++)
            {
                StopSingle(i);
            }
        }

        /// <summary>
        /// 關閉並釋放所有相機
        /// </summary>
        public void Close()
        {
            for (int i = 0; i < app.MaxCameraCount; i++)
            {
                if (m_BaslerCameras[i] != null)
                {
                    try
                    {
                        StopSingle(i);
                        m_BaslerCameras[i].Close();
                        m_BaslerCameras[i].Dispose();
                    }
                    catch { }
                    m_BaslerCameras[i] = null;
                }
            }
        }

        /// <summary>
        /// 檢查目前實際偵測到的相機數量是否等於 app.MaxCameraCount
        /// </summary>
        public bool checkCameraCount()
        {
            return (CameraCount == app.MaxCameraCount);
        }

        /// <summary>
        /// 連接相機並完成事件綁定
        /// </summary>
        /// <returns>若有無法連接之相機，回傳名稱，或空字串</returns>
        public bool IsOpened(int cameraIndex)
        {
            if (cameraIndex < 0 || cameraIndex >= app.MaxCameraCount) return false;
            return (m_BaslerCameras[cameraIndex] != null &&
                    m_BaslerCameras[cameraIndex].IsOpen);
        }

        public string AddrLink()
        {
            string Name = "";
            var allDevices = CameraFinder.Enumerate();

            // 创建序列号与索引的映射
            var cameraIndexMap = new Dictionary<string, int>
            {
                { "25091590", 0 },
                { "25061245", 1 },
                { "25094092", 2 },
                { "25091594", 3 }
            };

            foreach (var cameraInfo in allDevices)
            {
                try
                {
                    string serialNumber = cameraInfo[CameraInfoKey.SerialNumber];

                    if (cameraIndexMap.TryGetValue(serialNumber, out int index))
                    {
                        // 创建相机实例
                        if (m_BaslerCameras[index] == null)
                        {
                            m_BaslerCameras[index] = new Basler.Pylon.Camera(cameraInfo);

                            // 注册事件
                            //m_BaslerCameras[index].CameraOpened += Configuration.AcquireContinuous;

                            m_BaslerCameras[index].ConnectionLost += (sender, e) =>
                            {
                                StopSingle(index);
                            };

                            // 捕获正确的索引
                            int camIndex = index;
                            m_BaslerCameras[index].StreamGrabber.ImageGrabbed += (sender, e) =>
                            {
                                OnImageGrabbed(e, camIndex);
                            };

                            // 打开相机
                            m_BaslerCameras[index].Open();

                            // 标记相机已连接
                            app.CameraLinked[index] = true;
                        }
                        else if (!m_BaslerCameras[index].IsOpen)
                        {
                            // 相機已存在但未開啟的情況

                            // 對於StreamGrabber.ImageGrabbed事件，最安全的方法是停止並重新開始StreamGrabber
                            // 這會間接清除所有註冊的事件處理器

                            try
                            {
                                if (m_BaslerCameras[index].StreamGrabber.IsGrabbing)
                                {
                                    m_BaslerCameras[index].StreamGrabber.Stop();
                                }
                            }
                            catch { }

                            // 重新註冊事件處理器
                            m_BaslerCameras[index].ConnectionLost += (sender, e) =>
                            {
                                StopSingle(index);
                            };

                            // 捕获正确的索引
                            int camIndex = index;
                            m_BaslerCameras[index].StreamGrabber.ImageGrabbed += (sender, e) =>
                            {
                                OnImageGrabbed(e, camIndex);
                            };
                            // 相機已創建但未開啟，則開啟
                            m_BaslerCameras[index].Open();
                            app.CameraLinked[index] = true;
                        }
                    }
                    else
                    {
                        // 未找到匹配的序列号
                        Name += $"Unknown Camera [{serialNumber}] ";
                    }
                }
                catch (Exception ex)
                {
                    Name += $"Camera[{cameraInfo[CameraInfoKey.SerialNumber]}]: {ex.Message} ";
                }
            }

            return Name;
        }


        /// <summary>
        /// 單一相機開始擷取
        /// </summary>
        private void StartSingle(int i)
        {
            if (m_BaslerCameras[i] == null) return;
            try
            {
                // 連續模式、硬體模式或軟體模式已在 Setting(...) 設好
                // 這裡只要開始 StreamGrabber
                m_BaslerCameras[i].StreamGrabber.Start(GrabStrategy.OneByOne, GrabLoop.ProvidedByStreamGrabber);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"StartSingle camera {i} error: {ex.Message}");
            }
        }

        /// <summary>
        /// 單一相機停止擷取
        /// </summary>
        private void StopSingle(int i)
        {
            if (m_BaslerCameras[i] == null) return;
            try
            {
                m_BaslerCameras[i].StreamGrabber.Stop();
            }
            catch { }
        }

        /// <summary>
        /// 核心：收到 Basler 影像事件後，轉成 OpenCvSharp.Mat，呼叫 form1.Receiver(...)
        /// </summary>
        private void OnImageGrabbed(ImageGrabbedEventArgs e, int cameraIndex)
        {

            IGrabResult grabResult = e.GrabResult;
            if (!grabResult.IsValid) return;
            if (!app.Run) return;

            DateTime time_start = DateTime.Now; // 拍照開始時間（計時）
            Mat currentImage = null;
            try
            {
                currentImage = GrabResultToMat(grabResult);

                if (cameraIndex == 1 || cameraIndex == 0 || cameraIndex == 3)
                {
                    bool isvalid = form1.CheckWhitePixelRatio(currentImage, cameraIndex + 1);
                    if (!isvalid)
                    {
                        LogMessage($"相機 {cameraIndex} 偵測到可能的重複觸發，執行進階檢查", true);
                        Mat imageForSave = currentImage.Clone();
                        SaveTouchedImage(imageForSave, cameraIndex, 0, "position_check");
                        return;
                    }           

                    #region 時間過濾

                    // 檢查是否為連續觸發（需要先檢查，再更新時間）
                    if (_lastGrabTimes.ContainsKey(cameraIndex))
                    {
                        DateTime lastTime = _lastGrabTimes[cameraIndex];
                        TimeSpan timeDiff = time_start - lastTime;

                        // 如果時間間隔過短（少於200ms），可能是重複觸發
                        if (timeDiff.TotalMilliseconds < 200) //有效取像間隔至少為
                        {
                            LogMessage($"相機 {cameraIndex} 偵測到可能的重複觸發，間隔僅 {timeDiff.TotalMilliseconds:F3}ms，執行進階檢查", true);
                            Mat imageForSave = currentImage.Clone();
                            SaveTouchedImage(imageForSave, cameraIndex, timeDiff.TotalMilliseconds, "position_check");
                            currentImage.Dispose();

                            // 判斷為誤觸發，直接跳過
                            return;
                        }
                    }
                    // 該圖像通過了所有檢查，是有效圖像，更新時間記錄
                    // 只記錄有效圖像的時間
                    _lastGrabTimes[cameraIndex] = time_start;
                    currentImage.Dispose();

                    #endregion
                }

            
                // 將 GrabResult 轉成 Mat
                Mat src = GrabResultToMat(grabResult);

                // 呼叫原程式中的 Receiver
                form1.Receiver(cameraIndex, src, time_start);

                // 釋放或歸零
                src = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OnImageGrabbed error: {ex.Message}");
            }
        }

        /// <summary>
        /// 將 Basler 的 IGrabResult 轉成 OpenCvSharp 的 Mat
        /// </summary>
        private Mat GrabResultToMat(IGrabResult grabResult)
        {
            // 參考 Basler SDK
            PixelDataConverter converter = new PixelDataConverter();
            converter.OutputPixelFormat = PixelType.BGR8packed;

            // 依影像大小建一個 Mat
            Mat mat = new Mat((int)grabResult.Height, (int)grabResult.Width, MatType.CV_8UC3);

            // 計算需求緩衝大小
            long bufferSize = converter.GetBufferSizeForConversion(grabResult);
            converter.Convert(mat.Data, bufferSize, grabResult);

            return mat;
        }

        /// <summary>
        /// 儲存誤觸發照片到指定資料夾
        /// </summary>
        /// <param name="image">要儲存的圖像</param>
        /// <param name="cameraIndex">相機索引</param>
        /// <param name="timeDiff">與上次觸發的時間間隔(ms)</param>
        /// <param name="reason">誤觸發原因</param>
        private void SaveTouchedImage(Mat image, int cameraIndex, double timeDiff, string reason)
        {
            try
            {
                // 計算儲存路徑
                string baseDir = $@".\image\{st.ToString("yyyy-MM")}\{st.ToString("MMdd")}\{form1.GetAppFolderName()}\touched";

                // 確保資料夾存在
                if (!Directory.Exists(baseDir))
                {
                    Directory.CreateDirectory(baseDir);
                }

                // 使用自增的計數器來命名誤觸發照片
                int touchedCounter = GetNextTouchedCounter(cameraIndex);

                // 檔案名稱包含時間、相機索引、觸發間隔和原因 
                string fileName = $"{touchedCounter:D4}_{cameraIndex}_interval{timeDiff:F0}ms_{reason}.jpg";
                string fullPath = Path.Combine(baseDir, fileName);

                // 儲存圖像
                form1.SaveImageAsync(image, fullPath);

                LogMessage($"已儲存誤觸發照片: {fullPath}", false);
            }
            catch (Exception ex)
            {
                LogMessage($"儲存誤觸發照片失敗: {ex.Message}", true);
            }
        }
        private int GetNextTouchedCounter(int cameraIndex)
        {
            lock (_logLock) // 使用已有的鎖來保護計數器
            {
                if (!_touchedCounters.ContainsKey(cameraIndex))
                {
                    _touchedCounters[cameraIndex] = 0;
                }

                int counter = _touchedCounters[cameraIndex];
                _touchedCounters[cameraIndex]++;

                return counter;
            }
        }
        public void softon(int CameraID)
        {
            m_BaslerCameras[CameraID].ExecuteSoftwareTrigger();
        }
        public void setback(int CameraID)
        {
            m_BaslerCameras[CameraID].Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
            m_BaslerCameras[CameraID].Parameters[PLCamera.TriggerMode].SetValue(PLCamera.TriggerMode.Off);
        }
        public Mat GetCurrentFrame(int cameraID, bool preserveCurrentMode = false)
        {
            if (cameraID < 0 || cameraID >= app.MaxCameraCount || m_BaslerCameras[cameraID] == null)
                return null;

            // 備份當前模式
            byte currentMode = 1; // 預設為硬體觸發
            if (preserveCurrentMode)
            {
                // 這裡可以讀取當前模式，但沒有簡單的API可取得，所以用參數控制
            }

            try
            {
                // 只有在不保留當前模式時才切換到軟體觸發
                if (!preserveCurrentMode)
                {
                    // 設置為軟體觸發模式
                    Setting(2);
                }

                // 先啟動 StreamGrabber (這是關鍵步驟)
                m_BaslerCameras[cameraID].StreamGrabber.Start(GrabStrategy.OneByOne, GrabLoop.ProvidedByUser);

                // 執行軟體觸發
                softon(cameraID);

                // 等待一小段時間
                System.Threading.Thread.Sleep(50);

                // 獲取影像
                using (IGrabResult grabResult = m_BaslerCameras[cameraID].StreamGrabber.RetrieveResult(500, TimeoutHandling.ThrowException))
                {
                    if (grabResult != null && grabResult.IsValid)
                    {
                        // 獲取完畫面後停止 StreamGrabber
                        m_BaslerCameras[cameraID].StreamGrabber.Stop();

                        // 如果有切換模式，則恢復原來的模式
                        if (!preserveCurrentMode)
                        {
                            Setting(currentMode);
                        }

                        // 轉換並返回影像
                        return GrabResultToMat(grabResult);
                    }
                }

                // 確保停止 StreamGrabber
                m_BaslerCameras[cameraID].StreamGrabber.Stop();
            }
            catch (Exception ex)
            {
                // 確保在發生例外時也會停止 StreamGrabber
                try { m_BaslerCameras[cameraID].StreamGrabber.Stop(); } catch { }

                Console.WriteLine($"GetCurrentFrame error: {ex.Message}");
            }
            finally
            {
                // 恢復原來的模式
                if (!preserveCurrentMode)
                {
                    Setting(currentMode);
                }
            }
            return null;
        }
    }

    /// <summary>
    /// 你原先的 app class，保持不動
    /// </summary>
    public class app
    {
        public const byte MaxCameraCount = 4; //最大攝影機數量
        public static bool Run = true; //app運作狀態
        public static bool SoftTriggerMode = true;
        public static byte Mode = 1; // Mode 0:攝影模式 1:拍照模式
        public static bool[] CameraLinked = new bool[MaxCameraCount];

        #region PLC COM點
        public static string Comport1 = string.Empty;
        public static string Comport2 = string.Empty;
        #endregion

        #region 顏色
        public static int[,] ColorValue = new int[,] {
            {   0,  50,  30, 100,  60, 150 },
            { 150, 180,  30, 100,  60, 150 },
            {  80, 120,   0,  60, 110, 220 },
            {  80, 120,   0,  60, 110, 220 },
            {   0,  50,  30, 100,  60, 150 },
            { 150, 180,  30, 100,  60, 150 },
            {   0,  50,  30, 100,  60, 150 },
            { 150, 180,  30, 100,  60, 150 },
            {  80, 120,   0,  60, 110, 220 },
            {  80, 120,   0,  60, 110, 220 },
            {  80, 120,   0,  60, 110, 220 },
            {  80, 120,   0,  60, 110, 220 },
            {   0,  50,  30, 100,  60, 150 },
            { 150, 180,  30, 100,  60, 150 },
            {   0,  50,  30, 100,  60, 150 },
            { 150, 180,  30, 100,  60, 150 }
        };
        #endregion

        #region 計時
        public static Stopwatch RunningSW = new Stopwatch();
        public static Stopwatch TotalSW = new Stopwatch();
        public static Stopwatch SingleRunningSW = new Stopwatch();
        #endregion

        #region 暫存照片按鈕
        public static int Column = 10;
        public static int Row = 10;
        public static int RowHeight = 37;
        public static int ColumeWidth = 26;
        #endregion

    }
}
