using System;
using System.Runtime;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Point = OpenCvSharp.Point;
using Size = OpenCvSharp.Size;
using PLC;
using Serilog;
using LinqToDB;
using System.Data.SQLite;
using System.Diagnostics;
using System.Reflection;
using CherngerTools.SKIIProx64;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.VisualBasic;

using Microsoft.ML.OnnxRuntime.Tensors;
using Microsoft.ML.OnnxRuntime;

using AnomalyTensorRT;
using basler;

using peilin;
using CherngerUI;

using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;

#region NPOI 
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using static LinqToDB.SqlQuery.SqlPredicate;
#endregion NPOI
// 修改紀錄
// 2023/12/14 
// 1.註解showMat Task 所有picuturebox 都改用自動排隊的BeginInvoke() 

namespace peilin
{
    public partial class Form1 : Form
    {
        DelayButton button2 = new DelayButton();
        DelayButton button6 = new DelayButton();
        #region 資料庫簡介
        /*
         * Cameras 相機參數          
         * DefectChecks 料號對應檢測瑕疵項目                 
         * DefectCounts 料號對應檢測瑕疵項目計數(page2、3用)
         * DefectTypes 瑕疵項目總表
         * params 演算法參數
         * Parameters 存圖勾選
         * Totals 總數量
         * Types 料號資訊 
         * Users 使用者清單         
        */
        #endregion
        public Label[] p2label;
        #region dll
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

        #endregion        
        #region 參數
        Mat testMat = Cv2.ImRead(@".\test.jpg");
        private SmartKeyPro smartKey;
        //SensorTechnology.Camera cam = new SensorTechnology.Camera();
        basler.Camera cam = new basler.Camera();
        DateTime st = DateTime.Now;
        public static List<Task> tasks = new List<Task>();

        private static readonly SemaphoreSlim _disposeSemaphore = new SemaphoreSlim(4); // 限制2個並行任務
        private static readonly ConcurrentQueue<IDisposable> _disposeQueue = new ConcurrentQueue<IDisposable>();

        private YoloDetection _yoloDetection;

        // 用於追蹤樣品檢測時間的佇列
        private Queue<DateTime> sampleDetectionTimes = new Queue<DateTime>();
        // 近一分鐘內檢測的樣品數量
        private int recentMinuteSampleCount = 0;

        private alert alertForm = null;
        // 新增: button17 防連點旗標
        private bool _isButton17Processing = false;
        private readonly object _button17Lock = new object();

        private DateTime? _lastProductiveTime = null;

        // 監控 PLC OK 計數上升的前次值（只記憶，不寫檔）
        private int _prevD803 = -1;
        private int _prevD805 = -1;
        private int _prevD807 = -1;

        // 監控相關的靜態變數
        private DateTime lastPlcLogTime = DateTime.MinValue;
        private System.Threading.Timer _plcMonitorTimer; //D值
        private readonly object _monitorLock = new object();
        private System.Threading.Timer _plcQueueMonitorTimer; //佇列長度
        private readonly object _queueMonitorLock = new object();
        private static bool _systemInitialized = false;
        private static readonly object _initLock = new object();

        #region sqlite
        static string mydb = @".\setting\mydb.sqlite";
        // 由 GitHub Copilot 產生
        // 修正: 增加 SQLite 並發支援,避免 "database is locked" 錯誤
        // BusyTimeout=5000: 等待鎖定最多 5 秒
        // Journal Mode=WAL: 使用 Write-Ahead Logging,支援多讀者/單寫者
        // Pooling=true: 啟用連線池
        static string mydbStr = $"data source={mydb};BusyTimeout=5000;Journal Mode=WAL;Pooling=true;Max Pool Size=100";
        #endregion
        #endregion
        #region Form
        public Form1()
        {
            InitializeComponent();
            InitializeResourceDisposer();
            //PerformanceProfiler.Initialize(@".\logs\ai_performance_log.txt");
            _yoloDetection = new YoloDetection();
            ResultManager.InitializeStats();
        }
        private string keyhash = "14cc05c6baf41b9ed7bc92bc5b1b78ea94e12703048d8c6967942052d42a53de";
        void KeyRemove(object sender, KeyRemoveArgs e)
        {
            if (app.usekey)
            {
                if (e.isRemove)
                {
                    var w = new Form() { Size = new System.Drawing.Size(0, 0) };
                    //Task.Delay(TimeSpan.FromSeconds(3))
                    //    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());

                    MessageBox.Show(w, "未偵測到智慧卡，程式即將關閉", "警告");
                    Log.Error("未偵測到智慧卡，程式即將關閉");
                    BeginInvoke(new Action(Close));
                }
            }
            
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            var keypro = new SmartKeyPro(new byte[] { 0x12, 0x34, 0x56, 0x78 }, keyhash, KeyRemove);

            #region 新增記憶體測試按鈕（開發用）
            if (false) // 只有工程師能看到
            {
                Button btnMemoryTest = new Button
                {
                    Name = "btnMemoryTest",
                    Text = "記憶體測試",
                    Font = new Font("微軟正黑體", 14F, FontStyle.Bold),
                    Location = new System.Drawing.Point(1100, 10),
                    Size = new System.Drawing.Size(150, 50),
                    BackColor = Color.Orange,
                    ForeColor = Color.White
                };
                btnMemoryTest.Visible = false;
                btnMemoryTest.Click += async (s, ev) =>
                {
                    btnMemoryTest.Enabled = false;
                    btnMemoryTest.Text = "測試中...";

                    try
                    {
                        await MemoryLeakTest.RunFullPipelineTest(100, 100); // 快速測試 100 個樣品

                        MessageBox.Show("快速測試完成！", "提示", MessageBoxButtons.OK);
                    }
                    finally
                    {
                        btnMemoryTest.Enabled = true;
                        btnMemoryTest.Text = "記憶體測試";
                    }
                };

                this.tabPage1.Controls.Add(btnMemoryTest);
            }
            #endregion

            #region 預設
            app.lastIn = DateTime.Now;
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;

            // 初始化統計管理器
            ScoreStatisticsManager.Initialize();
            // 初始化系統狀態
            app.currentState = app.SystemState.UpdatedNeedReset;
            UpdateButtonStates();

            // 初始化測量選單狀態
            站1ToolStripMenuItem.Enabled = false;
            站2ToolStripMenuItem.Enabled = false;
            站3ToolStripMenuItem.Enabled = false;
            站4ToolStripMenuItem.Enabled = false;


            // 初始化 PLC 佇列監控
            StartPLCQueueMonitoring();
            // 顯示啟動提示
            lbAdd("程式啟動完成，請先執行異常復歸後再開始檢測", "inf", "");

            #endregion

            #region Log Create
            var template = "{Timestamp:HH:mm:ss} [{Level:u3}] {Message}{NewLine}{Exception}";
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File(@".\logs\peilin_log-.txt", rollingInterval: RollingInterval.Day, outputTemplate: template)
                .CreateLogger();
            lbAdd("開啟檢測程式", "inf", "");
            #endregion
            #region 空間不足警告
            if (100 > GetHardDiskFreeSpace("C"))
            {
                MessageBox.Show("剩餘磁碟空間小於100G!");
                lbAdd("剩餘磁碟空間小於100G。", "err", "");
            }
            #endregion
            #region sqlite
            checkTable();
            #endregion
            #region 參數設置
            if (!app.start_error)
            {
                try
                {
                    TypeSetting();
                    setDictionary();
                }
                catch (Exception e1)
                {
                    lbAdd("參數初始化失敗", "err", e1.ToString());
                }
            }
            #endregion
            #region PLC初始值設定
            if (!app.start_error)
            {
                try
                {
                    //PLC_ModBus.PLC_On();

                    Log.Information("開始 PLC 初始化");
                    if (!app.offline)
                    {
                        Log.Information("呼叫 PLC_ModBus.PLC_On()");
                        PLC_ModBus.PLC_On();
                        Log.Information("PLC_ModBus.PLC_On() 完成");

                        Thread.Sleep(50);

                        Log.Information("呼叫 PLC_ServoOn()");
                        PLC_ServoOn();
                        Log.Information("PLC_ServoOn() 完成");

                        PLC_SetM(0, false);
                        PLC_SetM(1, false);
                        PLC_SetM(2, false);
                        PLC_SetM(5, false);
                        PLC_SetY(20, false); //NG推料
                        PLC_SetY(21, false); //OK1推料
                        PLC_SetY(22, false); //OK2推料
                        PLC_SetM(333, false); //轉盤
                        //PLC_SetM(22, true); //門禁
                        PLC_SetM(3, true);
                        //Thread.Sleep(15000);
                        PLC_SetM(3, false);

                        PLC_SetM(401, true);

                        PLC_SetD(803, 0);
                        label3.Text = "0";

                        PLC_SetD(805, 0);
                        label58.Text = "0";
                        PLC_SetD(100, 0); // 第一站先歸0

                        updateLabel();

                        PLC_SetM(30, true);   // 先開燈                    
                    }
                }
                catch (Exception e1)
                {
                    lbAdd("PLC初始化失敗", "err", e1.ToString());
                    app.start_error = true;
                }
            }
            #endregion
            #region cam 
            if (!app.start_error)
            {
                if (!app.offline)
                //if (!app.testc)
                {

                    if (!cam.Open(this))
                    {
                        MessageBox.Show("未偵測到攝影機", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        lbAdd("未偵測到攝影機", "err", "");
                        app.start_error = true;
                    }
                    
                    else if (!cam.checkCameraCount())
                    {
                        // 只有在Open成功後才檢查相機數量
                        MessageBox.Show("攝影機數量不正確", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        lbAdd("攝影機數量不正確", "err", "");
                        app.start_error = true;

                        /*
                         * string camName = cam.AddrLink();
                        if (!string.IsNullOrEmpty(camName))
                        {
                            MessageBox.Show("攝影機 " + camName + " 名稱錯誤", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            lbAdd("未偵測到攝影機", "err", "");
                            app.start_error = true;
                        }
                        */
                    }

                    /*
                    if (!cam.checkCameraCount())
                    {
                        MessageBox.Show("攝影機數量不正確", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        lbAdd("攝影機數量不正確", "err", "");
                        app.start_error = true;
                    }
                    */
                    cam.Setting(1);
                    cam.Start();
                    
                }
            }
            #endregion
            #region 介面
            // ...existing button2 and button6 initialization...

            // 新增：初始化 button26 button12
            button26.Text = "ON";
            button26.BackColor = Color.FromArgb(128, 255, 128); // 綠色
            if (!app.offline)
            {
                PLC_SetM(6, true);
            }
            button12.Text = "OFF";
            button12.BackColor = Color.FromArgb(255, 128, 128);
            if (!app.offline)
            {
                PLC_SetM(9, true);
            }
            button49.Text = "正常檢測";
            button49.BackColor = Color.FromArgb(128, 255, 128); // 綠色

            if (!app.start_error)
            {
                #region 介面                           
                button2.BackColor = System.Drawing.SystemColors.Control;
                button2.Enabled = false;
                button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
                button2.Font = new System.Drawing.Font("微軟正黑體", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                button2.ForeColor = System.Drawing.Color.Red;
                button2.Location = new System.Drawing.Point(804, 0);
                button2.Name = "button2";
                button2.Size = new System.Drawing.Size(180, 60);
                button2.TabIndex = 321;
                button2.Text = "停止檢測";
                button2.UseVisualStyleBackColor = true;
                button2.after = true;
                button2.Interval = 0;
                button2.Click += new System.EventHandler(this.button2_Click);
                this.Controls.Add(button2);
                button2.BringToFront();

                button6.Font = new System.Drawing.Font("微軟正黑體", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                button6.Location = new System.Drawing.Point(935, 105);
                button6.Name = "button6";
                button6.Size = new System.Drawing.Size(135, 55);
                button6.TabIndex = 385;
                button6.Text = "設定";
                button6.UseVisualStyleBackColor = true;
                button6.before = false;
                button6.Interval = 5000;
                button6.Click += new System.EventHandler(this.button6_Click);
                this.tabPage1.Controls.Add(button6);

                檢測模式ToolStripMenuItem.Checked = true;

                //userChange(false);
                if (app.param["origin"] == "true") 原圖ToolStripMenuItem.Checked = true;
                if (app.param["OK"] == "true") oKToolStripMenuItem.Checked = true;
                if (app.param["NG"] == "true") nGToolStripMenuItem.Checked = true;
                if (app.param["NULL"] == "true") nULLToolStripMenuItem.Checked = true;
                if (app.param["report"] == "true") 報表ToolStripMenuItem1.Checked = true;
                if (app.param["saveROI"] == "true")  rOIToolStripMenuItem.Checked = true;
                if (app.param["saveStations"] == "true") stationsToolStripMenuItem.Checked = true;
                if (app.param["saveVIS"] == "true") stationsToolStripMenuItem.Checked = true;
                if (app.param["alert"] == "true")
                {
                    蜂鳴器ToolStripMenuItem.Checked = true;
                    if (!app.offline)
                    {
                        PLC_SetM(20, true);
                    }
                }

                if (app.param["dooralert"] == "true")
                {
                    button27.Text = "ON";
                    button27.BackColor = Color.Lime;
                    if (!app.offline)
                    {
                        PLC_SetM(22, true);
                    }
                }
                else
                {
                    button27.Text = "OFF";
                    button27.BackColor = Color.FromArgb(255, 128, 128);
                    if (!app.offline)
                    {
                        PLC_SetM(22, false);
                    }
                }

                #endregion
                #region 當日計數
                //dailyCheck();
                #endregion
                #region 刪檔作業
                //try
                //{
                //    var dir = @".\image\";
                //    var days = int.Parse(app.param["KeepDay"]);

                //    if (!Directory.Exists(dir) || days < 1) return;

                //    var now = DateTime.Now;
                //    foreach (var f in Directory.GetFileSystemEntries(dir).Where(f => Directory.Exists(f)))
                //    {
                //        var t = File.GetCreationTime(f);

                //        var elapsedTicks = now.Ticks - t.Ticks;
                //        var elapsedSpan = new TimeSpan(elapsedTicks);

                //        if (elapsedSpan.TotalDays > days) Directory.Delete(f, true);
                //    }
                //}
                //catch (Exception e1)
                //{
                //    lbAdd("刪檔錯誤", "err");
                //    lbAdd(e1.ToString(), "err");
                //}

                //try
                //{
                //    var dir = @".\logs\";
                //    var days = int.Parse(app.param["KeepDay"]);

                //    if (!Directory.Exists(dir) || days < 1) return;

                //    var now = DateTime.Now;
                //    foreach (var f in Directory.GetFileSystemEntries(dir).Where(f => File.Exists(f)))
                //    {
                //        var t = File.GetCreationTime(f);

                //        var elapsedTicks = now.Ticks - t.Ticks;
                //        var elapsedSpan = new TimeSpan(elapsedTicks);

                //        if (elapsedSpan.TotalDays > days) File.Delete(f);
                //    }
                //}
                //catch (Exception e1)
                //{
                //    lbAdd("刪檔錯誤", "err");
                //    lbAdd(e1.ToString(), "err");
                //}
                #endregion
                //setNet();
                //tasks.Add(Task.Factory.StartNew(() => buttonM9()));


            }

            if (app.start_error)
            {
                Close();
            }
            #endregion
            #region smartKey
            if (!app.offline)
            {
                /*
                var keyhash = "9661a66c8bdfe640294116569cccff2696eaddae041285d9f53d99e15decc4a7";
                smartKey = new SmartKeyPro(new byte[] { 0x12, 0x34, 0x56, 0x78 }, keyhash, KeyRemove);


                void KeyRemove(object sender1, KeyRemoveArgs e1)
                {
                    if (e1.isRemove)
                    {
                        var w = new Form() { Size = new System.Drawing.Size(0, 0) };
                        Task.Delay(TimeSpan.FromSeconds(3))
                            .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
                        MessageBox.Show(w, "未偵測到智慧卡，程式即將關閉", "警告");
                        lbAdd("錯誤--未偵測到智慧卡。", "err", "");
                        BeginInvoke(new Action(Close));
                    }
                }
                */
            }
            #endregion
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)  // ✅ 改用 FormClosing
        {
            // 執行關閉確認邏輯
            if (!PerformCloseSequence())
            {
                e.Cancel = true; // ✅ FormClosingEventArgs 有 Cancel 屬性
                return;
            }

            // 儲存最後設定
            using (var db = new MydbDB())
            {
                db.LastSettings.Delete();
                db.LastSettings
                    .Value(p => p.LastNoIndex, comboBox1.SelectedIndex.ToString())
                    .Value(p => p.LastUserIndex, comboBox2.SelectedIndex.ToString())
                    .Value(p => p.LastLotid, textBox1.Text)
                    .Value(p => p.LastOrder, textBox2.Text)
                    .Value(p => p.LastPack, textBox3.Text)
                    .Value(p => p.LastNGLimit, textBox4.Text)
                    .Value(p => p.LastNULLLimit, textBox5.Text)
                    .Insert();
            }

            _plcQueueMonitorTimer?.Dispose();
            lbAdd("關閉檢測程式", "inf", "");
            Log.CloseAndFlush();
        }
        #endregion

        #region D值監控
        private void StartPLCMonitoring()
        {
            // 每 500 毫秒執行一次監控
            _plcMonitorTimer = new System.Threading.Timer(
                MonitorPLCValues,
                null,
                TimeSpan.FromSeconds(1),  // 延遲 1 秒後開始
                TimeSpan.FromMilliseconds(500)  // 每 0.5 秒執行一次
            );
        }
        private void MonitorPLCValues(object state)
        {
            // 只在系統運行時監控
            if (!app.status || app.offline)
                return;

            try
            {
                lock (_monitorLock)
                {
                    // 讀取 D803 (OK1)
                    int d803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803);

                    // 讀取 D805 (OK2)
                    int d805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805);

                    // 記錄到 LOG
                    //Log.Debug($"[PLC監控] D803(OK1)={d803}, D805(OK2)={d805}");

                    // 可選：如果需要更詳細的資訊，可以加上時間戳記
                    Log.Information($"[PLC監控] {DateTime.Now:HH:mm:ss.fff} D803(OK1)={d803}, D805(OK2)={d805}");
                }
            }
            catch (Exception ex)
            {
                // 避免監控錯誤影響主程式
                Log.Warning($"[PLC監控] 讀取失敗: {ex.Message}");
            }
        }
        #endregion
        #region PLC 訊號數量監控
        private void StartPLCQueueMonitoring()
        {
            // 由 GitHub Copilot 產生
            // 修正: 計時器週期設定為 500 毫秒，提供足夠的監控頻率
            _plcQueueMonitorTimer = new System.Threading.Timer(
                MonitorPLCQueueLengths,
                null,
                TimeSpan.FromSeconds(1),                // 首次延遲：1秒後啟動
                TimeSpan.FromMilliseconds(500)          // 週期：每 500 毫秒檢查一次
            );
        }

        private DateTime _lastQueueLogTime = DateTime.MinValue;

        private void MonitorPLCQueueLengths(object state)
        {
            // 由 GitHub Copilot 產生
            // 修正: 使用時間戳判斷，確保每秒最多寫一次 LOG

            // 只在系統運行時監控
            if (!app.status)
                return;

            try
            {
                lock (_queueMonitorLock)
                {
                    // 檢查距離上次寫 LOG 是否已超過 1 秒
                    if ((DateTime.Now - _lastQueueLogTime).TotalSeconds >= 1.0)
                    {
                        // 獲取佇列計數
                        int normalQueueCount = PLC_ModBus.SendData.Count;
                        int emergencyQueueCount = PLC_ModBus.SendData_Emergency.Count;
                        int totalCount = normalQueueCount + emergencyQueueCount;

                        // 寫入 LOG
                        Log.Information($"[PLC佇列監控] {DateTime.Now:HH:mm:ss.fff} | 普通佇列={normalQueueCount}, 緊急佇列={emergencyQueueCount}, 總計={totalCount}");

                        // 更新時間戳
                        _lastQueueLogTime = DateTime.Now;

                        // 可選：檢查佇列是否超過設定的閾值
                        if (normalQueueCount > 5)
                        {
                            Log.Warning($"[PLC佇列監控] 普通佇列超過5: {normalQueueCount}");
                        }

                        if (emergencyQueueCount > 5)
                        {
                            Log.Warning($"[PLC佇列監控] 緊急佇列超過5: {emergencyQueueCount}");
                        }

                        if (totalCount > 10)
                        {
                            Log.Warning($"[PLC佇列監控] 總佇列超過10: {totalCount}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error($"MonitorPLCQueueLengths 錯誤: {ex.Message}");
            }
        }
        #endregion

        #region Receiver
        public void Receiver(int camID, Mat Src, DateTime dt)
        {
            // 由 GitHub Copilot 產生 - 調機模式防禦性檢查
            if (app.isAdjustmentMode)
            {
                // 理論上不應該到這裡（因為 Camera0.cs 的 OnImageGrabbed 已經阻擋）
                // 但保留這個檢查作為雙重保險
                Log.Warning($"Receiver: 調機模式下收到影像 (CamID={camID})，已忽略");
                Src?.Dispose();  // 釋放影像
                return;
            }

            if (!app.SoftTriggerMode)
            {
                if (app.status)
                {
                    if (app.DetectMode == 0)
                    {
                        Mat clonedMat = null; //clonedMat 在getmat1~4 釋放
                        bool transferred = false;  // 追蹤所有權是否已轉移
                        try
                        {
                            #region 第一個樣品是否已經過光纖
                            if (!_systemInitialized)
                            {
                                Log.Debug($"系統尚未初始化（等待 D100 > 0），忽略 Camera {camID} 的影像");
                                Src?.Dispose(); // 記得釋放記憶體
                                return;
                            }
                            #endregion

                            clonedMat = Src.Clone();

                            #region 記憶體監控
                            /*
                            lock (app.counter)
                            {
                                //app.counter.AddOrUpdate("stop" + camID, 1, (key, oldValue) => oldValue + 1);
                                SAMPLE_ID = app.counter["stop3"];
                            }

                            // 由 GitHub Copilot 產生
                            // 新增：每 100 個樣品監控一次記憶體（統一在 Receiver 監控）
                            lock (_memoryMonitorLock)
                            {
                                if (SAMPLE_ID % 100 == 0 && SAMPLE_ID != _lastMemoryCheckSampleId)
                                {
                                    _lastMemoryCheckSampleId = SAMPLE_ID;
                                    _currentProcess.Refresh();

                                    long privateMemory = _currentProcess.PrivateMemorySize64 / (1024 * 1024);
                                    long workingSet = _currentProcess.WorkingSet64 / (1024 * 1024);
                                    long managedMemory = GC.GetTotalMemory(false) / (1024 * 1024);

                                    int q1 = app.Queue_Bitmap1.Count;
                                    int q2 = app.Queue_Bitmap2.Count;
                                    int q3 = app.Queue_Bitmap3.Count;
                                    int q4 = app.Queue_Bitmap4.Count;
                                    int qSave = app.Queue_Save.Count;
                                    
                                    Log.Information($"[記憶體監控] 站{camID + 1} 樣品ID={SAMPLE_ID}, " +
                                                  $"Private={privateMemory} MB, WorkingSet={workingSet} MB, Managed={managedMemory} MB, " +
                                                  $"佇列: Q1={q1}, Q2={q2}, Q3={q3}, Q4={q4}, QSave={qSave}");
                                    
                                    // 警告閾值檢查
                                    if (privateMemory > 4096)
                                    {
                                        Log.Warning($"[記憶體監控] ⚠️ 記憶體使用過高: {privateMemory} MB (站{camID + 1}, 樣品{SAMPLE_ID})");
                                    }

                                    if (q1 + q2 + q3 + q4 > 200)
                                    {
                                        Log.Warning($"[記憶體監控] ⚠️ 佇列積壓過多: {q1 + q2 + q3 + q4} 張影像 (站{camID + 1}, 樣品{SAMPLE_ID})");
                                    }
                                }
                            }*/
                   
                            #endregion

                            app.lastIn = DateTime.Now;
                            if (camID == 0)
                            {
                                app.Queue_Bitmap1.Enqueue(new ImagePosition(clonedMat, 1, app.counter["stop" + camID]));
                                app._wh1.Set();
                                transferred = true;  // ✅ 標記已轉移
                                clonedMat = null;
                            }
                            else if (camID == 1)
                            {
                                if(app.counter["stop1"] == app.counter["stop0"])
                                {
                                    clonedMat?.Dispose(); // 不處理的話立即釋放
                                    return;
                                }
                                app.Queue_Bitmap2.Enqueue(new ImagePosition(clonedMat, 2, app.counter["stop" + camID]));
                                app._wh2.Set();
                                transferred = true;  // ✅ 標記已轉移
                                clonedMat = null;
                            }
                            else if (camID == 2)
                            {
                                if (app.counter["stop2"] == app.counter["stop1"])
                                {
                                    clonedMat?.Dispose(); // 不處理的話立即釋放
                                    return;
                                }
                                app.Queue_Bitmap3.Enqueue(new ImagePosition(clonedMat, 3, app.counter["stop" + camID]));
                                app._wh3.Set();
                                transferred = true;  // ✅ 標記已轉移
                                clonedMat = null;
                            }
                            else if (camID == 3)
                            {
                                if (app.counter["stop3"] == app.counter["stop2"])
                                {
                                    clonedMat?.Dispose(); // 不處理的話立即釋放
                                    return;
                                }
                                // 以 PLC D98 前進為放行門檻，避免站4誤觸造成 sampleID 超前
                                // 說明：D98 為站4光纖 X23 的節拍索引（模 15），每次真顆通過必然前進一次
                                // 第四站拍照時，記錄時間戳
                                int currentSampleId = app.counter["stop" + camID]; 
                                // 四站取像完成，不管後續是否急停，可視為完整取像，代表1~4站sampleID統一不會亂掉，所以可以寫進資料庫

                                UpdateDetectionRate();

                                // 記錄在樣品拍攝時間字典中
                                app.samplePhotoTimes[currentSampleId] = dt;

                                app.Queue_Bitmap4.Enqueue(new ImagePosition(clonedMat, 4, currentSampleId, dt));
                                app._wh4.Set();
                                transferred = true;  // ✅ 標記已轉移
                                clonedMat = null;

                                //Log.Debug($"樣品 {currentSampleId} 站4拍攝時間記錄: {dt.ToString("HH:mm:ss.fff")}");
                            }
                            app.counter["stop" + camID]++; //重要!!! 計數來源!!!
                        }
                        catch (Exception e1)
                        {
                            lbAdd("取像發生錯誤", "err", e1.ToString());
                        }
                        finally
                        {
                            // ✅ 只有未轉移的 clonedMat 才釋放
                            if (!transferred && clonedMat != null)
                            {
                                clonedMat.Dispose();
                                Log.Debug($"Receiver: clonedMat 未入隊已釋放 (CamID: {camID})");
                            }
                        }
                    }
                    else if (app.DetectMode == 1)
                    {
                        try
                        {
                            // 由 GitHub Copilot 產生
                            // 修正: 調機模式也需要 Clone,因為 OnImageGrabbed 返回後 Src 可能被 GC
                            Mat clonedMat = null;
                            try
                            {
                                clonedMat = Src.Clone();
                            }
                            catch (ObjectDisposedException ex)
                            {
                                Log.Error($"Receiver (調機模式): Clone Src 時已被釋放 (CamID: {camID}): {ex.Message}");
                                return;
                            }

                            if (camID == 0)
                            {
                                Cv2.Circle(clonedMat, new Point(989, 610), 220, Scalar.Red, 3);
                                app.Queue_Bitmap1.Enqueue(new ImagePosition(clonedMat, app.counter["stop" + camID]));
                                app._wh1.Set();
                            }
                            else if (camID == 1)
                            {
                                Cv2.Circle(clonedMat, new Point(943, 627), 500, Scalar.Red, 3);
                                app.Queue_Bitmap2.Enqueue(new ImagePosition(clonedMat, app.counter["stop" + camID]));
                                app._wh2.Set();
                            }
                            else if (camID == 2)
                            {
                                Cv2.Rectangle(clonedMat, new Rect(659, 106, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(1018, 88, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(435, 340, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(1260, 310, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(453, 680, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(1280, 680, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(683, 917, 200, 200), Scalar.Red, 3);
                                Cv2.Rectangle(clonedMat, new Rect(1024, 917, 200, 200), Scalar.Red, 3);
                                app.Queue_Bitmap3.Enqueue(new ImagePosition(clonedMat, app.counter["stop" + camID]));
                                app._wh3.Set();
                            }
                            else if (camID == 3)
                            {
                                app.Queue_Bitmap4.Enqueue(new ImagePosition(clonedMat, app.counter["stop" + camID]));
                                app._wh4.Set();
                            }
                        }
                        catch (Exception ex)
                        {
                            lbAdd($"DetectMode=1 取像錯誤 (Cam {camID}): {ex.Message}", "err", "");
                        }
                    }
                }
            }
            else
            {
                // 在軟體觸發模式下，可以選擇記錄日誌但不放入佇列處理
                Console.WriteLine($"收到軟體觸發的影像，不進行處理。Camera ID: {camID}, Time: {dt}");
            }
        }
        #endregion
        #region 演算法

        /// <summary>
        /// 參考流程: 存原圖 > 檢查白色占比(是否為有效取像) > 檢查變形 > 找非檢測區域 > 檢查倒角 > 正式檢測
        /// </summary>
        async void getMat1()
        {

            while (true)
            {
                if (app.Queue_Bitmap1.Count > 0)
                {
                    ImagePosition input;
                    app.Queue_Bitmap1.TryDequeue(out input);
                    if (app.status && input != null)
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: Receiver 已經 Clone,這裡只需檢查有效性
                        if (input.image == null || input.image.IsDisposed)
                        {
                            Log.Warning($"getMat1: input.image 已被釋放或為 null，跳過處理 (SampleID: {input.count})");
                            continue;
                        }
                        
                        try
                        {
                            if (app.DetectMode == 0)
                            {
                                #region 存原圖
                                var fname = "";
                                fname = (input.count).ToString() + "-1.jpg";
                                if (原圖ToolStripMenuItem.Checked)
                                {
                                    try
                                    {
                                        // 由 GitHub Copilot 產生
                                        // 緊急修正 (ObjectDisposedException): 直接 Enqueue 必須 Clone
                                        // 原因: input.image 後續還要給 findGap/DetectAndExtractROI 等函數使用
                                        // 只有透過 SaveImageAsync() 函數才會內部自動 Clone
                                        app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), @".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\origin\\" + fname));
                                        app._sv.Set();
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"存圖失敗: {ex.Message}");
                                    }
                                }
                                #endregion
                                //continue;
                                #region  app.detect_result create
                                /*if (!app.detect_result.ContainsKey(input.count))
                                {
                                    app.detect_result.Add(input.count, "OK");
                                    app.detect_result_check.Add(input.count, new bool[4] { false, false, false, false });
                                }*/
                                #endregion

                                string sampleIdentifier = $"Sample_{input.count}";
                                if (app.enableProfiling && app.profilingStations.Contains(1))
                                {
                                    PerformanceProfiler.StartMeasure($"{input.count}_getmat1");
                                }

                                //Log.Debug($"樣品 {input.count} 第1站處理開始時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");

                                #region 是否為有效取像（白比率）停用
                                /*
                                // 檢查白色像素占比 (NULL)
                                Mat whiteCheckImage = null;
                                try
                                {
                                    whiteCheckImage = input.image.Clone();
                                    bool isValidImage = CheckWhitePixelRatio(whiteCheckImage, input.stop);
                                    if (!isValidImage)
                                    {
                                        // 顯示檢測結果
                                        showResultMat(whiteCheckImage, input.stop);

                                        // 由 GitHub Copilot 產生
                                        // 修正: FinalMap 必須使用 Clone，避免與 input.image 共享同一物件
                                        // 建立檢測結果物件
                                        StationResult WhitePixelResult = new StationResult
                                        {
                                            Stop = input.stop,
                                            IsNG = false,
                                            OkNgScore = 0.0f,
                                            FinalMap = input.image.Clone(),  // 使用 Clone 避免 double-free
                                            DefectName = "NULL_invalid_White",
                                            DefectScore = 1.0f,
                                            OriName = input.name
                                        };

                                        // 添加結果到管理器
                                        app.resultManager.AddResult(input.count, WhitePixelResult);
                                        //updateLabel();
                                        continue;
                                    }
                                }
                                finally
                                {
                                    whiteCheckImage?.Dispose();
                                }
                                */
                                #endregion

                                #region 是否變形，無效取像也有可能在這裡檢出，因為使用固定座標
                                /*
                                PerformanceProfiler.StartMeasure($"{input.count}_DetectCircles1");
                                var (outerCircles, innerCircles) = DetectCircles(input.image, input.stop);
                                PerformanceProfiler.StopMeasure($"{input.count}_DetectCircles1");
                                */
                                // 根本解決方案: 傳遞 Clone 副本,避免 finally 塊釋放 input.image 時影響函數內部
                                // 原因: async 函數的 finally 塊可能在函數呼叫完成前執行
                                Mat gapInputImage = null;
                                bool gapIsNG = false;
                                Mat gapResult = null;
                                List<Point> non = null;
                            
                                try
                                {
                                    gapInputImage = input.image.Clone();
                                    (gapIsNG, gapResult, non) = findGapWidth(gapInputImage, input.stop);

                                    if (gapIsNG)
                                    {
                                        showResultMat(gapResult, input.stop);

                                        StationResult GapResult = new StationResult
                                        {
                                            Stop = input.stop,
                                            IsNG = gapIsNG,
                                            OkNgScore = 1.0f,
                                            FinalMap = gapResult.Clone(),  // ✅ Clone 避免記憶體共享
                                            DefectName = "deform",
                                            DefectScore = 1.0f,
                                            OriName = input.name
                                        };

                                        app.resultManager.AddResult(input.count, GapResult);
                                        continue;
                                    }
                                }
                                finally
                                {
                                    gapInputImage?.Dispose();
                                    gapResult?.Dispose();  // ✅ 確保釋放
                                }
                                // 保存開口位置供後續使用
                                //List<Point> detectedGapPositions = gapPositions;
                                #endregion

                                
                                // 由 GitHub Copilot 產生
                                // 根本解決方案: 傳遞 Clone 副本給 DetectAndExtractROI
                                Mat roiInputImage = null;
                                Mat roi = null;
                                Mat nong = null;
                                try
                                {
                                    roiInputImage = input.image.Clone();
                                    roi = DetectAndExtractROI(roiInputImage, input.stop, input.count);
                                    // 由 GitHub Copilot 產生
                                    // 緊急修正: 檢查 roi 是否有效，避免後續操作失敗
                                    if (roi == null || roi.IsDisposed || roi.Empty())
                                    {
                                        Log.Warning($"getMat1: roi 為 null、已被釋放或為空 (SampleID: {input.count})");
                                        continue;
                                    }

                                    #region 找NROI
                                    (bool nonb, Mat nongTemp, List<Point> detectedGapPositions) = findGap(roiInputImage, input.stop);
                                    nong = nongTemp;

                                    bool performNonRoiDetection = app.param.ContainsKey($"find_NROI_{input.stop}") && int.Parse(app.param[$"find_NROI_{input.stop}"]) == 1;
                                    //List<Rect> nonRoiRects = new List<Rect>();
                                    List<(Rect rect, string className, double score)> nonRoiRects = new List<(Rect, string, double)>();

                                    if (performNonRoiDetection)
                                    {
                                        if (app.enableProfiling && app.profilingStations.Contains(1))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_nonRoiDetection1");
                                        }
                                        // 執行非 ROI 區域檢測
                                        string nonRoiServerUrl = app.produce_inner_NROI_ServerUrl;
                                        DetectionResponse nonRoiDetection = await _yoloDetection.PerformObjectDetection(roi, $"{nonRoiServerUrl}/detect");

                                        // 檢查非 ROI 檢測結果
                                        if (nonRoiDetection.detections != null && nonRoiDetection.detections.Count > 0)
                                        {
                                            foreach (var detection in nonRoiDetection.detections)
                                            {
                                                //nonRoiRects.Add(new Rect(detection.box[0], detection.box[1], detection.box[2] - detection.box[0], detection.box[3] - detection.box[1]));
                                                Rect rect = new Rect(detection.box[0], detection.box[1], detection.box[2] - detection.box[0], detection.box[3] - detection.box[1]);
                                                nonRoiRects.Add((rect, detection.class_name, detection.score));
                                            }
                                        }
                                        if (app.enableProfiling && app.profilingStations.Contains(1))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_nonRoiDetection1");
                                        }
                                    }

                                    #endregion

                                    #region 倒角
                                    // 由 GitHub Copilot 產生
                                    // 修正: 使用快取檢查是否需要倒角檢測,避免資料庫鎖定
                                    string chamferCacheKey = $"{app.produce_No}_{input.stop}";
                                    bool needChamferDetection = app.chamferDetectionCache.TryGetValue(chamferCacheKey, out bool cached) 
                                        ? cached 
                                        : false;

                                    // 如果不需要檢測倒角，則跳過此區塊
                                    if (needChamferDetection)
                                    {

                                        //
                                        // 由 GitHub Copilot 產生
                                        // 修正 P1-1: 為 chamferRoi 加 using，確保釋放 (15-20 MB)
                                        //檢查倒角
                                        using (Mat chamferRoi = DetectAndExtractROI(roiInputImage, input.stop, input.count, true))
                                        {
                                            string chamferServerUrl = app.produce_chamferServerUrl;

                                            if (app.enableProfiling && app.profilingStations.Contains(1))
                                            {
                                                PerformanceProfiler.StartMeasure($"{input.count}_yolo_chamferRoi1");
                                            }
                                            DetectionResponse chamferDetection = await _yoloDetection.PerformObjectDetection(chamferRoi.Clone(), $"{chamferServerUrl}/detect");
                                            if (app.enableProfiling && app.profilingStations.Contains(1))
                                            {
                                                PerformanceProfiler.StopMeasure($"{input.count}_yolo_chamferRoi1");
                                            }
                                            if (chamferDetection.error != null)
                                            {
                                                lbAdd($"站點1檢測錯誤: {chamferDetection.error}", "err", "");
                                                continue;
                                            }

                                            List<DetectionResult> chamferDefects = new List<DetectionResult>(); // 取檢測結果
                                            if (chamferDetection.detections != null && chamferDetection.detections.Count > 0)
                                            {
                                                foreach (var defect in chamferDetection.detections)
                                                {
                                                    bool isOverlapping = false;
                                                    if (performNonRoiDetection && nonRoiRects.Count > 0)
                                                    {
                                                        Rect defectRect = new Rect(defect.box[0], defect.box[1], defect.box[2] - defect.box[0], defect.box[3] - defect.box[1]);

                                                        foreach (var nonRoiRect in nonRoiRects)
                                                        {
                                                            double iou = CalculateIoU(defectRect, nonRoiRect.rect);
                                                            int expand_in = 0;
                                                            int expand_out = 0;
                                                            if (nonRoiRect.className == "cyg")
                                                            {
                                                                expand_in = GetIntParam(app.param, $"expandNROI_in_{input.stop}", 0);
                                                                expand_out = GetIntParam(app.param, $"expandNROI_out_{input.stop}", 0);
                                                            }
                                                            else if (nonRoiRect.className == "mouth")
                                                            {
                                                                expand_in = GetIntParam(app.param, $"expandNROI_in_3", 0);
                                                                expand_out = GetIntParam(app.param, $"expandNROI_out_4", 0);
                                                            }

                                                            bool inNonRoi = IsDefectInNonRoiRegion_in(defectRect, nonRoiRect.rect, input.stop, expand_in, expand_out);

                                                            if (iou > GetDoubleParam(app.param, $"IOU_{input.stop}", 0.2) || inNonRoi)
                                                            {
                                                                isOverlapping = true;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (!isOverlapping)
                                                    {
                                                        chamferDefects.Add(defect);
                                                    }
                                                }
                                            }

                                            if (chamferDefects.Count > 0) //若有檢到
                                            {
                                                bool has_chamfer_Defect = false;
                                                string chamfer_defectName = "OK";
                                                float chamfer_highestScore = 0;
                                                float chamfer_threshold = 0.5f; // 默認閾值

                                                var highestScoreDefect = chamferDefects
                                                    .OrderByDescending(d => d.score)
                                                    .FirstOrDefault();

                                                if (highestScoreDefect != null)
                                                {
                                                    chamfer_highestScore = (float)highestScoreDefect.score;

                                                    // 取得對應的閥值 在 defect_check
                                                    if (app.param.TryGetValue(highestScoreDefect.class_name + "_threshold", out string thresholdStr))
                                                    {
                                                        float.TryParse(thresholdStr, out chamfer_threshold);
                                                    }
                                                    has_chamfer_Defect = true;
                                                    chamfer_defectName = highestScoreDefect.class_name;

                                                    if (chamfer_highestScore > chamfer_threshold) //若超過閥值
                                                    {
                                                        // 繪製檢測結果
                                                        using (Mat chamfer_resultImage = _yoloDetection.DrawDetectionResults(roiInputImage, new DetectionResponse { detections = chamferDefects }, chamfer_threshold))
                                                        {
                                                            // 使用現有的 showResultMat 方法顯示結果
                                                            showResultMat(chamfer_resultImage, 1);

                                                            // 創建StationResult物件
                                                            StationResult chamferResult = new StationResult
                                                            {
                                                                Stop = 1, // 站點1
                                                                IsNG = true,
                                                                OkNgScore = chamfer_highestScore > 0 ? (float?)chamfer_highestScore : 0.0f,
                                                                FinalMap = chamfer_resultImage.Clone(), // ✅ P0-2 修正: Clone 避免 using 釋放後 FinalMap 無效
                                                                DefectName = chamfer_defectName,
                                                                DefectScore = has_chamfer_Defect ? (float?)chamfer_highestScore : 0.0f, // 只有NG才有瑕疵分數
                                                                OriName = Path.GetFileName(input.name)
                                                            };

                                                            // 添加結果到結果管理器
                                                            app.resultManager.AddResult(input.count, chamferResult);
                                                        } // 結束 chamfer_resultImage 的 using 語句
                                                        continue;
                                                    }   
                                                }
                                            }
                                        } // 由 GitHub Copilot 產生 - 結束 chamferRoi 的 using 語句
                                    }
                    
                                    #endregion   #region 正常yolo

                                    #region YOLO
                      
                                    using (Mat visualizationImage = input.image.Clone())
                                    {
                                        List<string> defectsToDetect = GetDefectNameListForThisStop(app.produce_No, input.stop);

                                        string defectServerUrl = "";
                                        // AI 推論 OK/NG
                                        if (!string.IsNullOrEmpty(app.produce_innerServerUrl))
                                        {
                                            // 已經有值（被更新過）
                                            defectServerUrl = app.produce_innerServerUrl;
                                        }
                                        else
                                        {
                                            defectServerUrl = app.produce_station1ServerUrl;
                                        }

                                        if (app.enableProfiling && app.profilingStations.Contains(1))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_yolo_Inference1");
                                        }
                                        DetectionResponse defectDetection = await _yoloDetection.PerformObjectDetection(roi, $"{defectServerUrl}/detect");
                                        if (app.enableProfiling && app.profilingStations.Contains(1))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_yolo_Inference1");
                                        }

                                        if (defectDetection.error != null)
                                        {
                                            lbAdd($"站點1檢測錯誤: {defectDetection.error}", "err", "");
                                            continue;
                                        }
                                        // 處理瑕疵檢測結果
                                        if (app.enableProfiling && app.profilingStations.Contains(1))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_holdResult1");
                                        }

                                        List<DetectionResult> validDefects = new List<DetectionResult>();
                                        if (defectDetection.detections != null && defectDetection.detections.Count > 0)
                                        {

                                            float minArea = 500.0f; // 預設閥值
                                            if (app.param.TryGetValue("minArea" + input.stop.ToString() + "_threshold", out string thresholdStr))
                                            {
                                                float.TryParse(thresholdStr, out minArea);
                                            }    
                                    
                                            foreach (var defect in defectDetection.detections)
                                            {
                                                // 檢查這個瑕疵是否需要檢測
                                                if (!defectsToDetect.Contains(defect.class_name))
                                                {
                                                    continue;  // 跳過不需要檢測的瑕疵
                                                }

                                                // 新增：特殊處理 INP 瑕疵：檢查是否位於開口位置
                                                if (defect.class_name == "INP" && detectedGapPositions.Count > 0)
                                                {
                                                    Rect defectINP = new Rect(defect.box[0], defect.box[1],
                                                                              defect.box[2] - defect.box[0],
                                                                              defect.box[3] - defect.box[1]);

                                                    Point defectCenter = new Point(
                                                        defectINP.X + defectINP.Width / 2,
                                                        defectINP.Y + defectINP.Height / 2
                                                    );

                                                    // 檢查瑕疵中心是否接近任何開口位置
                                                    bool isNearGap = false;
                                                    foreach (var gapPos in detectedGapPositions)
                                                    {
                                                        double distance = Math.Sqrt(
                                                            Math.Pow(defectCenter.X - gapPos.X, 2) +
                                                            Math.Pow(defectCenter.Y - gapPos.Y, 2)
                                                        );

                                                        if (distance <= 80) // 80像素容忍範圍，可調整
                                                        {
                                                            isNearGap = true;
                                                            break;
                                                        }
                                                    }

                                                    if (isNearGap)
                                                    {
                                                        //Log.Debug($"INP瑕疵位於開口位置附近，跳過檢測: 瑕疵中心({defectCenter.X}, {defectCenter.Y})");
                                                        continue; // 跳過這個 INP 瑕疵
                                                    }
                                                }

                                                Rect defectRect = new Rect(defect.box[0], defect.box[1], defect.box[2] - defect.box[0], defect.box[3] - defect.box[1]);

                                                // 計算瑕疵面積
                                                int width = defect.box[2] - defect.box[0];
                                                int height = defect.box[3] - defect.box[1];
                                                int area = width * height;
                                                float aspectRatio = Math.Max(width, height) / (float)Math.Min(width, height);

                                                // 如果長寬比小於等於2.5且面積小於500，跳過此瑕疵
                                                if (aspectRatio <= 2.5f && area < minArea)
                                                {
                                                    continue;
                                                }

                                                // 如果執行了非 ROI 區域檢測，則檢查與非 ROI 區域的 IoU
                                                bool isOverlapping = false;
                                                if (performNonRoiDetection && nonRoiRects.Count > 0)
                                                {
                                                    using (Mat nonRoiVisImage = input.image.Clone())
                                                    {
                                                        foreach (var nonRoiRect in nonRoiRects)
                                                        {
                                                            double iou = CalculateIoU(defectRect, nonRoiRect.rect);
                                                            int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                            int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);
                                                            bool inNonRoi = IsDefectInNonRoiRegion_in(defectRect, nonRoiRect.rect, input.stop, expand_in, expand_out);
                                                            //nonRoiRect.class_name;
                                                            // 使用DrawNonRoiRegion繪製擴展後的非ROI區域
                                                            /*
                                                            nonRoiVisImage = DrawNonRoiRegion_in(nonRoiVisImage, nonRoiRect.rect, input.stop, expand_in, expand_out);

                                                            // 在非ROI視覺化影像上添加IoU值的顯示
                                                            Cv2.PutText(nonRoiVisImage, $"IoU: {iou:F3}",
                                                                new Point(nonRoiRect.rect.X, nonRoiRect.rect.Y + nonRoiRect.rect.Height + 20),
                                                                HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);

                                                            Cv2.PutText(nonRoiVisImage, $"inNROI: {inNonRoi}",
                                                                new Point(nonRoiRect.rect.X, nonRoiRect.rect.Y + nonRoiRect.rect.Height + 50),
                                                                HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);
                                                            */

                                                            if (iou > GetDoubleParam(app.param, $"IOU_{input.stop}", 0.2) || inNonRoi)
                                                            {
                                                                isOverlapping = true;
                                                                break;
                                                            }                                             
                                                        }
                                                        /*
                                                        // 保存視覺化影像（如果需要）
                                                        string visImgPath = $@".\image\{st.ToString("yyyy-MM")}\{st.ToString("MMdd")}\{app.foldername}\vis\{input.count}_{input.stop}_vis.jpg";
                                                        try
                                                        {
                                                            app.Queue_Save.Enqueue(new ImageSave(nonRoiVisImage, visImgPath));
                                                            app._sv.Set();
                                                            //Directory.CreateDirectory(Path.GetDirectoryName(visImgPath));
                                                            //Cv2.ImWrite(visImgPath, nonRoiVisImage);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Console.WriteLine($"保存視覺化影像失敗: {ex.Message}");
                                                        }
                                                        */
                                                    } // 結束 nonRoiVisImage 的 using 語句
                                                }

                                                if (!isOverlapping)
                                                {
                                                    validDefects.Add(defect);
                                                }
                                            }
                                        }
                                

                                        // 找出最高分的瑕疵檢測結果
                                        bool hasDefect = false;
                                        string defectName = "OK";
                                        float highestScore = 0;
                                        float threshold = 0.5f; // 默認閾值

                                        if (validDefects.Count > 0)
                                        {
                                            // 依分數從高到低排序所有有效瑕疵
                                            var sortedDefects = validDefects
                                                .OrderByDescending(d => d.score)
                                                .ToList();

                                            // 逐一檢查每個瑕疵，看是否有任何一個超過其閥值
                                            foreach (var defect in sortedDefects)
                                            {
                                                // 獲取此瑕疵類型對應的閥值
                                                float defectThreshold = 0.5f; // 預設閥值
                                                if (app.param.TryGetValue(defect.class_name + input.stop.ToString() + "_threshold", out string thresholdStr))
                                                {
                                                    float.TryParse(thresholdStr, out defectThreshold);
                                                }

                                                // 如果此瑕疵分數超過其閥值，記錄下來並跳出循環
                                                if (defect.score > defectThreshold)
                                                {
                                                    hasDefect = true;
                                                    defectName = defect.class_name;
                                                    highestScore = (float)defect.score;
                                                    threshold = defectThreshold; // 記錄此瑕疵的閥值供後續使用
                                                    break; // 找到一個超過閥值的瑕疵就停止
                                                }
                                            }

                                            // 如果沒有找到超過閥值的瑕疵，仍然記錄最高分瑕疵的資訊
                                            if (!hasDefect && sortedDefects.Count > 0)
                                            {
                                                var highestScoreDefect = sortedDefects[0];
                                                highestScore = (float)highestScoreDefect.score;
                                                defectName = highestScoreDefect.class_name;

                                                // 獲取閥值供後續使用
                                                if (app.param.TryGetValue(defectName + "_threshold", out string thresholdStr))
                                                {
                                                    float.TryParse(thresholdStr, out threshold);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // 沒有檢測到任何物件，也視為OK
                                            hasDefect = false;
                                            defectName = "OK";
                                            highestScore = 0;
                                        }
                                        // 繪製檢測結果
                                        using (Mat resultImage = _yoloDetection.DrawDetectionResults(visualizationImage.Clone(), new DetectionResponse { detections = validDefects }, threshold))
                                        {
                                            // 使用現有的 showResultMat 方法顯示結果
                                            showResultMat(resultImage, 1);

                                            // 創建StationResult物件
                                            StationResult stationResult = new StationResult
                                            {
                                                Stop = 1, // 站點1
                                                IsNG = hasDefect,
                                                OkNgScore = highestScore > 0 ? (float?)highestScore : 0.0f,
                                                FinalMap = resultImage.Clone(), // ✅ P0-1 修正: Clone 避免 using 釋放後 FinalMap 無效
                                                DefectName = defectName, // 若有瑕疵超過閥值，則為該瑕疵名稱；否則為最高分瑕疵名稱
                                                DefectScore = highestScore > 0 ? (float?)highestScore : 0.0f, // 若有檢測到物件，保留分數
                                                OriName = Path.GetFileName(input.name)
                                            };

                                            //Log.Debug($"樣品 {input.count} 第1站處理結束時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");

                                            // 添加結果到結果管理器
                                            app.resultManager.AddResult(input.count, stationResult);
                                        } // 結束 resultImage 的 using 語句

                                        // 添加到統計管理器
                                        //ScoreStatisticsManager.AddScore(1, highestScore, hasDefect, defectName);

                                        // 更新檢測率
                                        //updateLabel();
                                        //UpdateDetectionRate();
                                        if (app.enableProfiling && app.profilingStations.Contains(1))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_holdResult1");
                                        }
                                    } // 結束 visualizationImage 的 using 語句

                                    #endregion

                                    if (app.enableProfiling && app.profilingStations.Contains(1))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_getmat1");
                                    }
                                    //var totalTime = PerformanceProfiler.StopMeasure("AI_singleProcess1");
                                    //PerformanceProfiler.FlushToDisk();
                                } // 結束 roi try
                                finally
                                {
                                    // 由 GitHub Copilot 產生
                                    // 修正: 釋放 roi, nong 和 roiInputImage Mat (總計 35-50 MB)
                                    roiInputImage?.Dispose();
                                    roi?.Dispose();
                                    nong?.Dispose();
                                }
                            }

                            /*
                            else if (app.DetectMode == 1)
                            {
                                Cv2.Resize(input.image, input.image, new Size(345, 345));
                                BeginInvoke(new Action(() => cherngerPictureBox1.Image = input.image.ToBitmap()));
                                BeginInvoke(new Action(() => cherngerPictureBox1.Refresh()));
                            }
                            */
                        } // 結束 try
                        finally
                        {
                            // 修正: 使用 finally 確保 input.image 一定會被釋放，即使中途 continue
                            input.image?.Dispose();
                        }
                    }
                }
                else
                {
                    app._wh1.WaitOne();
                    if (app.user == 0)
                    {
                        Console.WriteLine("gm1 on");
                    }
                }
            }
        }
        async void getMat2()
        {
            while (true)
            {
                if (app.Queue_Bitmap2.Count > 0)
                {
                    ImagePosition input;
                    app.Queue_Bitmap2.TryDequeue(out input);
                    if (app.status && input != null)
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: Receiver 已經 Clone,這裡只需檢查有效性
                        if (input.image == null || input.image.IsDisposed)
                        {
                            Log.Warning($"getMat2: input.image 已被釋放或為 null，跳過處理 (SampleID: {input.count})");
                            continue;
                        }
                        
                        try
                        {
                            if (app.DetectMode == 0)
                            {
                                #region 存原圖
                                var fname = "";
                                fname = (input.count).ToString() + "-2.jpg";
                                if (原圖ToolStripMenuItem.Checked)
                                {
                                    try
                                    {
                                        // 由 GitHub Copilot 產生
                                        // 緊急修正 (ObjectDisposedException): 直接 Enqueue 必須 Clone
                                        // 原因: input.image 後續還要給 findGap/DetectAndExtractROI 等函數使用
                                        // 只有透過 SaveImageAsync() 函數才會內部自動 Clone
                                        app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), @".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\origin\\" + fname));
                                        app._sv.Set();
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"存圖失敗: {ex.Message}");
                                    }
                                }
                                #endregion
                                //continue;
                                #region  app.detect_result create
                                /*if (!app.detect_result.ContainsKey(input.count))
                                {
                                    app.detect_result.Add(input.count, "OK");
                                    app.detect_result_check.Add(input.count, new bool[4] { false, false, false, false });
                                }*/
                                #endregion

                                string sampleIdentifier = $"Sample_{input.count}";
                                if (app.enableProfiling && app.profilingStations.Contains(2))
                                {
                                    PerformanceProfiler.StartMeasure($"{input.count}_getmat2");
                                }
                                //Log.Debug($"樣品 {input.count} 第2站處理開始時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");

                                #region 是否為有效取像（白比率）停用
                                /*
                                // 檢查白色像素占比 (NULL)
                                Mat whiteCheckImage = null;

                                try
                                {
                                    whiteCheckImage = input.image.Clone();
                                    bool isValidImage = CheckWhitePixelRatio(whiteCheckImage, input.stop);
                                    if (!isValidImage)
                                    {
                                        // 顯示檢測結果
                                        showResultMat(whiteCheckImage, input.stop);

                                        // 由 GitHub Copilot 產生
                                        // 修正: FinalMap 必須使用 Clone，避免與 input.image 共享同一物件
                                        // 建立檢測結果物件
                                        StationResult WhitePixelResult = new StationResult
                                        {
                                            Stop = input.stop,
                                            IsNG = false,
                                            OkNgScore = 0.0f,
                                            FinalMap = input.image.Clone(),  // 使用 Clone 避免 double-free
                                            DefectName = "NULL_invalid_White",
                                            DefectScore = 1.0f,
                                            OriName = input.name
                                        };

                                        // 添加結果到管理器
                                        app.resultManager.AddResult(input.count, WhitePixelResult);
                                        //updateLabel();
                                        continue;
                                    }
                                }
                                finally
                                {
                                    whiteCheckImage?.Dispose();
                                }
                                */
                                #endregion

                                #region 是否為有效取像（物體位置），只在這裡做!

                                // 0801: 需新增區分真NULL/變形功能
                                Mat ObjCheckimage = null;
                                Mat ObjResult = null;
                                bool isObjectNG = false;
                                try
                                {
                                    ObjCheckimage = input.image.Clone();
                                    (isObjectNG, ObjResult) = CheckObjectPosition(ObjCheckimage, input.stop);

                                    if (isObjectNG)
                                    {
                                        // 顯示檢測結果
                                        showResultMat(ObjResult, input.stop);

                                        // 建立檢測結果物件
                                        StationResult CheckObjectResult = new StationResult
                                        {
                                            Stop = input.stop,
                                            IsNG = isObjectNG,
                                            OkNgScore = 0.0f,
                                            FinalMap = ObjResult.Clone(),
                                            DefectName = "NULL_Invalid_Obj",
                                            DefectScore = 1.0f,
                                            OriName = input.name
                                        };

                                        // 添加結果到管理器
                                        app.resultManager.AddResult(input.count, CheckObjectResult);
                                        //updateLabel();
                                        continue;
                                    }
                                }
                                finally
                                {
                                    ObjCheckimage?.Dispose();
                                    ObjResult?.Dispose();
                                }

                                #endregion

                                #region 是否變形，無效取像也有可能在這裡檢出，因為使用固定座標
                                /*
                                PerformanceProfiler.StartMeasure($"{input.count}_DetectCircles2");
                                var (outerCircles, innerCircles) = DetectCircles(input.image, input.stop);
                                PerformanceProfiler.StopMeasure($"{input.count}_DetectCircles2");
                                */
                                // 由 GitHub Copilot 產生
                                // 根本解決方案: 傳遞 Clone 副本給 findGapWidth
                                Mat gapInputImage = null;
                                bool gapIsNG = false;
                                Mat gapResult = null;
                                List<Point> gapPositions = null;

                                try
                                {
                                    gapInputImage = input.image.Clone();
                                    (gapIsNG, gapResult, gapPositions) = findGapWidth(gapInputImage, input.stop);

                                    if (gapIsNG)
                                    {
                                        // 顯示檢測結果
                                        showResultMat(gapResult, input.stop);

                                        // 建立檢測結果物件
                                        StationResult GapResult = new StationResult
                                        {
                                            Stop = input.stop,
                                            IsNG = gapIsNG,
                                            OkNgScore = 1.0f,
                                            FinalMap = gapResult.Clone(),
                                            DefectName = "deform",
                                            DefectScore = 1.0f,
                                            OriName = input.name
                                        };

                                        // 添加結果到管理器
                                        app.resultManager.AddResult(input.count, GapResult);
                                        //updateLabel();
                                        continue;
                                    }
                                }

                                finally
                                {
                                    gapInputImage?.Dispose();
                                    gapResult?.Dispose();
                                }
                                #endregion

                                
                                // 由 GitHub Copilot 產生
                                // 根本解決方案: 傳遞 Clone 副本給 DetectAndExtractROI
                                Mat roiInputImage = null;
                                Mat roi = null;
                                try
                                {
                                    roiInputImage = input.image.Clone();
                                    roi = DetectAndExtractROI(roiInputImage, input.stop, input.count);
                                    // 由 GitHub Copilot 產生
                                    // 緊急修正: 檢查 roi 是否有效，避免後續操作失敗
                                    if (roi == null || roi.IsDisposed || roi.Empty())
                                    {
                                        Log.Warning($"getMat2: roi 為 null、已被釋放或為空 (SampleID: {input.count})");
                                        continue;
                                    }

                                    #region 找NROI
                                    bool performNonRoiDetection = app.param.ContainsKey($"find_NROI_{input.stop}") && int.Parse(app.param[$"find_NROI_{input.stop}"]) == 1;
                                    List<Rect> nonRoiRects = new List<Rect>();

                                    if (performNonRoiDetection)
                                    {
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_nonRoiDetection2");
                                        }
                                        // 執行非 ROI 區域檢測
                                        string nonRoiServerUrl = app.produce_inner_NROI_ServerUrl;
                                        DetectionResponse nonRoiDetection = await _yoloDetection.PerformObjectDetection(roi, $"{nonRoiServerUrl}/detect");

                                        // 檢查非 ROI 檢測結果
                                        if (nonRoiDetection.detections != null && nonRoiDetection.detections.Count > 0)
                                        {
                                            foreach (var detection in nonRoiDetection.detections)
                                            {
                                                nonRoiRects.Add(new Rect(detection.box[0], detection.box[1], detection.box[2] - detection.box[0], detection.box[3] - detection.box[1]));
                                            }
                                        }
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_nonRoiDetection2");
                                        }
                                    }
                                    #endregion

                                    #region 倒角
                                    // 由 GitHub Copilot 產生
                                    // 修正: 使用快取檢查是否需要倒角檢測,避免資料庫鎖定
                                    string chamferCacheKey = $"{app.produce_No}_{input.stop}";
                                    bool needChamferDetection = app.chamferDetectionCache.TryGetValue(chamferCacheKey, out bool cached) 
                                        ? cached 
                                        : false;

                                    // 如果不需要檢測倒角，則跳過此區塊
                                    if (needChamferDetection)
                                    {
                                        //
                                        // 由 GitHub Copilot 產生
                                        // 修正 P1-2: 為 chamferRoi 加 using，確保釋放 (15-20 MB)
                                        //檢查倒角
                                        using (Mat chamferRoi = DetectAndExtractROI(roiInputImage, input.stop, input.count, true))
                                        {
                                        string chamferServerUrl = app.produce_chamferServerUrl;
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_yolo_chamferRoi2");
                                        }
                                        DetectionResponse chamferDetection = await _yoloDetection.PerformObjectDetection(chamferRoi.Clone(), $"{chamferServerUrl}/detect");
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_yolo_chamferRoi2");
                                        }
                                        if (chamferDetection.error != null)
                                        {
                                            lbAdd($"站點2檢測錯誤: {chamferDetection.error}", "err", "");
                                            continue;
                                        }

                                        List<DetectionResult> chamferDefects = new List<DetectionResult>(); // 取檢測結果
                                        if (chamferDetection.detections != null && chamferDetection.detections.Count > 0)
                                        {
                                            foreach (var defect in chamferDetection.detections)
                                            {
                                                bool isOverlapping = false;
                                                if (performNonRoiDetection && nonRoiRects.Count > 0)
                                                {
                                                    Rect defectRect = new Rect(defect.box[0], defect.box[1], defect.box[2] - defect.box[0], defect.box[3] - defect.box[1]);

                                                    foreach (var nonRoiRect in nonRoiRects)
                                                    {
                                                        double iou = CalculateIoU(defectRect, nonRoiRect);
                                                        int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                        int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);

                                                        bool inNonRoi = IsDefectInNonRoiRegion_in(defectRect, nonRoiRect, input.stop, expand_in, expand_out);

                                                        if (iou > GetDoubleParam(app.param, $"IOU_{input.stop}", 0.2) || inNonRoi)
                                                        {
                                                            isOverlapping = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (!isOverlapping)
                                                {
                                                    chamferDefects.Add(defect);
                                                }
                                            }
                                        }

                                        if (chamferDefects.Count > 0) //若有檢到
                                        {
                                            bool has_chamfer_Defect = false;
                                            string chamfer_defectName = "OK";
                                            float chamfer_highestScore = 0;
                                            float chamfer_threshold = 0.5f; // 默認閾值

                                            var highestScoreDefect = chamferDefects
                                                .OrderByDescending(d => d.score)
                                                .FirstOrDefault();

                                            if (highestScoreDefect != null)
                                            {
                                                chamfer_highestScore = (float)highestScoreDefect.score;

                                                // 取得對應的閥值 在 defect_check
                                                if (app.param.TryGetValue(highestScoreDefect.class_name + "_threshold", out string thresholdStr))
                                                {
                                                    float.TryParse(thresholdStr, out chamfer_threshold);
                                                }
                                                has_chamfer_Defect = true;
                                                chamfer_defectName = highestScoreDefect.class_name;

                                                if (chamfer_highestScore > chamfer_threshold) //若超過閥值
                                                {
                                                    // 繪製檢測結果
                                                    using (Mat chamfer_resultImage = _yoloDetection.DrawDetectionResults(roiInputImage, new DetectionResponse { detections = chamferDefects }, chamfer_threshold))
                                                    {
                                                        // 使用現有的 showResultMat 方法顯示結果
                                                        showResultMat(chamfer_resultImage, input.stop);

                                                        // 創建StationResult物件
                                                        StationResult chamferResult = new StationResult
                                                        {
                                                            Stop = 1, // 站點1
                                                            IsNG = true,
                                                            OkNgScore = chamfer_highestScore > 0 ? (float?)chamfer_highestScore : 0.0f,
                                                            FinalMap = chamfer_resultImage.Clone(), // ✅ P0-4 修正: Clone 避免 using 釋放後 FinalMap 無效
                                                            DefectName = chamfer_defectName,
                                                            DefectScore = has_chamfer_Defect ? (float?)chamfer_highestScore : 0.0f, // 只有NG才有瑕疵分數
                                                            OriName = Path.GetFileName(input.name)
                                                        };

                                                        // 添加結果到結果管理器
                                                        app.resultManager.AddResult(input.count, chamferResult);
                                                    } // 結束 chamfer_resultImage 的 using 語句
                                                    continue;
                                                }
                                            }
                                        }
                                        } // 由 GitHub Copilot 產生 - 結束 chamferRoi 的 using 語句
                                    }
                                    #endregion

                                    #region 正常yolo
                            
                                
                                    using (Mat visualizationImage = input.image.Clone())
                                    {
                                        List<string> defectsToDetect = GetDefectNameListForThisStop(app.produce_No, input.stop);

                                        string defectServerUrl = "";
                                        // AI 推論 OK/NG
                                        if (!string.IsNullOrEmpty(app.produce_innerServerUrl))
                                        {
                                            // 已經有值（被更新過）
                                            defectServerUrl = app.produce_innerServerUrl;
                                        }
                                        else
                                        {
                                            defectServerUrl = app.produce_station1ServerUrl;
                                        }
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_yolo_Inference2");
                                        }
                                        DetectionResponse defectDetection = await _yoloDetection.PerformObjectDetection(roi, $"{defectServerUrl}/detect");
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_yolo_Inference2");
                                        }

                                        if (defectDetection.error != null)
                                        {
                                            lbAdd($"站點2檢測錯誤: {defectDetection.error}", "err", "");
                                            continue;
                                        }

                                        // 處理瑕疵檢測結果
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StartMeasure($"{input.count}_holdResult2");
                                        }
                            
                                        List<DetectionResult> validDefects = new List<DetectionResult>();
                                        if (defectDetection.detections != null && defectDetection.detections.Count > 0)
                                        {
                                            float minArea = 500.0f; // 預設閥值
                                            if (app.param.TryGetValue("minArea" + input.stop.ToString() + "_threshold", out string thresholdStr))
                                            {
                                                    float.TryParse(thresholdStr, out minArea);
                                                }

                                                foreach (var defect in defectDetection.detections)
                                                {
                                                    // 檢查這個瑕疵是否需要檢測
                                                    if (!defectsToDetect.Contains(defect.class_name))
                                                    {
                                                        continue;  // 跳過不需要檢測的瑕疵
                                                    }                                    Rect defectRect = new Rect(defect.box[0], defect.box[1], defect.box[2] - defect.box[0], defect.box[3] - defect.box[1]);

                                                // 計算瑕疵面積
                                                int width = defect.box[2] - defect.box[0];
                                                int height = defect.box[3] - defect.box[1];
                                                int area = width * height;
                                                float aspectRatio = Math.Max(width, height) / (float)Math.Min(width, height);

                                                // 如果長寬比小於等於2.5且面積小於500，跳過此瑕疵
                                                if (aspectRatio <= 2.5f && area < minArea)
                                                {
                                                    continue;
                                                }

                                                if (defect.class_name == "bp") // 跳過鏡頭髒污
                                                {
                                                    int centerX = defectRect.X + defectRect.Width / 2;
                                                    int centerY = defectRect.Y + defectRect.Height / 2;
                                                    int camarea = defectRect.Width * defectRect.Height;

                                                    if (Math.Abs(centerX - 630) <= 20 &&
                                                        Math.Abs(centerY - 1115) <= 20 &&
                                                        camarea <= 2500)
                                                    {
                                                        // 跳過這個瑕疵
                                                        continue;
                                                    }
                                                }

                                                // 如果執行了非 ROI 區域檢測，則檢查與非 ROI 區域的 IoU
                                                bool isOverlapping = false;
                                                if (performNonRoiDetection && nonRoiRects.Count > 0)
                                                {
                                                    using (Mat nonRoiVisImage = input.image.Clone())
                                                    {
                                                        foreach (var nonRoiRect in nonRoiRects)
                                                        {

                                                            double iou = CalculateIoU(defectRect, nonRoiRect);
                                                            int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                            int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);
                                                            bool inNonRoi = IsDefectInNonRoiRegion_in(defectRect, nonRoiRect, input.stop, expand_in, expand_out);

                                                            /*
                                                            // 使用DrawNonRoiRegion繪製擴展後的非ROI區域
                                                            PerformanceProfiler.StartMeasure($"{input.count}_DrawNonRoiRegion_in");
                                                            nonRoiVisImage = DrawNonRoiRegion_in(visualizationImage.Clone(), nonRoiRect, input.stop, expand_in, expand_out);
                                                            PerformanceProfiler.StopMeasure($"{input.count}_DrawNonRoiRegion_in");
                                                            // 在非ROI視覺化影像上添加IoU值的顯示
                                                            Cv2.PutText(nonRoiVisImage, $"IoU: {iou:F3}",
                                                                new Point(nonRoiRect.X, nonRoiRect.Y + nonRoiRect.Height + 20),
                                                                HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);

                                                            // 將是否為重疊區域的結果顯示在影像上
                                                            Cv2.PutText(nonRoiVisImage, $"inNROI: {inNonRoi}",
                                                                new Point(nonRoiRect.X, nonRoiRect.Y + nonRoiRect.Height + 50),
                                                                HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);
                                                            */
                                                            if (iou > GetDoubleParam(app.param, $"IOU_{input.stop}", 0.2) || inNonRoi)
                                                            {
                                                                isOverlapping = true;
                                                                break;
                                                            }
                                                
                                                        }
                                                    } // 結束 nonRoiVisImage 的 using 語句
                                                }
                                                if (!isOverlapping)
                                                {
                                                    //Console.WriteLine($"{input.count}_{defectRect.X}_{defectRect.Y}_{defectRect.Width * defectRect.Height}");
                                                    validDefects.Add(defect);
                                                }
                                    
                                            }
                                        }
                            
                            
                                        // 找出最高分的瑕疵檢測結果
                                        bool hasDefect = false;
                                        string defectName = "OK";
                                        float highestScore = 0;
                                        float threshold = 0.5f; // 默認閾值

                                        if (validDefects.Count > 0)
                                        {
                                            // 依分數從高到低排序所有有效瑕疵
                                            var sortedDefects = validDefects
                                                .OrderByDescending(d => d.score)
                                                .ToList();

                                            // 逐一檢查每個瑕疵，看是否有任何一個超過其閥值
                                            foreach (var defect in sortedDefects)
                                            {
                                                // 獲取此瑕疵類型對應的閥值
                                                float defectThreshold = 0.5f; // 預設閥值
                                                if (app.param.TryGetValue(defect.class_name + input.stop.ToString() + "_threshold", out string thresholdStr))
                                                {
                                                    float.TryParse(thresholdStr, out defectThreshold);
                                                }

                                                // 如果此瑕疵分數超過其閥值，記錄下來並跳出循環
                                                if (defect.score > defectThreshold)
                                                {
                                                    hasDefect = true;
                                                    defectName = defect.class_name;
                                                    highestScore = (float)defect.score;
                                                    threshold = defectThreshold; // 記錄此瑕疵的閥值供後續使用
                                                    break; // 找到一個超過閥值的瑕疵就停止
                                                }
                                            }

                                            // 如果沒有找到超過閥值的瑕疵，仍然記錄最高分瑕疵的資訊
                                            if (!hasDefect && sortedDefects.Count > 0)
                                            {
                                                var highestScoreDefect = sortedDefects[0];
                                                highestScore = (float)highestScoreDefect.score;
                                                defectName = highestScoreDefect.class_name;

                                                // 獲取閥值供後續使用
                                                if (app.param.TryGetValue(defectName + "_threshold", out string thresholdStr))
                                                {
                                                    float.TryParse(thresholdStr, out threshold);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // 沒有檢測到任何物件，也視為OK
                                            hasDefect = false;
                                            defectName = "OK";
                                            highestScore = 0;
                                        }
                                        // 繪製檢測結果
                                        using (Mat resultImage = _yoloDetection.DrawDetectionResults(visualizationImage, new DetectionResponse { detections = validDefects }, threshold))
                                        {
                                            // 使用現有的 showResultMat 方法顯示結果
                                            showResultMat(resultImage, input.stop);

                                            // 創建StationResult物件
                                            StationResult stationResult = new StationResult
                                            {
                                                Stop = 2, // 站點2
                                                IsNG = hasDefect,
                                                OkNgScore = highestScore > 0 ? (float?)highestScore : 0.0f,
                                                FinalMap = resultImage.Clone(), // ✅ P0-3 修正: Clone 避免 using 釋放後 FinalMap 無效
                                                DefectName = defectName, // 若有瑕疵超過閥值，則為該瑕疵名稱；否則為最高分瑕疵名稱
                                                DefectScore = highestScore > 0 ? (float?)highestScore : 0.0f, // 若有檢測到物件，保留分數
                                                OriName = Path.GetFileName(input.name)
                                            };
                            
                                            //Log.Debug($"樣品 {input.count} 第2站處理結束時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");
                                            // 添加結果到結果管理器
                                            app.resultManager.AddResult(input.count, stationResult);
                                        } // 結束 resultImage 的 using 語句
                                        if (app.enableProfiling && app.profilingStations.Contains(2))
                                        {
                                            PerformanceProfiler.StopMeasure($"{input.count}_holdResult2");
                                        }
                                    } // 結束 visualizationImage 的 using 語句

                                    // 添加到統計管理器
                                    //ScoreStatisticsManager.AddScore(1, highestScore, hasDefect, defectName);

                                    // 更新檢測率
                                    //updateLabel();
                                    //UpdateDetectionRate();

                                    #endregion
                                    if (app.enableProfiling && app.profilingStations.Contains(2))
                                    {

                                        PerformanceProfiler.StopMeasure($"{input.count}_getmat2");
                                    }
                                //PerformanceProfiler.FlushToDisk();
                                } // 由 GitHub Copilot 產生 - 結束 roi 的 try 語句
                                finally
                                {
                                    // 由 GitHub Copilot 產生
                                    // 修正: 釋放 roiInputImage 和 roi Mat (總計 30-40 MB)
                                    roiInputImage?.Dispose();
                                    roi?.Dispose();
                                }
                            }
                        /*
                        else if (app.DetectMode == 1)
                        {
                            Cv2.Resize(input.image, input.image, new Size(345, 345));
                            BeginInvoke(new Action(() => cherngerPictureBox2.Image = input.image.ToBitmap()));
                            BeginInvoke(new Action(() => cherngerPictureBox2.Refresh()));
                        }
                        */
                        } // 結束 try
                        finally
                        {
                            // 由 GitHub Copilot 產生
                            // 修正: 使用 finally 確保 input.image 一定會被釋放，即使中途 continue
                            input.image?.Dispose();
                        }
                    }
                }
                else
                {
                    app._wh2.WaitOne();
                    if (app.user == 0)
                    {
                        Console.WriteLine("gm2 on");
                    }
                }
            }
        }
        async void getMat3()
        {
            while (true)
            {
                if (app.Queue_Bitmap3.Count > 0)
                {
                    ImagePosition input;
                    app.Queue_Bitmap3.TryDequeue(out input);
                    if (app.status && input != null)
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: Receiver 已經 Clone,這裡只需檢查有效性
                        if (input.image == null || input.image.IsDisposed)
                        {
                            Log.Warning($"getMat3: input.image 已被釋放或為 null，跳過處理 (SampleID: {input.count})");
                            continue;
                        }
                        
                        try
                        {
                        if (app.DetectMode == 0)
                        {
                            #region 存原圖
                            var fname = "";
                            fname = (input.count).ToString() + "-3.jpg";
                            if (原圖ToolStripMenuItem.Checked)
                            {
                                try
                                {
                                    // 由 GitHub Copilot 產生
                                    // 緊急修正 (ObjectDisposedException): 直接 Enqueue 必須 Clone
                                    // 原因: input.image 後續還要給 findGap/DetectAndExtractROI 等函數使用
                                    // 只有透過 SaveImageAsync() 函數才會內部自動 Clone
                                    app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), @".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\origin\\" + fname));
                                    app._sv.Set();
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"存圖失敗: {ex.Message}");
                                }
                            }
                            #endregion
                            //continue;
                            #region  app.detect_result create
                            /*if (!app.detect_result.ContainsKey(input.count))
                            {
                                app.detect_result.Add(input.count, "OK");
                                app.detect_result_check.Add(input.count, new bool[4] { false, false, false, false });
                            }*/
                            #endregion

                            string sampleIdentifier = $"Sample_{input.count}";
                            if (app.enableProfiling && app.profilingStations.Contains(3))
                            {
                                PerformanceProfiler.StartMeasure($"{input.count}_getmat3");
                            }
                                //Log.Debug($"樣品 {input.count} 第3站處理開始時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");
                            #region 是否為有效取像（白比率）停用
                            /*
                            // 檢查白色像素占比 (NULL)
                            Mat whiteCheckImage = null;
                            try
                            {
                                whiteCheckImage = input.image.Clone();
                                bool isValidImage = CheckWhitePixelRatio(whiteCheckImage, input.stop);
                                if (!isValidImage)
                                {
                                    // 顯示檢測結果
                                    showResultMat(whiteCheckImage, input.stop);

                                    // 由 GitHub Copilot 產生
                                    // 修正: FinalMap 必須使用 Clone，避免與 input.image 共享同一物件
                                    // 建立檢測結果物件
                                    StationResult WhitePixelResult = new StationResult
                                    {
                                        Stop = input.stop,
                                        IsNG = false,
                                        OkNgScore = 0.0f,
                                        FinalMap = input.image.Clone(),  // 使用 Clone 避免 double-free
                                        DefectName = "NULL_invalid_White",
                                        DefectScore = 1.0f,
                                        OriName = input.name
                                    };

                                    // 添加結果到管理器
                                    app.resultManager.AddResult(input.count, WhitePixelResult);
                                    //updateLabel();
                                    continue;
                                }
                            }
                            finally
                            {
                                whiteCheckImage?.Dispose();
                            }
                            */
                            #endregion

                            #region 是否為有效取像 只在第二站做
                            /*
                            // 檢查白色像素占比 (NULL)
                            PerformanceProfiler.StartMeasure($"{input.count}_CheckWhitePixelRatio3");
                            bool isValidImage = CheckWhitePixelRatio(input.image, input.stop);
                            PerformanceProfiler.StopMeasure($"{input.count}_CheckWhitePixelRatio3");

                            if (!isValidImage)
                            {
                                // 顯示檢測結果
                                showResultMat(input.image, input.stop);

                                // 建立檢測結果物件
                                StationResult WhitePixelResult = new StationResult
                                {
                                    Stop = input.stop,
                                    IsNG = false,
                                    OkNgScore = 0.0f,
                                    FinalMap = input.image,
                                    DefectName = "NULL_invalid",
                                    DefectScore = 1.0f,
                                    OriName = input.name
                                };

                                // 添加結果到管理器
                                app.resultManager.AddResult(input.count, WhitePixelResult);
                                //updateLabel();
                                continue;
                            }
                            */
                            #endregion

                            #region 變形先不檢，目前用內圓檢，但外環鏡要用外圓檢
                            /*
                            PerformanceProfiler.StartMeasure($"{input.count}_DetectCircles3");
                            var (outerCircles, innerCircles) = DetectCircles(input.image, input.stop);
                            PerformanceProfiler.StopMeasure($"{input.count}_DetectCircles3");

                            PerformanceProfiler.StartMeasure($"{input.count}_findGapWidth3");
                            (bool gapIsNG, Mat gapResult) = findGapWidth(input.image, input.stop, outerCircles, innerCircles);
                            PerformanceProfiler.StopMeasure($"{input.count}_findGapWidth3");

                            if (gapIsNG)
                            {
                                // 顯示檢測結果
                                showResultMat(gapResult, input.stop);

                                // 建立檢測結果物件
                                StationResult stationResult = new StationResult
                                {
                                    Stop = input.stop,
                                    IsNG = gapIsNG,
                                    OkNgScore = 1.0f,
                                    FinalMap = gapResult,
                                    DefectName = "deform",
                                    DefectScore = 1.0f,
                                    OriName = input.name
                                };

                                // 添加結果到管理器
                                app.resultManager.AddResult(input.count, stationResult);
                                //updateLabel();
                                continue;
                            }
                            */
                            #endregion

                            
                            // 由 GitHub Copilot 產生
                            // 修正 P1-3: 手動管理 roi 生命週期，避免 using 在 async 期間提早釋放 (15-20 MB)

                            // 由 GitHub Copilot 產生
                            // 緊急修正: 在呼叫 DetectAndExtractROI 前檢查 input.image 狀態
                            // 注意：必須先檢查 null 和 IsDisposed，不能在已釋放的物件上呼叫 Empty()
                            if (input.image == null || input.image.IsDisposed)
                            {
                                Log.Error($"getMat3: 呼叫 DetectAndExtractROI 前 input.image 已無效 (SampleID: {input.count}, IsNull: {input.image == null}, IsDisposed: {input.image?.IsDisposed})");
                                continue;
                            }
                            
                            if (input.image.Empty())
                            {
                                Log.Error($"getMat3: 呼叫 DetectAndExtractROI 前 input.image 為空 (SampleID: {input.count})");
                                continue;
                            }

                            // 由 GitHub Copilot 產生
                            // 根本解決方案: 傳遞 Clone 副本給 DetectAndExtractROI
                            Mat roiInputImage = null;
                            Mat roi = null;
                            try
                            {
                                roiInputImage = input.image.Clone();
                                roi = DetectAndExtractROI(roiInputImage, input.stop, input.count);
                                // 由 GitHub Copilot 產生
                                // 緊急修正: 檢查 roi 是否有效，避免後續操作失敗
                                // 注意：必須先檢查 null 和 IsDisposed，不能在已釋放的物件上呼叫 Empty()
                                if (roi == null)
                                {
                                    Log.Warning($"getMat3: DetectAndExtractROI 返回 null (SampleID: {input.count})");
                                    continue;
                                }
                            
                                if (roi.IsDisposed)
                                {
                                    Log.Warning($"getMat3: DetectAndExtractROI 返回已釋放的 roi (SampleID: {input.count})");
                                    continue;
                                }
                            
                                if (roi.Empty())
                                {
                                    Log.Warning($"getMat3: DetectAndExtractROI 返回空的 roi (SampleID: {input.count})");
                                    continue;
                                }
                                List<string> defectsToDetect = GetDefectNameListForThisStop(app.produce_No, input.stop);

                                #region 找NROI
                                bool performNonRoiDetection = app.param.ContainsKey($"find_NROI_{input.stop}") && int.Parse(app.param[$"find_NROI_{input.stop}"]) == 1; //存在且是1
                                List<Rect> nonRoiRects = new List<Rect>();

                                if (performNonRoiDetection)
                                {
                                    if (app.enableProfiling && app.profilingStations.Contains(3))
                                    {
                                        PerformanceProfiler.StartMeasure($"{input.count}_nonRoiDetection3");
                                    }
                                    // 執行非 ROI 區域檢測
                                    string nonRoiServerUrl = app.produce_outer_NROI_ServerUrl;
                                    DetectionResponse nonRoiDetection = await _yoloDetection.PerformObjectDetection(roi, $"{nonRoiServerUrl}/detect");

                                    // 檢查非 ROI 檢測結果
                                    if (nonRoiDetection.detections != null && nonRoiDetection.detections.Count > 0)
                                    {
                                        foreach (var detection in nonRoiDetection.detections)
                                        {
                                            nonRoiRects.Add(new Rect(detection.box[0], detection.box[1], detection.box[2] - detection.box[0], detection.box[3] - detection.box[1]));
                                        }
                                    }
                                    if (app.enableProfiling && app.profilingStations.Contains(3))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_nonRoiDetection3");
                                    }
                                }

                                    #endregion

                                #region 小黑點AOI檢測
                                if (defectsToDetect.Contains("blackDot")) { 

                                    bool hasBlackDotsDefect = false;
                                    Mat blackDotResultImage = null;
                                    List<Point[]> blackDotContours = null;

                                    try
                                    {
                                        var blackDotResult = DetectBlackDots(roi, input.stop, input.count);
                                        hasBlackDotsDefect = blackDotResult.hasBlackDots;
                                        blackDotResultImage = blackDotResult.resultImage;
                                        blackDotContours = blackDotResult.detectedContours;

                                        int expandPixels = 5;

                                        if (hasBlackDotsDefect && performNonRoiDetection && nonRoiRects.Count > 0)
                                        {
                                            bool isOverlapping = false;

                                            // 將黑點輪廓轉換為最小矩形並檢查與NROI的重疊
                                            foreach (var contour in blackDotContours)
                                            {
                                                // 計算輪廓的最小外接矩形
                                                Rect contourRect = Cv2.BoundingRect(contour);

                                                // 檢查與每個NROI區域的重疊
                                                foreach (var nonRoiRect in nonRoiRects)
                                                {
                                                    Rect expandedNonRoiRect = ExpandRect(nonRoiRect, expandPixels, roi.Width, roi.Height);

                                                    double iou = CalculateIoU(contourRect, expandedNonRoiRect);
                                                    int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                    int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);
                                                    bool inNonRoi = IsDefectInNonRoiRegion_out(contourRect, expandedNonRoiRect, input.stop, expand_in, expand_out);

                                                    if (iou > 0 || inNonRoi)
                                                    {
                                                        isOverlapping = true;
                                                        //Log.Debug($"黑點輪廓與NROI重疊: IoU={iou:F3}, inNonRoi={inNonRoi}");
                                                        break;
                                                    }
                                                }

                                                if (isOverlapping)
                                                    break;
                                            }
                                            // 如果所有黑點都與NROI重疊，則不視為瑕疵
                                            if (isOverlapping)
                                            {
                                                hasBlackDotsDefect = false;
                                                //Log.Debug($"站{input.stop} 黑點檢測結果被NROI過濾");
                                            }
                                        }

                                        if (hasBlackDotsDefect)
                                        {
                                            // 顯示檢測結果
                                            showResultMat(blackDotResultImage, input.stop);

                                            // 建立檢測結果物件
                                            StationResult GapResult = new StationResult
                                            {
                                                Stop = input.stop,
                                                IsNG = hasBlackDotsDefect,
                                                OkNgScore = 1.0f,
                                                FinalMap = blackDotResultImage.Clone(), // ✅ P1 修正: Clone 避免原始 Mat 被釋放後 FinalMap 無效
                                                DefectName = "blackDot",
                                                DefectScore = 1.0f,
                                                OriName = input.name
                                            };

                                            // 添加結果到管理器
                                            app.resultManager.AddResult(input.count, GapResult);

                                            //updateLabel();
                                            continue;
                                        }
                                    }
                                    finally
                                    {
                                        // ✅ P1-1 修正: try-finally 確保無論 hasBlackDotsDefect 為何都會釋放
                                        blackDotResultImage?.Dispose();
                                    }

                                }
                                #endregion

                                #region 正常yolo

                                using (Mat visualizationImage = input.image.Clone())
                                {
                                    // AI 推論 OK/NG
                                    string defectServerUrl = app.produce_outerServerUrl;

                                    if (app.enableProfiling && app.profilingStations.Contains(3))
                                    {
                                        PerformanceProfiler.StartMeasure($"{input.count}_yolo_Inference3");
                                    }
                                    DetectionResponse defectDetection = await _yoloDetection.PerformObjectDetection(roi, $"{defectServerUrl}/detect");
                                    if (app.enableProfiling && app.profilingStations.Contains(3))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_yolo_Inference3");
                                    }

                                    if (defectDetection.error != null)
                                    {
                                        lbAdd($"站點3檢測錯誤: {defectDetection.error}", "err", "");
                                        continue;
                                    }

                                    // 處理瑕疵檢測結果
                                    if (app.enableProfiling && app.profilingStations.Contains(3))
                                    {
                                        PerformanceProfiler.StartMeasure($"{input.count}_holdResult3");
                                    }
                                    List<DetectionResult> validDefects = new List<DetectionResult>();
                                    if (defectDetection.detections != null && defectDetection.detections.Count > 0)
                                    {
                                        float minArea = 500.0f; // 預設閥值
                                        if (app.param.TryGetValue("minArea" + input.stop.ToString() + "_threshold", out string thresholdStr))
                                        {
                                            float.TryParse(thresholdStr, out minArea);
                                        }

                                        foreach (var defect in defectDetection.detections)
                                        {
                                            // 檢查這個瑕疵是否需要檢測
                                            if (!defectsToDetect.Contains(defect.class_name))
                                            {
                                                continue;  // 跳過不需要檢測的瑕疵
                                            }
                                            Rect defectRect = new Rect(defect.box[0], defect.box[1], defect.box[2] - defect.box[0], defect.box[3] - defect.box[1]);

                                            // 計算瑕疵面積
                                            int width = defect.box[2] - defect.box[0];
                                            int height = defect.box[3] - defect.box[1];
                                            int area = width * height;
                                            float aspectRatio = Math.Max(width, height) / (float)Math.Min(width, height);

                                            // 如果長寬比小於等於2.5且面積小於500，跳過此瑕疵
                                            if (aspectRatio <= 2.5f && area < minArea)
                                            {
                                                //continue;
                                            }

                                            // 如果執行了非 ROI 區域檢測，則檢查與非 ROI 區域的 IoU
                                            bool isOverlapping = false;
                                            if (performNonRoiDetection && nonRoiRects.Count > 0)
                                            {
                                                using (Mat nonRoiVisImage = input.image.Clone())
                                                {
                                                    foreach (var nonRoiRect in nonRoiRects)
                                                    {
                                                        double iou = CalculateIoU(defectRect, nonRoiRect);
                                                        int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                        int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);
                                                        bool inNonRoi = IsDefectInNonRoiRegion_out(defectRect, nonRoiRect, input.stop, expand_in, expand_out);

                                                        /*
                                                        // 使用DrawNonRoiRegion繪製擴展後的非ROI區域
                                                        nonRoiVisImage = DrawNonRoiRegion_out(nonRoiVisImage, nonRoiRect, input.stop, expand_in, expand_out);

                                                        // 在非ROI視覺化影像上添加IoU值的顯示
                                                        Cv2.PutText(nonRoiVisImage, $"IoU: {iou:F3}",
                                                            new Point(nonRoiRect.X, nonRoiRect.Y + nonRoiRect.Height + 20),
                                                            HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);

                                                        // 將是否為重疊區域的結果顯示在影像上
                                                        Cv2.PutText(nonRoiVisImage, $"inNROI: {inNonRoi}",
                                                            new Point(nonRoiRect.X, nonRoiRect.Y + nonRoiRect.Height + 50),
                                                            HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);
                                                        */

                                                        if (iou > GetDoubleParam(app.param, $"IOU_{input.stop}", 0.2) || inNonRoi)
                                                        {
                                                            isOverlapping = true;
                                                            break;
                                                        }                                              
                                                    }

                                                    /*
                                                    string visImgPath = $@".\image\{st.ToString("yyyy-MM")}\{st.ToString("MMdd")}\{app.foldername}\vis\{input.count}_{input.stop}_vis.jpg";
                                                    try
                                                    {
                                                        app.Queue_Save.Enqueue(new ImageSave(nonRoiVisImage, visImgPath));
                                                        app._sv.Set();
                                                        //Directory.CreateDirectory(Path.GetDirectoryName(visImgPath));
                                                        //Cv2.ImWrite(visImgPath, nonRoiVisImage);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        Console.WriteLine($"保存視覺化影像失敗: {ex.Message}");
                                                    }
                                                    */
                                                } // 結束 nonRoiVisImage 的 using 語句
                                            }

                                            // 如果沒有與非 ROI 區域高度重疊，則保留該瑕疵檢測結果
                                            if (!isOverlapping)
                                            {
                                                validDefects.Add(defect);
                                            }
                                        }
                                    }

                                    // 對 outsc 瑕疵進行過殺降低處理
                                    List<DetectionResult> processedDefects = ApplyOutscOverkillReduction(validDefects, input.stop);

                                    // 對 OTP 瑕疵進行色彩檢測處理
                                    //processedDefects = ApplyOtpColorDetection(processedDefects, roi, input.stop);

                                    // 在 getMat3() 和 getMat4() 函數中替換色彩檢測部分

                                    #region 五彩鋅色彩複檢 (針對OTP瑕疵)
                                    // 只對站3和站4的OTP瑕疵進行色彩、面積複檢
                                    string areaText = "non";
                                    string roiText = "non";
                                    string ratioText = "non";
                                    bool isAreaNG = false;
                                    if (processedDefects.Count > 0)
                                    {
                                        // 計算ROI面積 (外圓面積 - 內圓面積)
                                        int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{input.stop}"]);
                                        int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{input.stop}"]);

                                        double outerAreaPixels = Math.PI * Math.Pow(knownOuterRadius, 2);
                                        double innerAreaPixels = Math.PI * Math.Pow(knownInnerRadius, 2);
                                        double roiAreaPixels = outerAreaPixels - innerAreaPixels;

                                        // 轉換為物理單位
                                        double pixelToMmSquared = Math.Pow(double.Parse(app.param[$"PixelToMM_{input.stop}"]), 2);
                                        double roiAreaMmSquared = roiAreaPixels * pixelToMmSquared;
                                        roiText = $"ROI: {roiAreaPixels:F1} pixels)";

                                        // 分離OTP瑕疵和其他瑕疵
                                        var otpDefects = processedDefects.Where(d => d.class_name == "OTP").ToList();
                                        var otherDefects = processedDefects.Where(d => d.class_name != "OTP").ToList();

                                        // 如果有OTP瑕疵，才進行色彩複檢
                                        if (otpDefects.Count > 0)
                                        {
                                            var colorVerifier = new SimplifiedColorVerifier();

                                            // 使用簡化的3+1特徵模型進行色彩複檢
                                            // 主要特徵 (2.5σ): G通道、V通道、R通道
                                            // 輔助特徵 (3.0σ): B通道
                                            var verifiedOtpDefects = colorVerifier.VerifyDefectsByColor(otpDefects, roi, input.stop, 2.5, 5.0);

                                            if (verifiedOtpDefects.Count > 0)
                                            {
                                                double defectToRoiRatio = 0.0;
                                                // 將OTP瑕疵框轉換為矩形列表
                                                List<Rect> otpRects = verifiedOtpDefects.Select(defect =>
                                                new Rect(defect.box[0], defect.box[1],
                                                        defect.box[2] - defect.box[0],
                                                        defect.box[3] - defect.box[1])).ToList();

                                                // 使用方案二：精確計算聯合面積（處理重疊問題）
                                                int totalOtpAreaPixels = CalculateTotalAreaWithoutOverlap(otpRects);

                                                // 轉換為物理單位（如果需要）
                                                double totalOtpAreaMmSquared = totalOtpAreaPixels * pixelToMmSquared;

                                                // 計算瑕疵面積佔ROI面積的百分比
                                                defectToRoiRatio = (totalOtpAreaPixels / roiAreaPixels);

                                                // 如果需要在結果圖像上顯示總面積信息
                                                areaText = $"OTP Total Area: {totalOtpAreaPixels:F2} pixels ({verifiedOtpDefects.Count:F2} defects)";
                                                ratioText = $"defectToRoiRatio: {defectToRoiRatio:F2}";
                                                if (app.param.ContainsKey($"OTPratio_{input.stop}"))
                                                {
                                                    double otpRatioThreshold = 0.4; // 預設值
                                                    otpRatioThreshold = double.Parse(app.param[$"OTPratio_{input.stop}"]);
                                                    isAreaNG = defectToRoiRatio > otpRatioThreshold;
                                                }
                                                /*
                                                if (resultImage != null)
                                                {
                                                    // 在圖像上添加OTP總面積信息
                                                    areaText = $"OTP Total Area: {totalOtpAreaMmSquared:F1}mm² ({verifiedOtpDefects.Count} defects)";
                                                    Cv2.PutText(resultImage, areaText,
                                                               new Point(10, 100), // 避免與其他文字重疊，放在稍低的位置
                                                               HersheyFonts.HersheyDuplex, 0.6, Scalar.Cyan, 2);
                                                }
                                                */

                                                // 可選：將面積信息記錄到資料庫或檔案
                                                /*
                                                try
                                                {
                                                    string areaLogPath = $@".\logs\otp_area_log_{st.ToString("yyyy-MM")}.csv";
                                                    string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss},{input.count},{input.stop},{verifiedOtpDefects.Count},{totalOtpAreaPixels},{totalOtpAreaMmSquared:F3}";

                                                    // 如果檔案不存在，先寫入標題行
                                                    if (!File.Exists(areaLogPath))
                                                    {
                                                        File.WriteAllText(areaLogPath, "時間,樣品編號,站點,OTP瑕疵數,總面積(像素),總面積(mm²)\n");
                                                    }

                                                    File.AppendAllText(areaLogPath, logEntry + "\n");
                                                }
                                                catch (Exception ex)
                                                {
                                                    Log.Warning($"記錄OTP面積日誌失敗: {ex.Message}");
                                                }
                                                */
                                            }
                                            else
                                            {
                                                //Log.Debug($"站{input.stop} 色彩複檢後無確認OTP瑕疵，總面積為0");
                                            }

                                            // 合併驗證後的OTP瑕疵和其他瑕疵
                                            processedDefects = new List<DetectionResult>();
                                            processedDefects.AddRange(verifiedOtpDefects);
                                            processedDefects.AddRange(otherDefects);

                                            //Log.Debug($"站{input.stop} OTP色彩複檢完成，原OTP瑕疵: {otpDefects.Count}, 確認OTP瑕疵: {verifiedOtpDefects.Count}, 其他瑕疵: {otherDefects.Count}");
                                        }
                                        else
                                        {
                                            //Log.Debug($"站{input.stop} 無OTP瑕疵，跳過色彩複檢");
                                        }
                                    }
                                    else
                                    {
                                        //Log.Debug($"站{input.stop} 不進行色彩複檢或無瑕疵");
                                    }
                                    #endregion

                                    // 找出最高分的瑕疵檢測結果
                                    bool hasDefect = false;
                                    string defectName = "OK";
                                    float highestScore = 0;
                                    float threshold = 0.5f; // 默認閾值

                                    if (processedDefects.Count > 0)
                                    {
                                        // 依分數從高到低排序所有有效瑕疵
                                        var sortedDefects = processedDefects
                                            .OrderByDescending(d => d.score)
                                            .ToList();

                                        // 逐一檢查每個瑕疵，看是否有任何一個超過其閥值
                                        foreach (var defect in sortedDefects)
                                        {
                                            // 獲取此瑕疵類型對應的閥值
                                            float defectThreshold = 0.5f; // 預設閥值
                                            if (app.param.TryGetValue(defect.class_name + input.stop.ToString() + "_threshold", out string thresholdStr))
                                            {
                                                float.TryParse(thresholdStr, out defectThreshold);
                                            }

                                            // 如果此瑕疵分數超過其閥值，記錄下來並跳出循環
                                            if (defect.score > defectThreshold)
                                            {
                                                hasDefect = true;
                                                defectName = defect.class_name;
                                                highestScore = (float)defect.score;
                                                threshold = defectThreshold; // 記錄此瑕疵的閥值供後續使用
                                                break; // 找到一個超過閥值的瑕疵就停止
                                            }
                                        }

                                        // 如果沒有找到超過閥值的瑕疵，仍然記錄最高分瑕疵的資訊
                                        if (!hasDefect && sortedDefects.Count > 0)
                                        {
                                            var highestScoreDefect = sortedDefects[0];
                                            highestScore = (float)highestScoreDefect.score;
                                            defectName = highestScoreDefect.class_name;

                                            // 獲取閥值供後續使用
                                            if (app.param.TryGetValue(defectName + "_threshold", out string thresholdStr))
                                            {
                                                float.TryParse(thresholdStr, out threshold);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // 沒有檢測到任何物件，也視為OK
                                        hasDefect = false;
                                        defectName = "OK";
                                        highestScore = 0;
                                    }
                                    // 繪製檢測結果
                                    using (Mat resultImage = _yoloDetection.DrawDetectionResults(visualizationImage, new DetectionResponse { detections = processedDefects }, threshold))
                                    {
                                        if (defectName == "OTP")
                                        {
                                
                                            Cv2.PutText(resultImage, areaText,
                                                                   new Point(10, 300), // 避免與其他文字重疊，放在稍低的位置
                                                                   HersheyFonts.HersheyDuplex, 1, Scalar.Yellow, 2);
                                            Cv2.PutText(resultImage, roiText,
                                                                   new Point(10, 360), // 第二行顯示ROI信息
                                                                   HersheyFonts.HersheyDuplex, 1, Scalar.Yellow, 2);
                                            Cv2.PutText(resultImage, ratioText,
                                                                    new Point(10, 420), // 第三行顯示比例信息
                                                                    HersheyFonts.HersheyDuplex, 1, Scalar.Yellow, 2);
                                            hasDefect = isAreaNG; //如果最後驗到最高分是OTP 那就看OTP總面積是否超過閥值
                                        }
                                        // 使用現有的 showResultMat 方法顯示結果
                                        showResultMat(resultImage, input.stop);

                                        // 創建StationResult物件
                                        StationResult stationResult = new StationResult
                                        {
                                            Stop = 3, // 站點3
                                            IsNG = hasDefect,
                                            OkNgScore = highestScore > 0 ? (float?)highestScore : 0.0f,
                                            FinalMap = resultImage.Clone(), // ✅ P0-5 修正: Clone 避免 using 釋放後 FinalMap 無效
                                            DefectName = defectName, // 若有瑕疵超過閥值，則為該瑕疵名稱；否則為最高分瑕疵名稱
                                            DefectScore = highestScore > 0 ? (float?)highestScore : 0.0f, // 若有檢測到物件，保留分數
                                            OriName = Path.GetFileName(input.name)
                                        };
                                        //Log.Debug($"樣品 {input.count} 第3站處理結束時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");
                                        // 添加結果到結果管理器
                                        app.resultManager.AddResult(input.count, stationResult);
                                    } // 結束 resultImage 的 using 語句
                                    if (app.enableProfiling && app.profilingStations.Contains(3))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_holdResult3");
                                    }
                                } // 結束 visualizationImage 的 using 語句

                                // 添加到統計管理器
                                //ScoreStatisticsManager.AddScore(1, highestScore, hasDefect, defectName);

                                // 更新檢測率
                                //updateLabel();
                                //UpdateDetectionRate();

                                #endregion

                                if (app.enableProfiling && app.profilingStations.Contains(3))
                                {
                                    PerformanceProfiler.StopMeasure($"{input.count}_getmat3");
                                }
                                //var totalTime = PerformanceProfiler.StopMeasure("AI_singleProcess3");
                                //PerformanceProfiler.FlushToDisk();
                                } // 由 GitHub Copilot 產生 - 結束 roi 的 try 語句
                            finally
                            {
                                // 由 GitHub Copilot 產生
                                // 修正: 釋放 roiInputImage 和 roi Mat (總計 30-40 MB)
                                roiInputImage?.Dispose();
                                roi?.Dispose();
                            }
                        }
                        /*
                        else if (app.DetectMode == 1)
                        {
                            Cv2.Resize(input.image, input.image, new Size(345, 345));
                            BeginInvoke(new Action(() => cherngerPictureBox3.Image = input.image.ToBitmap()));
                            BeginInvoke(new Action(() => cherngerPictureBox3.Refresh()));
                        }
                        */
                        } // 結束 try
                        finally
                        {
                            // 由 GitHub Copilot 產生
                            // 修正: 使用 finally 確保 input.image 一定會被釋放，即使中途 continue
                            input.image?.Dispose();
                        }
                    }
                }
                else
                {
                    app._wh3.WaitOne();
                    if (app.user == 0)
                    {
                        Console.WriteLine("gm3 on");
                    }
                }
            }
        }
        async void getMat4()
        {
            while (true)
            {
                if (app.Queue_Bitmap4.Count > 0)
                {
                    ImagePosition input;
                    app.Queue_Bitmap4.TryDequeue(out input);

                    if (app.status && input != null)
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: Receiver 已經 Clone,這裡只需檢查有效性
                        if (input.image == null || input.image.IsDisposed)
                        {
                            Log.Warning($"getMat4: input.image 已被釋放或為 null，跳過處理 (SampleID: {input.count})");
                            continue;
                        }
                        
                        try
                        {
                        if (app.DetectMode == 0)
                        {
                            #region 存原圖
                            var fname = "";
                            fname = (input.count).ToString() + "-4.jpg";
                            if (原圖ToolStripMenuItem.Checked)
                            {
                                try
                                {
                                    // 由 GitHub Copilot 產生
                                    // 緊急修正 (ObjectDisposedException): 直接 Enqueue 必須 Clone
                                    // 原因: input.image 後續還要給 findGap/DetectAndExtractROI 等函數使用
                                    // 只有透過 SaveImageAsync() 函數才會內部自動 Clone
                                    app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), @".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\origin\\" + fname));
                                    app._sv.Set();
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"存圖失敗: {ex.Message}");
                                }
                            }
                            #endregion
                            //continue;
                            #region  app.detect_result create
                            /*if (!app.detect_result.ContainsKey(input.count))
                            {
                                app.detect_result.Add(input.count, "OK");
                                app.detect_result_check.Add(input.count, new bool[4] { false, false, false, false });
                            }*/
                            #endregion

                            // 記錄樣本ID，方便追蹤
                            string sampleIdentifier = $"Sample_{input.count}";
                            if (app.enableProfiling && app.profilingStations.Contains(4))
                            {
                                PerformanceProfiler.StartMeasure($"{input.count}_getmat4");
                            }

                            #region 是否為有效取像（白比率）停用
                            /*
                            // 檢查白色像素占比 (NULL)
                            Mat whiteCheckImage = null;
                            try
                            {
                                whiteCheckImage = input.image.Clone();
                                bool isValidImage = CheckWhitePixelRatio(whiteCheckImage, input.stop);
                                if (!isValidImage)
                                {
                                    // 顯示檢測結果
                                    showResultMat(whiteCheckImage, input.stop);

                                    // 由 GitHub Copilot 產生
                                    // 修正: FinalMap 必須使用 Clone，避免與 input.image 共享同一物件
                                    // 建立檢測結果物件
                                    StationResult WhitePixelResult = new StationResult
                                    {
                                        Stop = input.stop,
                                        IsNG = false,
                                        OkNgScore = 0.0f,
                                        FinalMap = input.image.Clone(),  // 使用 Clone 避免 double-free
                                        DefectName = "NULL_invalid_White",
                                        DefectScore = 1.0f,
                                        OriName = input.name
                                    };

                                    // 添加結果到管理器
                                    app.resultManager.AddResult(input.count, WhitePixelResult);
                                    //updateLabel();
                                    continue;
                                }
                            }
                            finally
                            {
                                whiteCheckImage?.Dispose();
                            }
                            */
                                #endregion

                            #region 找開口不做
                            /*
                            PerformanceProfiler.StartMeasure($"{input.count}_DetectCircles4");
                            //var (outerCircles, innerCircles) = DetectCircles(input.image, input.stop);
                            PerformanceProfiler.StopMeasure($"{input.count}_DetectCircles4");


                            PerformanceProfiler.StartMeasure($"{input.count}_findGapWidth4");
                            (bool gapIsNG, Mat gapResult) = findGapWidth(input.image, input.stop, outerCircles, innerCircles);
                            PerformanceProfiler.StopMeasure($"{input.count}_findGapWidth4");

                            if (gapIsNG)
                            {
                                // 顯示檢測結果
                                showResultMat(gapResult, input.stop);

                                // 建立檢測結果物件
                                StationResult stationResult = new StationResult
                                {
                                    Stop = input.stop,
                                    IsNG = gapIsNG,
                                    OkNgScore = 1.0f,
                                    FinalMap = gapResult,
                                    DefectName = "deform",
                                    DefectScore = 1.0f,
                                    OriName = input.name
                                };

                                // 添加結果到管理器
                                app.resultManager.AddResult(input.count, stationResult);
                                updateLabel();
                                continue;
                            }

                            else
                            */
                            #endregion

                            
                            // 由 GitHub Copilot 產生
                            // 修正 P1-4: 手動管理 roi 生命週期，避免 using 在 async 期間提早釋放 (15-20 MB)

                            // 由 GitHub Copilot 產生
                            // 根本修正: 傳遞 Clone 給 DetectAndExtractROI，避免 finally 執行時釋放正在使用中的影像
                            // 建立獨立生命週期的 Clone (15-20 MB)，傳遞給函數後立即釋放

                            Mat roiInputImage = null;
                            Mat roi = null;                            
                                                      
                            try
                            {
                                // 由 GitHub Copilot 產生
                                // 緊急修正: 檢查 roi 是否有效，避免後續操作失敗
                                // 注意：必須先檢查 null 和 IsDisposed，不能在已釋放的物件上呼叫 Empty()
                                roiInputImage = input.image.Clone();
                                roi = DetectAndExtractROI(roiInputImage, input.stop, input.count);
                                if (roi == null)
                                {
                                    Log.Warning($"getMat4: DetectAndExtractROI 返回 null (SampleID: {input.count})");
                                    continue;
                                }
                            
                                if (roi.IsDisposed)
                                {
                                    Log.Warning($"getMat4: DetectAndExtractROI 返回已釋放的 roi (SampleID: {input.count})");
                                    continue;
                                }
                            
                                if (roi.Empty())
                                {
                                    Log.Warning($"getMat4: DetectAndExtractROI 返回空的 roi (SampleID: {input.count})");
                                    continue;
                                }

                                #region 找NROI
                                bool performNonRoiDetection = app.param.ContainsKey($"find_NROI_{input.stop}") && int.Parse(app.param[$"find_NROI_{input.stop}"]) == 1;
                                List<Rect> nonRoiRects = new List<Rect>();

                                if (performNonRoiDetection)
                                {
                                    if (app.enableProfiling && app.profilingStations.Contains(4))
                                    {
                                        PerformanceProfiler.StartMeasure($"{input.count}_nonRoiDetection4");
                                    }
                                    // 執行非 ROI 區域檢測
                                    string nonRoiServerUrl = app.produce_outer_NROI_ServerUrl;
                                    DetectionResponse nonRoiDetection = await _yoloDetection.PerformObjectDetection(roi, $"{nonRoiServerUrl}/detect");

                                    // 檢查非 ROI 檢測結果
                                    if (nonRoiDetection.detections != null && nonRoiDetection.detections.Count > 0)
                                    {
                                        foreach (var detection in nonRoiDetection.detections)
                                        {
                                            nonRoiRects.Add(new Rect(detection.box[0], detection.box[1], detection.box[2] - detection.box[0], detection.box[3] - detection.box[1]));
                                        }
                                    }
                                    if (app.enableProfiling && app.profilingStations.Contains(4))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_nonRoiDetection4");
                                    }
                                }
                                    #endregion

                                #region 小黑點AOI檢測
                                /*
                                bool hasBlackDotsDefect = false;
                                Mat blackDotResultImage = null;
                                List<Point[]> blackDotContours = null;

                                try
                                {
                                    var blackDotResult = DetectBlackDots(roi, input.stop, input.count);
                                    hasBlackDotsDefect = blackDotResult.hasBlackDots;
                                    blackDotResultImage = blackDotResult.resultImage;
                                    blackDotContours = blackDotResult.detectedContours;

                                    int expandPixels = 5;

                                    if (hasBlackDotsDefect && performNonRoiDetection && nonRoiRects.Count > 0)
                                    {
                                        bool isOverlapping = false;

                                        // 將黑點輪廓轉換為最小矩形並檢查與NROI的重疊
                                        foreach (var contour in blackDotContours)
                                        {
                                            // 計算輪廓的最小外接矩形
                                            Rect contourRect = Cv2.BoundingRect(contour);

                                            // 檢查與每個NROI區域的重疊
                                            foreach (var nonRoiRect in nonRoiRects)
                                            {
                                                Rect expandedNonRoiRect = ExpandRect(nonRoiRect, expandPixels, roi.Width, roi.Height);

                                                double iou = CalculateIoU(contourRect, expandedNonRoiRect);
                                                int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);
                                                bool inNonRoi = IsDefectInNonRoiRegion_out(contourRect, expandedNonRoiRect, input.stop, expand_in, expand_out);

                                                if (iou > 0 || inNonRoi)
                                                {
                                                    isOverlapping = true;
                                                    //Log.Debug($"黑點輪廓與NROI重疊: IoU={iou:F3}, inNonRoi={inNonRoi}");
                                                    break;
                                                }
                                            }

                                            if (isOverlapping)
                                                break;
                                        }
                                        // 如果所有黑點都與NROI重疊，則不視為瑕疵
                                        if (isOverlapping)
                                        {
                                            hasBlackDotsDefect = false;
                                            //Log.Debug($"站{input.stop} 黑點檢測結果被NROI過濾");
                                        }
                                    }

                                    if (hasBlackDotsDefect)
                                    {
                                        // 顯示檢測結果
                                        showResultMat(blackDotResultImage, input.stop);

                                        // 建立檢測結果物件
                                        StationResult GapResult = new StationResult
                                        {
                                            Stop = input.stop,
                                            IsNG = hasBlackDotsDefect,
                                            OkNgScore = 1.0f,
                                            FinalMap = blackDotResultImage.Clone(), // ✅ P1 修正: Clone 避免原始 Mat 被釋放後 FinalMap 無效
                                            DefectName = "blackDot",
                                            DefectScore = 1.0f,
                                            OriName = input.name
                                        };

                                        // 添加結果到管理器
                                        app.resultManager.AddResult(input.count, GapResult);

                                        //updateLabel();
                                        continue;
                                    }
                                }
                                finally
                                {
                                    // ✅ P1-1 修正: try-finally 確保無論 hasBlackDotsDefect 為何都會釋放
                                    blackDotResultImage?.Dispose();
                                }
                                    */
                                #endregion

                                #region 正常yolo

                                using (Mat visualizationImage = input.image.Clone())
                                {
                                    List<string> defectsToDetect = GetDefectNameListForThisStop(app.produce_No, input.stop);

                                    // AI 推論 OK/NG
                                    string defectServerUrl = app.produce_outerServerUrl;

                                    if (app.enableProfiling && app.profilingStations.Contains(4))
                                    {
                                        PerformanceProfiler.StartMeasure($"{input.count}_yolo_Inference4");
                                    }
                                    DetectionResponse defectDetection = await _yoloDetection.PerformObjectDetection(roi, $"{defectServerUrl}/detect");
                                    if (app.enableProfiling && app.profilingStations.Contains(4))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_yolo_Inference4");
                                    }

                                    if (defectDetection.error != null)
                                    {
                                        lbAdd($"站點4檢測錯誤: {defectDetection.error}", "err", "");
                                        continue;
                                    }

                                    // 處理瑕疵檢測結果
                                    if (app.enableProfiling && app.profilingStations.Contains(4))
                                    {
                                        PerformanceProfiler.StartMeasure("holdResult4");
                                    }
                                    List<DetectionResult> validDefects = new List<DetectionResult>();
                                    if (defectDetection.detections != null && defectDetection.detections.Count > 0)
                                    {
                                        float minArea = 500.0f; // 預設閥值
                                        if (app.param.TryGetValue("minArea" + input.stop.ToString() + "_threshold", out string thresholdStr))
                                        {
                                            float.TryParse(thresholdStr, out minArea);
                                        }
                                        foreach (var defect in defectDetection.detections)
                                        {
                                            // 檢查這個瑕疵是否需要檢測
                                            if (!defectsToDetect.Contains(defect.class_name))
                                            {
                                                continue;  // 跳過不需要檢測的瑕疵
                                            }
                                            Rect defectRect = new Rect(defect.box[0], defect.box[1], defect.box[2] - defect.box[0], defect.box[3] - defect.box[1]);

                                            // 計算瑕疵面積
                                            int width = defect.box[2] - defect.box[0];
                                            int height = defect.box[3] - defect.box[1];
                                            int area = width * height;
                                            float aspectRatio = Math.Max(width, height) / (float)Math.Min(width, height);

                                            // 如果長寬比小於等於2.5且面積小於500，跳過此瑕疵
                                            if (aspectRatio <= 2.5f && area < minArea)
                                            {
                                                //continue;
                                            }

                                            // 如果執行了非 ROI 區域檢測，則檢查與非 ROI 區域的 IoU
                                            bool isOverlapping = false;
                                            if (performNonRoiDetection && nonRoiRects.Count > 0)
                                            {
                                                using (Mat nonRoiVisImage = input.image.Clone())
                                                {
                                                    foreach (var nonRoiRect in nonRoiRects)
                                                    {

                                                        double iou = CalculateIoU(defectRect, nonRoiRect);
                                                        int expand_in = int.Parse(app.param[$"expandNROI_in_{input.stop}"]);
                                                        int expand_out = int.Parse(app.param[$"expandNROI_out_{input.stop}"]);
                                                        bool inNonRoi = IsDefectInNonRoiRegion_out(defectRect, nonRoiRect, input.stop, expand_in, expand_out);

                                                        /*
                                                        // 使用DrawNonRoiRegion繪製擴展後的非ROI區域
                                                        nonRoiVisImage = DrawNonRoiRegion_out(visualizationImage.Clone(), nonRoiRect, input.stop, expand_in, expand_out);

                                                        // 在非ROI視覺化影像上添加IoU值的顯示
                                                        Cv2.PutText(nonRoiVisImage, $"IoU: {iou:F3}",
                                                            new Point(nonRoiRect.X, nonRoiRect.Y + nonRoiRect.Height + 20),
                                                            HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);

                                                        // 將是否為重疊區域的結果顯示在影像上
                                                        Cv2.PutText(nonRoiVisImage, $"inNROI: {inNonRoi}",
                                                            new Point(nonRoiRect.X, nonRoiRect.Y + nonRoiRect.Height + 50),
                                                            HersheyFonts.HersheyDuplex, 0.6, new Scalar(255, 255, 0), 1);
                                                        */

                                                        if (iou > GetDoubleParam(app.param, $"IOU_{input.stop}", 0.2) || inNonRoi)
                                                        {
                                                            isOverlapping = true;
                                                            break;
                                                        }                                              
                                                    }
                                                } // 結束 nonRoiVisImage 的 using 語句
                                            }

                                            // 如果沒有與非 ROI 區域高度重疊，則保留該瑕疵檢測結果
                                            if (!isOverlapping)
                                            {
                                                validDefects.Add(defect);
                                            }
                                        }
                                    }

                                    // 對 outsc 瑕疵進行過殺降低處理
                                    List<DetectionResult> processedDefects = ApplyOutscOverkillReduction(validDefects, input.stop);
                                    // 在 getMat3() 和 getMat4() 函數中替換色彩檢測部分

                                    #region 五彩鋅色彩複檢 (針對OTP瑕疵)
                                    // 只對站3和站4的OTP瑕疵進行色彩複檢
                                    if ((input.stop == 3 || input.stop == 4) && processedDefects.Count > 0)
                                    {
                                        // 分離OTP瑕疵和其他瑕疵
                                        var otpDefects = processedDefects.Where(d => d.class_name == "OTP").ToList();
                                        var otherDefects = processedDefects.Where(d => d.class_name != "OTP").ToList();

                                        // 如果有OTP瑕疵，才進行色彩複檢
                                        if (otpDefects.Count > 0)
                                        {
                                            var colorVerifier = new SimplifiedColorVerifier();

                                            //Log.Debug($"=== 站{input.stop} OTP瑕疵色彩複檢開始，待檢OTP瑕疵數: {otpDefects.Count} ===");

                                            // 使用簡化的3+1特徵模型進行色彩複檢
                                            // 主要特徵 (2.5σ): G通道、V通道、R通道
                                            // 輔助特徵 (3.0σ): B通道
                                            var verifiedOtpDefects = colorVerifier.VerifyDefectsByColor(otpDefects, roi, input.stop, 2.5, 3.0);

                                            // 合併驗證後的OTP瑕疵和其他瑕疵
                                            processedDefects = new List<DetectionResult>();
                                            processedDefects.AddRange(verifiedOtpDefects);
                                            processedDefects.AddRange(otherDefects);

                                            //Log.Debug($"站{input.stop} OTP色彩複檢完成，原OTP瑕疵: {otpDefects.Count}, 確認OTP瑕疵: {verifiedOtpDefects.Count}, 其他瑕疵: {otherDefects.Count}");
                                        }
                                        else
                                        {
                                            //Log.Debug($"站{input.stop} 無OTP瑕疵，跳過色彩複檢");
                                        }
                                    }
                                    else
                                    {
                                        //Log.Debug($"站{input.stop} 不進行色彩複檢或無瑕疵");
                                    }
                                    #endregion
                                    // 找出最高分的瑕疵檢測結果
                                    bool hasDefect = false;
                                    string defectName = "OK";
                                    float highestScore = 0;
                                    float threshold = 0.5f; // 默認閾值

                                    if (processedDefects.Count > 0)
                                    {
                                        // 依分數從高到低排序所有有效瑕疵
                                        var sortedDefects = processedDefects
                                            .OrderByDescending(d => d.score)
                                            .ToList();

                                        // 逐一檢查每個瑕疵，看是否有任何一個超過其閥值
                                        foreach (var defect in sortedDefects)
                                        {
                                            // 獲取此瑕疵類型對應的閥值
                                            float defectThreshold = 0.5f; // 預設閥值
                                            if (app.param.TryGetValue(defect.class_name + input.stop.ToString() + "_threshold", out string thresholdStr))
                                            {
                                                float.TryParse(thresholdStr, out defectThreshold);
                                            }

                                            // 如果此瑕疵分數超過其閥值，記錄下來並跳出循環
                                            if (defect.score > defectThreshold)
                                            {
                                                hasDefect = true;
                                                defectName = defect.class_name;
                                                highestScore = (float)defect.score;
                                                threshold = defectThreshold; // 記錄此瑕疵的閥值供後續使用
                                                break; // 找到一個超過閥值的瑕疵就停止
                                            }
                                        }

                                        // 如果沒有找到超過閥值的瑕疵，仍然記錄最高分瑕疵的資訊
                                        if (!hasDefect && sortedDefects.Count > 0)
                                        {
                                            var highestScoreDefect = sortedDefects[0];
                                            highestScore = (float)highestScoreDefect.score;
                                            defectName = highestScoreDefect.class_name;

                                            // 獲取閥值供後續使用
                                            if (app.param.TryGetValue(defectName + "_threshold", out string thresholdStr))
                                            {
                                                float.TryParse(thresholdStr, out threshold);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // 沒有檢測到任何物件，也視為OK
                                        hasDefect = false;
                                        defectName = "OK";
                                        highestScore = 0;
                                    }
                                    // 繪製檢測結果
                                    using (Mat resultImage = _yoloDetection.DrawDetectionResults(visualizationImage, new DetectionResponse { detections = processedDefects }, threshold))
                                    {
                                        // 使用現有的 showResultMat 方法顯示結果
                                        showResultMat(resultImage, input.stop);

                                        // 創建StationResult物件
                                        StationResult stationResult = new StationResult
                                        {
                                            Stop = 4, // 站點4
                                            IsNG = hasDefect,
                                            OkNgScore = highestScore > 0 ? (float?)highestScore : 0.0f,
                                            FinalMap = resultImage.Clone(), // ✅ P0-6 修正: Clone 避免 using 釋放後 FinalMap 無效
                                            DefectName = defectName, // 若有瑕疵超過閥值，則為該瑕疵名稱；否則為最高分瑕疵名稱
                                            DefectScore = highestScore > 0 ? (float?)highestScore : 0.0f, // 若有檢測到物件，保留分數
                                            OriName = Path.GetFileName(input.name)
                                        };

                                        //Log.Debug($"樣品 {input.count} 第4站處理結束時間: {DateTime.Now.ToString("HH:mm:ss.fff")}");

                                        // 添加結果到結果管理器
                                        app.resultManager.AddResult(input.count, stationResult);
                                    } // 結束 resultImage 的 using 語句
                                    if (app.enableProfiling && app.profilingStations.Contains(4))
                                    {
                                        PerformanceProfiler.StopMeasure($"{input.count}_holdResult4");
                                    }
                                } // 結束 visualizationImage 的 using 語句

                                // 添加到統計管理器
                                //ScoreStatisticsManager.AddScore(1, highestScore, hasDefect, defectName);

                                #endregion
                                if (app.enableProfiling && app.profilingStations.Contains(4))
                                {
                                    PerformanceProfiler.StopMeasure($"{input.count}_getmat4");
                                }
                                PerformanceProfiler.FlushToDisk();
                            } // 由 GitHub Copilot 產生 - 結束 roi 的 try 語句
                            finally
                            {
                                // 由 GitHub Copilot 產生
                                // 修正 P1-4: 在 finally 釋放 roiInputImage 和 roi Mat (30-40 MB)
                                roiInputImage?.Dispose();
                                roi?.Dispose();
                            }
                        }
                        /*
                        else if (app.DetectMode == 1)
                        {
                            Cv2.Resize(input.image, input.image, new Size(345, 345));
                            BeginInvoke(new Action(() => cherngerPictureBox4.Image = input.image.ToBitmap()));
                            BeginInvoke(new Action(() => cherngerPictureBox4.Refresh()));
                        }
                        */
                        } // 結束 try
                        finally
                        {
                            // 由 GitHub Copilot 產生
                            // 修正: 使用 finally 確保 input.image 一定會被釋放，即使中途 continue
                            input.image?.Dispose();
                        }
                    }
                }
                else
                {
                    app._wh4.WaitOne();
                }
            }
        }

        #endregion
        #region EXCEL
        public void save2excel(string produce_No, string LotID, DateTime st, DateTime ed, bool svdialog)
        {
            Log.Information($"save2excel 開始: produce_No={produce_No}, LotID={LotID}, svdialog={svdialog}");
            #region 路徑
            string ReportPath = @".\report\" + produce_No + "\\";
            string ReportName = ReportPath + LotID + ".xlsx";

            if (svdialog)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.FileName = LotID;
                saveFileDialog.DefaultExt = ".xlsx";
                saveFileDialog.Filter = "xlsx | *.xlsx"; // Filter files by extension

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    ReportName = saveFileDialog.FileName;
                }
            }
            #endregion
            #region parameter
            var defectcheck_count = 0;
            var defectcheck = new List<string>();
            var defectdata = new List<DefectCount>();
            var firsttime = new DateTime();
            var lasttime = new DateTime();

            #endregion
            #region copy data from DB
            Log.Information($"save2excel: 開始查詢資料庫 DefectCounts");
            using (var db = new MydbDB())
            {
                var q = from c in db.DefectCounts
                        where c.Type == produce_No && c.LotId == LotID && c.Name != "OK" && c.Name != "NG" && c.Name != "NULL" && c.Name != "SAMPLE_ID"
                        orderby c.Time
                        select c;

                Log.Information($"save2excel: DefectCounts 查詢結果數量 = {q.Count()}");
                if (q.Count() > 0)
                {
                    foreach (var item in q)
                    {
                        // 【檢查點1】檢查計數是否為負值，如果是則記錄並設為0
                        if (item.Count < 0)
                        {
                            lbAdd($"發現負數計數值: {item.Name}={item.Count}，已自動轉換為0", "war", "");
                            item.Count = 0; // 將負數替換為0，避免顯示負數
                        }

                        defectdata.Add(item);
                        if (firsttime > item.Time || firsttime == new DateTime())
                        {
                            firsttime = item.Time;
                        }

                        if (lasttime < item.Time || lasttime == new DateTime())
                        {
                            lasttime = item.Time;
                        }

                        if (!defectcheck.Contains(item.Name))
                        {
                            defectcheck.Add(item.Name);
                        }
                    }
                }
            }
            defectcheck_count = defectcheck.Count();
            st = firsttime;
            ed = lasttime;
            Log.Information($"save2excel: defectcheck_count={defectcheck_count}, firsttime={firsttime}, lasttime={lasttime}");
            #endregion
            #region 建檔
            Log.Information($"save2excel: 開始建立 Excel 檔案: {ReportName}");
            try
            {
                var di = new DirectoryInfo(ReportPath);
                if (!di.Exists)
                {
                    di.Create();
                }
                var svFile = new FileInfo(ReportName);

                XSSFWorkbook wb = new XSSFWorkbook();
                wb.CreateSheet("缺陷統計");
                wb.CreateSheet("良率統計");
                FileStream file = new FileStream(ReportName, FileMode.Create);
                wb.Write(file);
                file.Close();
                wb.Close();
                Log.Information($"save2excel: Excel 空檔案建立成功");
            }
            catch (Exception ex1)
            {
                Log.Warning($"save2excel: Excel 建檔失敗，嘗試尋找新檔名: {ex1.Message}");
                for (int i = 0; i < 9999; i++)
                {
                    if (i == 0)
                    {
                        ReportName = ReportPath + LotID + ".xlsx";
                    }
                    else
                    {
                        ReportName = ReportPath + LotID + "-" + i + ".xlsx";
                    }


                    var di = new DirectoryInfo(ReportPath);
                    if (!di.Exists)
                    {
                        di.Create();
                    }
                    var svFile = new FileInfo(ReportName);
                    if (!svFile.Exists)
                    {
                        break;
                    }
                }

                XSSFWorkbook wb = new XSSFWorkbook();
                wb.CreateSheet("缺陷統計");
                wb.CreateSheet("良率統計");
                FileStream file = new FileStream(ReportName, FileMode.Create);
                wb.Write(file);
                file.Close();
                wb.Close();

                lbAdd("報表儲存錯誤，報表將存於" + ReportName + "。", "err", "");
            }


            #endregion
            #region 設定值
            IWorkbook workbook = null;
            FileStream fs = new FileStream(ReportName, FileMode.Open, FileAccess.Read);
            workbook = new XSSFWorkbook(fs);

            ISheet sheet = workbook.GetSheetAt(0);
            IRow worksheetRow;
            ICell cell;
            IFont font = workbook.CreateFont(); ;
            font.FontName = "微軟正黑體";
            font.FontHeightInPoints = 12;
            font.IsBold = true;
            #region style

            ICellStyle Style = workbook.CreateCellStyle();
            Style.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            Style.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            Style.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            Style.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            Style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Style.SetFont(font);

            ICellStyle Style2 = workbook.CreateCellStyle();
            Style2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Style2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Style2.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            Style2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Style2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Style2.SetFont(font);

            ICellStyle Style3 = workbook.CreateCellStyle();
            Style3.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Style3.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Style3.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            Style3.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Style3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Style3.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Style3.SetFont(font);

            ICellStyle Style4 = workbook.CreateCellStyle();
            Style4.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Style4.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Style4.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            Style4.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            Style4.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Style4.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Style4.SetFont(font);
            #endregion
            #region datastyle

            ICellStyle dataStyle = workbook.CreateCellStyle();
            ICellStyle dataStyle2 = workbook.CreateCellStyle();
            ICellStyle dataStyle3 = workbook.CreateCellStyle();

            dataStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            dataStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            dataStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            dataStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.None;

            dataStyle2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            dataStyle2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            dataStyle2.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            dataStyle2.BorderRight = NPOI.SS.UserModel.BorderStyle.None;

            dataStyle3.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            dataStyle3.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            dataStyle3.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            dataStyle3.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;

            dataStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            dataStyle2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            dataStyle3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            dataStyle.SetFont(font);
            dataStyle2.SetFont(font);
            dataStyle3.SetFont(font);
            #endregion
            #endregion
            #region sheet1-缺陷統計
            #region 表單資料
            sheet.CreateRow(0);
            worksheetRow = sheet.GetRow(0);

            cell = worksheetRow.CreateCell(0);
            cell.SetCellValue("料號:");
            worksheetRow.GetCell(0).CellStyle = Style;

            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 1, 2));
            cell = worksheetRow.CreateCell(1);
            cell.SetCellValue(produce_No);
            worksheetRow.GetCell(1).CellStyle = Style;
            worksheetRow.CreateCell(2).CellStyle = Style;

            cell = worksheetRow.CreateCell(4);
            cell.SetCellValue("Lot ID");
            worksheetRow.GetCell(4).CellStyle = Style;

            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 5, 6));
            cell = worksheetRow.CreateCell(5);
            cell.SetCellValue(LotID);
            worksheetRow.GetCell(5).CellStyle = Style;
            worksheetRow.CreateCell(6).CellStyle = Style;

            sheet.CreateRow(1);
            worksheetRow = sheet.GetRow(1);
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 4));

            cell = worksheetRow.CreateCell(0);
            cell.SetCellValue("數據統計時間區間：" + st.ToString("yyyy/MM/dd HH:mm") + " 至 " + ed.ToString("yyyy/MM/dd HH:mm"));
            worksheetRow.GetCell(0).CellStyle = Style;
            worksheetRow.CreateCell(1).CellStyle = Style;
            worksheetRow.CreateCell(2).CellStyle = Style;
            worksheetRow.CreateCell(3).CellStyle = Style;
            worksheetRow.CreateCell(4).CellStyle = Style;
            #endregion
            #region Header
            sheet.CreateRow(3);
            worksheetRow = sheet.GetRow(3);
            sheet.AddMergedRegion(new CellRangeAddress(3, 3, 1, Math.Max(defectcheck_count, 6)));




            cell = worksheetRow.CreateCell(0);
            cell.SetCellValue("時間");
            worksheetRow.GetCell(0).CellStyle = Style2;

            cell = worksheetRow.CreateCell(1);
            cell.SetCellValue("缺陷檢測累計總數");
            worksheetRow.GetCell(1).CellStyle = Style2;

            for (int i = 2; i <= Math.Max(defectcheck_count, 6); i++)
            {
                if (i == Math.Max(defectcheck_count, 6))
                {
                    worksheetRow.CreateCell(i).CellStyle = Style3;
                }
                else
                {
                    worksheetRow.CreateCell(i).CellStyle = Style4;
                }
            }
            sheet.CreateRow(4);
            worksheetRow = sheet.GetRow(4);

            worksheetRow.CreateCell(0).CellStyle = Style2;
            sheet.AddMergedRegion(new CellRangeAddress(3, 4, 0, 0));
            for (int i = 0; i < defectcheck_count; i++)
            {
                cell = worksheetRow.CreateCell(i + 1);
                cell.SetCellValue(defectcheck[i]);
                worksheetRow.GetCell(i + 1).CellStyle = Style2;
            }

            #endregion
            #region 資料寫入
            sheet.CreateRow(sheet.LastRowNum + 1);
            worksheetRow = sheet.GetRow(sheet.LastRowNum);

            var move = 0;
            var lastTime = new DateTime();
            for (int i = 0; i < defectdata.Count; i += move)
            {
                move = 0;
                if (i == 0)
                {
                    cell = worksheetRow.CreateCell(0);
                    cell.SetCellValue(defectdata[0].Time.ToString("yyyy/MM/dd HH:mm"));
                    worksheetRow.GetCell(0).CellStyle = Style2;
                }

                if ((defectdata[i].Time - lastTime).TotalMinutes >= 5 || (ed - defectdata[i].Time).TotalMinutes < 1)
                {
                    lastTime = defectdata[i].Time;

                    cell.SetCellValue(defectdata[i].Time.ToString("yyyy/MM/dd HH:mm"));

                    while (true)
                    {
                        if (i + move != defectdata.Count)
                        {
                            if ((defectdata[i + move].Time - lastTime).TotalMinutes < 1)
                            {
                                for (int j = 0; j < defectcheck_count; j++)
                                {
                                    if (defectcheck[j] == defectdata[i + move].Name)
                                    {
                                        cell = worksheetRow.CreateCell(j + 1);
                                        cell.SetCellValue(defectdata[i + move].Count);
                                        worksheetRow.GetCell(j + 1).CellStyle = Style2;

                                        move++;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (i + move != defectdata.Count)
                    {
                        sheet.CreateRow(sheet.LastRowNum + 1);
                        worksheetRow = sheet.GetRow(sheet.LastRowNum);
                        cell = worksheetRow.CreateCell(0);
                        //cell.SetCellValue(defectdata[i + move].Time.ToString("yyyy/MM/dd HH:mm"));
                        worksheetRow.GetCell(0).CellStyle = Style2;
                    }
                }
                else
                {
                    move = 1;
                }
            }

            sheet.SetColumnWidth(0, (int)((20 + 0.71) * 256));
            for (int i = 1; i < Math.Max(defectcheck_count + 1, 7); i++)
            {
                sheet.SetColumnWidth(i, (int)((15 + 0.71) * 256));
            }

            #endregion
            #endregion
            #region sheet2-良率統計
            Log.Information($"save2excel: 開始處理良率統計工作表");
            sheet = workbook.GetSheetAt(1);
            #region copy data from DB
            defectdata = new List<DefectCount>();
            using (var db = new MydbDB())
            {
                var q = from c in db.DefectCounts
                        where c.Type == produce_No && c.LotId == LotID && (c.Name == "OK" || c.Name == "NG" || c.Name == "NULL")
                        orderby c.Time
                        select c;

                Log.Information($"save2excel: 良率統計查詢結果數量 = {q.Count()}");
                if (q.Count() > 0)
                {
                    foreach (var item in q)
                    {
                        defectdata.Add(item);
                    }
                }
            }
            #endregion
            #region 表單資料
            sheet.CreateRow(0);
            worksheetRow = sheet.GetRow(0);

            cell = worksheetRow.CreateCell(0);
            cell.SetCellValue("料號:");
            worksheetRow.GetCell(0).CellStyle = Style;

            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 1, 2));
            cell = worksheetRow.CreateCell(1);
            cell.SetCellValue(produce_No);
            worksheetRow.GetCell(1).CellStyle = Style;
            worksheetRow.CreateCell(2).CellStyle = Style;

            cell = worksheetRow.CreateCell(4);
            cell.SetCellValue("Lot ID");
            worksheetRow.GetCell(4).CellStyle = Style;

            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 5, 6));
            cell = worksheetRow.CreateCell(5);
            cell.SetCellValue(LotID);
            worksheetRow.GetCell(5).CellStyle = Style;
            worksheetRow.CreateCell(6).CellStyle = Style;

            sheet.CreateRow(1);
            worksheetRow = sheet.GetRow(1);
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 4));

            cell = worksheetRow.CreateCell(0);
            cell.SetCellValue("數據統計時間區間：" + st.ToString("yyyy/MM/dd HH:mm") + " 至 " + ed.ToString("yyyy/MM/dd HH:mm"));
            worksheetRow.GetCell(0).CellStyle = Style;
            worksheetRow.CreateCell(1).CellStyle = Style;
            worksheetRow.CreateCell(2).CellStyle = Style;
            worksheetRow.CreateCell(3).CellStyle = Style;
            worksheetRow.CreateCell(4).CellStyle = Style;
            #endregion
            #region Header
            sheet.CreateRow(3);
            worksheetRow = sheet.GetRow(3);
            sheet.AddMergedRegion(new CellRangeAddress(3, 3, 1, 3));
            sheet.AddMergedRegion(new CellRangeAddress(3, 3, 4, 6));


            cell = worksheetRow.CreateCell(0);
            cell.SetCellValue("時間");
            worksheetRow.GetCell(0).CellStyle = Style2;

            cell = worksheetRow.CreateCell(1);
            cell.SetCellValue("成品檢測累計總數");
            worksheetRow.GetCell(1).CellStyle = Style2;
            worksheetRow.CreateCell(2).CellStyle = Style4;
            worksheetRow.CreateCell(3).CellStyle = Style3;

            cell = worksheetRow.CreateCell(4);
            cell.SetCellValue("成品檢測比例");
            worksheetRow.GetCell(4).CellStyle = Style2;
            worksheetRow.CreateCell(5).CellStyle = Style4;
            worksheetRow.CreateCell(6).CellStyle = Style3;

            sheet.CreateRow(4);
            worksheetRow = sheet.GetRow(4);

            worksheetRow.CreateCell(0).CellStyle = Style2;
            sheet.AddMergedRegion(new CellRangeAddress(3, 4, 0, 0));

            cell = worksheetRow.CreateCell(1);
            cell.SetCellValue("OK");
            worksheetRow.GetCell(1).CellStyle = Style2;
            cell = worksheetRow.CreateCell(2);
            cell.SetCellValue("NG");
            worksheetRow.GetCell(2).CellStyle = Style2;
            cell = worksheetRow.CreateCell(3);
            cell.SetCellValue("NULL");
            worksheetRow.GetCell(3).CellStyle = Style2;
            cell = worksheetRow.CreateCell(4);
            cell.SetCellValue("OK Rate");
            worksheetRow.GetCell(4).CellStyle = Style2;
            cell = worksheetRow.CreateCell(5);
            cell.SetCellValue("NG Rate");
            worksheetRow.GetCell(5).CellStyle = Style2;
            cell = worksheetRow.CreateCell(6);
            cell.SetCellValue("NULL Rate");
            worksheetRow.GetCell(6).CellStyle = Style2;

            #endregion
            #region 資料寫入
            sheet.CreateRow(sheet.LastRowNum + 1);
            worksheetRow = sheet.GetRow(sheet.LastRowNum);
            lastTime = new DateTime();

            for (int i = 0; i < defectdata.Count; i += move)
            {
                move = 0;
                var total = 0.0;
                if (i == 0)
                {
                    cell = worksheetRow.CreateCell(0);
                    cell.SetCellValue(defectdata[0].Time.ToString("yyyy/MM/dd HH:mm"));
                    worksheetRow.GetCell(0).CellStyle = Style2;
                }
                double[] count = new double[3];
                string[] name = { "OK", "NG", "NULL" };

                if ((defectdata[i].Time - lastTime).TotalMinutes >= 5 || (ed - defectdata[i].Time).TotalMinutes < 1)
                {
                    lastTime = defectdata[i].Time;

                    cell.SetCellValue(defectdata[i].Time.ToString("yyyy/MM/dd HH:mm"));

                    while (true)
                    {
                        if (i + move != defectdata.Count)
                        {
                            if ((defectdata[i + move].Time - lastTime).TotalMinutes < 1)
                            {
                                for (int j = 0; j < name.Length; j++)
                                {
                                    if (name[j] == defectdata[i + move].Name)
                                    {
                                        count[j] = defectdata[i + move].Count;

                                        cell = worksheetRow.CreateCell(j + 1);
                                        cell.SetCellValue(defectdata[i + move].Count);
                                        worksheetRow.GetCell(j + 1).CellStyle = Style2;
                                        move++;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    total = count[0] + count[1] + count[2];
                    if (total == 0)
                    {
                        cell = worksheetRow.CreateCell(4);
                        cell.SetCellValue("0.0%");
                        worksheetRow.GetCell(4).CellStyle = Style2;
                        cell = worksheetRow.CreateCell(5);
                        cell.SetCellValue("0.0%");
                        worksheetRow.GetCell(5).CellStyle = Style2;
                        cell = worksheetRow.CreateCell(6);
                        cell.SetCellValue("0.0%");
                        worksheetRow.GetCell(6).CellStyle = Style2;
                    }
                    else
                    {
                        cell = worksheetRow.CreateCell(4);
                        cell.SetCellValue(double.Parse((count[0] * 100 / total).ToString("f1")) + "%");
                        worksheetRow.GetCell(4).CellStyle = Style2;
                        cell = worksheetRow.CreateCell(5);
                        cell.SetCellValue(double.Parse((count[1] * 100 / total).ToString("f1")) + "%");
                        worksheetRow.GetCell(5).CellStyle = Style2;
                        cell = worksheetRow.CreateCell(6);
                        cell.SetCellValue(double.Parse((count[2] * 100 / total).ToString("f1")) + "%");
                        worksheetRow.GetCell(6).CellStyle = Style2;
                    }
                    if (i + move != defectdata.Count)
                    {
                        sheet.CreateRow(sheet.LastRowNum + 1);
                        worksheetRow = sheet.GetRow(sheet.LastRowNum);
                        cell = worksheetRow.CreateCell(0);
                        //cell.SetCellValue(defectdata[i + move].Time.ToString("yyyy/MM/dd HH:mm"));
                        worksheetRow.GetCell(0).CellStyle = Style2;
                    }
                }
                else
                {
                    move = 1;
                }
            }

            sheet.SetColumnWidth(0, (int)((20 + 0.71) * 256));
            for (int i = 1; i < 7; i++)
            {
                sheet.SetColumnWidth(i, (int)((15 + 0.71) * 256));
            }


            #endregion
            #endregion
            #region 寫檔
            Log.Information($"save2excel: 開始寫入最終檔案: {ReportName}");
            try
            {
                var di = new DirectoryInfo(ReportPath);
                if (!di.Exists)
                {
                    di.Create();
                }
                fs = new FileStream(ReportName, FileMode.Create);

                workbook.Write(fs);
                fs.Close();
                fs = null;
                Log.Information($"save2excel: 報表儲存成功: {ReportName}");
            }
            catch (Exception ex2)
            {
                Log.Warning($"save2excel: 最終寫檔失敗，嘗試尋找新檔名: {ex2.Message}");
                for (int i = 0; i < 9999; i++)
                {
                    ReportPath = @".\report\" + produce_No + "\\";
                    if (i == 0)
                    {
                        ReportName = ReportPath + LotID + ".xlsx";
                    }
                    else
                    {
                        ReportName = ReportPath + LotID + "-" + i + ".xlsx";
                    }

                    var di = new DirectoryInfo(ReportPath);
                    if (!di.Exists)
                    {
                        di.Create();
                    }
                    var fi = new FileInfo(ReportName);
                    if (!fi.Exists)
                    {
                        break;
                    }
                }
                fs = new FileStream(ReportName, FileMode.Create);

                workbook.Write(fs);
                fs.Close();
                fs = null;
                lbAdd("報表儲存錯誤，報表將存於" + ReportName + "。", "err", "");
            }
            #endregion
        }
        #endregion
        #region 介面事件
        #region uiControl
        #region combobox繪製
        private void comboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                e.DrawBackground();
                if (e.Index >= 0)
                {
                    //文字置中
                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;

                    Brush brush = new SolidBrush(cbx.ForeColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;

                    //重繪字串
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox5.Items.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectChecks
                    where c.Yn == 1 && c.Type == comboBox4.Text
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (!comboBox5.Items.Contains(c.Name)) // 檢查是否已存在
                        {
                            comboBox5.Items.Add(c.Name);
                        }
                    }
                }
            }
            if (comboBox5.Items.Count > 0)
            {
                comboBox5.SelectedIndex = 0;
            }

            comboBox7.Items.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectCounts
                    where c.Type == comboBox4.Text
                    orderby c.LotId
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (!comboBox7.Items.Contains(c.LotId))
                        {
                            comboBox7.Items.Add(c.LotId);
                        }
                    }
                }
            }
            if (comboBox7.Items.Count > 0)
            {
                comboBox7.SelectedIndex = 0;
            }
        }

        #endregion
        #region 執行中鎖UI
        void uiLock(bool status)
        {
            if (app.DetectMode == 0)
            {
                Invoke(new Action(() => button6.Enabled = status));    //料號設定
                Invoke(new Action(() => button8.Enabled = status));    //權限設定
                Invoke(new Action(() => button9.Enabled = status));    //LOT ID
                Invoke(new Action(() => button10.Enabled = status));    //工單數量
                
                Invoke(new Action(() => button11.Enabled = status));    //OK1
                Invoke(new Action(() => button42.Enabled = status));    //OK2
                Invoke(new Action(() => button43.Enabled = status));    //OK
                Invoke(new Action(() => button13.Enabled = status));    //NG
                Invoke(new Action(() => button14.Enabled = status));    //NG籃
                Invoke(new Action(() => button15.Enabled = status));    //NULL
                Invoke(new Action(() => button41.Enabled = status));    //NULL籃
                
                Invoke(new Action(() => button16.Enabled = status));    //包裝數量
                Invoke(new Action(() => button17.Enabled = status));    //異常復歸

                Invoke(new Action(() => button39.Enabled = status));    //NG籃計數
                Invoke(new Action(() => button40.Enabled = status));    //NULL籃計數
                Invoke(new Action(() => button47.Enabled = status));    //更新面板
                Invoke(new Action(() => button49.Enabled = status));    //全進OK
            }

            Invoke(new Action(() => button1.Enabled = status));     //開始
            Invoke(new Action(() => button2.Enabled = !status));    //停止
        }

        #endregion
        #region Listbox.Add
        void lbAdd(string s, string logType, string err_message)    //  資訊欄&&LOG紀錄
        {
            if (logType == "inf")
            {
                Log.Information(s);
            }
            else if (logType == "err")
            {
                BeginInvoke(new Action(() => listBox1.Items.Add(DateTime.Now.ToString() + "---" + s)));
                BeginInvoke(new Action(() => listBox1.TopIndex = listBox1.Items.Count - 1));
                Log.Error(s);
                Log.Error(err_message);
            }
            else if (logType == "war")
            {
                Log.Warning(s);
            }
        }
        #endregion
        #region 每日資料夾建立
        void dailyCheck()
        {
            if (app.foldername != "")
            {
                if (原圖ToolStripMenuItem.Checked || oKToolStripMenuItem.Checked || nGToolStripMenuItem.Checked)
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd"), false, true);
                if (原圖ToolStripMenuItem.Checked)
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\origin", false, true);
                if (oKToolStripMenuItem.Checked)
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\OK", false, true);
                if (nGToolStripMenuItem.Checked)
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\NG", false, true);
                if (nULLToolStripMenuItem.Checked)
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\NULL", false, true);
            }
        }
        #endregion        
        #region 計數面板刷新
        void updateLabel()
        {
            if (!app.offline)
            {              
                var handle1 = new ManualResetEvent(false);
                PLC_ModBus.CheckValue(1, ValueUnit.D, 801, 14, false, handle1);
                handle1.WaitOne();
                var d801 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 801);
                var d803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803);
                var d805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805);
                var d807 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 807);
                var d809 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 809);
                //var d809 = PLC_CheckD(809);
                var d811 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 811);
                //var d811 = PLC_CheckD(811);
                var d813 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 813);
                //var d813 = PLC_CheckD(813);

                var ngCount = d801;
                var ok1Count = d803;
                var ok2Count = d805;
                var okCount = d807;
                var nullCount = d809;

                var ngBox = d811;
                var nullBox = d813;

                var totalCount = ngCount + okCount + nullCount;

                BeginInvoke(new Action(() => label3.Text = ok1Count.ToString()));
                BeginInvoke(new Action(() => label58.Text = ok2Count.ToString()));
                BeginInvoke(new Action(() => label59.Text = okCount.ToString()));

                BeginInvoke(new Action(() => label6.Text = ngCount.ToString()));
                BeginInvoke(new Action(() => label26.Text = ngBox.ToString()));

                BeginInvoke(new Action(() => label14.Text = nullCount.ToString()));
                BeginInvoke(new Action(() => label55.Text = nullBox.ToString()));

                BeginInvoke(new Action(() => label63.Text = totalCount.ToString()));

                if (label59.Text == "0" && label6.Text == "0" && label14.Text == "0")
                {
                    BeginInvoke(new Action(() => label60.Text = "0.0%"));
                    BeginInvoke(new Action(() => label23.Text = "0.0%"));
                    BeginInvoke(new Action(() => label17.Text = "0.0%"));
                }
                else
                {
                    BeginInvoke(new Action(() => label60.Text = ((double)okCount * 100 / totalCount).ToString("f1") + "%"));
                    BeginInvoke(new Action(() => label23.Text = ((double)ngCount * 100 / totalCount).ToString("f1") + "%"));
                    BeginInvoke(new Action(() => label17.Text = ((double)nullCount * 100 / totalCount).ToString("f1") + "%"));
                }
                if ((DateTime.Now - lastPlcLogTime).TotalMilliseconds >= 500)
                {
                    Log.Information($"[PLC監控] {DateTime.Now:HH:mm:ss.fff},          D803(OK1):{d803}, D805(OK2):{d805}, D807(OK):{d807}, D801(NG):{d801}, D809(NULL):{d809}");
                    lastPlcLogTime = DateTime.Now;
                }
            }

            Thread.Sleep(20);
        }
        #endregion
        #endregion

        #endregion
        #region AI
        async void setNet()
        {

            try
            {
                // 重置總站台數
                ResultManager.totalStations = 0;

                // 模型目錄路徑
                string modelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "models");

                // 批次檔路徑
                string batFilePath = "start_server.bat";

                // 為內環和外環設定不同的伺服器URL
                string innerServerUrl = "http://localhost:5001";
                string inner_NROI_ServerUrl = "http://localhost:5003";
                string outerServerUrl = "http://localhost:5002";
                string outer_NROI_ServerUrl = "http://localhost:5004";
                string chamfer_ServerUrl = "http://localhost:5005";

                // 新增站1和站2的伺服器URL
                string station1_ServerUrl = "http://localhost:5006";
                string station2_ServerUrl = "http://localhost:5007";

                // 構建內環和外環模型路徑
                string innerModelName = app.produce_No + "_in.pt";
                string outerModelName = app.produce_No + "_out.pt";
                string inner_NROI_ModelName = app.produce_No + "_in_NROI.pt";
                string outer_NROI_ModelName = app.produce_No + "_out_NROI.pt";
                string chamfer_ModelName = app.produce_No + "_chamfer.pt";

                // 新增站1和站2模型名稱
                string station1_ModelName = app.produce_No + "_1.pt";
                string station2_ModelName = app.produce_No + "_2.pt";

                string innerModelPath = Path.Combine(modelPath, innerModelName);
                string outerModelPath = Path.Combine(modelPath, outerModelName);
                string inner_NROI_ModelPath = Path.Combine(modelPath, inner_NROI_ModelName);
                string outer_NROI_ModelPath = Path.Combine(modelPath, outer_NROI_ModelName);
                string chamfer_ModelPath = Path.Combine(modelPath, chamfer_ModelName);

                // 新增站1和站2模型路徑
                string station1_ModelPath = Path.Combine(modelPath, station1_ModelName);
                string station2_ModelPath = Path.Combine(modelPath, station2_ModelName);

                // 檢查模型文件是否存在
                bool hasInnerModel = File.Exists(innerModelPath);
                bool hasOuterModel = File.Exists(outerModelPath);
                app.has_NROI_InnerModel = File.Exists(inner_NROI_ModelPath);
                app.has_NROI_OuterModel = File.Exists(outer_NROI_ModelPath);
                bool haschamferModel = File.Exists(chamfer_ModelPath);

                // 檢查站1和站2模型是否存在
                bool hasStation1Model = File.Exists(station1_ModelPath);
                bool hasStation2Model = File.Exists(station2_ModelPath);

                //Console.WriteLine(innerModelName);
                //Console.WriteLine(hasInnerModel);
                // 確保已經初始化YOLO檢測器
                if (_yoloDetection == null)
                {
                    _yoloDetection = new YoloDetection();
                }

                // 初始化伺服器和載入模型
                if (hasInnerModel)
                {
                    lbAdd($"[setNet] 正在啟動內環伺服器並載入模型: {innerModelName}", "inf", "");

                    // 啟動伺服器並檢查連線
                    bool innerServerReady = await _yoloDetection.ServerOn(innerServerUrl, batFilePath, innerModelPath);

                    if (innerServerReady)
                    {
                        // 載入內環模型
                        var loadResult = await _yoloDetection.LoadYoloModel(innerModelPath, innerServerUrl);

                        if (loadResult.error == null)
                        {
                            // 成功載入
                            app.produce_innerModelPath = innerModelPath;
                            app.produce_innerServerUrl = innerServerUrl; // 使用URL替代json路徑
                            lbAdd($"[setNet] 成功載入內環模型: {innerModelName}", "inf", "");
                            ResultManager.totalStations += 2;

                            await WarmUpYoloModel(_yoloDetection, innerServerUrl, 2448, 2048); // 依你的模型輸入尺寸
                        }
                        else
                        {
                            lbAdd($"[setNet] 載入內環模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到內環伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到內環模型文件: {innerModelPath}", "war", "");
                }

                if (hasOuterModel)
                {
                    lbAdd($"[setNet] 正在啟動外環伺服器並載入模型: {outerModelName}", "inf", "");

                    // 啟動伺服器並檢查連線
                    bool outerServerReady = await _yoloDetection.ServerOn(outerServerUrl, batFilePath, outerModelPath);

                    if (outerServerReady)
                    {
                        // 載入外環模型
                        var loadResult = await _yoloDetection.LoadYoloModel(outerModelPath, outerServerUrl);

                        if (loadResult.error == null)
                        {
                            // 成功載入
                            app.produce_outerModelPath = outerModelPath;
                            app.produce_outerServerUrl = outerServerUrl; // 使用URL替代json路徑
                            ResultManager.totalStations += 2;

                            lbAdd($"[setNet] 成功載入外環模型: {outerModelName}", "inf", "");

                            await WarmUpYoloModel(_yoloDetection, outerServerUrl, 2448, 2048); // 依你的模型輸入尺寸
                        }
                        else
                        {
                            lbAdd($"[setNet] 載入外環模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到外環伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到外環模型文件: {outerModelPath}", "war", "");
                }
                // 初始化內環非ROI伺服器和載入模型
                if (app.has_NROI_InnerModel)
                {
                    lbAdd($"[setNet] 正在啟動內環非ROI伺服器並載入模型: {inner_NROI_ModelName}", "inf", "");

                    bool innerNROIServerReady = await _yoloDetection.ServerOn(inner_NROI_ServerUrl, batFilePath, inner_NROI_ModelPath);

                    if (innerNROIServerReady)
                    {
                        var loadResult = await _yoloDetection.LoadYoloModel(inner_NROI_ModelPath, inner_NROI_ServerUrl);

                        if (loadResult.error == null)
                        {
                            app.produce_inner_NROI_ServerUrl = inner_NROI_ServerUrl;
                            lbAdd($"[setNet] 成功載入內環非ROI模型: {inner_NROI_ModelName}", "inf", "");

                            await WarmUpYoloModel(_yoloDetection, inner_NROI_ServerUrl, 2448, 2048); // 依你的模型輸入尺寸

                        }
                        else
                        {
                            lbAdd($"[setNet] 載入內環非ROI模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到內環非ROI伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到內環非ROI模型文件: {inner_NROI_ModelPath}", "war", "");
                }

                // 初始化外環非ROI伺服器和載入模型
                if (app.has_NROI_OuterModel)
                {
                    lbAdd($"[setNet] 正在啟動外環非ROI伺服器並載入模型: {outer_NROI_ModelName}", "inf", "");

                    bool outerNROIServerReady = await _yoloDetection.ServerOn(outer_NROI_ServerUrl, batFilePath, outer_NROI_ModelPath);

                    if (outerNROIServerReady)
                    {
                        var loadResult = await _yoloDetection.LoadYoloModel(outer_NROI_ModelPath, outer_NROI_ServerUrl);

                        if (loadResult.error == null)
                        {
                            app.produce_outer_NROI_ServerUrl = outer_NROI_ServerUrl;
                            lbAdd($"[setNet] 成功載入外環非ROI模型: {outer_NROI_ModelName}", "inf", "");

                            await WarmUpYoloModel(_yoloDetection, outer_NROI_ServerUrl, 2448, 2048); // 依你的模型輸入尺寸

                        }
                        else
                        {
                            lbAdd($"[setNet] 載入外環非ROI模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到外環非ROI伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到外環非ROI模型文件: {outer_NROI_ModelPath}", "war", "");
                }

                if (haschamferModel)
                {
                    lbAdd($"[setNet] 正在啟動倒角伺服器並載入模型: {chamfer_ModelName}", "inf", "");

                    bool chamferServerReady = await _yoloDetection.ServerOn(chamfer_ServerUrl, batFilePath, chamfer_ModelPath);

                    if (chamferServerReady)
                    {
                        var loadResult = await _yoloDetection.LoadYoloModel(chamfer_ModelPath, chamfer_ServerUrl);

                        if (loadResult.error == null)
                        {
                            app.produce_chamferServerUrl = chamfer_ServerUrl;
                            lbAdd($"[setNet] 成功載入倒角模型: {chamfer_ModelName}", "inf", "");

                            await WarmUpYoloModel(_yoloDetection, chamfer_ServerUrl, 2448, 2048); // 依你的模型輸入尺寸

                        }
                        else
                        {
                            lbAdd($"[setNet] 載入倒角模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到倒角伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到倒角模型文件: {chamfer_ModelPath}", "war", "");
                }

                // 新增站1模型初始化
                if (hasStation1Model)
                {
                    lbAdd($"[setNet] 正在啟動站1伺服器並載入模型: {station1_ModelName}", "inf", "");

                    bool station1ServerReady = await _yoloDetection.ServerOn(station1_ServerUrl, batFilePath, station1_ModelPath);

                    if (station1ServerReady)
                    {
                        var loadResult = await _yoloDetection.LoadYoloModel(station1_ModelPath, station1_ServerUrl);

                        if (loadResult.error == null)
                        {
                            app.produce_station1ServerUrl = station1_ServerUrl;
                            lbAdd($"[setNet] 成功載入站1模型: {station1_ModelName}", "inf", "");
                            ResultManager.totalStations += 1;

                            await WarmUpYoloModel(_yoloDetection, station1_ServerUrl, 2448, 2048);
                        }
                        else
                        {
                            lbAdd($"[setNet] 載入站1模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到站1伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到站1模型文件: {station1_ModelPath}", "war", "");
                }

                // 新增站2模型初始化
                if (hasStation2Model)
                {
                    lbAdd($"[setNet] 正在啟動站2伺服器並載入模型: {station2_ModelName}", "inf", "");

                    bool station2ServerReady = await _yoloDetection.ServerOn(station2_ServerUrl, batFilePath, station2_ModelPath);

                    if (station2ServerReady)
                    {
                        var loadResult = await _yoloDetection.LoadYoloModel(station2_ModelPath, station2_ServerUrl);

                        if (loadResult.error == null)
                        {
                            app.produce_station2ServerUrl = station2_ServerUrl;
                            lbAdd($"[setNet] 成功載入站2模型: {station2_ModelName}", "inf", "");
                            ResultManager.totalStations += 1;

                            await WarmUpYoloModel(_yoloDetection, station2_ServerUrl, 2448, 2048);
                        }
                        else
                        {
                            lbAdd($"[setNet] 載入站2模型失敗: {loadResult.error}", "err", "");
                        }
                    }
                    else
                    {
                        lbAdd($"[setNet] 無法連接到站2伺服器，請檢查設定。", "err", "");
                    }
                }
                else
                {
                    lbAdd($"[setNet] 找不到站2模型文件: {station2_ModelPath}", "war", "");
                }

                Console.WriteLine($"[setNet] 已經為 {app.produce_No} 完成模型初始化");
                Console.WriteLine($"目前設置的總站台數： {ResultManager.totalStations}");
                lbAdd($"目前設置的總站台數： {ResultManager.totalStations}", "inf", "");

                Console.WriteLine("開始執行預熱功能");
                // 在所有模型載入完成後，啟動持續預熱
                if (hasInnerModel && !string.IsNullOrEmpty(app.produce_innerServerUrl))
                {
                    // 啟動持續預熱 (每2秒)
                    _yoloDetection.StartContinuousWarmup(app.produce_innerServerUrl, 2);
                }

                if (hasOuterModel && !string.IsNullOrEmpty(app.produce_outerServerUrl))
                {
                    _yoloDetection.StartContinuousWarmup(app.produce_outerServerUrl, 2);
                }

                if (app.has_NROI_InnerModel && !string.IsNullOrEmpty(app.produce_inner_NROI_ServerUrl))
                {
                    _yoloDetection.StartContinuousWarmup(app.produce_inner_NROI_ServerUrl, 2);
                }

                if (app.has_NROI_OuterModel && !string.IsNullOrEmpty(app.produce_outer_NROI_ServerUrl))
                {
                    _yoloDetection.StartContinuousWarmup(app.produce_outer_NROI_ServerUrl, 2);
                }

                if (haschamferModel && !string.IsNullOrEmpty(app.produce_chamferServerUrl))
                {
                    _yoloDetection.StartContinuousWarmup(app.produce_chamferServerUrl, 2);
                }

                if (hasStation1Model && !string.IsNullOrEmpty(app.produce_station1ServerUrl))
                {
                    _yoloDetection.StartContinuousWarmup(app.produce_station1ServerUrl, 2);
                }

                if (hasStation2Model && !string.IsNullOrEmpty(app.produce_station2ServerUrl))
                {
                    _yoloDetection.StartContinuousWarmup(app.produce_station2ServerUrl, 2);
                }

                lbAdd("[setNet] 已啟動所有模型的持續預熱功能", "inf", "");

            }
            catch (Exception e1)
            {
                lbAdd("AI Model載入失敗", "err", e1.ToString());
            }   
            
        }

        private async Task WarmUpYoloModel(YoloDetection yolo, string serverUrl, int width, int height, int count = 3)
        {
            for (int i = 0; i < count; i++)
            {
                using (var dummyMat = new Mat(new Size(width, height), MatType.CV_8UC3, Scalar.Black))
                {
                    try
                    {
                        await yolo.PerformObjectDetection(dummyMat, $"{serverUrl}/detect");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"模型預熱第{i + 1}張失敗: {ex.Message}");
                    }
                }
            }
        }


        #endregion
        #region set
        void setDictionary()    //  資料建置
        {
            #region Counter            
            var counterType = new[] { "OK1", "OK2", "OK", "NG", "NULL", "stop0", "stop1", "stop2", "stop3" };
            foreach (var s in counterType)
            {
                app.counter.TryAdd(s, 0);
            }

            #endregion

            #region _wh, Queue, Task                        
            app.T1 = new Task(() => getMat1(), TaskCreationOptions.LongRunning);
            app.T2 = new Task(() => getMat2(), TaskCreationOptions.LongRunning);
            app.T3 = new Task(() => getMat3(), TaskCreationOptions.LongRunning);
            app.T4 = new Task(() => getMat4(), TaskCreationOptions.LongRunning);
            app.Tsv = new Task(() => sv(), TaskCreationOptions.LongRunning);
            app.Treader = new Task(() => read_plc(), TaskCreationOptions.LongRunning);
            //app.TAI = new Task(() => AI(), TaskCreationOptions.LongRunning);
            //app.Tshow = new Task(() => showMat(), TaskCreationOptions.LongRunning);

            app.T1.Start();
            app.T2.Start();
            app.T3.Start();
            app.T4.Start();
            app.Tsv.Start();
            app.Treader.Start();
            //app.TAI.Start();
            //app.Tshow.Start();
            #endregion
            #region param
            ReadParam();
            #endregion
        }
        void ReadParam()    //  sqlite資料讀取
        {
            try
            {
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.Totals
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            //label39.Text = c.Num.ToString();
                        }
                    }
                }

                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.Users
                        orderby c.UserName
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            comboBox2.Items.Add(c.UserName);
                        }
                    }
                }
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.Types
                        orderby c.TypeColumn
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            comboBox1.Items.Add(c.TypeColumn);
                            comboBox3.Items.Add(c.TypeColumn);
                            comboBox4.Items.Add(c.TypeColumn);
                        }
                    }
                }
                comboBox3.SelectedIndex = 0;
                comboBox4.SelectedIndex = 0;

                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.LastSettings
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            comboBox1.SelectedIndex = int.Parse(c.LastNoIndex);
                            comboBox2.SelectedIndex = int.Parse(c.LastUserIndex);
                            textBox1.Text = c.LastLotid;
                            textBox2.Text = c.LastOrder;
                            textBox3.Text = c.LastPack;
                            textBox4.Text = c.LastNGLimit;
                            textBox5.Text = c.LastNULLLimit;
                        }
                    }
                }

                lbAdd("參數讀取完成", "inf", "");
            }
            catch (Exception e1)
            {
                lbAdd("參數讀取失敗2", "err", e1.ToString());
            }

        }
        #region UnusedTypeSetting
        //void TypeSetting()
        //{
        //    // 初始化為 Dictionary<(type, name, stop), value>
        //    app.dbParam = new Dictionary<(string type, string name, int stop), string>();

        //    using (var db = new MydbDB())
        //    {
        //        // 查詢資料庫中與目前 type (料號) 匹配的所有參數
        //        var query =
        //            from p in db.@params
        //            where p.type == app.produce_No
        //            select p;

        //        foreach (var p in query)
        //        {
        //            // 使用 (type, name, stop) 作為鍵，value 作為值
        //            var key = (p.type, p.name, p.stop);
        //            app.dbParam[key] = p.value.ToString();
        //        }
        //    }

        //    // 測試：如何調用
        //    /*
        //    var testKey = ("10215112360T", "outer_p1", 4);
        //    if (app.param.TryGetValue(testKey, out string value))
        //    {
        //        Console.WriteLine($"Key: {testKey}, Value: {value}");
        //    }
        //    else
        //    {
        //        Console.WriteLine($"Key {testKey} not found!");
        //    }
        //    */
        //}

        //int GetParamValue(string type, string name, int stop)
        //{
        //    if (app.param.TryGetValue(type, out var innerDict) &&
        //        innerDict.TryGetValue((name, stop), out int value))
        //    {
        //        return value; // 返回對應的參數值
        //    }
        //    else
        //    {
        //        throw new Exception($"參數 {type}-{name}-{stop} 不存在");
        //    }
        //}
        #endregion

        void TypeSetting()
        {
            Dictionary<string, string> param = new Dictionary<string, string>();

            if (button6.Text == "變更")
            {
                app.produce_No = "";
                app.foldername = "";
                button6.Text = "設定";
                comboBox1.Enabled = true;

                label32.Text = "";
                label33.Text = "";
                label34.Text = "";

                clear();
            }
            if (comboBox1.Text != "")
            {
                app.produce_No = comboBox1.Text;

                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.Cameras
                        where c.Type == app.produce_No
                        orderby c.Type, c.Stop
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            param[c.Name + c.Stop.ToString()] = c.Value;
                        }
                        if (!app.offline)
                        {
                            for (int i = 1; i < 5; i++)
                            {
                                if (!app.offline)
                                {
                                    cam.Stop();
                                    cam.SetExposureClock(i - 1, uint.Parse(param["exposure" + i]));
                                    cam.SetGain(i - 1, ushort.Parse(param["gain" + i]));
                                    //cam.SetDigitalGain(i - 1, ushort.Parse(app.param["digital_gain" + i]));
                                    PLC_SetD(i * 10, int.Parse(param["delay" + i]));
                                }
                                cam.Start();
                            }
                        }
                    }
                }

                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.@params
                        where c.Type == app.produce_No
                        orderby c.Type, c.Stop
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            param.Add(c.Name + "_" + c.Stop.ToString(), c.Value);
                            //Console.WriteLine(c.Name + "_" + c.Stop.ToString());
                            //Console.WriteLine(c.Value);
                        }
                    }
                }
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.DefectChecks
                        where c.Type == app.produce_No
                        orderby c.Type, c.Stop
                        select c;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            // 使用索引賦值，如果鍵已存在則會覆蓋現有值
                            param[c.Name + c.Stop.ToString() + "_threshold"] = c.Threshold.ToString();
                        }
                    }
                }

                setNet();

                /*
                //待更改
                app.defect_name.Clear();
                app.ng_rec.Clear();
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.DefectChecks
                        where c.Type == app.produce_No && c.Yn == 1
                        orderby c.Name
                        select c.Name;
                    if (q.Count() > 0)
                    {
                        foreach (var c in q)
                        {
                            param.Add(c.Name, c.Yn.ToString());
                            param.Add(c.Name + "_id", c.ClsId.ToString());
                            param.Add(c.Name + "_stop", c.Stop.ToString());
                            if (c.Yn == 1 && c.ClsId != -1)
                            {
                                app.defect_name.Add(c.Name);
                                app.ng_rec.Add(false);
                            }
                        }
                    }
                }
                */
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.DefectChecks
                        where c.Type == app.produce_No && c.Yn == 1 && (c.Stop == 1 || c.Stop == 2)
                        orderby c.Name
                        select c.Name;
                    if (q.Count() > 0)
                    {
                        app.defectLabels_in = q.Distinct().ToList();
                    }
                    Console.WriteLine($"已載入料號={app.produce_No} 內環鏡的瑕疵類別 {app.defectLabels_in.Count} 筆.");
                }
                using (var db = new MydbDB())
                {
                    var q =
                        from c in db.DefectChecks
                        where c.Type == app.produce_No && c.Yn == 1 && (c.Stop == 3 || c.Stop == 4)
                        orderby c.Name
                        select c.Name;
                    if (q.Count() > 0)
                    { 
                        app.defectLabels_out = q.Distinct().ToList();
                    }
                    Console.WriteLine($"已載入料號={app.produce_No} 外環鏡的瑕疵類別 {app.defectLabels_out.Count} 筆.");
                }
                
                // 由 GitHub Copilot 產生
                // 新增: 載入每個站點的瑕疵名稱快取,避免檢測時重複查詢資料庫
                app.defectNamesPerStop.Clear();
                app.chamferDetectionCache.Clear();
                using (var db = new MydbDB())
                {
                    for (int stop = 1; stop <= 4; stop++)
                    {
                        var defectNames = db.DefectChecks
                            .Where(dc => dc.Type == app.produce_No && dc.Stop == stop && dc.Yn == 1)
                            .OrderBy(dc => dc.Name)
                            .Select(dc => dc.Name)
                            .ToList();
                        
                        app.defectNamesPerStop[stop] = defectNames;
                        Console.WriteLine($"已載入站點 {stop} 的瑕疵類別 {defectNames.Count} 筆.");
                        
                        // 由 GitHub Copilot 產生
                        // 新增: 預載入倒角檢測設定
                        var chamferCheck = db.DefectChecks
                            .Where(dc => dc.Type == app.produce_No && 
                                       dc.Stop == stop && 
                                       dc.Name == "cf" && 
                                       dc.Yn == 1)
                            .FirstOrDefault();
                        
                        string cacheKey = $"{app.produce_No}_{stop}";
                        app.chamferDetectionCache[cacheKey] = (chamferCheck != null);
                        Console.WriteLine($"已載入站點 {stop} 的倒角檢測設定: {(chamferCheck != null ? "需要" : "不需要")}");
                    }
                }
            }
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Parameters
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        param.Add(c.Name, c.Value);
                    }
                }
            }
            /*
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Types
                    where c.TypeColumn == app.produce_No
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        param.Add("PTFE", c.PTFE);
                    }
                }
            }
            */
            // 由 GitHub Copilot 產生
            // 使用 ConcurrentDictionary 建構函式轉換
            app.param = new ConcurrentDictionary<string, string>(param);
        }

        #endregion
        #region PLC     

        public static void PLC_SetY(int num, bool status)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SetPoint(1, BoolUnit.Y, num, status, false, Handle);
                Handle.WaitOne();
            }
        }
        public static void PLC_SetM(int num, bool status, bool emer = false)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SetPoint(1, BoolUnit.M, num, status, emer, Handle);
                Handle.WaitOne();
            }
        }
        public static bool PLC_CheckX(int num)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.CheckPoint(1, BoolUnit.X, num, 1, false, Handle);
                Handle.WaitOne();
            }
            return PLC_Value.Point_X[num][0];
        }
        public static bool PLC_CheckM(int num)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.CheckPoint(1, BoolUnit.M, num, 1, false, Handle);
                Handle.WaitOne();
            }
            return PLC_Value.Point_M[num];
        }
        public static void PLC_SetD32(int num, int value)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SendValue_32bit(1, ValueUnit.D, num, value, false, Handle);
                Handle.WaitOne();
            }
        }
        public static void PLC_SetD(int num, int value, bool emer = false)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SendValue(1, ValueUnit.D, num, value, emer, Handle);
                Handle.WaitOne();
            }
        }
        int PLC_CheckD32(int num)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.CheckValue(1, ValueUnit.D, num, 2, false, Handle);
                Handle.WaitOne();
            }
            return PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, num);
        }
        public static int PLC_CheckD(int num)
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.CheckValue(1, ValueUnit.D, num, 1, false, Handle);
                Handle.WaitOne();
            }
            return PLC_Value.Value_D[num];
        }

        public static void PLC_Start()
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SetPoint(1, BoolUnit.M, 3, true, true, Handle);
                Handle.WaitOne();
            }
        }

        void PLC_Counter_Clear(ValueUnit unit, int[] num)
        {
            ManualResetEvent Handle = new ManualResetEvent(false);
            foreach (var n in num)
            {
                PLC_ModBus.SendValue(1, unit, n, 0, false, Handle);
                Handle.WaitOne();
            }
        }
        void PLC_ServoOn()
        {
            PLC_SetM(3, false);//作動
            Thread.Sleep(500);
        }
        void PLC_ServoOff()
        {
            if (!app.offline)
            {
                var Handle = new ManualResetEvent(false);
                PLC_ModBus.SetPoint(1, BoolUnit.M, 1, true, false, Handle);
                Handle.WaitOne();
            }
        }
        #endregion 
        #region createFile
        public void sv() //  用於影像儲存
        {
            while (true)
            {
                if (app.Queue_Save.Count > 0)
                {
                    ImageSave file;
                    app.Queue_Save.TryDequeue(out file);

                    try
                    {
                        string directoryPath = Path.GetDirectoryName(file.path);

                        // 確保資料夾存在
                        if (!Directory.Exists(directoryPath))
                        {
                            Directory.CreateDirectory(directoryPath);
                        }
                        Cv2.ImWrite(file.path, file.image);
                    }
                    catch (Exception e1)
                    {
                        lbAdd("存圖錯誤", "err", e1.ToString());
                    }
                    finally
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: 存檔完成後釋放圖像，避免記憶體洩漏
                        file.image?.Dispose();
                    }
                }
                else
                {
                    app._sv.WaitOne();
                }
            }
        }
        void FolderJob(string FilePath, bool delete, bool create)   //  用於資料夾創建
        {
            DirectoryInfo svDir;
            if (delete)
            {
                svDir = new DirectoryInfo(FilePath);

                if (svDir.Exists)
                    svDir.Delete(true);
            }
            if (create)
            {
                svDir = new DirectoryInfo(FilePath);
                if (!svDir.Exists)
                    svDir.Create();
            }
        }
        private void checkTable()  //檢查sqlite存在
        {
            if (!File.Exists(mydb))
            {
                using (var cn = new SQLiteConnection(mydbStr))
                {
                    lbAdd("資料檔案遺失，請確認setting路徑下存在mydb.sqlite", "err", "錯誤--sqlite。");
                    var result = MessageBox.Show("資料庫檔案遺失，請確認setting路徑下存在mydb.sqlite", "資料庫檔案遺失", MessageBoxButtons.OK);

                    app.start_error = true;
                }
            }
        }
        #endregion
        #region ON/OFF
        void switchButton(bool status)  //  用於機台狀態切換
        {
            if (status) //停止狀態的邏輯
            {
                if (app.DetectMode == 0)
                {
                    Invoke(new Action(() => label46.Text = "停止中"));
                    Invoke(new Action(() => label46.BackColor = Color.Red));
                    Invoke(new Action(() => label51.Text = "0 PCS"));

                    lbAdd("停止檢測", "inf", "");
                    
                    // 由 GitHub Copilot 產生 - 修正：明確停止相機
                    if (!app.offline)
                    {
                        // 1. 先停止相機抓取
                        cam.Stop();
                        System.Threading.Thread.Sleep(100); // 等待停止完成
                        lbAdd("相機已停止", "inf", "");
                    }
                    
                    if (!app.offline)
                    {
                        //app.status = false;
                        PLC_SetM(0, false); //進料停
                        PLC_SetM(2, false); //轉盤停
                        PLC_SetM(5, false); //拍照停
                        PLC_SetY(20, false); //推料
                        PLC_SetY(21, false); //推料
                        PLC_SetY(22, false); //推料

                        //_plcMonitorTimer?.Dispose(); // 停止監控OK1 OK2

                        #region 取像過程訊號清空
                        Log.Information("取像訊號清空開始");
                        for (int i = 150; i <= 159; i++) //NG出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 190; i <= 199; i++) //NG出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 335; i <= 354; i++) //NG出料
                        {
                            PLC_SetM(i, false);
                        }
                        
                        for (int i = 210; i <= 229; i++) //站4到NG佇列
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 301; i <= 320; i++) //站4到NG佇列
                        {
                            PLC_SetM(i, false);
                        }
                        
                        for (int i = 160; i <= 169; i++) //OK1出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 201; i <= 205; i++) //OK1出料
                        {
                            PLC_SetM(i, false);
                        }
                        
                        for (int i = 230; i <= 244; i++) //站4到OK1佇列
                        {
                            PLC_SetM(i, false);
                        }
                        
                        for (int i = 170; i <= 184; i++) //OK2出料
                        {
                            PLC_SetM(i, false);
                        }
                        
                        for (int i = 260; i <= 274; i++) //站4到OK2佇列
                        {
                            PLC_SetM(i, false);
                        }            

                        for (int i = 276; i <= 298; i++) //NULL佇列
                        {
                            if (i == 280 || i == 285 || i == 290) continue;
                            PLC_SetM(i, false);
                        }
                        for (int i = 360; i <= 379; i++) //NULL佇列
                        {
                            PLC_SetM(i, false);
                        }
                        Log.Information("取像訊號清空結束");
                        #endregion
                    }
                    //tasks.Clear();
                    try
                    {
                        dailyCheck();
                    }
                    catch (Exception e1)
                    {
                        lbAdd("資料夾建立失敗", "err", e1.ToString());
                    }

                    if (報表ToolStripMenuItem1.Checked)
                    {
                        if (app.produce_No != "" && app.LotID != "")
                        {
                            try
                            {
                                Log.Information($"switchButton: 準備呼叫 save2excel, produce_No={app.produce_No}, LotID={app.LotID}");
                                save2excel(app.produce_No, app.LotID, new DateTime(), new DateTime(), false);
                                Log.Information($"switchButton: save2excel 完成");
                                lbAdd("報表儲存", "inf", "");
                            }
                            catch (Exception exReport)
                            {
                                Log.Error($"switchButton: save2excel 異常: {exReport.Message}");
                                Log.Error($"switchButton: StackTrace: {exReport.StackTrace}");
                                lbAdd("報表儲存失敗", "err", exReport.ToString());
                            }
                        }
                        else
                        {
                            Log.Warning($"switchButton: produce_No 或 LotID 為空，跳過報表儲存 (produce_No={app.produce_No}, LotID={app.LotID})");
                        }
                    }
                }
                else if (app.DetectMode == 1)
                {
                    Invoke(new Action(() => label46.Text = "停止中"));
                    Invoke(new Action(() => label46.BackColor = Color.Red));

                    if (!app.offline)
                    {
                        PLC_SetM(1, false); //激磁
                        PLC_SetM(5, false); //拍照停
                        PLC_SetM(30, false);
                        PLC_SetM(333, false); 
                    }
                    lbAdd("停止調機", "inf", "");
                }
                
                #region 介面變更
                Invoke(new Action(() => 設定ToolStripMenuItem.Enabled = true));
                Invoke(new Action(() => 使用者ToolStripMenuItem.Enabled = true));
                Invoke(new Action(() => 模式切換ToolStripMenuItem.Enabled = true));
                //Invoke(new Action(() => label13.Text = "0 PCS/分"));
                BeginInvoke(new Action(() => label51.Text = "0 PCS"));

                uiLock(true);
                #endregion
                if (!app.offline)
                {
                    PLC_SetM(30, false); //燈具只有1個                       
                }
            }
            else //啟動狀態的邏輯
            {
                try
                {
                    app.status = true; // 系統整體狀態標記為「正在運行中」
                    var result = new DialogResult();
                    if (app.DetectMode == 0) //正式生產模式
                    {
                        result = MessageBox.Show("閥門即將開啟。", "啟動", MessageBoxButtons.OKCancel);
                        if (result == DialogResult.OK) //若使用者按下「OK」，表示同意正式開始檢測程序。
                        {
                            lock (_initLock)
                            {
                                _systemInitialized = false;
                                Log.Information("系統啟動，等待第一個物體經過第一站光纖（D100 > 0）");
                            }
                            dailyCheck();

                            #region 介面變更


                            Invoke(new Action(() => label46.Text = "檢測中"));
                            Invoke(new Action(() => label46.BackColor = Color.LimeGreen));
                            lbAdd("開始檢測", "inf", "");

                            Invoke(new Action(() => 設定ToolStripMenuItem.Enabled = false));
                            Invoke(new Action(() => 使用者ToolStripMenuItem.Enabled = false));
                            Invoke(new Action(() => 模式切換ToolStripMenuItem.Enabled = false));

                            uiLock(false);
                            #endregion
                            var delaycount = DateTime.Now;
                            if (!app.offline)  // 當系統不是離線模式(!app.offline)時，代表有實際的PLC及相機存在，程式會做一系列控制
                            {
                                trigger(1); //// 將相機觸發模式切換為參數1所代表的模式(可能是硬體觸發)
                                PLC_SetM(21, false); //急停
                                PLC_SetM(3, false); // 轉盤單獨運作、清空訊號

                                PLC_SetM(0, true);
                                PLC_SetM(2, true);
                                PLC_SetM(5, true);

                                //StartPLCMonitoring(); //開始監控OK1 OK2

                                #region 燈具控制
                                //另外拉 button34 開燈
                                PLC_SetM(30, true);
                                #endregion
                            }
                            #region 瑕疵計數
                            //--------------------初始化 接著設料號、LOT 會寫入
                            if (app.dc == null)
                            {
                                app.dc = new ConcurrentDictionary<string, int>();
                            }
                            using (var db = new MydbDB())
                            {
                                var q = from c in db.DefectChecks
                                        where c.Type == app.produce_No && c.Yn == 1
                                        select c;

                                if (q.Count() > 0)
                                {
                                    foreach (var c in q)
                                    {
                                        app.dc[c.Name] = 0;
                                    }
                                }
                            }
                            //--------------------初始化

                            // 嘗試從數據庫讀取上次的計數來初始化 把現有紀錄丟進app.dc
                            Dictionary<string, int> lastCounts = DefectCountManager.ReadLatestDefectCounts(app.produce_No, app.LotID);
                            if (lastCounts != null && lastCounts.Count > 0)
                            {
                                // 由 GitHub Copilot 產生 // 修改: 使用 AddOrUpdate 保證執行緒安全
                                foreach (var pair in lastCounts)
                                {
                                    app.dc.AddOrUpdate(pair.Key, pair.Value, (key, oldValue) => pair.Value);
                                }
                            }
                            // 如果沒有找到現有記錄，為當前料號和批次創建初始記錄
                            if (lastCounts == null || lastCounts.Count == 0)
                            {
                                using (var db = new MydbDB())
                                {
                                    // 為瑕疵類型創建記錄
                                    foreach (var item in app.dc)
                                    {
                                        db.DefectCounts
                                            .Value(p => p.Type, app.produce_No)
                                            .Value(p => p.Name, item.Key)
                                            .Value(p => p.Count, 0)
                                            .Value(p => p.LotId, app.LotID)
                                            .Value(p => p.Time, DateTime.Now)
                                            .Insert();
                                    }

                                    // 為 OK/NG/NULL/SAMPLE_ID 創建記錄
                                    string[] generalTypes = new[] { "OK", "NG", "NULL", "SAMPLE_ID" };
                                    foreach (var type in generalTypes)
                                    {
                                        db.DefectCounts
                                            .Value(p => p.Type, app.produce_No)
                                            .Value(p => p.Name, type)
                                            .Value(p => p.Count, 0)
                                            .Value(p => p.LotId, app.LotID)
                                            .Value(p => p.Time, DateTime.Now)
                                            .Insert();
                                    }
                                }
                            }
                            #endregion
                            #region 開始時間
                            app.rec_time = DateTime.Now;
                            app.lastIn = DateTime.Now;
                            #endregion
                            #region PLC出口計數
                            app._reader.Set();
                            //app._AI.Set();
                            #endregion
                            app.continue_NULL = 0;
                            app.continue_NG = 0;

                            app.detect_result.Clear();
                            app.detect_result_check.Clear();
                        }
                    }
                    if (app.DetectMode == 1)
                    {
                        Invoke(new Action(() => label46.Text = "手動調機中"));
                        Invoke(new Action(() => label46.BackColor = Color.LimeGreen));
                        lbAdd("開始調機", "inf", "");
                        if (!app.offline)
                        {
                            tasks.Add(Task.Factory.StartNew(() => trigger(0)));
                            //await Task.Run(() => trigger(0));
                            PLC_SetM(1, true);
                            PLC_SetM(30, true);
                            //PLC_SetM(5, true);
                            PLC_SetM(333, true);
                            調機設定ToolStripMenuItem_Click(null, null);
                            
                        }
                        uiLock(false);
                    }
                }
                catch (Exception e1)
                {
                    lbAdd("啟動程序異常", "inf", e1.ToString());
                }
            }
        }

        void trigger(int i) //  用於相機觸發模式切換
        {
            if (!app.offline)
            {
                cam.Stop();
                cam.Setting((byte)i);
                cam.Start();
            }
        }
        #endregion
        #region 工具列
        #region 檔案
        private void 圖檔ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lbAdd("開啟圖檔資料夾", "inf", "");

            string exefile = @"C:\Windows\explorer.exe";
            if (Directory.Exists(@".\image"))
                System.Diagnostics.Process.Start(exefile, @".\image");
            else
                BeginInvoke(new System.Action(() => MessageBox.Show("尚無資料", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)));
        }

        private void log檔ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lbAdd("開啟Log檔資料夾", "inf", "");
            string exefile = @"C:\Windows\explorer.exe";
            if (Directory.Exists(@".\logs"))
                System.Diagnostics.Process.Start(exefile, @".\logs");
            else
                BeginInvoke(new System.Action(() => MessageBox.Show("尚無資料", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)));
        }

        private void 報表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lbAdd("開啟報表資料夾", "inf", "");
            string exefile = @"C:\Windows\explorer.exe";
            if (Directory.Exists(@".\Statistics"))
                System.Diagnostics.Process.Start(exefile, @".\Statistics");
            else
                BeginInvoke(new System.Action(() => MessageBox.Show("尚無資料", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)));
        }
        #endregion
        #region 使用者
        private void 管理使用者ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            user_info user_info = new user_info();
            user_info.ShowDialog();

            var register = comboBox2.Text;
            comboBox2.Items.Clear();

            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Users
                    orderby c.UserName
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        comboBox2.Items.Add(c.UserName);
                    }
                }
            }
            comboBox2.Text = register;
        }
        #endregion
        #region 模式
        private void 檢測模式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clear();

            檢測模式ToolStripMenuItem.Checked = true;
            調機模式ToolStripMenuItem.Checked = false;

            if (app.DetectMode != 0)
            {
                if (!app.offline)
                {
                    cam.Setting(1);
                    cam.Start();
                }
            }
            app.DetectMode = 0;

            Invoke(new Action(() => button1.Text = "開始檢測"));
            Invoke(new Action(() => button2.Text = "停止檢測"));
            PLC_SetM(333, false);
            lbAdd("更改為檢測模式", "inf", "");
        }

        private void 調機模式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            檢測模式ToolStripMenuItem.Checked = false;
            調機模式ToolStripMenuItem.Checked = true;

            app.DetectMode = 1;
            button1.Enabled = true;
            button2.Enabled = true;
            Invoke(new Action(() => button1.Text = "開始調機"));
            Invoke(new Action(() => button2.Text = "停止調機"));
            PLC_SetM(333, true);
            lbAdd("更改為調機模式", "inf", "");
        }
        #endregion
        #region 資料儲存
        private void 原圖ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!原圖ToolStripMenuItem.Checked)
            {
                FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + "origin", false, true);
                lbAdd("原圖儲存開啟", "inf", "");
                原圖ToolStripMenuItem.Checked = true;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "origin").Set(p => p.Value, "true").Update();
                }
            }
            else
            {
                原圖ToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "origin").Set(p => p.Value, "false").Update();
                }
            }
        }

        private void oKToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!oKToolStripMenuItem.Checked)
            {
                FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + "OK", false, true);
                lbAdd("OK儲存開啟", "inf", "");
                oKToolStripMenuItem.Checked = true;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "OK").Set(p => p.Value, "true").Update();
                }
            }
            else
            {
                oKToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "OK").Set(p => p.Value, "false").Update();
                }
            }
        }

        private void nGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!nGToolStripMenuItem.Checked)
            {
                FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + "NG", false, true);
                lbAdd("NG儲存開啟", "inf", "");
                nGToolStripMenuItem.Checked = true;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "NG").Set(p => p.Value, "true").Update();
                }
            }
            else
            {
                nGToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "NG").Set(p => p.Value, "false").Update();
                }
            }
        }
        private void nULLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!nULLToolStripMenuItem.Checked)
            {
                FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + "NULL", false, true);
                lbAdd("NULL儲存開啟", "inf", "");
                nULLToolStripMenuItem.Checked = true;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "NULL").Set(p => p.Value, "true").Update();
                }
            }
            else
            {
                nULLToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "NULL").Set(p => p.Value, "false").Update();
                }
            }
        }

        private void 報表ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (!報表ToolStripMenuItem1.Checked)
            {
                lbAdd("報表儲存開啟", "inf", "");
                報表ToolStripMenuItem1.Checked = true;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "report").Set(p => p.Value, "true").Update();
                }
            }
            else
            {
                報表ToolStripMenuItem1.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "report").Set(p => p.Value, "false").Update();
                }
            }
        }
        private void rOIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!rOIToolStripMenuItem.Checked)
            {
                // 為每個站點建立 ROI 資料夾
                for (int station = 1; station <= 4; station++)
                {
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + $"ROI_{station}", false, true);
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + $"chamferROI_{station}", false, true);
                }

                lbAdd("ROI儲存開啟", "inf", "");
                rOIToolStripMenuItem.Checked = true;

                using (var db = new MydbDB())
                {
                    // 檢查參數是否存在，如果不存在則新增
                    var existingParam = db.Parameters.Where(p => p.Name == "saveROI").FirstOrDefault();
                    if (existingParam != null)
                    {
                        db.Parameters.Where(p => p.Name == "saveROI").Set(p => p.Value, "true").Update();
                    }
                    else
                    {
                        db.Parameters
                            .Value(p => p.Name, "saveROI")
                            .Value(p => p.Value, "true")
                            .Insert();
                    }
                }
            }
            else
            {
                rOIToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "saveROI").Set(p => p.Value, "false").Update();
                }
                lbAdd("ROI儲存關閉", "inf", "");
            }
        }

        private void stationsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!stationsToolStripMenuItem.Checked)
            {
                // 為每個站點建立 Stations 結果資料夾
                for (int station = 1; station <= 4; station++)
                {
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + $"Station_{station}_Results", false, true);
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + $"Station_{station}_OK", false, true);
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + $"Station_{station}_NG", false, true);
                }

                lbAdd("各站結果儲存開啟", "inf", "");
                stationsToolStripMenuItem.Checked = true;

                using (var db = new MydbDB())
                {
                    // 檢查參數是否存在，如果不存在則新增
                    var existingParam = db.Parameters.Where(p => p.Name == "saveStations").FirstOrDefault();
                    if (existingParam != null)
                    {
                        db.Parameters.Where(p => p.Name == "saveStations").Set(p => p.Value, "true").Update();
                    }
                    else
                    {
                        db.Parameters
                            .Value(p => p.Name, "saveStations")
                            .Value(p => p.Value, "true")
                            .Insert();
                    }
                }
            }
            else
            {
                stationsToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "saveStations").Set(p => p.Value, "false").Update();
                }
                lbAdd("各站結果儲存關閉", "inf", "");
            }
        }
        private void visToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!visToolStripMenuItem.Checked)
            {
                // 為視覺化影像建立資料夾
                for (int station = 1; station <= 4; station++)
                {
                    FolderJob(@".\image\" + st.ToString("yyyy-MM") + "\\" + st.ToString("MMdd") + "\\" + app.foldername + "\\" + $"VIS_{station}", false, true);
                }

                lbAdd("視覺化影像儲存開啟", "inf", "");
                visToolStripMenuItem.Checked = true;

                using (var db = new MydbDB())
                {
                    // 檢查參數是否存在，如果不存在則新增
                    var existingParam = db.Parameters.Where(p => p.Name == "saveVIS").FirstOrDefault();
                    if (existingParam != null)
                    {
                        db.Parameters.Where(p => p.Name == "saveVIS").Set(p => p.Value, "true").Update();
                    }
                    else
                    {
                        db.Parameters
                            .Value(p => p.Name, "saveVIS")
                            .Value(p => p.Value, "true")
                            .Insert();
                    }
                }
            }
            else
            {
                visToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "saveVIS").Set(p => p.Value, "false").Update();
                }
                lbAdd("視覺化影像儲存關閉", "inf", "");
            }
        }
        private void 蜂鳴器ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!蜂鳴器ToolStripMenuItem.Checked)
            {
                lbAdd("蜂鳴器開啟", "inf", "");
                蜂鳴器ToolStripMenuItem.Checked = true;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "alert").Set(p => p.Value, "true").Update();
                }
                PLC_SetM(20, true);
            }
            else
            {
                蜂鳴器ToolStripMenuItem.Checked = false;
                using (var db = new MydbDB())
                {
                    db.Parameters.Where(p => p.Name == "alert").Set(p => p.Value, "false").Update();
                }
                PLC_SetM(20, false);
            }

        }

        #endregion
        #region 設定

        private void 調機設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            parameter_info parameter_info = new parameter_info(cam);
            parameter_info.ShowDialog();

            TypeSetting();
        }
        private void 料號設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            type_info type_info = new type_info();
            type_info.ShowDialog();

            var register1 = comboBox1.Text;
            var register2 = comboBox3.Text;
            var register3 = comboBox4.Text;
            comboBox1.Items.Clear();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();

            using (var db = new MydbDB())
            {
                var q =
                    from c in db.Types
                    orderby c.TypeColumn
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        comboBox1.Items.Add(c.TypeColumn);
                        comboBox3.Items.Add(c.TypeColumn);
                        comboBox4.Items.Add(c.TypeColumn);
                    }
                }
            }
            comboBox1.Text = register1;
            comboBox3.Text = register2;
            comboBox4.Text = register3;

            TypeSetting();
        }
        private void 相機參數設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            camera_info camera_info = new camera_info();
            camera_info.ShowDialog();

            TypeSetting();

        }
        private void 瑕疵種類設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            defect_type_info defect_type__info = new defect_type_info();
            defect_type__info.ShowDialog();
        }
        private void 檢測瑕疵設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            defect_check_info defect_check_info = new defect_check_info();
            defect_check_info.ShowDialog();
        }
        private void 檔案留存天數設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //  參數視窗，作業員無權限，管理者工程師介面有別
            if (app.user == 1 || app.user == 0)
            {
                keepday keepday = new keepday();
                keepday.ShowDialog();
            }
            if (app.paramUpdate)
            {
                app.paramUpdate = false;
                ReadParam();
            }
        }
        private void 出料卡料時間設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                // 創建設定視窗
                using (Form settingForm = new Form())
                {
                    settingForm.Text = "出料卡料時間設定";
                    settingForm.Size = new System.Drawing.Size(400, 250);
                    settingForm.StartPosition = FormStartPosition.CenterParent;
                    settingForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    settingForm.MaximizeBox = false;
                    settingForm.MinimizeBox = false;

                    // 創建標籤和文字框
                    Label lblD8 = new Label();
                    lblD8.Text = "卡料倒數時間長度 (D8):";
                    lblD8.Location = new System.Drawing.Point(20, 30);
                    lblD8.Size = new System.Drawing.Size(150, 25);
                    lblD8.Font = new Font("微軟正黑體", 12F);

                    TextBox txtD8 = new TextBox();
                    txtD8.Location = new System.Drawing.Point(180, 30);
                    txtD8.Size = new System.Drawing.Size(100, 25);
                    txtD8.Font = new Font("微軟正黑體", 12F);
                    txtD8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;

                    Label lblD8Unit = new Label();
                    lblD8Unit.Text = "秒";
                    lblD8Unit.Location = new System.Drawing.Point(290, 30);
                    lblD8Unit.Size = new System.Drawing.Size(30, 25);
                    lblD8Unit.Font = new Font("微軟正黑體", 12F);

                    Label lblD9 = new Label();
                    lblD9.Text = "沒料倒數時間長度 (D9):";
                    lblD9.Location = new System.Drawing.Point(20, 80);
                    lblD9.Size = new System.Drawing.Size(150, 25);
                    lblD9.Font = new Font("微軟正黑體", 12F);

                    TextBox txtD9 = new TextBox();
                    txtD9.Location = new System.Drawing.Point(180, 80);
                    txtD9.Size = new System.Drawing.Size(100, 25);
                    txtD9.Font = new Font("微軟正黑體", 12F);
                    txtD9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;

                    Label lblD9Unit = new Label();
                    lblD9Unit.Text = "秒";
                    lblD9Unit.Location = new System.Drawing.Point(290, 80);
                    lblD9Unit.Size = new System.Drawing.Size(30, 25);
                    lblD9Unit.Font = new Font("微軟正黑體", 12F);

                    // 狀態標籤
                    Label lblStatus = new Label();
                    lblStatus.Text = "正在讀取PLC資料...";
                    lblStatus.Location = new System.Drawing.Point(20, 130);
                    lblStatus.Size = new System.Drawing.Size(350, 25);
                    lblStatus.Font = new Font("微軟正黑體", 10F);
                    lblStatus.ForeColor = Color.Blue;

                    // 按鈕
                    Button btnConfirm = new Button();
                    btnConfirm.Text = "確認";
                    btnConfirm.Location = new System.Drawing.Point(180, 170);
                    btnConfirm.Size = new System.Drawing.Size(80, 35);
                    btnConfirm.Font = new Font("微軟正黑體", 12F);
                    btnConfirm.BackColor = Color.LightGreen;

                    Button btnCancel = new Button();
                    btnCancel.Text = "取消";
                    btnCancel.Location = new System.Drawing.Point(280, 170);
                    btnCancel.Size = new System.Drawing.Size(80, 35);
                    btnCancel.Font = new Font("微軟正黑體", 12F);
                    btnCancel.BackColor = Color.LightCoral;
                    btnCancel.DialogResult = DialogResult.Cancel;

                    // 將控件加入視窗
                    settingForm.Controls.Add(lblD8);
                    settingForm.Controls.Add(txtD8);
                    settingForm.Controls.Add(lblD8Unit);
                    settingForm.Controls.Add(lblD9);
                    settingForm.Controls.Add(txtD9);
                    settingForm.Controls.Add(lblD9Unit);
                    settingForm.Controls.Add(lblStatus);
                    settingForm.Controls.Add(btnConfirm);
                    settingForm.Controls.Add(btnCancel);

                    // 讀取PLC資料的Task
                    Task.Run(() =>
                    {
                        try
                        {
                            if (!app.offline)
                            {
                                // 讀取D8和D9的值 (單位：0.1秒)
                                int d8RawValue = PLC_CheckD(8);
                                int d9RawValue = PLC_CheckD(9);
                                if (d8RawValue == 0)
                                {
                                    PLC_SetD(8, 6000);
                                    d8RawValue = 6000;
                                }
                                if (d9RawValue == 0)
                                {
                                    PLC_SetD(9, 6000);
                                    d9RawValue = 6000;
                                }
                                // 轉換為秒顯示 (除以10)
                                double d8Seconds = d8RawValue / 10.0;
                                double d9Seconds = d9RawValue / 10.0;

                                // 更新UI (需要在UI線程執行)
                                settingForm.BeginInvoke(new Action(() =>
                                {
                                    txtD8.Text = d8Seconds.ToString("0.0");
                                    txtD9.Text = d9Seconds.ToString("0.0");
                                    lblStatus.Text = "資料讀取完成";
                                    lblStatus.ForeColor = Color.Green;
                                    btnConfirm.Enabled = true;
                                }));
                            }
                            else
                            {
                                // 離線模式
                                settingForm.BeginInvoke(new Action(() =>
                                {
                                    txtD8.Text = "0.0";
                                    txtD9.Text = "0.0";
                                    lblStatus.Text = "離線模式 - 請手動輸入值";
                                    lblStatus.ForeColor = Color.Orange;
                                    btnConfirm.Enabled = true;
                                }));
                            }
                        }
                        catch (Exception ex)
                        {
                            settingForm.BeginInvoke(new Action(() =>
                            {
                                lblStatus.Text = $"讀取失敗: {ex.Message}";
                                lblStatus.ForeColor = Color.Red;
                                btnConfirm.Enabled = true;
                            }));
                        }
                    });

                    // 初始化時先禁用確認按鈕
                    btnConfirm.Enabled = false;

                    // 確認按鈕事件
                    btnConfirm.Click += (s, args) =>
                    {
                        try
                        {
                            // 驗證輸入
                            if (!double.TryParse(txtD8.Text, out double newD8Seconds) || newD8Seconds < 0)
                            {
                                MessageBox.Show("請輸入有效的卡料倒數時間長度 (>=0秒)", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            if (!double.TryParse(txtD9.Text, out double newD9Seconds) || newD9Seconds < 0)
                            {
                                MessageBox.Show("請輸入有效的沒料倒數時間長度 (>=0秒)", "輸入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            // 轉換為PLC單位 (乘以10，轉為0.1秒單位)
                            int newD8Value = (int)(newD8Seconds * 10);
                            int newD9Value = (int)(newD9Seconds * 10);

                            // 確認對話框
                            DialogResult result = MessageBox.Show(
                                $"確定要將參數更新為:\n" +
                                $"卡料倒數時間長度 (D8): {newD8Seconds:0.0} 秒\n" +
                                $"沒料倒數時間長度 (D9): {newD9Seconds:0.0} 秒\n\n" +
                                $"PLC內部值將設為:\n" +
                                $"D8: {newD8Value} (單位: 0.1秒)\n" +
                                $"D9: {newD9Value} (單位: 0.1秒)",
                                "確認更新",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.Yes)
                            {
                                lblStatus.Text = "正在更新PLC參數...";
                                lblStatus.ForeColor = Color.Blue;
                                btnConfirm.Enabled = false;

                                // 在背景線程更新PLC
                                Task.Run(() =>
                                {
                                    try
                                    {
                                        if (!app.offline)
                                        {
                                            // 更新PLC的D8和D9 (以0.1秒為單位)
                                            PLC_SetD(8, newD8Value);
                                            PLC_SetD(9, newD9Value);

                                            // 驗證更新是否成功
                                            int verifyD8 = PLC_CheckD(8);
                                            int verifyD9 = PLC_CheckD(9);

                                            settingForm.BeginInvoke(new Action(() =>
                                            {
                                                if (verifyD8 == newD8Value && verifyD9 == newD9Value)
                                                {
                                                    lblStatus.Text = "參數更新成功";
                                                    lblStatus.ForeColor = Color.Green;
                                                    lbAdd($"出料卡料時間設定更新完成 - D8:{newD8Seconds:0.0}秒({newD8Value}), D9:{newD9Seconds:0.0}秒({newD9Value})", "inf", "");

                                                    // 1秒後關閉視窗
                                                    System.Windows.Forms.Timer closeTimer = new System.Windows.Forms.Timer();
                                                    closeTimer.Interval = 1000;
                                                    closeTimer.Tick += (ts, te) =>
                                                    {
                                                        closeTimer.Stop();
                                                        settingForm.DialogResult = DialogResult.OK;
                                                        settingForm.Close();
                                                    };
                                                    closeTimer.Start();
                                                }
                                                else
                                                {
                                                    lblStatus.Text = "參數更新後驗證失敗";
                                                    lblStatus.ForeColor = Color.Red;
                                                    btnConfirm.Enabled = true;
                                                }
                                            }));
                                        }
                                        else
                                        {
                                            settingForm.BeginInvoke(new Action(() =>
                                            {
                                                lblStatus.Text = "離線模式 - 無法更新PLC";
                                                lblStatus.ForeColor = Color.Orange;
                                                btnConfirm.Enabled = true;
                                            }));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        settingForm.BeginInvoke(new Action(() =>
                                        {
                                            lblStatus.Text = $"更新失敗: {ex.Message}";
                                            lblStatus.ForeColor = Color.Red;
                                            btnConfirm.Enabled = true;
                                            lbAdd($"出料卡料時間設定更新失敗: {ex.Message}", "err", ex.ToString());
                                        }));
                                    }
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"操作失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            lbAdd($"出料卡料時間設定操作失敗: {ex.Message}", "err", ex.ToString());
                        }
                    };

                    // 取消按鈕事件
                    btnCancel.Click += (s, args) =>
                    {
                        settingForm.DialogResult = DialogResult.Cancel;
                        settingForm.Close();
                    };

                    // 數字輸入限制 (允許小數點)
                    txtD8.KeyPress += (s, args) =>
                    {
                        if (!char.IsControl(args.KeyChar) && !char.IsDigit(args.KeyChar) && args.KeyChar != '.')
                        {
                            args.Handled = true;
                        }
                        // 防止多個小數點
                        if (args.KeyChar == '.' && (s as TextBox).Text.IndexOf('.') > -1)
                        {
                            args.Handled = true;
                        }
                    };

                    txtD9.KeyPress += (s, args) =>
                    {
                        if (!char.IsControl(args.KeyChar) && !char.IsDigit(args.KeyChar) && args.KeyChar != '.')
                        {
                            args.Handled = true;
                        }
                        // 防止多個小數點
                        if (args.KeyChar == '.' && (s as TextBox).Text.IndexOf('.') > -1)
                        {
                            args.Handled = true;
                        }
                    };

                    // 設定預設按鈕
                    settingForm.AcceptButton = btnConfirm;
                    settingForm.CancelButton = btnCancel;

                    // 顯示視窗
                    settingForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"開啟出料卡料時間設定失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lbAdd($"出料卡料時間設定開啟失敗: {ex.Message}", "err", ex.ToString());
            }
        }
        private void 圓心校正工具ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowCircleCalibrationTool();
        }
        private void 檢測參數設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 檢查是否已經有料號設定
            if (string.IsNullOrEmpty(app.produce_No))
            {
                MessageBox.Show("請先設定料號後再進行參數設定", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 建立並顯示參數設定表單
                using (var parameterConfigForm = new ParameterConfigForm(app.produce_No))
                {
                    parameterConfigForm.ShowDialog();
                }

                // 參數設定完成後，重新載入參數設定
                TypeSetting();

                lbAdd($"料號 {app.produce_No} 的參數設定已完成", "inf", "");
            }
            catch (Exception ex)
            {
                lbAdd("開啟參數設定介面時發生錯誤", "err", ex.ToString());
                MessageBox.Show($"開啟參數設定介面時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 由 GitHub Copilot 產生
        private void 時間測量ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!時間測量ToolStripMenuItem.Checked)
            {
                時間測量ToolStripMenuItem.Checked = true;
                app.enableProfiling = true;
                PerformanceProfiler.Initialize(@".\logs\ai_performance_log.txt"); // 啟用時計一次初始化

                // 啟用子選單
                站1ToolStripMenuItem.Enabled = true;
                站2ToolStripMenuItem.Enabled = true;
                站3ToolStripMenuItem.Enabled = true;
                站4ToolStripMenuItem.Enabled = true;

                lbAdd("時間測量已啟用", "inf", "");
            }
            else
            {
                時間測量ToolStripMenuItem.Checked = false;
                app.enableProfiling = false;
                PerformanceProfiler.Shutdown(); // 停用時釋放

                // 清空選中的站點並禁用子選單
                app.profilingStations.Clear();
                站1ToolStripMenuItem.Checked = false;
                站1ToolStripMenuItem.Enabled = false;
                站2ToolStripMenuItem.Checked = false;
                站2ToolStripMenuItem.Enabled = false;
                站3ToolStripMenuItem.Checked = false;
                站3ToolStripMenuItem.Enabled = false;
                站4ToolStripMenuItem.Checked = false;
                站4ToolStripMenuItem.Enabled = false;

                lbAdd("時間測量已停用", "inf", "");
            }
        }
        private void 站1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!站1ToolStripMenuItem.Checked)
            {
                站1ToolStripMenuItem.Checked = true;
                app.profilingStations.Add(1);
                lbAdd("站1時間測量已啟用", "inf", "");
            }
            else
            {
                站1ToolStripMenuItem.Checked = false;
                app.profilingStations.Remove(1);
                lbAdd("站1時間測量已停用", "inf", "");
            }
        }
        private void 站2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!站2ToolStripMenuItem.Checked)
            {
                站2ToolStripMenuItem.Checked = true;
                app.profilingStations.Add(2);
                lbAdd("站2時間測量已啟用", "inf", "");
            }
            else
            {
                站2ToolStripMenuItem.Checked = false;
                app.profilingStations.Remove(2);
                lbAdd("站2時間測量已停用", "inf", "");
            }
        }
        private void 站3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!站3ToolStripMenuItem.Checked)
            {
                站3ToolStripMenuItem.Checked = true;
                app.profilingStations.Add(3);
                lbAdd("站3時間測量已啟用", "inf", "");
            }
            else
            {
                站3ToolStripMenuItem.Checked = false;
                app.profilingStations.Remove(3);
                lbAdd("站3時間測量已停用", "inf", "");
            }
        }
        private void 站4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!站4ToolStripMenuItem.Checked)
            {
                站4ToolStripMenuItem.Checked = true;
                app.profilingStations.Add(4);
                lbAdd("站4時間測量已啟用", "inf", "");
            }
            else
            {
                站4ToolStripMenuItem.Checked = false;
                app.profilingStations.Remove(4);
                lbAdd("站4時間測量已停用", "inf", "");
            }
        }

        /// <summary>
        /// 檢查出料卡料時間是否已設定
        /// </summary>
        /// <returns>如果D8和D9都有設定合理值則返回true，否則返回false</returns>
        private bool CheckDeliveryTimeSettings()
        {
            try
            {
                if (app.offline)
                {
                    // 離線模式下假設已設定
                    return true;
                }

                // 讀取D8和D9的值
                int d8Value = PLC_CheckD(8);  // 卡料倒數時間 (單位: 0.1秒)
                int d9Value = PLC_CheckD(9);  // 沒料倒數時間 (單位: 0.1秒)

                // 檢查是否為合理的設定值
                // D8和D9應該都大於0且小於一個合理的上限 (例如10分鐘 = 6000個0.1秒)
                bool d8Valid = d8Value > 0;
                bool d9Valid = d9Value > 0;

                if (!d8Valid || !d9Valid)
                {
                    lbAdd($"出料卡料時間檢查失敗 - D8:{d8Value / 10.0:F1}秒, D9:{d9Value / 10.0:F1}秒", "err",
                          $"D8有效:{d8Valid}, D9有效:{d9Valid}");
                    return false;
                }

                lbAdd($"出料卡料時間檢查通過 - D8:{d8Value / 10.0:F1}秒, D9:{d9Value / 10.0:F1}秒", "inf", "");
                return true;
            }
            catch (Exception ex)
            {
                lbAdd($"檢查出料卡料時間設定時發生錯誤: {ex.Message}", "err", ex.ToString());
                // 發生錯誤時，為了不阻止生產，返回true並記錄錯誤
                return true;
            }
        }
        #endregion
        #region 關閉程式
        // 在 Form1 類別中新增共用方法
        private bool PerformCloseSequence()
        {
            try
            {
                // 第一步：詢問是否真的要關閉程式
                DialogResult closeConfirm = MessageBox.Show(
                    "確定要關閉程式嗎？",
                    "確認關閉",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);

                if (closeConfirm != DialogResult.Yes)
                {
                    return false; // 使用者取消
                }
                #region 匯出瑕疵統計已改到button2做
                /*
                // 第二步：匯出瑕疵統計詢問（如果需要）
                DialogResult exportResult = MessageBox.Show(
                    "是否要匯出瑕疵統計資料？",
                    "匯出統計",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (exportResult == DialogResult.Cancel)
                {
                    return false; // 使用者取消
                }

                if (exportResult == DialogResult.Yes)
                {
                    string statsFolder = Path.Combine(Environment.CurrentDirectory, "Statistics");
                    if (!Directory.Exists(statsFolder))
                    {
                        Directory.CreateDirectory(statsFolder);
                    }

                    string fileName = string.Format("瑕疵統計_{0}_{1}_{2:yyyyMMdd_HHmmss}.csv",
                                                   app.produce_No,
                                                   app.LotID,
                                                   DateTime.Now);
                    string filePath = Path.Combine(statsFolder, fileName);

                    bool success = ResultManager.ExportStatsToCsv(filePath);

                    if (success)
                    {
                        lbAdd($"瑕疵統計已成功匯出至: {filePath}", "inf", "");
                    }
                    else
                    {
                        lbAdd("匯出瑕疵統計失敗", "war", "");
                    }
                }
                */
                #endregion
                // 第三步：最後確認
                DialogResult finalConfirm = MessageBox.Show(
                    "即將關閉程式，此操作無法復原。\n確定要繼續嗎？",
                    "最後確認",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                return finalConfirm == DialogResult.Yes;
            }
            catch (Exception ex)
            {
                lbAdd("關閉程式時發生錯誤", "err", ex.Message);
                return true; // 即使發生錯誤也允許關閉
            }
        }
        private void 關閉程式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion
        #endregion
        #region 介面事件
        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) ||
                (e.KeyChar >= (Char)65 && e.KeyChar <= (Char)90) ||
                (e.KeyChar >= (Char)97 && e.KeyChar <= (Char)122) ||
               e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
        private void textBox_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= (Char)48 && e.KeyChar <= (Char)57) ||
               e.KeyChar == (Char)13 || e.KeyChar == (Char)8)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
        void clear()    //  全清
        {
            #region Data
            #region 資料及計數清除
            app.detect_result.Clear();
            app.detect_result_check.Clear();
            app.counter.Clear();
            var counterType = new[] { "OK1", "OK2", "OK", "SAMPLE_ID", "NG", "NULL", "stop0", "stop1", "stop2", "stop3" };
            foreach (var s in counterType)
            {
                app.counter.TryAdd(s, 0);
            }

            Invoke(new Action(() => label3.Text = "0"));
            Invoke(new Action(() => label58.Text = "0"));
            Invoke(new Action(() => label59.Text = "0"));
            Invoke(new Action(() => label6.Text = "0"));
            Invoke(new Action(() => label14.Text = "0"));
            Invoke(new Action(() => label60.Text = "0.0%"));
            Invoke(new Action(() => label23.Text = "0.0%"));
            Invoke(new Action(() => label17.Text = "0.0%"));
            #endregion
            #endregion
            #region Folder
            st = DateTime.Now;
            dailyCheck();
            #endregion
            lbAdd("清空計數", "inf", "");
        }
        #region 計算空間
        public static long GetSize(DirectoryInfo dirInfo)
        {
            System.Type tp = System.Type.GetTypeFromProgID("Scripting.FileSystemObject");
            object fso = Activator.CreateInstance(tp);
            object fd = tp.InvokeMember("GetFolder", BindingFlags.InvokeMethod, null, fso, new object[] { dirInfo.FullName });
            long ret = Convert.ToInt64(tp.InvokeMember("Size", BindingFlags.GetProperty, null, fd, null));
            Marshal.ReleaseComObject(fso);
            return ret;
        }
        public static double GetHardDiskFreeSpace(string DiskName)
        {
            double freeSpace = new double();
            DiskName = DiskName.ToUpper() + ":\\";
            System.IO.DriveInfo[] drives = System.IO.DriveInfo.GetDrives();
            foreach (System.IO.DriveInfo drive in drives)
            {
                if (drive.Name == DiskName)
                {
                    freeSpace = drive.TotalFreeSpace / (double)(1024 * 1024 * 1024);
                }
            }
            return freeSpace;
        }
        #endregion
        #region 按鈕事件
        #region On/Off

        private void button1_Click(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 調機模式優先檢查，不受狀態限制
            if (app.DetectMode == 1)
            {
                // 調機模式：直接開始調機，不檢查狀態和其他參數
                // 由 GitHub Copilot 產生 - 立即設定調機模式旗標，避免影像進入檢測流程
                app.isAdjustmentMode = true;
                switchButton(false);
                lbAdd("開始調機", "inf", "");
                return;
            }

            // 檢測模式：需要檢查狀態
            if (app.currentState != app.SystemState.Stopped)
            {
                string message = GetStateMessage();
                CustomMessageBox.Show(
                    message,
                    "操作順序提醒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
                );
                return;
            }
            if (app.DetectMode == 0)
            {
                if (!comboBox1.Enabled)
                {
                    if (!textBox1.Enabled)
                    {
                        if (!textBox2.Enabled)
                        {
                            if (!textBox3.Enabled)
                            {
                                if (!textBox4.Enabled)
                                {
                                    if (!textBox5.Enabled)
                                    {
                                        if (!comboBox2.Enabled)
                                        {
                                            // 檢查出料卡料時間設定
                                            if (CheckDeliveryTimeSettings()) // CheckDeliveryTimeSettings()
                                            {
                                                // **新增：檢查閘門狀態（M6 和 M9）**
                                                if (!app.offline)
                                                {
                                                    bool m16Status = PLC_CheckM(16);
                                                    bool m18Status = PLC_CheckM(18);

                                                    if (m16Status && !m18Status)
                                                    {
                                                        // **新增：狀態變更為運行中**
                                                        app.currentState = app.SystemState.Running;
                                                        switchButton(false);

                                                        // **新增：更新UI狀態**
                                                        UpdateButtonStates();
                                                    }
                                                    else
                                                    {
                                                        string gateStatusMessage = "閘門狀態異常，請檢查：\n";
                                                        if (!m16Status) gateStatusMessage += "- OK1 未開啟\n";
                                                        if (m18Status) gateStatusMessage += "- OK2 未關閉\n";

                                                        CustomMessageBox.Show(
                                                            gateStatusMessage,
                                                            "請檢查閘門狀態",
                                                            MessageBoxButtons.OK,
                                                            MessageBoxIcon.Warning,
                                                            new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
                                                        );
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                CustomMessageBox.Show("尚未設定出料卡料時間。\n請先至 【設定->出料/卡料時間設定】\n設定D8(卡料倒數時間)和D9(沒料倒數時間)。",
                                                               "出料卡料時間未設定",
                                                               MessageBoxButtons.OK,
                                                               MessageBoxIcon.Warning,
                                                                new Font("微軟正黑體", 16F, FontStyle.Bold)
                                                                );
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("尚未登入使用者。");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("尚未設定NULL籃上限數量。");
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("尚未設定NG籃上限數量。");
                                }
                            }
                            else
                            {
                                MessageBox.Show("尚未設定包裝數量。");
                            }
                        }
                        else
                        {
                            MessageBox.Show("尚未設定工單數量。");
                        }
                    }
                    else
                    {
                        MessageBox.Show("尚未設定Lot_ID。");
                    }
                }
                else
                {
                    MessageBox.Show("尚未設定料號。");
                }
            }
            // 由 GitHub Copilot 產生 - 移除此處的調機模式處理，已在函式開頭處理
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // 由 GitHub Copilot 產生 - 調機模式優先檢查，不受狀態限制
            if (app.DetectMode == 1)
            {
                // 調機模式：直接停止調機
                // 由 GitHub Copilot 產生 - 重置調機模式旗標
                app.isAdjustmentMode = false;
                switchButton(true);
                lbAdd("停止調機", "inf", "");
                return;
            }

            // 使用 DefectCountManager 執行寫入
            // **新增：檢查狀態，只有在運行中才能停止**
            if (app.currentState != app.SystemState.Running)
            {
                CustomMessageBox.Show(
                    "系統未在運行中，無法執行停止操作。",
                    "操作提醒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    new System.Drawing.Font("微軟正黑體", 16F)
                );
                return;
            }

            try
            {
                // 停止時強制寫入最終計數
                bool writeSuccess = DefectCountManager.PerformPeriodicWrite(forceWrite: true);
                if (!writeSuccess)
                {
                    Log.Error("停止檢測時 DefectCount 寫入失敗");
                }
                else
                {
                    Log.Information("停止檢測時 DefectCount 強制寫入成功");
                }
            }
            catch (Exception ex)
            {
                Log.Error($"背景寫入失敗: {ex.Message}");
            }

            // 調用 switchButton 切換狀態
            Log.Information("button2_Click: 開始執行 switchButton");
            switchButton(true);
            Log.Information("button2_Click: switchButton 完成");

            // 輸送帶先復歸
            Log.Information("button2_Click: 開始執行 PLC_SetM(401)");
            PLC_SetM(401, true);
            Log.Information("button2_Click: PLC_SetM(401) 完成");

            //通知
            Log.Information("button2_Click: 開始顯示停止通知");
            ShowStopDetectionNotification();
            Log.Information("button2_Click: 停止通知顯示完成");

            _lastProductiveTime = null;   // 重置為null，等待第一個樣品推料
            // **新增：狀態變更為停止後需要更新計數**
            app.currentState = app.SystemState.StoppedNeedUpdate;
            UpdateButtonStates();

            #region 匯出檢測統計
            DialogResult exportResult = MessageBox.Show(
                    "是否要匯出瑕疵統計資料？",
                    "匯出統計",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

            if (exportResult == DialogResult.Cancel)
            {
                return; // 使用者選擇取消，直接返回
            }

            if (exportResult == DialogResult.Yes)
            {
                // 建立資料夾路徑
                string statsFolder = Path.Combine(Environment.CurrentDirectory, "Statistics");
                if (!Directory.Exists(statsFolder))
                {
                    Directory.CreateDirectory(statsFolder);
                }

                // 生成檔案名稱
                string fileName = string.Format("瑕疵統計_{0}_{1}_{2:yyyyMMdd_HHmmss}.csv",
                                               app.produce_No,
                                               app.LotID,
                                               DateTime.Now);
                string filePath = Path.Combine(statsFolder, fileName);

                // 匯出統計資料
                bool success = ResultManager.ExportStatsToCsv(filePath);

                if (success)
                {
                    lbAdd($"瑕疵統計已成功匯出至: {filePath}", "inf", "");
                }
                else
                {
                    lbAdd("匯出瑕疵統計失敗", "war", "");
                }
            }
            #endregion
        }

        // **新增：獲取狀態提示訊息的輔助方法**
        private string GetStateMessage()
        {
            switch (app.currentState)
            {
                case app.SystemState.Running:
                    return "系統正在運行中，請先停止檢測。";
                case app.SystemState.StoppedNeedUpdate:
                    return "檢測已停止，請先點擊「更新計數」按鈕更新數據。";
                case app.SystemState.UpdatedNeedReset:
                    return "計數已更新，請先點擊「異常復歸」按鈕進行復歸操作。";
                default:
                    return "系統狀態異常，請聯絡技術人員。";
            }
        }
        #endregion
        #region 分頁切換
        private void button3_Click(object sender, EventArgs e)
        {
            button3.BackColor = Color.FromArgb(255, 128, 128);
            button4.BackColor = Color.FromArgb(224, 224, 224);
            button5.BackColor = Color.FromArgb(224, 224, 224);
            Invoke(new Action(() => tabControl1.SelectedTab = tabPage1));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            button18_Click(null, null);
            button3.BackColor = Color.FromArgb(224, 224, 224);
            button4.BackColor = Color.FromArgb(255, 128, 128);
            button5.BackColor = Color.FromArgb(224, 224, 224);
            Invoke(new Action(() => tabControl1.SelectedTab = tabPage2));
        }
        private void button5_Click(object sender, EventArgs e)
        {
            button3.BackColor = Color.FromArgb(224, 224, 224);
            button4.BackColor = Color.FromArgb(224, 224, 224);
            button5.BackColor = Color.FromArgb(255, 128, 128);
            Invoke(new Action(() => tabControl1.SelectedTab = tabPage3));
        }
        #endregion
        #region 料號
        private void button6_Click(object sender, EventArgs e)
        {
            type_info typeInfoForm = new type_info();
            if (button6.Text == "設定")
            {
                if (comboBox1.Text != "")
                {
                    app.produce_No = comboBox1.Text;
                    if (app.produce_No != "" && app.LotID != "")
                    {
                        app.foldername = app.produce_No + "-" + app.LotID;
                    }
                    lbAdd("設定料號為:" + app.produce_No, "inf", "");

                    using (var db = new MydbDB())
                    {
                        var q =
                            from c in db.Types
                            where c.TypeColumn == app.produce_No
                            orderby c.TypeColumn
                            select c;
                        if (q.Count() > 0)
                        {
                            foreach (var c in q)
                            {
                                label32.Text = comboBox1.Text;
                                label33.Text = c.material.ToString() + c.thick.ToString() + c.PTFEColor.ToString();
                                label22.Text = c.boxorpack;
                                label54.Text = c.ID.ToString() + "ID x " + c.OD.ToString() + "OD x " + c.H.ToString() + "H";
                                label61.Text = c.package;
                            }
                        }
                    }
                    lbAdd("typesetting","inf","");
                    TypeSetting();
                    lbAdd("typesetting_End", "inf", "");
                    PLC_SetD(1, int.Parse(app.param["fourToNG_time_ms_4"]), true);
                    PLC_SetD(2, int.Parse(app.param["fourToNG_time_ms_4"]), true);
                    PLC_SetD(3, int.Parse(app.param["fourToOK1_time_ms_4"]), true);
                    PLC_SetD(4, int.Parse(app.param["fourToOK1_time_ms_4"]), true);
                    PLC_SetD(5, int.Parse(app.param["fourToOK2_time_ms_4"]), true);
                    PLC_SetD(6, int.Parse(app.param["fourToOK2_time_ms_4"]), true);
                    PLC_SetD(13, int.Parse(app.param["fourToNULL_time_ms_4"]), true);


                    button6.Text = "變更";
                    comboBox1.Enabled = false;

                    using (var db = new MydbDB())
                    {
                        var q =
                            from c in db.DefectCounts
                            where c.LotId == app.LotID && c.Type == app.produce_No
                            orderby c.Name, c.Time
                            select c;
                        if (q.Count() > 0)
                        {
                            foreach (var c in q)
                            {
                                if (c.Name == "OK")
                                {
                                    label59.Text = c.Count.ToString();
                                    app.counter["OK"] = c.Count;
                                }
                                else if (c.Name == "NG")
                                {
                                    label6.Text = c.Count.ToString();
                                    app.counter["NG"] = c.Count;
                                }
                                else if (c.Name == "NULL")
                                {
                                    label14.Text = c.Count.ToString();
                                    app.counter["NULL"] = c.Count;
                                }
                            }
                        }

                        #region 在BTN9
                        /*
                        //if (!app.offline)
                        {
                            int totalcount = app.counter["OK"] + app.counter["NG"] + app.counter["NULL"];
                            int Qfirst = totalcount % 15;

                            PLC_SetD(803, 0);
                            PLC_SetD(805, 0);
                            PLC_SetD(807, app.counter["OK"]);
                            PLC_SetD(801, app.counter["NG"]);
                            PLC_SetD(809, app.counter["NULL"]);
                            PLC_SetD(813, totalcount);

                            PLC_SetD(99, Qfirst);

                            app.counter["stop0"] = totalcount;
                            app.counter["stop1"] = totalcount;
                            app.counter["stop2"] = totalcount;
                            app.counter["stop3"] = totalcount;
                            app.okr = 0;

                            updateLabel();
                            //Invoke(new Action(() => label4.Text = (PLC_CheckD32(1025) - app.okr).ToString()));
                        }
                        if (label4.Text == "0" && label6.Text == "0" && label14.Text == "0")
                        {
                            BeginInvoke(new Action(() => label22.Text = "0.0%"));
                            BeginInvoke(new Action(() => label23.Text = "0.0%"));
                            BeginInvoke(new Action(() => label17.Text = "0.0%"));
                        }
                        else
                        {
                            BeginInvoke(new Action(() => label22.Text = (double.Parse(label4.Text) * 100 / (int.Parse(label4.Text) + int.Parse(label6.Text) + int.Parse(label14.Text))).ToString("f1") + "%"));
                            BeginInvoke(new Action(() => label23.Text = (double.Parse(label6.Text) * 100 / (int.Parse(label4.Text) + int.Parse(label6.Text) + int.Parse(label14.Text))).ToString("f1") + "%"));
                            BeginInvoke(new Action(() => label17.Text = (double.Parse(label14.Text) * 100 / (int.Parse(label4.Text) + int.Parse(label6.Text) + int.Parse(label14.Text))).ToString("f1") + "%"));
                        }
                        */
                        #endregion
                    }

                }
                else
                {
                    MessageBox.Show("尚未選擇料號。");
                }
            }
            else
            {
                app.produce_No = "";
                app.foldername = "";
                button6.Text = "設定";
                comboBox1.Enabled = true;

                label32.Text = "";
                label33.Text = "";
                label34.Text = "";


                clear();
            }
        }
        #endregion
        #region 登入/登出
        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
            {
                login login = new login();
                var pw = "";
                login.ShowDialog();

                pw = login.TextBoxMsg;
                if (pw != "無")
                {
                    using (var db = new MydbDB())
                    {
                        var q =
                            from c in db.Users
                            where c.UserName == comboBox2.Text
                            orderby c.UserName
                            select c;
                        if (q.Count() > 0)
                        {
                            foreach (var c in q)
                            {
                                if (c.Password == pw)
                                {
                                    app.user = c.Level;
                                    app.username = c.UserName;
                                    if (app.user == 0)
                                    {
                                        Invoke(new Action(() => label35.Text = c.UserName + "(工程師)"));
                                        lbAdd(app.username + "已登入，權限為:工程師", "inf", "");

                                        button26.Visible = true;
                                        label52.Visible = true;
                                    }
                                    else if (app.user == 1)
                                    {
                                        Invoke(new Action(() => label35.Text = c.UserName + "(管理者)"));
                                        lbAdd(app.username + "已登入，權限為:管理者", "inf", "");
                                    }
                                    else
                                    {
                                        Invoke(new Action(() => label35.Text = c.UserName + "(作業員)"));
                                        lbAdd(app.username + "已登入，權限為:作業員", "inf", "");
                                    }
                                    button8.Enabled = true;
                                    button7.Enabled = false;
                                    comboBox2.Enabled = false;
                                    設定ToolStripMenuItem.Enabled = true;
                                }
                                else
                                {
                                    MessageBox.Show("密碼輸入錯誤。");
                                }
                            }
                        }
                    }

                    if (app.user < 2)
                    {
                        管理使用者ToolStripMenuItem.Enabled = true;
                    }
                    else
                    {
                        管理使用者ToolStripMenuItem.Enabled = false;
                    }

                }
            }
            else
            {
                MessageBox.Show("請先選擇使用者身分。");
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            Invoke(new Action(() => label35.Text = "未登入"));
            lbAdd("使用者:" + app.username + "已登出", "inf", "");
            app.user = 3;
            app.username = "";
            button8.Enabled = false;
            button7.Enabled = true;
            comboBox2.Enabled = true;
            管理使用者ToolStripMenuItem.Enabled = false;
            設定ToolStripMenuItem.Enabled = false;

            label52.Visible = false;
        }
        #endregion
        #region Lot_ID
        private void button9_Click(object sender, EventArgs e)
        {
            if (button9.Text == "設定")
            {
                if (textBox1.Text != "")
                {
                    app.LotID = textBox1.Text;
                    if (app.produce_No != "" && app.LotID != "")
                    {
                        app.foldername = app.produce_No + "-" + app.LotID;
                    }
                    lbAdd("設定Lot ID為:" + app.LotID.ToString(), "inf", "");
                    button9.Text = "變更";
                    textBox1.Enabled = false;

                    using (var db = new MydbDB())
                    {
                        var q =
                            from c in db.DefectCounts
                            where c.LotId == app.LotID && c.Type == app.produce_No
                            orderby c.Name, c.Time
                            select c;
                        if (q.Count() > 0)
                        {
                            foreach (var c in q)
                            {
                                if (c.Name == "OK")
                                {
                                    label59.Text = c.Count.ToString();
                                    app.counter["OK"] = c.Count;
                                }
                                else if (c.Name == "NG")
                                {
                                    label6.Text = c.Count.ToString();
                                    app.counter["NG"] = c.Count;
                                }
                                else if (c.Name == "NULL")
                                {
                                    label14.Text = c.Count.ToString();
                                    app.counter["NULL"] = c.Count;
                                }

                                // 導入sampleID，理論上全部都一樣
                                else if (c.Name == "SAMPLE_ID")
                                {
                                    app.counter["stop0"] = c.Count+1; //假設上次測到100個 這次就要從101開始
                                    app.counter["stop1"] = c.Count+1;
                                    app.counter["stop2"] = c.Count+1;
                                    app.counter["stop3"] = c.Count+1;

                                    // 由 GitHub Copilot 產生
                                    // 同步初始化 ResultManager.counter["SAMPLE_ID"] 為上次完成的最後一個 ID
                                    ResultManager.counter["SAMPLE_ID"] = c.Count;

                                    Log.Information($"sampleID：{c.Count}，已同步至 ResultManager.counter");
                                }
                            }
                        }

                        if (!app.offline)
                        {
                            int totalcount = app.counter["OK"] + app.counter["NG"] + app.counter["NULL"];
                            int Qfirst = totalcount % 15;

                            PLC_SetD(803, 0);
                            PLC_SetD(805, 0);
                            PLC_SetD(807, app.counter["OK"]);
                            PLC_SetD(801, app.counter["NG"]);
                            PLC_SetD(809, app.counter["NULL"]);
                            //PLC_SetD(813, totalcount);

                            PLC_SetD(97, app.counter["stop3"]%40); //直接取餘數就好，因為PLC那邊會先+1
                            PLC_SetD(98, app.counter["stop3"]%15);
                            PLC_SetD(99, app.counter["stop3"]%15);
                            PLC_SetD(101, app.counter["stop3"]-1);
                            Log.Information($"D97：{app.counter["stop3"] % 40}");
                            Log.Information($"D98：{app.counter["stop3"] % 15}");
                            Log.Information($"D99：{app.counter["stop3"] % 15}");
                            Log.Information($"D101：{app.counter["stop3"]-1}");

                            /*
                            button26.Text = "ON";
                            button26.BackColor = Color.FromArgb(128, 255, 128);
                            PLC_SetM(6, true);
                            button12.Text = "OFF";
                            button12.BackColor = Color.FromArgb(255, 128, 128);
                            PLC_SetM(9, true);
                            updateLabel();
                            */
                            /*
                            app.counter["stop0"] = totalcount;
                            app.counter["stop1"] = totalcount;
                            app.counter["stop2"] = totalcount;
                            app.counter["stop3"] = totalcount;
                            app.okr = 0;
                            */

                            //updateLabel();
                            //Invoke(new Action(() => label4.Text = (PLC_CheckD32(1025) - app.okr).ToString()));
                        }
                        if (label59.Text == "0" && label6.Text == "0" && label14.Text == "0")
                        {
                            BeginInvoke(new Action(() => label60.Text = "0.0%"));
                            BeginInvoke(new Action(() => label23.Text = "0.0%"));
                            BeginInvoke(new Action(() => label17.Text = "0.0%"));
                        }
                        else
                        {
                            BeginInvoke(new Action(() => label60.Text = (double.Parse(label59.Text) * 100 / (int.Parse(label59.Text) + int.Parse(label6.Text) + int.Parse(label14.Text))).ToString("f1") + "%"));
                            BeginInvoke(new Action(() => label23.Text = (double.Parse(label6.Text) * 100 / (int.Parse(label59.Text) + int.Parse(label6.Text) + int.Parse(label14.Text))).ToString("f1") + "%"));
                            BeginInvoke(new Action(() => label17.Text = (double.Parse(label14.Text) * 100 / (int.Parse(label59.Text) + int.Parse(label6.Text) + int.Parse(label14.Text))).ToString("f1") + "%"));
                        }
                    }
                }
                else
                {
                    MessageBox.Show("尚未輸入Lot_ID。");
                }
            }
            else
            {
                app.LotID = "";
                app.foldername = "";
                button9.Text = "設定";
                textBox1.Enabled = true;
                clear();
            }
        }
        #endregion
        #region 工單數量
        private void button10_Click(object sender, EventArgs e)
        {
            if (button10.Text == "設定")
            {
                if (textBox2.Text != "")
                {
                    if (int.Parse(textBox2.Text) != 0)
                    {
                        app.order = int.Parse(textBox2.Text);
                        lbAdd("設定工單數量為:" + app.order.ToString(), "inf", "");
                        button10.Text = "變更";
                        textBox2.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("最小工單數量為1。");
                    }
                }
                else
                {
                    MessageBox.Show("尚未輸入工單數量。");
                }
            }
            else
            {
                app.order = 0;
                button10.Text = "設定";
                textBox2.Enabled = true;
            }
        }

        #endregion
        #region Counter
        private void button11_Click(object sender, EventArgs e)
        {
            if (!app.offline)
            {
                PLC_SetD(803, 0);
                //app.counter["OK1"] = 0;
                label3.Text = "0";
                updateLabel();
            }
        }
        private void button42_Click(object sender, EventArgs e)
        {
            if (!app.offline)
            {
                PLC_SetD(805, 0);
                //app.counter["OK2"] = 0;
                label58.Text = "0";
                updateLabel();
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            if (!app.offline)
            {
                //app.counter["NG"] = 0;
                PLC_SetD(801, 0);
                label6.Text = "0";
                label23.Text = "0.0%";
                updateLabel();
            }
        }
        private void button14_Click(object sender, EventArgs e) //NG籃
        {
            if (!app.offline)
            {
                PLC_SetD(811, 0);
                //app.counter["NG"] = 0;
                //label6.Text = "0";
                label26.Text = "0";
                PLC_SetM(36, true);
                updateLabel();
            }
        }
        private void button15_Click(object sender, EventArgs e)
        {
            if (!app.offline)
            {
                //NULL
                PLC_SetD(809, 0);
                label14.Text = "0";
                label17.Text = "0.0%";
                updateLabel();
            }
        }
        private void button43_Click(object sender, EventArgs e)
        {
            if (!app.offline)
            {
                PLC_SetD(807, 0);
                label59.Text = "0";
                label60.Text = "0.0%";
                updateLabel();
            }
        }

        private void button41_Click(object sender, EventArgs e) // NULL籃
        {
            if (!app.offline)
            {
                PLC_SetD(813, 0);
                label55.Text = "0";
                PLC_SetM(36, true);
                updateLabel();
            }
        }
        #endregion
        #region 包裝數量
        private void button16_Click(object sender, EventArgs e)
        {
            if (button16.Text == "設定")
            {
                if (textBox3.Text != "")
                {
                    if (int.Parse(textBox3.Text) != 0)
                    {
                        app.pack = int.Parse(textBox3.Text);
                        PLC_SetD(7, int.Parse(textBox3.Text));
                        lbAdd("設定包裝數量為:" + app.pack.ToString(), "inf", "");
                        button16.Text = "變更";
                        textBox3.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("最小包裝數量為1。");
                    }
                }
                else
                {
                    MessageBox.Show("尚未輸入包裝數量。");
                }
            }
            else
            {
                app.pack = 0;
                button16.Text = "設定";
                textBox3.Enabled = true;
            }
        }
        
        #endregion
        #region 圖表頁面
        private void button18_Click(object sender, EventArgs e)
        {
            var data = new Dictionary<string, int>();
            var data_name = new List<string>();
            var last_ID = "";
            var count = 0;
            var stNum = 0;
            var notZero = 0;
            var total_count = 0;

            app.p2label_name.Clear();

            app.p2label_name.Add("料號");
            app.p2label_name.Add(comboBox3.Text);
            app.p2label_name.Add("Lot ID");
            app.p2label_name.Add(comboBox6.Text);
            app.p2label_name.Add("缺陷種類");
            app.p2label_name.Add("NG/OK數量");
            app.p2label_name.Add("百分比");

            List<DefectCount> dc = new List<DefectCount>();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectCounts
                    where c.LotId == app.LotID && c.Type == app.produce_No && (c.Name != "SAMPLE_ID" && c.Name != "NG" && c.Name != "NULL")
                    orderby c.Name, c.Time
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dc.Add(c);
                    }
                }
            }

            if (dc.Count == 0)
            {
                MessageBox.Show("無資料");
            }
            else //if (dc.Count <= 9)
            {
                for (int i = 0; i < dc.Count; i++)
                {
                    if (data.ContainsKey(dc[i].Name))
                    {
                        if (data[dc[i].Name] < dc[i].Count - stNum)
                        {
                            data[dc[i].Name] = dc[i].Count - stNum;
                            count = data[dc[i].Name];
                            last_ID = dc[i].LotId;
                        }
                    }
                    else
                    {
                        data.Add(dc[i].Name, 0);
                        data_name.Add(dc[i].Name);
                        last_ID = dc[i].LotId;
                        stNum = dc[i].Count;
                        count = data[dc[i].Name];
                    }
                }

                foreach (var i in data)
                {
                    total_count += i.Value;
                }

                for (int i = 0; i < data_name.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        app.p2label_name.Add(data_name[i]);
                        app.p2label_name.Add(data[data_name[i]].ToString());
                        app.p2label_name.Add(((double)data[data_name[i]] * 100 / total_count).ToString("f1") + "%");
                        notZero++;
                    }
                }
            }

            for (int i = 0; i < Math.Max(app.p2label_name.Count, app.p2label.Count); i++)
            {
                try
                {
                    if (app.p2label_name.Count > i)
                    {
                        app.p2label[i].Text = app.p2label_name[i];
                    }
                    else
                    {
                        app.p2label[i].Text = "";
                    }
                }
                catch
                {
                    app.p2label.Add(new Label());
                    app.p2label[i].Text = app.p2label_name[i];
                    this.tabPage2.Controls.Add(app.p2label[i]);
                    app.p2label[i].ForeColor = System.Drawing.Color.Black;
                    if (i < 4)
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + i % 2 * 200, 170 + i / 2 * 50);
                        app.p2label[i].Size = new System.Drawing.Size(200, 50);
                    }
                    else if (i < 7)
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + (i - 4) % 3 * 150, 210 + (i - 4) / 3 * 50 + 100);
                        app.p2label[i].Size = new System.Drawing.Size(150, 50);
                    }
                    else
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + (i - 7) % 3 * 150, 220 + (i - 7) / 3 * 50 + 150);
                        app.p2label[i].Size = new System.Drawing.Size(150, 50);
                    }
                    app.p2label[i].Name = "p2label" + i.ToString();
                    app.p2label[i].TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                }
            }

            Title Title = new Title
            {
                Text = "NG/OK數量",
                Alignment = ContentAlignment.MiddleCenter,
                Font = new Font("微軟正黑體", 18F, FontStyle.Bold)
            };

            if (chart1.Titles.Count > 0)
            {
                chart1.Titles.RemoveAt(0);
                chart1.Series.Clear();
            }
            if (dc.Count > 0)
            {
                chart1.Titles.Add(Title);

                if (chart1.Series.Count == 0)
                {
                    chart1.Series.Add("瑕疵個數");
                }

                chart1.Series["瑕疵個數"].ChartType = SeriesChartType.Column;
                chart1.Series["瑕疵個數"].Color = Color.Blue;
                chart1.Series["瑕疵個數"].BorderWidth = 3;
                chart1.Series["瑕疵個數"].XValueType = ChartValueType.String;
                chart1.Series["瑕疵個數"].YValueType = ChartValueType.Int64;
            }
            string[] X_title = new string[dc.Count];
            int[] Y_count = new int[dc.Count];

            if (dc.Count == 0)
            {
            }
            else //if (dc.Count <= 9)
            {
                X_title = new string[notZero];
                Y_count = new int[notZero];

                var c = 0;

                for (int i = 0; i < data.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        X_title[c] = data_name[i];
                        Y_count[c] = data[data_name[i]];
                        c++;
                    }
                }
            }

            if (dc.Count > 0)
            {
                chart1.Series["瑕疵個數"].Points.DataBindXY(X_title, Y_count);
                chart1.Series["瑕疵個數"].ToolTip = "#VALX:#VAL (個)";

                ChartArea area = chart1.ChartAreas[0];
                area.AxisX.MajorGrid.LineWidth = 0;
                area.AxisY.Title = "個";
                area.AxisY.TextOrientation = TextOrientation.Horizontal;
                if (dc.Count > 9)
                {
                    //設置X軸座標的間隔為1
                    chart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                    //設置X軸座標偏移為1
                    chart1.ChartAreas["ChartArea1"].AxisX.IntervalOffset = 1;
                    //設置是否交錯顯示，比如數據多時分成兩行來顯示
                    chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.IsStaggered = true; //如果 X 軸要設成文字轉向，這一行就要 MARK 起來不作交錯顯示
                }
            }
        }
        private void button19_Click(object sender, EventArgs e)
        {
            var data = new Dictionary<string, int>();
            var data_name = new List<string>();
            var count = 0;
            var stNum = 0;
            var total_count = 0;
            var notZero = 0;

            double start = (dateTimePicker1.Value.AddMinutes(0)).ToOADate();
            double end = (dateTimePicker2.Value.AddMinutes(0)).ToOADate();

            // 存儲每種瑕疵類型的基準值
            Dictionary<string, int> baselineCounts = new Dictionary<string, int>();

            List<DefectCount> dc = new List<DefectCount>();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectCounts
                    where c.Time > DateTime.FromOADate(start) && c.Time < DateTime.FromOADate(end) 
                                    && c.Type == comboBox3.Text && c.LotId == comboBox6.Text 
                                    && (c.Name != "SAMPLE_ID" && c.Name != "NG" && c.Name != "NULL")
                    orderby c.Name, c.Time
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dc.Add(c);
                    }
                    // 對於每種瑕疵類型，尋找時間範圍之前的最後一筆記錄
                    var defectTypes = dc.Select(d => d.Name).Distinct().ToList();
                    foreach (var defectType in defectTypes)
                    {
                        // 查找此瑕疵類型在時間範圍前的最後一筆記錄
                        var baseline = db.DefectCounts
                            .Where(c => c.Time < DateTime.FromOADate(start)
                                   && c.Type == comboBox3.Text
                                   && c.LotId == comboBox6.Text
                                   && c.Name == defectType)
                            .OrderByDescending(c => c.Time)
                            .FirstOrDefault();

                        // 若存在基準記錄，使用其計數值；否則使用0
                        baselineCounts[defectType] = baseline?.Count ?? 0;
                    }
                }
            }
            app.p2label_name.Clear();

            app.p2label_name.Add("料號");
            app.p2label_name.Add(comboBox3.Text);
            app.p2label_name.Add("Lot ID");
            app.p2label_name.Add(comboBox6.Text);
            app.p2label_name.Add("缺陷種類");
            app.p2label_name.Add("NG/OK數量");
            app.p2label_name.Add("百分比");

            if (dc.Count == 0)
            {
                MessageBox.Show("無資料");
            }
            else //if (dc.Count <= 9)
            {
                for (int i = 0; i < dc.Count; i++)
                {
                    // 獲取當前瑕疵類型
                    string defectType = dc[i].Name;

                    if (data.ContainsKey(defectType))
                    {
                        if (dc[i].Count - stNum >= count)
                        {
                            if (i == 0)
                            {
                                //stNum = dc[i].Count;

                                // 使用時間範圍前的基準值，而不是第一筆記錄的值
                                stNum = baselineCounts.ContainsKey(defectType) ? baselineCounts[defectType] : 0;                            
                            }
                            data[defectType] = dc[i].Count - stNum;
                            count = data[defectType];
                        }
                        else
                        {
                            data[defectType] = dc[i].Count - dc[i - 1].Count + count;
                            count += dc[i].Count - dc[i - 1].Count;
                        }
                    }
                    else
                    {
                        data.Add(defectType, 0);
                        data_name.Add(defectType);
                        //stNum = dc[i].Count;
                        // 使用時間範圍前的基準值
                        stNum = baselineCounts.ContainsKey(defectType) ? baselineCounts[defectType] : 0;
                        data[defectType] = dc[i].Count - stNum;
                        count = data[defectType];
                    }
                }
                foreach (var i in data)
                {
                    total_count += i.Value;
                }
                for (int i = 0; i < data_name.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        app.p2label_name.Add(data_name[i]);
                        app.p2label_name.Add(data[data_name[i]].ToString());
                        app.p2label_name.Add(((double)data[data_name[i]] * 100 / total_count).ToString("f1") + "%");
                        notZero++;
                    }
                }
            }

            for (int i = 0; i < Math.Max(app.p2label_name.Count, app.p2label.Count); i++)
            {
                try
                {
                    if (app.p2label_name.Count > i)
                    {
                        app.p2label[i].Text = app.p2label_name[i];
                    }
                    else
                    {
                        app.p2label[i].Text = "";
                    }
                }
                catch
                {
                    app.p2label.Add(new Label());
                    app.p2label[i].Text = app.p2label_name[i];
                    this.tabPage2.Controls.Add(app.p2label[i]);
                    app.p2label[i].ForeColor = System.Drawing.Color.Black;

                    if (i < 4)
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + i % 2 * 200, 170 + i / 2 * 50);
                        app.p2label[i].Size = new System.Drawing.Size(200, 50);
                    }
                    else if (i < 7)
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + (i - 4) % 3 * 150, 210 + (i - 4) / 3 * 50 + 100);
                        app.p2label[i].Size = new System.Drawing.Size(150, 50);
                    }
                    else
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + (i - 7) % 3 * 150, 220 + (i - 7) / 3 * 50 + 150);
                        app.p2label[i].Size = new System.Drawing.Size(150, 50);
                    }

                    app.p2label[i].Name = "p2label" + i.ToString();
                    app.p2label[i].TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                }
            }
            Title Title = new Title
            {
                Text = "NG/OK數量",
                Alignment = ContentAlignment.MiddleCenter,
                Font = new Font("微軟正黑體", 18F, FontStyle.Bold)
            };

            if (chart1.Titles.Count > 0)
            {
                chart1.Titles.RemoveAt(0);
                chart1.Series.Clear();
            }
            if (dc.Count > 0)
            {
                chart1.Titles.Add(Title);

                if (chart1.Series.Count == 0)
                {
                    chart1.Series.Add("瑕疵個數");
                }

                chart1.Series["瑕疵個數"].ChartType = SeriesChartType.Column;
                chart1.Series["瑕疵個數"].Color = Color.Blue;
                chart1.Series["瑕疵個數"].BorderWidth = 5;
                chart1.Series["瑕疵個數"].XValueType = ChartValueType.String;
                chart1.Series["瑕疵個數"].YValueType = ChartValueType.Int64;
            }
            string[] X_title = new string[dc.Count];
            int[] Y_count = new int[dc.Count];

            if (dc.Count == 0)
            {
            }
            else //if (dc.Count <= 9)
            {
                X_title = new string[notZero];
                Y_count = new int[notZero];
                var c = 0;

                for (int i = 0; i < data.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        X_title[c] = data_name[i];
                        Y_count[c] = data[data_name[i]];
                        c++;
                    }
                }
            }

            if (dc.Count > 0)
            {
                chart1.Series["瑕疵個數"].Points.DataBindXY(X_title, Y_count);
                chart1.Series["瑕疵個數"].ToolTip = "#VALX:#VAL (個)";

                ChartArea area = chart1.ChartAreas[0];
                area.AxisX.MajorGrid.LineWidth = 0;
                area.AxisY.Title = "個";
                area.AxisY.TextOrientation = TextOrientation.Horizontal;
                if (dc.Count > 9)
                {
                    //設置X軸座標的間隔為1
                    chart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                    //設置X軸座標偏移為1
                    chart1.ChartAreas["ChartArea1"].AxisX.IntervalOffset = 1;
                    //設置是否交錯顯示，比如數據多時分成兩行來顯示
                    chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.IsStaggered = true; //如果 X 軸要設成文字轉向，這一行就要 MARK 起來不作交錯顯示
                }
            }
        }
        private void button20_Click(object sender, EventArgs e)
        {
            var data = new Dictionary<string, int>();
            var data_name = new List<string>();
            var count = 0;
            var stNum = 0;
            var total_count = 0;
            var notZero = 0;

            app.p2label_name.Clear();

            app.p2label_name.Add("料號");
            app.p2label_name.Add(comboBox3.Text);
            app.p2label_name.Add("Lot ID");
            app.p2label_name.Add(comboBox6.Text);
            app.p2label_name.Add("缺陷種類");
            app.p2label_name.Add("NG/OK數量");
            app.p2label_name.Add("百分比");

            double start = (dateTimePicker1.Value.AddMinutes(0)).ToOADate();
            double end = (dateTimePicker2.Value.AddMinutes(0)).ToOADate();
            string selectedType = comboBox3.Text;
            string selectedLotId = comboBox6.Text;

            // 存儲基準值（時間範圍之前的記錄）
            Dictionary<string, int> baselineCounts = new Dictionary<string, int>();

            List<DefectCount> dc = new List<DefectCount>();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectCounts
                    where c.Time > DateTime.FromOADate(start) && c.Time < DateTime.FromOADate(end) && c.Type == comboBox3.Text && c.LotId == comboBox6.Text && (c.Name != "SAMPLE_ID" && c.Name != "NG" && c.Name != "NULL")
                    orderby c.Name, c.Time
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dc.Add(c);
                    }
                    // 獲取瑕疵類型列表
                    var defectTypes = dc.Select(d => d.Name).Distinct().ToList();

                    // 對每種瑕疵，查找時間範圍之前的最後一筆記錄
                    foreach (var defectType in defectTypes)
                    {
                        var baselineQuery = from c in db.DefectCounts
                                            where c.Time < DateTime.FromOADate(start)
                                                  && c.Type == selectedType && c.LotId == selectedLotId
                                                  && c.Name == defectType
                                            orderby c.Time descending
                                            select c;

                        var baselineRecord = baselineQuery.FirstOrDefault();
                        if (baselineRecord != null)
                        {
                            baselineCounts[defectType] = baselineRecord.Count;
                        }
                        else
                        {
                            baselineCounts[defectType] = 0;
                        }
                    }
                }
            }
            

            if (dc.Count == 0)
            {
                MessageBox.Show("無資料");
            }
            else //if (dc.Count <= 9)
            {
                for (int i = 0; i < dc.Count; i++)
                {
                    if (data.ContainsKey(dc[i].Name))
                    {
                        if (dc[i].Count - stNum >= count)
                        {
                            if (i == 0)
                            {
                                // 使用基準值作為起點
                                stNum = baselineCounts.ContainsKey(dc[i].Name) ? baselineCounts[dc[i].Name] : 0;
                            }
                            data[dc[i].Name] = dc[i].Count - stNum;
                            count = data[dc[i].Name];
                        }
                        else
                        {
                            data[dc[i].Name] = dc[i].Count - dc[i - 1].Count + count;
                            count += dc[i].Count - dc[i - 1].Count;
                        }
                    }
                    else
                    {
                        data.Add(dc[i].Name, 0);
                        data_name.Add(dc[i].Name);
                        // 使用基準值作為起點
                        stNum = baselineCounts.ContainsKey(dc[i].Name) ? baselineCounts[dc[i].Name] : 0;
                        data[dc[i].Name] = dc[i].Count - stNum;
                        count = data[dc[i].Name];
                    }
                }

                foreach (var i in data)
                {
                    total_count += i.Value;
                }

                for (int i = 0; i < data_name.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        app.p2label_name.Add(data_name[i]);
                        app.p2label_name.Add(data[data_name[i]].ToString());
                        app.p2label_name.Add(((double)data[data_name[i]] * 100 / total_count).ToString("f1") + "%");
                        notZero++;
                    }
                }
                #region old
                /*
                for (int i = 0; i < dc.Count; i++)
                {
                    if (data.ContainsKey(dc[i].Name))
                    {
                        if (dc[i].Count - stNum >= count)
                        {
                            if (i == 0)
                            {
                                stNum = dc[i].Count;
                            }
                            data[dc[i].Name] = dc[i].Count - stNum;
                            count = data[dc[i].Name];
                        }
                        else
                        {
                            data[dc[i].Name] = dc[i].Count - dc[i - 1].Count + count;
                            count += dc[i].Count - dc[i - 1].Count;
                        }
                    }
                    else
                    {
                        data.Add(dc[i].Name, 0);
                        data_name.Add(dc[i].Name);
                        stNum = dc[i].Count;
                        count = data[dc[i].Name];
                    }
                }
                foreach (var i in data)
                {
                    total_count += i.Value;
                }

                for (int i = 0; i < data_name.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        app.p2label_name.Add(data_name[i]);
                        app.p2label_name.Add(data[data_name[i]].ToString());
                        app.p2label_name.Add(((double)data[data_name[i]] * 100 / total_count).ToString("f1") + "%");
                        notZero++;
                    }
                }
                */
                #endregion
            }

            for (int i = 0; i < Math.Max(app.p2label_name.Count, app.p2label.Count); i++)
            {
                try
                {
                    if (app.p2label_name.Count > i)
                    {
                        app.p2label[i].Text = app.p2label_name[i];
                    }
                    else
                    {
                        app.p2label[i].Text = "";
                    }
                }
                catch
                {
                    app.p2label.Add(new Label());
                    app.p2label[i].Text = app.p2label_name[i];
                    this.tabPage2.Controls.Add(app.p2label[i]);
                    app.p2label[i].ForeColor = System.Drawing.Color.Black;
                    if (i < 4)
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + i % 2 * 200, 170 + i / 2 * 50);
                        app.p2label[i].Size = new System.Drawing.Size(200, 50);
                    }
                    else if (i < 7)
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + (i - 4) % 3 * 150, 210 + (i - 4) / 3 * 50 + 100);
                        app.p2label[i].Size = new System.Drawing.Size(150, 50);
                    }
                    else
                    {
                        app.p2label[i].Font = new System.Drawing.Font("新細明體", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
                        app.p2label[i].Location = new System.Drawing.Point(5 + (i - 7) % 3 * 150, 220 + (i - 7) / 3 * 50 + 150);
                        app.p2label[i].Size = new System.Drawing.Size(150, 50);
                    }
                    app.p2label[i].Name = "p2label" + i.ToString();
                    app.p2label[i].TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                }
            }

            Title Title = new Title
            {
                Text = "NG/OK數量",
                Alignment = ContentAlignment.MiddleCenter,
                Font = new Font("微軟正黑體", 18F, FontStyle.Bold)
            };

            if (chart1.Titles.Count > 0)
            {
                chart1.Titles.RemoveAt(0);
                chart1.Series.Clear();
            }
            if (dc.Count > 0)
            {
                chart1.Titles.Add(Title);

                if (chart1.Series.Count == 0)
                {
                    chart1.Series.Add("瑕疵個數");
                }

                chart1.Series["瑕疵個數"].ChartType = SeriesChartType.Column;
                chart1.Series["瑕疵個數"].Color = Color.Blue;
                chart1.Series["瑕疵個數"].BorderWidth = 5;
                chart1.Series["瑕疵個數"].XValueType = ChartValueType.String;
                chart1.Series["瑕疵個數"].YValueType = ChartValueType.Int64;
            }
            string[] X_title = new string[dc.Count];
            int[] Y_count = new int[dc.Count];

            if (dc.Count == 0)
            {
            }
            else //if (dc.Count <= 9)
            {
                X_title = new string[notZero];
                Y_count = new int[notZero];
                var c = 0;
                for (int i = 0; i < data.Count; i++)
                {
                    if (data[data_name[i]] != 0)
                    {
                        X_title[c] = data_name[i];
                        Y_count[c] = data[data_name[i]];
                        c++;
                    }
                }
            }

            if (dc.Count > 0)
            {
                chart1.Series["瑕疵個數"].Points.DataBindXY(X_title, Y_count);
                chart1.Series["瑕疵個數"].ToolTip = "#VALX:#VAL (個)";

                ChartArea area = chart1.ChartAreas[0];
                area.AxisX.MajorGrid.LineWidth = 0;
                area.AxisY.Title = "個";
                area.AxisY.TextOrientation = TextOrientation.Horizontal;
                if (dc.Count > 9)
                {
                    //設置X軸座標的間隔為1
                    chart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                    //設置X軸座標偏移為1
                    chart1.ChartAreas["ChartArea1"].AxisX.IntervalOffset = 1;
                    //設置是否交錯顯示，比如數據多時分成兩行來顯示
                    chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.IsStaggered = true; //如果 X 軸要設成文字轉向，這一行就要 MARK 起來不作交錯顯示
                }
            }
        }
        private void button21_Click(object sender, EventArgs e)
        {
            double start = (dateTimePicker3.Value.AddMinutes(0)).ToOADate();
            double end = (dateTimePicker4.Value.AddMinutes(0)).ToOADate();

            string selectedType = comboBox4.Text;
            string selectedDefect = comboBox5.Text;
            string selectedLotId = comboBox7.Text;

            // 查詢時間範圍內的資料
            List<DefectCount> dc = new List<DefectCount>();

            // 查找時間範圍之前的最後一筆記錄作為基準值
            int baselineCount = 0;


            using (var db = new MydbDB())
            {
                // 查找基準值（時間範圍之前的最後一筆記錄）
                var baseline = db.DefectCounts
                                .Where(c => c.Time < DateTime.FromOADate(start)
                                       && c.Type == selectedType
                                       && c.Name == selectedDefect
                                       && c.LotId == selectedLotId)
                                .OrderByDescending(c => c.Time)
                                .FirstOrDefault();

                // 如果找到基準值，使用它；否則使用0
                if (baseline != null)
                {
                    baselineCount = baseline.Count;
                }

                // 查詢時間範圍內的數據
                var q = from c in db.DefectCounts
                        where c.Time >= DateTime.FromOADate(start) && c.Time <= DateTime.FromOADate(end)
                              && c.Type == selectedType && c.Name == selectedDefect && c.LotId == selectedLotId
                        orderby c.Time
                        select c;

                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        dc.Add(c);
                    }
                }
            }

            Title Title = new Title
            {
                Text = selectedDefect,
                Alignment = ContentAlignment.MiddleCenter,
                Font = new Font("微軟正黑體", 18F, FontStyle.Bold)
            };
            if (chart2.Titles.Count > 0)
            {
                chart2.Titles.RemoveAt(0);
                chart2.Series.Clear();
            }

            if (dc.Count > 0)
            {
                chart2.Titles.Add(Title);

                if (chart2.Series.Count == 0)
                {
                    chart2.Series.Add("Line");
                }

                chart2.Series["Line"].ChartType = SeriesChartType.Line;
                chart2.Series["Line"].Color = Color.Blue;
                chart2.Series["Line"].BorderWidth = 5;
                chart2.Series["Line"].XValueType = ChartValueType.String;
                chart2.Series["Line"].YValueType = ChartValueType.Int64;

            }
            string[] X_title = new string[9];
            int[] Y_count = new int[9];

            X_title = new string[dc.Count];
            Y_count = new int[dc.Count];

            if (dc.Count == 0)
            {
                MessageBox.Show("無資料");
            }
            else //if (dc.Count <= 9)
            {
                X_title = new string[dc.Count];
                Y_count = new int[dc.Count];
                var last_ID = "";
                var total = 0;

                //var stNum = 0;
                // 使用查詢到的基準值作為起點
                var stNum = baselineCount;

                for (int i = 0; i < dc.Count; i++)
                {
                    if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year && dateTimePicker3.Value.Month == dateTimePicker4.Value.Month && dateTimePicker3.Value.Date == dateTimePicker4.Value.Date)
                    {
                        X_title[i] = dc[i].Time.ToString("HH:mm");
                    }
                    else if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year && dateTimePicker3.Value.Month == dateTimePicker4.Value.Month)
                    {
                        X_title[i] = dc[i].Time.ToString("MM/dd HH:mm");
                    }
                    else if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year)
                    {
                        X_title[i] = dc[i].Time.ToString("MM/dd HH:mm");
                    }
                    else
                    {
                        X_title[i] = dc[i].Time.ToString("yyyy/MM/dd");
                    }
                    /*
                    if (dc[i].Count - stNum >= total)
                    {
                        if (i == 0)
                        {
                            stNum = dc[i].Count;
                        }*/
                    if (i == 0)
                    {
                        Y_count[i] = dc[i].Count - stNum;
                        total = Y_count[i];
                        last_ID = dc[i].LotId;
                    }
                    else
                    {
                        if (dc[i].LotId != last_ID)
                        {
                            last_ID = dc[i].LotId;
                            Y_count[i] = dc[i].Count + total;
                            total += dc[i].Count;
                        }
                        else
                        {
                            Y_count[i] = dc[i].Count - dc[i - 1].Count + total;
                            total += dc[i].Count - dc[i - 1].Count;
                        }
                    }
                }
            }
            #region 固定9組(目前沒用)
            //else
            //{
            //    var st = dc[0].Time;
            //    var ed = dc[dc.Count - 1].Time;

            //    var last = 0;
            //    var last_ID = "";
            //    var total = 0;
            //    var same = false;
            //    for (int i = 0; i < 9; i++)
            //    {
            //        var time = st + TimeSpan.FromSeconds((ed - st).TotalSeconds / 9 * (i + 1));
            //        for (int j = last; j < dc.Count; j++)
            //        {
            //            if (dc[j].Count >= total)
            //            {
            //                Y_count[i] = dc[j].Count;
            //                total = dc[j].Count;
            //                Console.WriteLine("-"+1);
            //            }
            //            else
            //            {
            //                if (dc[j].LotId != dc[j - 1].LotId)
            //                {
            //                    if (same)
            //                    {
            //                        Y_count[i] = total;
            //                        Console.WriteLine("-" + 2);
            //                    }
            //                    else
            //                    {
            //                        Y_count[i] = dc[j].Count + total;
            //                        total += dc[j].Count;
            //                        Console.WriteLine("-" + 3);
            //                    }                                
            //                }
            //                else
            //                {
            //                    if (same)
            //                    {
            //                        Y_count[i] = total;
            //                        Console.WriteLine("-" + 2);
            //                    }
            //                    else
            //                    {
            //                        Y_count[i] = dc[j].Count - dc[j - 1].Count + total;
            //                        total += dc[j].Count - dc[j - 1].Count;
            //                        Console.WriteLine("-" + 4);
            //                    }                                
            //                }
            //            }
            //            if (j == dc.Count - 1)
            //            {
            //                if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year && dateTimePicker3.Value.Month == dateTimePicker4.Value.Month && dateTimePicker3.Value.Date == dateTimePicker4.Value.Date)
            //                {
            //                    X_title[i] = time.ToString("HH:mm");
            //                }
            //                else if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year && dateTimePicker3.Value.Month == dateTimePicker4.Value.Month)
            //                {
            //                    X_title[i] = time.ToString("MM/dd HH:mm");
            //                }
            //                else if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year)
            //                {
            //                    X_title[i] = time.ToString("MM/dd HH:mm");
            //                }
            //                else
            //                {
            //                    X_title[i] = time.ToString("yyyy/MM/dd");
            //                }

            //                last = j;
            //                same = true;
            //                break;
            //            }
            //            else
            //            {
            //                if (dc[j].Time < time && dc[j + 1].Time > time)
            //                {
            //                    if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year && dateTimePicker3.Value.Month == dateTimePicker4.Value.Month && dateTimePicker3.Value.Date == dateTimePicker4.Value.Date)
            //                    {
            //                        X_title[i] = time.ToString("HH:mm");
            //                    }
            //                    else if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year && dateTimePicker3.Value.Month == dateTimePicker4.Value.Month)
            //                    {
            //                        X_title[i] = time.ToString("MM/dd HH:mm");
            //                    }
            //                    else if (dateTimePicker3.Value.Year == dateTimePicker4.Value.Year)
            //                    {
            //                        X_title[i] = time.ToString("MM/dd HH:mm");
            //                    }
            //                    else
            //                    {
            //                        X_title[i] = time.ToString("yyyy/MM/dd");
            //                    }
            //                    Console.WriteLine(j);
            //                    last = j;
            //                    same = true;
            //                    break;
            //                }
            //            }
            //            same = false;
            //        }
            //    }
            //}
            #endregion

            if (dc.Count > 0)
            {
                chart2.Series["Line"].Points.DataBindXY(X_title, Y_count);
                chart2.Series["Line"].ToolTip = "#VALX:#VAL (個)";

                ChartArea area = chart2.ChartAreas[0];
                area.AxisX.MajorGrid.LineWidth = 0;
                area.AxisY.Title = "個";
                area.AxisY.TextOrientation = TextOrientation.Horizontal;
            }
        }
        #endregion
        #endregion

        #endregion

        void read_plc()
        {
            while (true)
            {
                if (!app.offline)
                {
                    if (app.status)
                    {
                        try
                        {
                            #region 監控第一個樣品
                            if (!_systemInitialized && !app.offline)
                            {
                                lock (_initLock)
                                {
                                    if (!_systemInitialized)
                                    {
                                        int currentD100 = PLC_CheckD(100);

                                        // D100 > 0 代表第一個物體經過第一站光纖
                                        if (currentD100 > 0)
                                        {
                                            _systemInitialized = true;
                                            Log.Information($"首次偵測到物體（D100 = {currentD100}），系統初始化完成，開始正常收圖");
                                        }
                                    }
                                }
                            }
                            #endregion

                            if (app.plc_stop)
                            {
                                this.Enabled = true;
                                button2_Click(null, null);
                                app.plc_stop = false;
                                app.alertTriggered = false; // 重置標記
                                this.Refresh();
                            }

                            else
                            {
                                updateLabel();
                                #region 監控OK1 OK2 未啟用 （高頻率讀取無法監控）
                                /*
                                // 讀 D803/D805/D807（32bit）
                                int curD803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803);
                                int curD805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805);
                                int curD807 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 807);

                                if (_prevD803 < 0) _prevD803 = curD803;
                                if (_prevD805 < 0) _prevD805 = curD805;
                                if (_prevD807 < 0) _prevD807 = curD807;

                                bool hasProduction = false;

                                // === 監控 D803 (OK1) ===
                                if (curD803 > _prevD803)
                                {
                                    DateTime plcPushTime = DateTime.Now; // **實際推料時間**
                                    int delta = curD803 - _prevD803;

                                    if (delta > 3) 
                                        Log.Warning($"[PLC推料] D803 連跳 {delta}，可能有遺漏");

                                    for (int i = 0; i < delta; i++)
                                    {
                                        if (app.pendingOK1.TryDequeue(out int sid))
                                        {
                                            // 從字典取得拍照時間
                                            if (app.samplePhotoTimes.TryGetValue(sid, out DateTime photoTime))
                                            {
                                                double totalMs = (plcPushTime - photoTime).TotalMilliseconds;

                                                // **統一Log格式 - OK1實際推料**
                                                Log.Information($"[樣品 {sid}] PLC實際推料:{plcPushTime:HH:mm:ss.fff} | 通道:OK1 | 總耗時:{totalMs:F0}ms");
                                            }
                                            else
                                            {
                                                Log.Warning($"[樣品 {sid}] PLC推料完成，但找不到拍照時間記錄，此次推料時間:{plcPushTime:HH:mm:ss.fff}");
                                            }
                                        }
                                        else
                                        {
                                            Log.Warning($"[PLC推料] D803 增加但 pendingOK1 佇列為空");
                                        }
                                    }
                                    _prevD803 = curD803;
                                    hasProduction = true;
                                }

                                // === 監控 D805 (OK2) ===
                                if (curD805 > _prevD805)
                                {
                                    DateTime plcPushTime = DateTime.Now;
                                    int delta = curD805 - _prevD805;

                                    if (delta > 3) 
                                        Log.Warning($"[PLC推料] D805 連跳 {delta}，可能有遺漏");

                                    for (int i = 0; i < delta; i++)
                                    {
                                        if (app.pendingOK2.TryDequeue(out int sid))
                                        {
                                            if (app.samplePhotoTimes.TryGetValue(sid, out DateTime photoTime))
                                            {
                                                double totalMs = (plcPushTime - photoTime).TotalMilliseconds;

                                                // **統一Log格式 - OK2實際推料**
                                                Log.Information($"[樣品 {sid}] PLC實際推料:{plcPushTime:HH:mm:ss.fff} | 通道:OK2 | 總耗時:{totalMs:F0}ms");
                                            }
                                            else
                                            {
                                                Log.Warning($"[樣品 {sid}] PLC推料完成，但找不到拍照時間記錄");
                                            }
                                        }
                                        else
                                        {
                                            Log.Warning($"[PLC推料] D805 增加但 pendingOK2 佇列為空");
                                        }
                                    }
                                    _prevD805 = curD805;
                                    hasProduction = true;
                                }

                                // === 監控 D807 (OK總數) - 僅作驗證 ===
                                if (curD807 > _prevD807)
                                {
                                    int delta = curD807 - _prevD807;
                                    if (delta > 3) 
                                        Log.Warning($"[PLC推料] D807 連跳 {delta}");
                                    _prevD807 = curD807;
                                }

                                // 若有生產輸出，重置NULL計數與時間
                                if (hasProduction)
                                {

                                    if (_consecutiveNullCount > 0)
                                    {
                                        Log.Information($"[Production] 恢復正常生產，清除NULL計數(原={_consecutiveNullCount})");
                                    }

                                    //_consecutiveNullCount = 0;
                                    //_lastTotalCount = currentTotalCount;
                                    _lastProductiveTime = DateTime.Now;
                                }
                                else if(_lastProductiveTime != null)
                                {

                                    // 檢查無生產輸出逾時
                                    TimeSpan idleSpan = DateTime.Now - _lastProductiveTime.Value;
                                    double idleSeconds = idleSpan.TotalSeconds;
                                    if (idleSeconds > app.ProductiveTimeoutSec)
                                    {
                                        Log.Error($"[Production] 無生產輸出逾時({idleSeconds:F0}秒 > {app.ProductiveTimeoutSec}秒)，觸發停機");
                                        this.Invoke(new Action(() =>
                                        {
                                            MessageBox.Show($"已連續{idleSeconds:F0}秒NULL，系統將停機檢查，請進入復歸流程", "生產異常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            button2_Click(null, null); // 停止檢測
                                        }));
                                        _lastProductiveTime = null; // 重置避免重複觸發
                                    }

                                }
                                */
                                #endregion

                                bool hasError = !PLC_CheckX(0) || PLC_CheckM(27) || PLC_CheckM(21);

                                if (hasError && !app.alertTriggered) // 只在首次檢測到錯誤時觸發
                                {
                                    app.alertTriggered = true;
                                    this.BeginInvoke(new Action(() => ShowAlert()));
                                }
                                else if (!hasError && app.alertTriggered)
                                {
                                    app.alertTriggered = false; // 錯誤解除時重置標記
                                }
                                else
                                {
                                    // 在介面更新檔板狀態
                                    //if ((PLC_CheckM(6) || PLC_CheckM(7)) && button26.Text != "ON")
                                    /*
                                    if ((PLC_Value.Point_Y[2][3] || PLC_Value.Point_Y[2][4]) && button26.Text != "ON") //OK檔板ON
                                    {
                                        BeginInvoke(new Action(() => button26.Text = "ON"));
                                        BeginInvoke(new Action(() => button26.BackColor = button26.BackColor = Color.Lime));
                                    }
                                    */

                                    #region 每分紀錄
                                    try
                                    {
                                        if ((DateTime.Now - app.rec_time).TotalSeconds >= 65)
                                        {
                                            var handle1 = new ManualResetEvent(false);
                                            PLC_ModBus.CheckValue(1, ValueUnit.D, 801, 14, false, handle1);
                                            handle1.WaitOne();
                                            var d801 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 801);
                                            var d803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803);
                                            var d805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805);
                                            var d807 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 807);
                                            var d809 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 809);
                                            //var d809 = PLC_CheckD(809);
                                            var d811 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 811);
                                            //var d811 = PLC_CheckD(811);
                                            var d813 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 813);

                                            // 修正：將 OK 數量取整到 50 的倍數（向下取整）
                                            int q = d807 / 50;
                                            d807 = q * 50;  // 正確：例如 d807=123 -> q=2 -> d807=100

                                            var ngCount = d801;
                                            var okCount = d807;
                                            var nullCount = d809;

                                            Dictionary<string, int> generalCounts = new Dictionary<string, int>
                                            {
                                                { "OK", okCount },
                                                { "NG", ngCount },
                                                { "NULL", nullCount }
                                            };

                                            // 使用 DefectCountManager 寫入數據
                                            bool writeSuccess = DefectCountManager.WriteAllDefectCounts(
                                                app.produce_No,
                                                app.LotID,
                                                app.dc,
                                                generalCounts
                                            );

                                            if (!writeSuccess)
                                            {
                                                Log.Error("定期寫入 DefectCount 失敗");
                                            }

                                            app.rec_time = DateTime.Now;
                                        }

                                    }
                                    catch (Exception e1)
                                    {
                                        Log.Error("readplc error:" + e1.ToString());
                                    }
                                    #endregion
                                    /*
                                    try
                                    {
                                        if ((DateTime.Now - app.lastIn).TotalSeconds > int.Parse(app.param["empty_th0"]))
                                        {
                                            PLC_SetM(999, true);
                                            //PLC_SetM(20, true);
                                            button2_Click(null, null);
                                            lbAdd("無進料時間超出門檻，自動停機。", "err", "");
                                        }
                                    }
                                    catch (Exception e1)
                                    {
                                        Log.Error("readplc error:" + e1.ToString());
                                        //lbAdd("readplc wrong2", "inf", e1.ToString());
                                    }
                                    */
                                    /*
                                    #region 速度
                                    try
                                    {
                                        if (app.finishtime.Count > 0)
                                        {
                                            while ((DateTime.Now - app.finishtime[0]).TotalMinutes >= 1)
                                            {
                                                app.finishtime.RemoveAt(0);
                                                if (app.finishtime.Count == 0)
                                                {
                                                    break;
                                                }
                                            }

                                            if ((DateTime.Now - app.finishtime[app.finishtime.Count - 1]).TotalSeconds > 1)
                                            {
                                                BeginInvoke(new Action(() => label13.Text = "0 PCS/分"));
                                            }
                                        }

                                        var oneminspeed = app.finishtime.Count;
                                        BeginInvoke(new Action(() => label51.Text = oneminspeed + " PCS"));
                                    }
                                    catch (Exception e1)
                                    {
                                        Log.Error("readplc error" + e1.ToString());
                                        // lbAdd("readplc wrong3", "inf", e1.ToString());
                                    }
                                    #endregion
                                    */
                                }
                            }
                        }
                        catch (Exception e1)
                        {
                            lbAdd("readplc wrong", "inf", e1.ToString());
                        }
                        //Thread.Sleep(20);
                        //Thread.Sleep(100);
                    }
                    else
                    {
                        app._reader.WaitOne();
                    }
                }
            }
        }

        List<string> GetDefectNameListForThisStop(string produceNo, int stopVal)
        {
            // 由 GitHub Copilot 產生
            // 修正: 使用快取避免每次都查詢資料庫,減少 "database is locked" 錯誤
            
            // 先檢查快取
            if (app.defectNamesPerStop.TryGetValue(stopVal, out var cachedNames))
            {
                return cachedNames;
            }

            // 快取不存在時才查詢資料庫(通常只在初始化或換料號時發生)
            List<string> defectNames = new List<string>();

            using (var db = new MydbDB())
            {
                var defectInfos = db.DefectChecks
                    .Where(dc => dc.Type == produceNo && dc.Stop == stopVal && dc.Yn == 1)  // 只選 YN == 1 的記錄
                    .OrderBy(dc => dc.Name)
                    .Select(dc => dc.Name)
                    .ToList();

                defectNames.AddRange(defectInfos);
            }

            // 存入快取
            app.defectNamesPerStop[stopVal] = defectNames;

            return defectNames;
        }

        public static (CircleSegment[] outerCircles, CircleSegment[] innerCircles) DetectCircles(Mat inputImage, int stop)
        {
            // 從 app.param 中讀取霍夫圓參數 (如果有app.param，否則使用預設值)
            int outerMinRadius, outerMaxRadius, innerMinRadius, innerMaxRadius;
            int outerP1, outerP2, innerP1, innerP2;
            int outerMinDist, innerMinDist;

            // 如果是在Form1中且有app.param可用

            outerMinRadius = int.Parse(app.param[$"outer_minRadius_{stop}"]);
            outerMaxRadius = int.Parse(app.param[$"outer_maxRadius_{stop}"]);
            innerMinRadius = int.Parse(app.param[$"inner_minRadius_{stop}"]);
            innerMaxRadius = int.Parse(app.param[$"inner_maxRadius_{stop}"]);
            outerP1 = int.Parse(app.param[$"outer_p1_{stop}"]);
            outerP2 = int.Parse(app.param[$"outer_p2_{stop}"]);
            innerP1 = int.Parse(app.param[$"inner_p1_{stop}"]);
            innerP2 = int.Parse(app.param[$"inner_p2_{stop}"]);
            outerMinDist = int.Parse(app.param[$"outer_minDist_{stop}"]);
            innerMinDist = int.Parse(app.param[$"inner_minDist_{stop}"]);

            // 灰階轉換和高斯模糊
            Mat gray = inputImage.CvtColor(ColorConversionCodes.BGR2GRAY);
            Mat blurred = gray.GaussianBlur(new Size(5, 5), 1);

            // 霍夫圓檢測：外圈
            var outerCircles = Cv2.HoughCircles(
                blurred,
                HoughModes.Gradient,
                dp: 1,
                minDist: outerMinDist,
                param1: outerP1,
                param2: outerP2,
                minRadius: outerMinRadius,
                maxRadius: outerMaxRadius);

            // 霍夫圓檢測：內圈
            var innerCircles = Cv2.HoughCircles(
                blurred,
                HoughModes.Gradient,
                dp: 1,
                minDist: innerMinDist,
                param1: innerP1,
                param2: innerP2,
                minRadius: innerMinRadius,
                maxRadius: innerMaxRadius);

            // 釋放不再需要的資源
            gray.Dispose();
            blurred.Dispose();

            return (outerCircles, innerCircles);
        }

        private Mat DetectAndExtractROI(Mat inputImage, int stop, int count, bool chamfer = false /*, CircleSegment?[] outerCircles, CircleSegment?[] innerCircles*/)
        {
            try
            {
                Mat mask = new Mat(inputImage.Size(), MatType.CV_8UC1, Scalar.Black);
                //Mat roi_full_color = new Mat();
                Mat roi_final = new Mat();
                Mat roi_full = new Mat();
                try
                {


                    // 從資料庫讀取預設圓心和半徑
                    int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
                    int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
                    int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);
                    int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                    int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                    int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);


                    #region 不霍夫
                    /*
                    logInfo.AppendLine($"參數設置:");
                    logInfo.AppendLine($"  外圈半徑範圍: {outerMinRadius}-{outerMaxRadius}, 內圈半徑範圍: {innerMinRadius}-{innerMaxRadius}");
                    logInfo.AppendLine($"  預設外圓圓心: ({knownOuterCenterX}, {knownOuterCenterY})");
                    logInfo.AppendLine($"  預設內圓圓心: ({knownInnerCenterX}, {knownInnerCenterY})");
                    logInfo.AppendLine($"  預設半徑: 外圈={knownOuterRadius}, 內圈={knownInnerRadius}");

                    int centerTolerance = 50; // 圓心容許偏差
                    */


                    /*
                    // 篩選和驗證檢測到的圓
                    CircleSegment? bestOuterCircle = null;
                    CircleSegment? bestInnerCircle = null;
                    double minCenterDistance = double.MaxValue;

                    foreach (var outerCircle in outerCircles)
                    {
                        foreach (var innerCircle in innerCircles)
                        {
                            // 計算圓心距離
                            double centerDistance = Math.Sqrt(Math.Pow(outerCircle.Center.X - innerCircle.Center.X, 2) + Math.Pow(outerCircle.Center.Y - innerCircle.Center.Y, 2));

                            // 檢查圓心距離和半徑比
                            if (centerDistance <= centerTolerance && centerDistance < minCenterDistance)
                            {
                                double radiusRatio = outerCircle.Radius / innerCircle.Radius;
                                //if (radiusRatio > 1.5 && radiusRatio < 3.0) // 假設合理的半徑比範圍
                                {
                                    bestOuterCircle = outerCircle;
                                    bestInnerCircle = innerCircle;
                                    minCenterDistance = centerDistance;
                                }
                            }
                        }
                    }

                    if (bestOuterCircle == null || bestInnerCircle == null)
                    {
                        throw new Exception("找不到有效的內外圓！");
                    }

                    // 直接寫入圓心半徑

                    // 檢測到的圓心和半徑
                    Point detectedOuterCenter = new Point((int)bestOuterCircle.Value.Center.X, (int)bestOuterCircle.Value.Center.Y);
                    int detectedOuterRadius = (int)bestOuterCircle.Value.Radius;
                    Point detectedInnerCenter = new Point((int)bestInnerCircle.Value.Center.X, (int)bestInnerCircle.Value.Center.Y);
                    int detectedInnerRadius = (int)bestInnerCircle.Value.Radius;

                    detectedInnerCenter.X = int.Parse(app.param[]);
                    detectedInnerCenter.Y = 943;
                    detectedInnerRadius = 545;

                    if (stop == 1)
                    {
                        detectedInnerCenter.X = 1157;
                        detectedInnerCenter.Y = 943;
                        detectedInnerRadius = 545;
                    }
                    else if (stop == 3)
                    {

                    }
                    */
                    /*
                    // 記錄檢測到的圓形信息
                    logInfo.AppendLine("檢測到的圓形信息:");
                    logInfo.AppendLine($"  外圈圓心: ({detectedOuterCenter.X}, {detectedOuterCenter.Y}), 半徑: {detectedOuterRadius}");
                    logInfo.AppendLine($"  內圈圓心: ({detectedInnerCenter.X}, {detectedInnerCenter.Y}), 半徑: {detectedInnerRadius}");

                    // 計算與預設值的偏差
                    int outerCenterXDiff = Math.Abs(detectedOuterCenter.X - knownOuterCenterX);
                    int outerCenterYDiff = Math.Abs(detectedOuterCenter.Y - knownOuterCenterY);
                    int innerCenterXDiff = Math.Abs(detectedInnerCenter.X - knownInnerCenterX);
                    int innerCenterYDiff = Math.Abs(detectedInnerCenter.Y - knownInnerCenterY);
                    int outerRadiusDiff = Math.Abs(detectedOuterRadius - knownOuterRadius);
                    int innerRadiusDiff = Math.Abs(detectedInnerRadius - knownInnerRadius);

                    logInfo.AppendLine("與預設值的偏差:");
                    logInfo.AppendLine($"  外圈圓心偏差: X={outerCenterXDiff}, Y={outerCenterYDiff}, 半徑偏差: {outerRadiusDiff}");
                    logInfo.AppendLine($"  內圈圓心偏差: X={innerCenterXDiff}, Y={innerCenterYDiff}, 半徑偏差: {innerRadiusDiff}");

                    // 圓心間距離
                    double centerDistance0 = Math.Sqrt(Math.Pow(detectedOuterCenter.X - detectedInnerCenter.X, 2) + Math.Pow(detectedOuterCenter.Y - detectedInnerCenter.Y, 2));
                    logInfo.AppendLine($"  內外圓圓心距離: {centerDistance0:F2} 像素");

                    // 預設外圓圓心 (用於整體圖像平移)
                    Point knownOuterCenter = new Point(knownOuterCenterX, knownOuterCenterY);

                    // 預設內圓圓心 (用於特殊內環檢測模式)
                    Point knownInnerCenter = new Point(knownInnerCenterX, knownInnerCenterY);

                    // 計算平移向量 (從檢測圓心到預設圓心)
                    int shiftX = knownInnerCenterX - detectedInnerCenter.X; //如果檢測不準確 平移向量就不對 造成圖片歪斜
                    int shiftY = knownInnerCenterY - detectedInnerCenter.Y; //其實檢測圓心/半徑 可以預設定值
                    logInfo.AppendLine($"應用平移向量: ({shiftX}, {shiftY})"); //只是一為訓練一為檢測

                    // 建立平移矩陣
                    Mat translationMatrix = new Mat(2, 3, MatType.CV_64FC1);
                    double[] translationData = new double[] { 1, 0, shiftX, 0, 1, shiftY };
                    Marshal.Copy(translationData, 0, translationMatrix.Data, translationData.Length);

                    // 使用平移矩陣進行整個圖像平移
                    Mat shiftedImage = new Mat();
                    Cv2.WarpAffine(inputImage, shiftedImage, translationMatrix, inputImage.Size());

                    // 建立遮罩
                    Mat mask = new Mat(shiftedImage.Size(), MatType.CV_8UC1, Scalar.Black);

                    // 取實際檢測半徑和預設半徑中較小者
                    //int finalOuterRadius = Math.Min(detectedOuterRadius, knownOuterRadius);
                    //int finalInnerRadius = Math.Max(detectedInnerRadius, knownInnerRadius);
                    int finalOuterRadius = detectedOuterRadius;
                    int finalInnerRadius = detectedInnerRadius;
                    logInfo.AppendLine($"最終使用的半徑: 外圈={finalOuterRadius}, 內圈={finalInnerRadius}");
                    */
                    #endregion

                    //直接用座標去被 如果動到光纖、延遲時間、轉盤轉速 會直接跑掉
                    // 如果跑掉 但是ROI範圍尚可 可以用平移補救 但要記錄訓練時的座標
                    // 預設外圓圓心 
                    Point knownOuterCenter = new Point(knownOuterCenterX, knownOuterCenterY);

                    // 預設內圓圓心 
                    Point knownInnerCenter = new Point(knownInnerCenterX, knownInnerCenterY);

                    // 建立遮罩
                    //Mat mask = new Mat(inputImage.Size(), MatType.CV_8UC1, Scalar.Black);

                    if (stop == 1 || stop == 2)
                    {
                        int chamferCenterX = int.Parse(app.param[$"known_chamfer_center_x_{stop}"]);
                        int chamferCenterY = int.Parse(app.param[$"known_chamfer_center_y_{stop}"]);
                        int chamferRadius = int.Parse(app.param[$"known_chamfer_radius_{stop}"]);
                        Point chamferCenter = new Point(chamferCenterX, chamferCenterY);

                        if (chamfer == false) // 主要檢測 
                        {
                            Cv2.Circle(mask, knownOuterCenter, knownOuterRadius, Scalar.White, -1);
                            Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.White, -1);
                        }
                        else //撿倒角 若未來倒角模型與主要模型合併 留這裡 上面註解掉
                        {
                            Cv2.Circle(mask, chamferCenter, chamferRadius, Scalar.White, -1);
                            Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.White, -1);
                        }
                    }
                    else if (stop == 3 || stop == 4)
                    {
                        if (chamfer == false) // 主要檢測 
                        {
                            Cv2.Circle(mask, knownOuterCenter, knownOuterRadius, Scalar.White, -1);
                            Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.Black, -1);
                        }
                        else //撿倒角 若未來倒角模型與主要模型合併 留這裡 上面註解掉 目前34站不檢倒角
                        {
                            //Cv2.Circle(mask, chamferCenter, chamferRadius, Scalar.White, -1);
                            Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.Black, -1);
                        }
                    }
                    #region 舊模式
                    /*
                    // 判斷是否為特殊內環檢測模式 //這邊會改掉
                    bool isInnerHoughMode = false;
                    if (app.param.ContainsKey($"innerHough_{stop}"))
                    {
                        isInnerHoughMode = (int.Parse(app.param[$"innerHough_{stop}"]) == 1);
                    }

                    int y = knownInnerCenter.Y;
                    if (isInnerHoughMode)
                    {
                        // 特殊內環檢測模式 - 內圓變成白色遮罩，且外圓用內圓的延伸
                        int offsetY = app.param.ContainsKey($"innerHoughOffsetY_{stop}") ? //無用
                                      int.Parse(app.param[$"innerHoughOffsetY_{stop}"]) : 0;
                        int roiRadius = app.param.ContainsKey($"innerHoughRoiRadius_{stop}") ?
                                       int.Parse(app.param[$"innerHoughRoiRadius_{stop}"]) : 0;

                        int calculatedOuterRadius = knownInnerRadius + roiRadius;
                        knownInnerCenter.Y = y + offsetY;

                        // 繪製白色圓形遮罩
                        Cv2.Circle(mask, knownInnerCenter, calculatedOuterRadius, Scalar.White, -1);
                        //Cv2.Circle(mask, knownCenter, finalInnerRadius, Scalar.White, -1);
                    }
                    else
                    {
                        // 標準模式 - 外圓白色，內圓黑色
                        Cv2.Circle(mask, knownOuterCenter, knownOuterRadius, Scalar.White, -1);
                        if (stop == 3 || stop == 4)
                        {
                            Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.Black, -1);
                        }
                        else
                        {
                            Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.White, -1);
                        }
                    }
                    */
                    #endregion

                    // 應用遮罩
                    //Mat roi_full = new Mat();
                    Cv2.BitwiseAnd(inputImage, inputImage, roi_full, mask);
                    // 如果是第1、2站，將內環區域填充為白色
                    if (stop == 1 || stop == 2)
                    {
                        Cv2.Circle(roi_full, knownInnerCenter, knownInnerRadius, Scalar.White, -1);
                    }
                    //Mat roi_final = new Mat();
                    // 轉灰階
                    if (!(app.param.ContainsKey($"color_{stop}") && app.param[$"color_{stop}"] == "1")) //not 有且是1
                    {
                        Cv2.CvtColor(roi_full, roi_final, ColorConversionCodes.BGR2GRAY);
                    }
                    else
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: 必須克隆 roi_full,否則在 finally 中釋放 roi_full 會導致 roi_final 也被釋放
                        roi_final = roi_full.Clone();
                    }

                    // 根據站點進行不同的後處理
                    if (stop == 1 || stop == 2)
                    {
                        int contrastOffset = app.param.ContainsKey($"deepenContrast_{stop}") ?
                                            int.Parse(app.param[$"deepenContrast_{stop}"]) : 30;
                        int brightnessOffset = app.param.ContainsKey($"deepenBrightness_{stop}") ?
                                              int.Parse(app.param[$"deepenBrightness_{stop}"]) : 5;
                        roi_final = ContrastAndClose(roi_final, contrastOffset, brightnessOffset, stop);
                    }
                    
                    // 轉回 BGR 以便顯示和保存
                    if (!(app.param.ContainsKey($"color_{stop}") && app.param[$"color_{stop}"] == "1"))
                    {
                        Cv2.CvtColor(roi_final, roi_final, ColorConversionCodes.GRAY2BGR);
                    }

                    // 由 GitHub Copilot 產生
                    // 修正: 使用快取的參數避免資料庫鎖定,不需每次都查詢資料庫
                    bool shouldSaveROI = app.param.ContainsKey("saveROI") && app.param["saveROI"] == "true";

                    if (shouldSaveROI)
                    {
                        if (chamfer == false)
                        {
                            string visImgPath = $@".\image\{st.ToString("yyyy-MM")}\{st.ToString("MMdd")}\{app.foldername}\ROI_{stop}\{count}_{stop}_roi.png";
                            app.Queue_Save.Enqueue(new ImageSave(roi_final, visImgPath));
                            app._sv.Set();
                        }
                        else
                        {
                            string visImgPath = $@".\image\{st.ToString("yyyy-MM")}\{st.ToString("MMdd")}\{app.foldername}\chamferROI_{stop}\{count}_{stop}_chamferRoi.png";
                            app.Queue_Save.Enqueue(new ImageSave(roi_final, visImgPath));
                            app._sv.Set();
                        }
                    }
                    return roi_final;
                }
                finally
                {
                    // 由 GitHub Copilot 產生
                    // 修正: 釋放所有臨時 Mat 物件，防止記憶體洩漏
                    mask?.Dispose();
                    roi_full?.Dispose();  // ✅ 修正: 必須釋放 roi_full（15 MB）
                    // roi_final 不能釋放，因為要返回
                }
            }
            catch (Exception ex)
            {
                Log.Error($"DetectAndExtractROI 錯誤: {ex.Message}, {stop}");
                return inputImage; // 錯誤時返回原始圖像
            }
        }

        public List<Rect> MergeRectangles(List<Rect> rects, double iouThreshold)
        {
            if (rects == null || rects.Count == 0)
                return new List<Rect>();

            // 先依面積排序(大到小)，可依需求改成小到大
            var sorted = rects.OrderByDescending(r => r.Width * r.Height).ToList();
            List<Rect> merged = new List<Rect>();

            while (sorted.Count > 0)
            {
                Rect current = sorted[0];
                sorted.RemoveAt(0);

                bool mergedFlag = false;
                for (int i = 0; i < merged.Count; i++)
                {
                    double iou = IOU(current, merged[i]);
                    if (iou >= iouThreshold)
                    {
                        // union
                        Rect unionRect = UnionRect(current, merged[i]);
                        merged[i] = unionRect;
                        mergedFlag = true;
                        break;
                    }
                }

                if (!mergedFlag)
                {
                    merged.Add(current);
                }
            }

            return merged;
        }
        public double IOU(Rect a, Rect b)
        {
            int interX1 = Math.Max(a.Left, b.Left);
            int interY1 = Math.Max(a.Top, b.Top);
            int interX2 = Math.Min(a.Right, b.Right);
            int interY2 = Math.Min(a.Bottom, b.Bottom);

            int interW = Math.Max(0, interX2 - interX1);
            int interH = Math.Max(0, interY2 - interY1);
            double interArea = interW * interH;

            double areaA = a.Width * a.Height;
            double areaB = b.Width * b.Height;
            double unionArea = areaA + areaB - interArea;

            if (unionArea <= 0) return 0.0;
            return interArea / unionArea;
        }
        public Rect UnionRect(Rect a, Rect b)
        {
            int x1 = Math.Min(a.Left, b.Left);
            int y1 = Math.Min(a.Top, b.Top);
            int x2 = Math.Max(a.Right, b.Right);
            int y2 = Math.Max(a.Bottom, b.Bottom);

            return new Rect(x1, y1, x2 - x1, y2 - y1);
        }

        public (bool hasDefect, Mat resultImg) DetectScratch(Mat src, int stop)
        {
            int deepenContrast = int.Parse(app.param[$"deepenContrast_{stop}"]);
            int deepenBrightness = int.Parse(app.param[$"deepenBrightness_{stop}"]);
            int deepenThresh = 0; // int.Parse(app.param[$"deepenThresh_{stop}"]);  //已刪
            int minScratch = int.Parse(app.param[$"minScratch_{stop}"]);
            int maxScratch = int.Parse(app.param[$"maxScratch_{stop}"]);
            try
            {
                // 1) 建立 darkenImage
                using (Mat darkenImage = AdjustContrast(src, deepenContrast, deepenBrightness))
                using (Mat gray = new Mat())
                using (Mat bin = new Mat())
                using (Mat eroded = new Mat())
                using (Mat element = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3)))
                {
                    // 2) 灰階 -> Threshold
                    Cv2.CvtColor(darkenImage, gray, ColorConversionCodes.BGR2GRAY);

                    // 固定閾值
                    Cv2.Threshold(gray, bin, deepenThresh, 255, ThresholdTypes.Binary);

                    // 可做些形態學 (如 Erode, Dilate 等)
                    Cv2.Erode(bin, eroded, element);

                    // 3) findContours
                    Point[][] contours;
                    HierarchyIndex[] hierarchy;
                    Cv2.FindContours(eroded, out contours, out hierarchy,
                                     RetrievalModes.List,
                                     ContourApproximationModes.ApproxSimple);

                    // 篩選 boundingRect高度 ∈ [minDefect, maxDefect]
                    bool hasDefect = false;
                    List<Point[]> defectContours = new List<Point[]>();
                    foreach (var contour in contours)
                    {
                        Rect br = Cv2.BoundingRect(contour);
                        if (br.Height >= minScratch && br.Height <= maxScratch)
                        {
                            hasDefect = true;
                            defectContours.Add(contour);
                        }
                    }
                    string predicted = hasDefect ? "NG" : "OK";
                    
                    // 4) 畫結果圖
                    Mat resultImg = darkenImage.Clone();
                    string labelText = predicted; // "NG" / "OK"
                    Scalar labelColor = (predicted == "NG") ? Scalar.Red : Scalar.Green;
                    Cv2.PutText(resultImg, labelText, new Point(30, 60),
                                HersheyFonts.HersheySimplex, 2.0, labelColor, 3);

                    // 如果是NG → 用藍色畫輪廓
                    if (predicted == "NG")
                    {
                        Cv2.DrawContours(resultImg, defectContours, -1, Scalar.Blue, 2);
                    }

                    // 注意: resultImg 由呼叫端負責釋放
                    return (hasDefect, resultImg);
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"處理檔案 {filePath} 時發生錯誤：{ex.Message}");
                return (false, null);
            }
        }

        public static Mat AdjustContrast(Mat src, int contrast, int brightness)
        {
            if (src == null)
                throw new ArgumentNullException(nameof(src), "輸入影像不能為 null。");
            if (contrast < -100 || contrast > 100)
                throw new ArgumentOutOfRangeException(nameof(contrast), "對比度應在 -100 到 100 之間。");
            if (brightness < -100 || brightness > 100)
                throw new ArgumentOutOfRangeException(nameof(brightness), "亮度應在 -100 到 100 之間。");

            // 建立查表
            byte[] lookupTableData = new byte[256];
            double delta, a, b;
            int brightnessOffset = brightness;
            int contrastOffset = contrast;

            if (contrastOffset > 0)
            {
                // 對比度 > 0
                delta = 127 * contrastOffset / 100.0;
                a = 255.0 / (255.0 - delta * 2);
                b = a * (brightnessOffset - delta);
            }
            else
            {
                // 對比度 <= 0
                delta = -128 * contrastOffset / 100.0;
                a = (256.0 - delta * 2) / 255.0;
                b = a * brightnessOffset + delta;
            }

            for (int i = 0; i < 256; i++)
            {
                int y = (int)(a * i + b + 0.5);
                if (y < 0) y = 0;
                if (y > 255) y = 255;
                lookupTableData[i] = (byte)y;
            }

            // 建立查表矩陣
            Mat lookupTableMatrix = new Mat(256, 1, MatType.CV_8UC1);
            Marshal.Copy(lookupTableData, 0, lookupTableMatrix.Data, lookupTableData.Length);

            // 應用查表
            Mat dst = new Mat();
            Cv2.LUT(src, lookupTableMatrix, dst);

            lookupTableMatrix.Dispose();
            return dst;
        }

        // 當 isNG => 這個函式做 threshold=> findContours=> mergeRect=> crop => classification

        public static (int, float[]) InferByOnnxRuntime(Mat img, InferenceSession session, int size/*, bool centerCrop = true*/)
        {
            try
            {
                // Resize and optionally center crop image
                //var resizedImage = ResizeAndCropImage(img, size, centerCrop);

                // Normalize image
                var tensor = NormalizeImage(img, size);

                // Perform inference
                var inputs = new[] { NamedOnnxValue.CreateFromTensor(session.InputMetadata.Keys.First(), tensor) };
                using (var results = session.Run(inputs))
                {
                    var output = results.First().AsTensor<float>().ToArray();
                    var prediction = Array.IndexOf(output, output.Max());
                    return (prediction, output);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(" InferByOnnxRuntime ERR " + e);
                return (-1, null);
            }
        }
        private static DenseTensor<float> NormalizeImage(Mat img, int size)
        {
            var mean = new[] { 0.485f, 0.456f, 0.406f };
            var std = new[] { 0.229f, 0.224f, 0.225f };
            var tensor = new DenseTensor<float>(new[] { 1, 3, size, size });

            var pixelptr = img.GetUnsafeGenericIndexer<Vec3b>();
            for (int c = 0; c < 3; c++)
            {
                for (int y = 0; y < size; y++)
                {
                    for (int x = 0; x < size; x++)
                    {
                        var pixel = pixelptr[y, x];
                        tensor[0, c, y, x] = (pixel[c] / 255.0f - mean[c]) / std[c];
                    }
                }
            }

            return tensor;
        }
        public string InferenceGetLabel(int clsIdx, int stop)
        {
            //Console.WriteLine(app.defectLabels.Count);
            if (stop == 1 || stop == 2)
            {
                if (clsIdx < 0 || clsIdx >= app.defectLabels_in.Count)
                {
                    return "Unknown";
                }
                return app.defectLabels_in[clsIdx];
            }
            else if (stop == 3 || stop == 4)
            {
                if (clsIdx < 0 || clsIdx >= app.defectLabels_out.Count)
                {
                    return "Unknown";
                }
                return app.defectLabels_out[clsIdx];
            }
            else
            {
                return "Unknown";
            }
        }
        public (bool isNG, Mat img, List<Point> gapPositions) findGapWidth(Mat img, int stop/*, CircleSegment[] outerCircles, CircleSegment[] innerCircles*/)
        {
            // 無效站別的thresh會預設為0

            // 新增：用來存儲開口位置的列表
            List<Point> gapPositions = new List<Point>();

            // 由 GitHub Copilot 產生
            // 修正: 將 gray 和 ringThresh 納入 using 管理，確保記憶體釋放
            using (Mat ori = img.Clone())
            using (Mat visualImg = img.Clone())
            using (Mat gray = new Mat())
            using (Mat ringThresh = new Mat())
            {
                int minthresh = int.Parse(app.param[$"gapThresh_{stop}"]);
                if (minthresh == 0)
                {
                    return (false, visualImg.Clone(), gapPositions); // 0表示不做阈值，需要 Clone 因為會被 Dispose
                }

                Cv2.CvtColor(visualImg, gray, ColorConversionCodes.BGR2GRAY);
                Cv2.Threshold(gray, ringThresh, minthresh, 255, ThresholdTypes.Binary);

            #region 不霍夫
            /*
            int centerTolerance = 80;
            bool matched = false;
            Mat roi_blurred = null;
            CircleSegment? bestInner = null;

            // 1) 用霍夫圓找外圈+內圈 看是否匹配 匹配意義在於確認樣本沒有非常嚴重的變形
            // 2) 沒有非常嚴重變形再去計算內彎和外擴
            if (outerCircles.Length == 0) Console.WriteLine("外圈沒圓");
            if (innerCircles.Length == 0) Console.WriteLine("內圈沒圓");

            if (outerCircles.Length > 0 && innerCircles.Length > 0) //皆存在
            {
                foreach (var outer in outerCircles)
                {
                    foreach (var inner in innerCircles)
                    {
                        double radiusDistance = outer.Radius - inner.Radius;
                        double centerDistance = Math.Sqrt(
                            Math.Pow(outer.Center.X - inner.Center.X, 2) +
                            Math.Pow(outer.Center.Y - inner.Center.Y, 2));
                        if (centerDistance <= centerTolerance)
                        {
                            // 视为匹配
                            bestInner = inner;
                            // 建 mask
                            Mat mask = new Mat(visualImg.Size(), MatType.CV_8UC1, Scalar.Black);
                            // 外圈=white
                            Cv2.Circle(mask, (Point)outer.Center, (int)outer.Radius, Scalar.White, -1);
                            // 内圈=black
                            Cv2.Circle(mask, (Point)inner.Center, (int)inner.Radius, Scalar.Black, -1);

                            Mat roi_full = new Mat();
                            Cv2.BitwiseAnd(visualImg, visualImg, roi_full, mask);

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

            if (!matched || bestInner == null)
            {
                Console.WriteLine("內外圈未成功匹配, 判斷變形嚴重.");
                return (true, visualImg);
            }
            */
            #endregion

            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
            int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);
            int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
            int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
            int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);

            Point2f innerCenter = new Point2f(knownInnerCenterX, knownInnerCenterY);
            double innerRadius = knownInnerRadius;
            Point2f outerCenter = new Point2f(knownOuterCenterX, knownOuterCenterY);
            double outerRadius = knownOuterRadius;

            // 4) 對 ringThresh 執行極坐標掃描
            double angleStep = 0.5;
            int nSteps = (int)(360.0 / angleStep);

            bool[] inner_isHole_outward = new bool[nSteps]; // 內圓向外擴掃描的結果
            bool[] inner_isHole_inward = new bool[nSteps];  // 內圓向內彎掃描的結果

            double pixeltomm = double.Parse(app.param[$"PixelToMM_{stop}"]);
            double roiTolerance = (1 / pixeltomm) / (1 / 0.6) - 5; // 0.5mm, 基本落在15~23px
            //double roiTolerance = 10;
            // 向外擴掃描的半徑 (原來邏輯)
            double outerScanRadius = innerRadius + roiTolerance;

            // 向內彎掃描的半徑 (新增邏輯：內圓半徑向內縮小)
            double inwardScanRadius = innerRadius - roiTolerance;

            // =========== 向外擴掃描 (內圓向外) ===========
            // 可视化: 在 src 上画点(绿=環, 红=hole)
            for (int i = 0; i < nSteps; i++)
            {
                double angleDeg = i * angleStep;
                double rad = angleDeg * Math.PI / 180.0;

                double rx = innerCenter.X + outerScanRadius * Math.Cos(rad);
                double ry = innerCenter.Y + outerScanRadius * Math.Sin(rad);

                int px = (int)Math.Round(rx);
                int py = (int)Math.Round(ry);

                if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                {
                    // hole
                    inner_isHole_outward[i] = true;
                    Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Red, -1);
                }
                else
                {
                    byte val = ringThresh.Get<byte>(py, px);
                    if (val < 127)
                    {
                        // ring
                        inner_isHole_outward[i] = false;
                        Cv2.Circle(visualImg, new Point(px, py), 1, Scalar.Green, -1);
                    }
                    else
                    {
                        // hole
                        inner_isHole_outward[i] = true;
                        Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Red, -1);

                        // 新增：記錄開口位置
                        gapPositions.Add(new Point(px, py));
                    }
                }
            }

            // =========== 向內彎掃描 (內圓向內) ===========
            for (int i = 0; i < nSteps; i++)
            {
                double angleDeg = i * angleStep;
                double rad = angleDeg * Math.PI / 180.0;

                double rx = innerCenter.X + inwardScanRadius * Math.Cos(rad);
                double ry = innerCenter.Y + inwardScanRadius * Math.Sin(rad);

                int px = (int)Math.Round(rx);
                int py = (int)Math.Round(ry);

                if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                {
                    // 超出圖像邊界，標記為hole
                    inner_isHole_inward[i] = true;
                    Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Blue, -1); // 藍色標記
                }
                else
                {
                    byte val = ringThresh.Get<byte>(py, px);
                    if (val < 127)
                    {
                        // 檢測到環狀物體，說明內彎！這是異常情況
                        inner_isHole_inward[i] = false; // 非hole代表檢測到內彎
                        Cv2.Circle(visualImg, new Point(px, py), 1, Scalar.Yellow, -1); // 黃色標記
                    }
                    else
                    {
                        // 正常情況，內側應該是空白的
                        inner_isHole_inward[i] = true;
                        Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Blue, -1); // 藍色標記
                    }
                }
            }

            // 5) 計算外擴尺寸 - 找最大連續缺口
            double outerMaxGapAngleDeg = 0;
            int outerStartIdx = -1;
            for (int idx = 0; idx < nSteps * 2; idx++)
            {
                int realIdx = idx % nSteps;
                if (inner_isHole_outward[realIdx])
                {
                    if (outerStartIdx < 0) outerStartIdx = idx;
                }
                else
                {
                    if (outerStartIdx >= 0)
                    {
                        int length = idx - outerStartIdx;
                        double gapDeg = length * angleStep;
                        if (gapDeg > outerMaxGapAngleDeg) outerMaxGapAngleDeg = gapDeg;
                        outerStartIdx = -1;
                    }
                }
            }
            if (outerStartIdx >= 0)
            {
                int length = (nSteps * 2) - outerStartIdx;
                double gapDeg = length * angleStep;
                if (gapDeg > outerMaxGapAngleDeg) outerMaxGapAngleDeg = gapDeg;
            }

            // 6) 計算向內彎的尺寸 - 找連續檢測到"非空白"區域的長度
            double inwardBendAngleDeg = 0;
            int inwardStartIdx = -1;
            for (int idx = 0; idx < nSteps * 2; idx++)
            {
                int realIdx = idx % nSteps;
                // 如果檢測到非空白，即內彎
                if (!inner_isHole_inward[realIdx])
                {
                    if (inwardStartIdx < 0) inwardStartIdx = idx;
                }
                else
                {
                    if (inwardStartIdx >= 0)
                    {
                        int length = idx - inwardStartIdx;
                        double bendDeg = length * angleStep;
                        if (bendDeg > inwardBendAngleDeg) inwardBendAngleDeg = bendDeg;
                        inwardStartIdx = -1;
                    }
                }
            }
            if (inwardStartIdx >= 0)
            {
                int length = (nSteps * 2) - inwardStartIdx;
                double bendDeg = length * angleStep;
                if (bendDeg > inwardBendAngleDeg) inwardBendAngleDeg = bendDeg;
            }

            // 7) 將角度轉換為物理尺寸 (mm)
            double outerGapArcPx = outerScanRadius * (outerMaxGapAngleDeg * Math.PI / 180.0);
            double outerGapArcMm = outerGapArcPx * pixeltomm;

            double inwardBendArcPx = inwardScanRadius * (inwardBendAngleDeg * Math.PI / 180.0);
            double inwardBendArcMm = inwardBendArcPx * pixeltomm;

            //Console.WriteLine($"外擴情況: 最大開口角度={outerMaxGapAngleDeg:F2} deg => 弧長={outerGapArcMm:F2} mm");
            //Console.WriteLine($"內彎情況: 最大內彎角度={inwardBendAngleDeg:F2} deg => 弧長={inwardBendArcMm:F2} mm");

                // 由 GitHub Copilot 產生
                // 修正: 移除手動 Dispose（using 會自動處理）
                // gray.Dispose();
                // ringThresh.Dispose();

                // 8) 判斷是否NG (內彎或外擴超出閾值)
                bool isOutwardNG = outerGapArcMm >= 1.5;
                bool isInwardNG = inwardBendAngleDeg > 0; // 只要有內彎就算NG

                if (isOutwardNG || isInwardNG)
                {
                    // ori 會被 using 自動釋放
                    return (true, visualImg.Clone(), gapPositions);
                }
                else
                {
                    // ori, visualImg 都會被 using 自動釋放
                    return (false, visualImg.Clone(), gapPositions);
                }
            } // using 結束，ori, visualImg, gray, ringThresh 自動 Dispose
        }

        public (bool isNG, Mat img, List<Point> gapPositions) findGap(Mat img, int stop/*, CircleSegment[] outerCircles, CircleSegment[] innerCircles*/)
        {
            // 四站皆做 有一站超過閥值就算變形 
            // 無效站別的thresh會預設為0

            // 新增：用來存儲開口位置的列表
            List<Point> gapPositions = new List<Point>();

            using (Mat ori = img.Clone())
            using (Mat visualImg = img.Clone())
            using (Mat ringThresh = img.Clone())
            {
            #region 不霍夫
            /*
            int centerTolerance = 80;
            bool matched = false;
            Mat roi_blurred = null;
            CircleSegment? bestInner = null;

            // 1) 用霍夫圓找外圈+內圈 看是否匹配 匹配意義在於確認樣本沒有非常嚴重的變形
            // 2) 沒有非常嚴重變形再去計算內彎和外擴
            if (outerCircles.Length == 0) Console.WriteLine("外圈沒圓");
            if (innerCircles.Length == 0) Console.WriteLine("內圈沒圓");

            if (outerCircles.Length > 0 && innerCircles.Length > 0) //皆存在
            {
                foreach (var outer in outerCircles)
                {
                    foreach (var inner in innerCircles)
                    {
                        double radiusDistance = outer.Radius - inner.Radius;
                        double centerDistance = Math.Sqrt(
                            Math.Pow(outer.Center.X - inner.Center.X, 2) +
                            Math.Pow(outer.Center.Y - inner.Center.Y, 2));
                        if (centerDistance <= centerTolerance)
                        {
                            // 视为匹配
                            bestInner = inner;
                            // 建 mask
                            Mat mask = new Mat(visualImg.Size(), MatType.CV_8UC1, Scalar.Black);
                            // 外圈=white
                            Cv2.Circle(mask, (Point)outer.Center, (int)outer.Radius, Scalar.White, -1);
                            // 内圈=black
                            Cv2.Circle(mask, (Point)inner.Center, (int)inner.Radius, Scalar.Black, -1);

                            Mat roi_full = new Mat();
                            Cv2.BitwiseAnd(visualImg, visualImg, roi_full, mask);

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

            if (!matched || bestInner == null)
            {
                Console.WriteLine("內外圈未成功匹配, 判斷變形嚴重.");
                return (true, visualImg);
            }
            */
            #endregion

            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
            int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);
            int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
            int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
            int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);

            Point2f innerCenter = new Point2f(knownInnerCenterX, knownInnerCenterY);
            double innerRadius = knownInnerRadius;
            Point2f outerCenter = new Point2f(knownOuterCenterX, knownOuterCenterY);
            double outerRadius = knownOuterRadius;

            // 4) 對 ringThresh 執行極坐標掃描
            double angleStep = 0.5;
            int nSteps = (int)(360.0 / angleStep);

            bool[] inner_isHole_outward = new bool[nSteps]; // 內圓向外擴掃描的結果
            bool[] inner_isHole_inward = new bool[nSteps];  // 內圓向內彎掃描的結果

            double pixeltomm = double.Parse(app.param[$"PixelToMM_{stop}"]);
            double roiTolerance = 8; // 0.5mm, 基本落在15~23px

            // 向外擴掃描的半徑 (原來邏輯)
            double outerScanRadius = innerRadius + roiTolerance;

            // 向內彎掃描的半徑 (新增邏輯：內圓半徑向內縮小)
            double inwardScanRadius = innerRadius - roiTolerance;

            // =========== 向外擴掃描 (內圓向外) ===========
            // 可视化: 在 src 上画点(绿=環, 红=hole)
            for (int i = 0; i < nSteps; i++)
            {
                double angleDeg = i * angleStep;
                double rad = angleDeg * Math.PI / 180.0;

                double rx = innerCenter.X + outerScanRadius * Math.Cos(rad);
                double ry = innerCenter.Y + outerScanRadius * Math.Sin(rad);

                int px = (int)Math.Round(rx);
                int py = (int)Math.Round(ry);

                if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                {
                    // hole
                    inner_isHole_outward[i] = true;
                    Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Red, -1);
                }
                else
                {
                    byte val = ringThresh.Get<byte>(py, px);
                    if (val < 127)
                    {
                        // ring
                        inner_isHole_outward[i] = false;
                        Cv2.Circle(visualImg, new Point(px, py), 1, Scalar.Green, -1);
                    }
                    else
                    {
                        // hole
                        inner_isHole_outward[i] = true;
                        Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Red, -1);

                        // 新增：記錄開口位置
                        gapPositions.Add(new Point(px, py));
                    }
                }
            }

            // =========== 向內彎掃描 (內圓向內) ===========
            for (int i = 0; i < nSteps; i++)
            {
                double angleDeg = i * angleStep;
                double rad = angleDeg * Math.PI / 180.0;

                double rx = innerCenter.X + inwardScanRadius * Math.Cos(rad);
                double ry = innerCenter.Y + inwardScanRadius * Math.Sin(rad);

                int px = (int)Math.Round(rx);
                int py = (int)Math.Round(ry);

                if (px < 0 || px >= ringThresh.Cols || py < 0 || py >= ringThresh.Rows)
                {
                    // 超出圖像邊界，標記為hole
                    inner_isHole_inward[i] = true;
                    Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Blue, -1); // 藍色標記
                }
                else
                {
                    byte val = ringThresh.Get<byte>(py, px);
                    if (val < 127)
                    {
                        // 檢測到環狀物體，說明內彎！這是異常情況
                        inner_isHole_inward[i] = false; // 非hole代表檢測到內彎
                        Cv2.Circle(visualImg, new Point(px, py), 1, Scalar.Yellow, -1); // 黃色標記
                    }
                    else
                    {
                        // 正常情況，內側應該是空白的
                        inner_isHole_inward[i] = true;
                        Cv2.Circle(visualImg, new Point(px, py), 2, Scalar.Blue, -1); // 藍色標記
                    }
                }
            }

            // 5) 計算外擴尺寸 - 找最大連續缺口
            double outerMaxGapAngleDeg = 0;
            int outerStartIdx = -1;
            for (int idx = 0; idx < nSteps * 2; idx++)
            {
                int realIdx = idx % nSteps;
                if (inner_isHole_outward[realIdx])
                {
                    if (outerStartIdx < 0) outerStartIdx = idx;
                }
                else
                {
                    if (outerStartIdx >= 0)
                    {
                        int length = idx - outerStartIdx;
                        double gapDeg = length * angleStep;
                        if (gapDeg > outerMaxGapAngleDeg) outerMaxGapAngleDeg = gapDeg;
                        outerStartIdx = -1;
                    }
                }
            }
            if (outerStartIdx >= 0)
            {
                int length = (nSteps * 2) - outerStartIdx;
                double gapDeg = length * angleStep;
                if (gapDeg > outerMaxGapAngleDeg) outerMaxGapAngleDeg = gapDeg;
            }

            // 6) 計算向內彎的尺寸 - 找連續檢測到"非空白"區域的長度
            double inwardBendAngleDeg = 0;
            int inwardStartIdx = -1;
            for (int idx = 0; idx < nSteps * 2; idx++)
            {
                int realIdx = idx % nSteps;
                // 如果檢測到非空白，即內彎
                if (!inner_isHole_inward[realIdx])
                {
                    if (inwardStartIdx < 0) inwardStartIdx = idx;
                }
                else
                {
                    if (inwardStartIdx >= 0)
                    {
                        int length = idx - inwardStartIdx;
                        double bendDeg = length * angleStep;
                        if (bendDeg > inwardBendAngleDeg) inwardBendAngleDeg = bendDeg;
                        inwardStartIdx = -1;
                    }
                }
            }
            if (inwardStartIdx >= 0)
            {
                int length = (nSteps * 2) - inwardStartIdx;
                double bendDeg = length * angleStep;
                if (bendDeg > inwardBendAngleDeg) inwardBendAngleDeg = bendDeg;
            }

            // 7) 將角度轉換為物理尺寸 (mm)
            double outerGapArcPx = outerScanRadius * (outerMaxGapAngleDeg * Math.PI / 180.0);
            double outerGapArcMm = outerGapArcPx * pixeltomm;

            double inwardBendArcPx = inwardScanRadius * (inwardBendAngleDeg * Math.PI / 180.0);
            double inwardBendArcMm = inwardBendArcPx * pixeltomm;

            //Console.WriteLine($"外擴情況: 最大開口角度={outerMaxGapAngleDeg:F2} deg => 弧長={outerGapArcMm:F2} mm");
            //Console.WriteLine($"內彎情況: 最大內彎角度={inwardBendAngleDeg:F2} deg => 弧長={inwardBendArcMm:F2} mm");

            // 8) 判斷是否NG (內彎或外擴超出閾值)
            bool isOutwardNG = outerGapArcMm >= 1.5;
            bool isInwardNG = inwardBendAngleDeg > 0; // 只要有內彎就算NG

            if (isOutwardNG || isInwardNG)
            {
                // ori 會被 using 自動釋放
                return (true, visualImg.Clone(), gapPositions);
            }
            else
            {
                // ori, visualImg, ringThresh 都會被 using 自動釋放
                return (false, visualImg.Clone(), gapPositions);
            }
            } // using 結束，所有 Mat 自動 Dispose
        }
        public Mat ContrastAndClose(Mat src, int contrast, int brightness, int stop)
        {
            if (src == null)
                throw new ArgumentNullException(nameof(src), "輸入圖像不能為空。");
            if (contrast < -100 || contrast > 100)
                throw new ArgumentOutOfRangeException(nameof(contrast), "對比度應在 -100 到 100 之間。");
            if (brightness < -100 || brightness > 100)
                throw new ArgumentOutOfRangeException(nameof(brightness), "亮度應在 -100 到 100 之間。");

            // 创建查找表
            byte[] lookupTableData = new byte[256];
            double delta, a, b;
            int brightnessOffset = brightness;
            int contrastOffset = contrast;

            if (contrastOffset > 0)
            {
                delta = 127 * contrastOffset / 100.0;
                a = 255.0 / (255.0 - delta * 2);
                b = a * (brightnessOffset - delta);
            }
            else
            {
                delta = -128 * contrastOffset / 100.0;
                a = (256.0 - delta * 2) / 255.0;
                b = a * brightnessOffset + delta;
            }

            for (int i = 0; i < 256; i++)
            {
                int y = (int)(a * i + b);
                if (y < 0)
                    y = 0;
                if (y > 255)
                    y = 255;
                lookupTableData[i] = (byte)y;
            }

            // 创建查找表矩阵
            Mat lookupTableMatrix = new Mat(256, 1, MatType.CV_8UC1);

            // 将数据复制到矩阵中
            Marshal.Copy(lookupTableData, 0, lookupTableMatrix.Data, lookupTableData.Length);

            // 应用查找表
            Mat dst = new Mat();
            Cv2.LUT(src, lookupTableMatrix, dst);

            lookupTableMatrix.Dispose();

            //if (stop == 1 || stop == 2)
            {
                Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
                Cv2.MorphologyEx(dst, dst, MorphTypes.Close, kernel);
                kernel.Dispose();
            }

            return dst;
        }

        public (bool hasBlackDots, Mat resultImage, List<Point[]> detectedContours) DetectBlackDots(Mat roiImage, int stop, int count)
        {
            // 由 GitHub Copilot 產生
            // 緊急修正: 先檢查 IsDisposed，避免 ObjectDisposedException
            if (roiImage == null || roiImage.IsDisposed)
            {
                Log.Warning($"DetectBlackDots: roiImage 為 null 或已被釋放 (stop={stop}, count={count})");
                return (false, null, new List<Point[]>());
            }
            
            if (roiImage.Empty())
            {
                return (false, null, new List<Point[]>());
            }

            try
            {
                // 1. 轉灰階
                Mat grayImage = new Mat();
                if (roiImage.Channels() == 3)
                {
                    Cv2.CvtColor(roiImage, grayImage, ColorConversionCodes.BGR2GRAY);
                }
                else
                {
                    grayImage = roiImage.Clone();
                }
                /*
                // 2. 高斯模糊 kernel = (5,5)
                Mat blurredImage = new Mat();
                Cv2.GaussianBlur(grayImage, blurredImage, new Size(5, 5), 0);

                // 3. 做morphology中的open，形狀為cross，大小(3,3)，做三次迴圈
                Mat morphImage = blurredImage.Clone();
                Mat crossKernel = Cv2.GetStructuringElement(MorphShapes.Cross, new Size(3, 3));

                for (int i = 0; i < 3; i++)
                {
                    Cv2.MorphologyEx(morphImage, morphImage, MorphTypes.Open, crossKernel);
                }
                */

                // 4. 做二值化 (65, 255)
                Mat binaryImage = new Mat();
                Cv2.Threshold(grayImage, binaryImage, 105, 255, ThresholdTypes.Binary);

                // 5. 做findcontour，找面積為150像素以上的輪廓
                Point[][] contours;
                HierarchyIndex[] hierarchy;
                Cv2.FindContours(binaryImage, out contours, out hierarchy,
                                 RetrievalModes.List, ContourApproximationModes.ApproxSimple);

                // 篩選面積大於150像素的輪廓
                List<Point[]> validContours = new List<Point[]>();

                float minArea = 150;
                if (app.param.TryGetValue("blackDot" + stop.ToString() + "_threshold", out string thresholdStr))
                {
                    float.TryParse(thresholdStr, out minArea);
                }
                foreach (var contour in contours)
                {
                    double area = Cv2.ContourArea(contour);
                    if (area > minArea && area < 5000)
                    {
                        // 檢驗輪廓內部是否為黑色像素
                        //if (IsContourBlackInBinary(binaryImage, contour))
                        {
                            validContours.Add(contour);
                        }
                    }
                }

                // 創建結果圖像（彩色，用於視覺化）
                Mat resultImage = new Mat();
                if (roiImage.Channels() == 3)
                {
                    resultImage = roiImage.Clone();
                }
                else
                {
                    Cv2.CvtColor(roiImage, resultImage, ColorConversionCodes.GRAY2BGR);
                }

                // 在結果圖像上繪製檢測到的輪廓
                bool hasBlackDots = validContours.Count > 0;
                string resultFileName; // 新增變數來存放結果檔名

                if (hasBlackDots)
                {
                    // 用紅色繪製檢測到的輪廓
                    for (int i = 0; i < validContours.Count; i++)
                    {
                        Cv2.DrawContours(resultImage, new[] { validContours[i] }, -1, Scalar.Red, 2);

                        // 計算輪廓的邊界框並標記面積
                        Rect boundingRect = Cv2.BoundingRect(validContours[i]);
                        double area = Cv2.ContourArea(validContours[i]);

                        // 在輪廓旁邊標記面積信息
                        Cv2.PutText(resultImage, $"Area: {area:F0}",
                                   new Point(boundingRect.X, boundingRect.Y - 10),
                                   HersheyFonts.HersheyDuplex, 0.5, Scalar.Red, 1);
                    }

                    // 在圖像左上角標記檢測結果
                    Cv2.PutText(resultImage, $"Black Dots: {validContours.Count}",
                               new Point(10, 30), HersheyFonts.HersheyDuplex, 0.8, Scalar.Red, 2);

                    // 檢測到黑點的檔名
                    resultFileName = $"{count}_blackdot_NG.png";
                }
                else
                {
                    // 沒有檢測到黑點，用綠色標記
                    Cv2.PutText(resultImage, "No Black Dots",
                               new Point(10, 30), HersheyFonts.HersheyDuplex, 0.8, Scalar.Green, 2);

                    // 沒檢測到黑點的檔名
                    resultFileName = $"{count}_blackdot_OK.png";
                }

                // 可選：保存檢測過程的中間圖像（用於調試）
                if (false)
                {
                    try
                    {
                        string debugPath = $@".\image\{st.ToString("yyyy-MM")}\{st.ToString("MMdd")}\{app.foldername}\BlackDotDebug_{stop}";
                        Directory.CreateDirectory(debugPath);

                        //Cv2.ImWrite(Path.Combine(debugPath, $"{count}_1_gray.png"), grayImage);
                        //Cv2.ImWrite(Path.Combine(debugPath, $"{count}_2_blur.png"), blurredImage);
                        //Cv2.ImWrite(Path.Combine(debugPath, $"{count}_3_morph.png"), morphImage);
                        //Cv2.ImWrite(Path.Combine(debugPath, $"{count}_4_binary.png"), binaryImage);
                        Cv2.ImWrite(Path.Combine(debugPath, resultFileName), resultImage);
                    }
                    catch (Exception ex)
                    {
                        Log.Warning($"保存黑點檢測調試圖像失敗: {ex.Message}");
                    }
                }

                // 釋放中間處理的圖像
                grayImage.Dispose();
                //blurredImage.Dispose();
                //morphImage.Dispose();
                //crossKernel.Dispose();
                binaryImage.Dispose();

                return (hasBlackDots, resultImage, validContours);
            }
            catch (Exception ex)
            {
                Log.Error($"DetectBlackDots 錯誤: {ex.Message}");
                return (false, roiImage?.Clone(), new List<Point[]>());
            }
        }

        private List<DetectionResult> ApplyOutscOverkillReduction(List<DetectionResult> validDefects, int stop)
        {
            // 篩選出outsc瑕疵和其他瑕疵
            var outscDefects = validDefects.Where(d => d.class_name == "outsc").ToList();
            var otherDefects = validDefects.Where(d => d.class_name != "outsc").ToList();

            // 如果沒有outsc瑕疵，直接返回原始結果
            if (outscDefects.Count == 0)
            {
                return validDefects;
            }

            List<DetectionResult> filteredOutscDefects = new List<DetectionResult>();

            // 從資料庫讀取站點的長邊閾值
            double outscLongerSide = GetDoubleParam(app.param, $"outscLen_{stop}", 100.0);

            if (outscDefects.Count >= 8)
            {
                // ≥8個：確實有嚴重問題，全部保留
                filteredOutscDefects.AddRange(outscDefects);
                //Log.Debug($"站點{stop}: {outscDefects.Count}個outsc瑕疵（≥8），全部保留");
            }
            else if (outscDefects.Count >= 2)
            {
                // 2-7個：先合併相鄰矩形，再檢查合併後的長邊
                var mergedDefects = MergeAdjacentOutscDefects(outscDefects, mergeDistance: 50);

                // 檢查每個合併後的矩形
                foreach (var defect in mergedDefects)
                {
                    int width = defect.box[2] - defect.box[0];
                    int height = defect.box[3] - defect.box[1];
                    double longerSide = Math.Max(width, height);

                    if (longerSide >= outscLongerSide)
                    {
                        // 合併後長邊達標，這個合併群組的所有原始框都保留
                        // 找出這個合併框對應的原始框
                        var originalDefects = FindOriginalDefects(defect, outscDefects);
                        filteredOutscDefects.AddRange(originalDefects);

                        //Log.Debug($"站點{stop}: 合併後長邊{longerSide}px ≥ 閾值{outscLongerSide}px，保留{originalDefects.Count}個原始框");
                    }
                    else
                    {
                        //Log.Debug($"站點{stop}: 合併後長邊{longerSide}px < 閾值{outscLongerSide}px，過濾");
                    }
                }
            }
            else // outscDefects.Count == 1
            {
                // 1個：直接檢查長邊
                var defect = outscDefects[0];
                int width = defect.box[2] - defect.box[0];
                int height = defect.box[3] - defect.box[1];
                double longerSide = Math.Max(width, height);

                if (longerSide >= outscLongerSide)
                {
                    filteredOutscDefects.Add(defect);
                    //Log.Debug($"站點{stop}: 單一outsc長邊{longerSide}px ≥ 閾值{outscLongerSide}px，保留");
                }
                else
                {
                    //Log.Debug($"站點{stop}: 單一outsc長邊{longerSide}px < 閾值{outscLongerSide}px，過濾");
                }
            }

            // 返回篩選後的outsc瑕疵加上其他瑕疵
            var result = new List<DetectionResult>();
            result.AddRange(filteredOutscDefects);
            result.AddRange(otherDefects);

            return result;
        }

        private List<DetectionResult> FindOriginalDefects(DetectionResult mergedDefect, List<DetectionResult> originalDefects)
        {
            var result = new List<DetectionResult>();
            Rect mergedRect = new Rect(
                mergedDefect.box[0],
                mergedDefect.box[1],
                mergedDefect.box[2] - mergedDefect.box[0],
                mergedDefect.box[3] - mergedDefect.box[1]
            );

            foreach (var original in originalDefects)
            {
                Rect originalRect = new Rect(
                    original.box[0],
                    original.box[1],
                    original.box[2] - original.box[0],
                    original.box[3] - original.box[1]
                );

                // 如果原始框與合併框有重疊或包含關係，視為屬於這個合併群組
                if (mergedRect.IntersectsWith(originalRect) || mergedRect.Contains(originalRect))
                {
                    result.Add(original);
                }
            }

            return result;
        }

        private List<DetectionResult> MergeAdjacentOutscDefects(List<DetectionResult> defects, double mergeDistance = 50)
        {
            if (defects.Count <= 1)
                return defects;

            // 創建矩形列表，附帶原始索引
            var rectList = defects.Select((d, idx) => new
            {
                Index = idx,
                Rect = new Rect(d.box[0], d.box[1],
                               d.box[2] - d.box[0],
                               d.box[3] - d.box[1]),
                Original = d
            }).ToList();

            // 使用 Union-Find 演算法找出需要合併的群組
            var parent = Enumerable.Range(0, rectList.Count).ToArray();

            int Find(int x)
            {
                if (parent[x] != x)
                    parent[x] = Find(parent[x]);
                return parent[x];
            }

            void Union(int x, int y)
            {
                int px = Find(x);
                int py = Find(y);
                if (px != py)
                    parent[px] = py;
            }

            // 檢查所有矩形對，距離小於閾值則合併
            for (int i = 0; i < rectList.Count; i++)
            {
                for (int j = i + 1; j < rectList.Count; j++)
                {
                    double distance = CalculateRectDistance(rectList[i].Rect, rectList[j].Rect);
                    if (distance <= mergeDistance)
                    {
                        Union(i, j);
                    }
                }
            }

            // 按群組分組
            var groups = rectList
                .Select((item, idx) => new { Item = item, Group = Find(idx) })
                .GroupBy(x => x.Group)
                .Select(g => g.Select(x => x.Item).ToList())
                .ToList();

            // 合併每個群組
            var mergedDefects = new List<DetectionResult>();

            foreach (var group in groups)
            {
                if (group.Count == 1)
                {
                    // 單一矩形，直接加入
                    mergedDefects.Add(group[0].Original);
                }
                else
                {
                    // 多個矩形，合併成一個
                    var mergedRect = group[0].Rect;
                    double maxScore = group[0].Original.score;

                    foreach (var item in group.Skip(1))
                    {
                        mergedRect = MergeRects(mergedRect, item.Rect);
                        maxScore = Math.Max(maxScore, item.Original.score);
                    }

                    // 創建合併後的 DetectionResult
                    var mergedDefect = new DetectionResult
                    {
                        box = new List<int>
                        {
                            mergedRect.X,
                            mergedRect.Y,
                            mergedRect.X + mergedRect.Width,
                            mergedRect.Y + mergedRect.Height
                        },
                        class_id = group[0].Original.class_id,
                        class_name = group[0].Original.class_name,
                        score = maxScore // 使用最高分數
                    };

                    mergedDefects.Add(mergedDefect);
                }
            }

            return mergedDefects;
        }

        private Rect MergeRects(Rect rect1, Rect rect2)
        {
            int x1 = Math.Min(rect1.Left, rect2.Left);
            int y1 = Math.Min(rect1.Top, rect2.Top);
            int x2 = Math.Max(rect1.Right, rect2.Right);
            int y2 = Math.Max(rect1.Bottom, rect2.Bottom);

            return new Rect(x1, y1, x2 - x1, y2 - y1);
        }
        private double CalculateRectDistance(Rect rect1, Rect rect2)
        {
            // 檢查是否重疊
            if (rect1.IntersectsWith(rect2))
                return 0;

            // 計算水平和垂直距離
            double dx = Math.Max(0, Math.Max(rect1.Left - rect2.Right, rect2.Left - rect1.Right));
            double dy = Math.Max(0, Math.Max(rect1.Top - rect2.Bottom, rect2.Top - rect1.Bottom));

            return Math.Sqrt(dx * dx + dy * dy);
        }

        /// <summary>
        /// 基於實際分析數據的簡化色彩驗證器
        /// 使用3+1特徵模型：G、V、R為主要特徵，B為輔助特徵
        /// </summary>
        public class SimplifiedColorVerifier
        {
            // 基於實際分析報告的統計標準
            private readonly Dictionary<int, ColorStandard> standards = new Dictionary<int, ColorStandard>
            {
                {
                    3, new ColorStandard  // 站3 - 使用OK1數據
                    {
                        // 主要特徵：G, V, R
                        PrimaryMeans = new double[] { 160.81, 163.91, 137.84 },
                        PrimaryStds = new double[] { 7.49, 7.32, 9.48 },
                        // 輔助特徵：B
                        SecondaryMeans = new double[] { 151.45 },
                        SecondaryStds = new double[] { 9.08 },
                        StationInfo = "站3 (OK1標準)"
                    }
                },
                {
                    4, new ColorStandard  // 站4 - 使用OK2數據
                    {
                        // 主要特徵：G, V, R
                        PrimaryMeans = new double[] { 108.59, 112.63, 92.59 },
                        PrimaryStds = new double[] { 6.16, 5.57, 6.53 },
                        // 輔助特徵：B
                        SecondaryMeans = new double[] { 106.95 },
                        SecondaryStds = new double[] { 5.59 },
                        StationInfo = "站4 (OK2標準)"
                    }
                }
            };

            /// <summary>
            /// 對檢測瑕疵進行色彩複檢
            /// 替換原有的 ApplyOtpColorDetection 功能
            /// </summary>
            public List<DetectionResult> VerifyDefectsByColor(List<DetectionResult> detections, Mat roi, int station,
                double primaryThreshold = 2.5, double secondaryThreshold = 5.0)
            {
                if (!standards.ContainsKey(station))
                {
                    //Log.Warning($"站別{station}不支援色彩檢測，保留所有檢測結果");
                    return detections;
                }

                if (detections == null || detections.Count == 0)
                {
                    return detections;
                }

                var standard = standards[station];
                List<DetectionResult> confirmedDefects = new List<DetectionResult>();
                int originalCount = detections.Count;

                //Log.Debug($"=== {standard.StationInfo} 色彩複檢開始 ===");
                //Log.Debug($"待檢瑕疵數: {originalCount}");

                foreach (var defect in detections)
                {
                    // 擷取檢測框區域
                    int x = Math.Max(0, (int)defect.box[0]);
                    int y = Math.Max(0, (int)defect.box[1]);
                    int width = Math.Min(roi.Width - x, (int)defect.box[2]);
                    int height = Math.Min(roi.Height - y, (int)defect.box[3]);

                    Rect defectRect = new Rect(x, y, width, height);

                    if (defectRect.Width < 20 || defectRect.Height < 20)
                    {
                        // 檢測框太小，直接保留
                        confirmedDefects.Add(defect);
                        //Log.Debug($"✓ 保留小尺寸瑕疵: {defect.class_name} (檢測框 {defectRect.Width}x{defectRect.Height})");
                        continue;
                    }

                    try
                    {
                        Mat defectRoi = new Mat(roi, defectRect);

                        // 進行色彩驗證
                        var verificationResult = VerifyDefectRegion(defectRoi, standard, primaryThreshold, secondaryThreshold);

                        /*
                        // 記錄詳細分析結果
                        Log.Debug($"檢測框位置: ({defectRect.X}, {defectRect.Y}), 大小: {defectRect.Width}x{defectRect.Height}");
                        Log.Debug($"瑕疵類別: {defect.class_name}, YOLO信心度: {defect.score:F3}");
                        Log.Debug($"色彩檢測: {(verificationResult.IsNG ? "NG" : "OK")}, 信心度: {verificationResult.Confidence:F3}");
                        Log.Debug($"判定原因: {verificationResult.Reason}");
                        */

                        // 輸出特徵分析
                        foreach (var detail in verificationResult.AnalysisDetails)
                        {
                            //Log.Debug($"  {detail}");
                        }

                        // 只有色彩檢測也判定為NG的才確認為真正的瑕疵
                        if (verificationResult.IsNG)
                        {
                            confirmedDefects.Add(defect);
                            //Log.Debug($"✓ 確認瑕疵: {defect.class_name} (色彩異常)");
                        }
                        else
                        {
                            //Log.Debug($"✗ 排除疑似瑕疵: {defect.class_name} (色彩正常)");
                        }
                        defectRoi.Dispose();
                    }
                    catch (Exception ex)
                    {
                        //Log.Error($"色彩檢測發生錯誤: {ex.Message}");
                        // 發生錯誤時保留瑕疵
                        confirmedDefects.Add(defect);
                    }
                }

                int confirmedCount = confirmedDefects.Count;
                double filterRate = originalCount > 0 ? ((double)(originalCount - confirmedCount) / originalCount * 100) : 0;

                //Log.Debug($"=== {standard.StationInfo} 色彩複檢完成 ===");
                //Log.Debug($"原始瑕疵數: {originalCount}, 確認瑕疵數: {confirmedCount}, 過濾率: {filterRate:F1}%");

                return confirmedDefects;
            }

            /// <summary>
            /// 對單一檢測區域進行色彩驗證
            /// 使用3+1特徵模型
            /// </summary>
            private ColorVerificationResult VerifyDefectRegion(Mat defectRoi, ColorStandard standard,
                double primaryThreshold, double secondaryThreshold)
            {
                var result = new ColorVerificationResult();

                // 1. 過濾有效像素
                Mat validMask = CreateValidPixelMask(defectRoi);
                int validPixelCount = Cv2.CountNonZero(validMask);
                double validRatio = (double)validPixelCount / (defectRoi.Width * defectRoi.Height);

                if (validRatio < 0.05) // 有效像素少於30%
                {
                    validMask.Dispose();
                    return new ColorVerificationResult
                    {
                        IsNG = false,
                        Confidence = 0,
                        Reason = $"有效像素比例過低: {validRatio:P2}"
                    };
                }

                // 2. 計算RGB和HSV特徵
                Scalar rgbMean = Cv2.Mean(defectRoi, validMask);
                Mat hsvRoi = new Mat();
                Cv2.CvtColor(defectRoi, hsvRoi, ColorConversionCodes.BGR2HSV);
                Scalar hsvMean = Cv2.Mean(hsvRoi, validMask);

                // BGR -> RGB，提取關鍵特徵
                double gValue = rgbMean.Val1;  // G通道 (BGR中的G)
                double vValue = hsvMean.Val2;  // V通道 (亮度)
                double rValue = rgbMean.Val2;  // R通道 (BGR中的R)
                double bValue = rgbMean.Val0;  // B通道 (BGR中的B)

                // 3. 計算主要特徵的Z分數 (G, V, R)
                double[] primaryFeatures = { gValue, vValue, rValue };
                double[] primaryZScores = new double[3];

                for (int i = 0; i < 3; i++)
                {
                    primaryZScores[i] = Math.Abs(primaryFeatures[i] - standard.PrimaryMeans[i]) / standard.PrimaryStds[i];
                }

                // 4. 計算輔助特徵的Z分數 (B)
                double secondaryZScore = Math.Abs(bValue - standard.SecondaryMeans[0]) / standard.SecondaryStds[0];

                // 5. 判斷邏輯
                double maxPrimaryZ = primaryZScores.Max();
                bool primaryPassed = maxPrimaryZ <= primaryThreshold;
                bool secondaryPassed = secondaryZScore <= secondaryThreshold;

                // 主要特徵必須通過，輔助特徵可以放寬
                result.IsNG = !(primaryPassed && secondaryPassed);
                result.Confidence = primaryPassed ? (1.0 - maxPrimaryZ / (primaryThreshold + 1.0)) : 0.0;

                if (!result.IsNG)
                {
                    result.Reason = "色彩特徵正常";
                }
                else if (!primaryPassed)
                {
                    result.Reason = $"主要特徵異常 (最大Z分數: {maxPrimaryZ:F2})";
                }
                else
                {
                    result.Reason = $"輔助特徵異常 (B通道Z分數: {secondaryZScore:F2})";
                }

                // 6. 詳細分析資訊
                result.AnalysisDetails = new List<string>
                {
                    $"檢測特徵 - G:{gValue:F1}, V:{vValue:F1}, R:{rValue:F1}, B:{bValue:F1}",
                    $"標準特徵 - G:{standard.PrimaryMeans[0]:F1}, V:{standard.PrimaryMeans[1]:F1}, R:{standard.PrimaryMeans[2]:F1}, B:{standard.SecondaryMeans[0]:F1}",
                    $"主要Z分數 - G:{primaryZScores[0]:F2}, V:{primaryZScores[1]:F2}, R:{primaryZScores[2]:F2} (閾值:{primaryThreshold})",
                    $"輔助Z分數 - B:{secondaryZScore:F2} (閾值:{secondaryThreshold})",
                    $"有效像素比例: {validRatio:P2}"
                };

                // 清理資源
                validMask.Dispose();
                hsvRoi.Dispose();

                return result;
            }

            private Mat CreateValidPixelMask(Mat roi)
            {
                Mat mask = new Mat(roi.Size(), MatType.CV_8UC1, Scalar.Black);
                Mat grayRoi = new Mat();
                Cv2.CvtColor(roi, grayRoi, ColorConversionCodes.BGR2GRAY);

                // 排除過暗（<20）和過亮（>240）的像素
                Cv2.Threshold(grayRoi, mask, 20, 255, ThresholdTypes.Binary);
                Mat brightMask = new Mat();
                Cv2.Threshold(grayRoi, brightMask, 240, 255, ThresholdTypes.BinaryInv);
                Cv2.BitwiseAnd(mask, brightMask, mask);

                grayRoi.Dispose();
                brightMask.Dispose();

                return mask;
            }
        }

        /// <summary>
        /// 色彩標準結構
        /// </summary>
        public class ColorStandard
        {
            public double[] PrimaryMeans { get; set; }    // 主要特徵均值：G, V, R
            public double[] PrimaryStds { get; set; }     // 主要特徵標準差
            public double[] SecondaryMeans { get; set; }  // 輔助特徵均值：B
            public double[] SecondaryStds { get; set; }   // 輔助特徵標準差
            public string StationInfo { get; set; }       // 站別資訊
        }

        /// <summary>
        /// 色彩驗證結果
        /// </summary>
        public class ColorVerificationResult
        {
            public bool IsNG { get; set; }                     // 是否為NG
            public double Confidence { get; set; }             // 信心度
            public string Reason { get; set; }                 // 判定原因
            public List<string> AnalysisDetails { get; set; } = new List<string>(); // 詳細分析
        }

        #region OTP五彩鋅色彩檢測相關函數0


        /// <summary>
        /// 對OTP瑕疵進行色彩檢測分析
        /// </summary>
        /// <param name="validDefects">有效的瑕疵檢測結果</param>
        /// <param name="roiImage">ROI圖像</param>
        /// <param name="stop">站點編號</param>
        /// <returns>經過色彩檢測篩選後的瑕疵檢測結果</returns>
        private List<DetectionResult> ApplyOtpColorDetection(List<DetectionResult> validDefects, Mat roiImage, int stop)
        {
            // 篩選出OTP瑕疵和其他瑕疵
            var otpDefects = validDefects.Where(d => d.class_name == "OTP").ToList();
            var otherDefects = validDefects.Where(d => d.class_name != "OTP").ToList();

            // 如果沒有OTP瑕疵，直接返回原始結果
            if (otpDefects.Count == 0)
            {
                return validDefects;
            }

            List<DetectionResult> filteredOtpDefects = new List<DetectionResult>();

            if (otpDefects.Count == 1)
            {
                // 單個框：嚴格AOI檢測
                var defect = otpDefects[0];

                // 提取瑕疵區域進行色彩分析
                Mat defectRoi = ExtractDefectRegion(roiImage, defect);
                var colorAnalysis = AnalyzeOtpColorFeatures(defectRoi, stop);

                // 嚴格模式判斷
                if (IsOtpDefectByColorAnalysis(colorAnalysis, stop, OtpDetectionMode.Strict))
                {
                    filteredOtpDefects.Add(defect);
                }
                // 如果色彩分析認為不是瑕疵，就過濾掉這個AI檢測結果

                defectRoi.Dispose();
            }
            else if (otpDefects.Count == 2)
            {
                // 兩個框：中等AOI檢測
                foreach (var defect in otpDefects)
                {
                    Mat defectRoi = ExtractDefectRegion(roiImage, defect);
                    var colorAnalysis = AnalyzeOtpColorFeatures(defectRoi, stop);

                    // 中等模式判斷
                    if (IsOtpDefectByColorAnalysis(colorAnalysis, stop, OtpDetectionMode.Moderate))
                    {
                        filteredOtpDefects.Add(defect);
                    }

                    defectRoi.Dispose();
                }
            }
            else if (otpDefects.Count >= 3)
            {
                // 三個框以上：暫時不做AOI檢測，直接通過
                filteredOtpDefects.AddRange(otpDefects);

                // 保留寬鬆檢測的程式碼供未來使用
                /*
                foreach (var defect in otpDefects)
                {
                    Mat defectRoi = ExtractDefectRegion(roiImage, defect);
                    var colorAnalysis = AnalyzeOtpColorFeatures(defectRoi, stop);

                    if (IsOtpDefectByColorAnalysis(colorAnalysis, stop, OtpDetectionMode.Lenient))
                    {
                        filteredOtpDefects.Add(defect);
                    }

                    defectRoi.Dispose();
                }
                */
            }

            // 返回篩選後的OTP瑕疵加上其他瑕疵
            var result = new List<DetectionResult>();
            result.AddRange(filteredOtpDefects);
            result.AddRange(otherDefects);

            return result;
        }

        /// <summary>
        /// 從ROI圖像中提取瑕疵區域
        /// </summary>
        private Mat ExtractDefectRegion(Mat roiImage, DetectionResult defect)
        {
            // 計算瑕疵框的座標
            int x = defect.box[0];
            int y = defect.box[1];
            int width = defect.box[2] - defect.box[0];
            int height = defect.box[3] - defect.box[1];

            // 確保座標在圖像範圍內
            x = Math.Max(0, Math.Min(x, roiImage.Width - 1));
            y = Math.Max(0, Math.Min(y, roiImage.Height - 1));
            width = Math.Min(width, roiImage.Width - x);
            height = Math.Min(height, roiImage.Height - y);

            // 提取瑕疵區域
            Rect defectRect = new Rect(x, y, width, height);
            Mat defectRoi = new Mat(roiImage, defectRect);

            return defectRoi.Clone(); // 返回副本避免記憶體問題
        }

        /// <summary>
        /// 分析OTP的色彩特徵
        /// </summary>
        /// <param name="roiImage">ROI圖像</param>
        /// <param name="stop">站點編號</param>
        /// <returns>色彩分析結果</returns>
        private OtpColorAnalysis AnalyzeOtpColorFeatures(Mat roiImage, int stop)
        {
            var result = new OtpColorAnalysis();

            // 創建更嚴格的非黑色遮罩
            Mat mask = CreateNonBlackMask(roiImage);

            // 計算有效像素數量
            result.ValidPixelCount = Cv2.CountNonZero(mask);

            if (result.ValidPixelCount < 1000) // 最低有效像素要求
            {
                mask.Dispose();
                return result;
            }

            // 轉換到不同色彩空間
            Mat hsvImage = new Mat();
            Mat labImage = new Mat();
            Cv2.CvtColor(roiImage, hsvImage, ColorConversionCodes.BGR2HSV);
            Cv2.CvtColor(roiImage, labImage, ColorConversionCodes.BGR2Lab);

            // 分析RGB（排除黑色像素）
            AnalyzeColorSpace(roiImage, mask, result.RgbMean, result.RgbStdDev);

            // 分析HSV（排除黑色像素）
            AnalyzeColorSpace(hsvImage, mask, result.HsvMean, result.HsvStdDev);

            // 分析Lab（排除黑色像素）
            AnalyzeLabColorSpace2(labImage, mask, result.LabMean, result.LabStdDev);

            // 計算自定義特徵（排除黑色像素）
            CalculateOtpCustomFeatures(roiImage, hsvImage, mask, result);

            // 釋放資源
            mask.Dispose();
            hsvImage.Dispose();
            labImage.Dispose();

            return result;
        }

        /// <summary>
        /// 創建更嚴格的非黑色遮罩
        /// </summary>
        /// <param name="image">輸入圖像</param>
        /// <returns>遮罩圖像</returns>
        private Mat CreateNonBlackMask(Mat image)
        {
            Mat mask = new Mat(image.Size(), MatType.CV_8UC1, Scalar.Black);

            // 方法1: 使用RGB總和來判斷
            Mat[] bgr = image.Split();
            Mat sumMat = new Mat();
            Cv2.Add(bgr[0], bgr[1], sumMat);
            Cv2.Add(sumMat, bgr[2], sumMat);

            // RGB總和大於30的像素視為有效
            Cv2.Threshold(sumMat, mask, 30, 255, ThresholdTypes.Binary);

            // 方法2: 結合HSV的飽和度和明度
            Mat hsv = new Mat();
            Cv2.CvtColor(image, hsv, ColorConversionCodes.BGR2HSV);
            Mat[] hsvChannels = hsv.Split();

            // 明度大於10或飽和度大於5的像素視為有效
            Mat valueMask = new Mat();
            Mat satMask = new Mat();
            Cv2.Threshold(hsvChannels[2], valueMask, 10, 255, ThresholdTypes.Binary);
            Cv2.Threshold(hsvChannels[1], satMask, 5, 255, ThresholdTypes.Binary);

            Mat combinedMask = new Mat();
            Cv2.BitwiseOr(valueMask, satMask, combinedMask);

            // 最終遮罩：RGB總和 AND (明度 OR 飽和度)
            Cv2.BitwiseAnd(mask, combinedMask, mask);

            // 形態學操作去除噪點
            Mat kernel = Cv2.GetStructuringElement(MorphShapes.Ellipse, new Size(3, 3));
            Cv2.MorphologyEx(mask, mask, MorphTypes.Open, kernel);

            // 釋放暫時資源
            foreach (var channel in bgr) channel.Dispose();
            sumMat.Dispose();
            hsv.Dispose();
            foreach (var channel in hsvChannels) channel.Dispose();
            valueMask.Dispose();
            satMask.Dispose();
            combinedMask.Dispose();
            kernel.Dispose();

            return mask;
        }

        /// <summary>
        /// 分析色彩空間統計特徵
        /// </summary>
        /// <param name="image">圖像</param>
        /// <param name="mask">遮罩</param>
        /// <param name="means">均值數組</param>
        /// <param name="stdDevs">標準差數組</param>
        private void AnalyzeColorSpace(Mat image, Mat mask, double[] means, double[] stdDevs)
        {
            Mat[] channels = image.Split();

            for (int c = 0; c < 3; c++)
            {
                Scalar meanScalar, stdDevScalar;
                Cv2.MeanStdDev(channels[c], out meanScalar, out stdDevScalar, mask);
                means[c] = meanScalar.Val0;
                stdDevs[c] = stdDevScalar.Val0;

                channels[c].Dispose();
            }
        }

        /// <summary>
        /// 分析Lab色彩空間
        /// </summary>
        /// <param name="labImage">Lab圖像</param>
        /// <param name="mask">遮罩</param>
        /// <param name="means">均值數組</param>
        /// <param name="stdDevs">標準差數組</param>
        private void AnalyzeLabColorSpace(Mat labImage, Mat mask, double[] means, double[] stdDevs)
        {
            Mat[] channels = labImage.Split();

            for (int c = 0; c < 3; c++)
            {
                Scalar meanScalar, stdDevScalar;
                Cv2.MeanStdDev(channels[c], out meanScalar, out stdDevScalar, mask);
                means[c] = meanScalar.Val0;
                stdDevs[c] = stdDevScalar.Val0;

                channels[c].Dispose();
            }
        }

        /// <summary>
        /// 計算OTP自定義特徵
        /// </summary>
        /// <param name="bgrImage">BGR圖像</param>
        /// <param name="hsvImage">HSV圖像</param>
        /// <param name="mask">遮罩</param>
        /// <param name="result">結果對象</param>
        private void CalculateOtpCustomFeatures(Mat bgrImage, Mat hsvImage, Mat mask, OtpColorAnalysis result)
        {
            Mat[] hsvChannels = hsvImage.Split();

            // 計算平均飽和度和亮度
            Scalar satMean, valMean, satStdDev, valStdDev;
            Cv2.MeanStdDev(hsvChannels[1], out satMean, out satStdDev, mask);
            Cv2.MeanStdDev(hsvChannels[2], out valMean, out valStdDev, mask);

            result.Saturation = satMean.Val0;
            result.Brightness = valMean.Val0;
            result.ColorfulnessIndex = satStdDev.Val0 + valStdDev.Val0;

            // 計算主導色調
            Mat hueHist = new Mat();
            int[] histSize = { 180 };
            Rangef[] ranges = { new Rangef(0, 180) };
            Cv2.CalcHist(new[] { hsvChannels[0] }, new[] { 0 }, mask, hueHist, 1, histSize, ranges);

            Point minLoc, maxLoc;
            double minVal, maxVal;
            Cv2.MinMaxLoc(hueHist, out minVal, out maxVal, out minLoc, out maxLoc);
            result.DominantHue = maxLoc.Y * 2;

            // 計算色彩分布均勻性
            result.ColorUniformity = CalculateColorUniformity(hsvChannels[0], mask);

            // 計算RGB灰階化程度
            result.GrayishLevel = CalculateGrayishLevel(result.RgbMean);

            // 釋放資源
            foreach (var channel in hsvChannels)
                channel.Dispose();
            hueHist.Dispose();
        }

        /// <summary>
        /// 計算色彩分布均勻性
        /// </summary>
        /// <param name="hueChannel">色調通道</param>
        /// <param name="mask">遮罩</param>
        /// <returns>均勻性指數</returns>
        private double CalculateColorUniformity(Mat hueChannel, Mat mask)
        {
            try
            {
                Mat hist = new Mat();
                int[] histSize = { 36 };
                Rangef[] ranges = { new Rangef(0, 180) };
                Cv2.CalcHist(new[] { hueChannel }, new[] { 0 }, mask, hist, 1, histSize, ranges);

                float[] histData = new float[36];
                Marshal.Copy(hist.Data, histData, 0, 36);

                double mean = histData.Average();
                double variance = histData.Select(x => Math.Pow(x - mean, 2)).Average();
                double uniformity = Math.Sqrt(variance);

                hist.Dispose();
                return uniformity;
            }
            catch
            {
                return 0.0;
            }
        }

        /// <summary>
        /// 計算RGB灰階化程度
        /// </summary>
        /// <param name="rgbMean">RGB均值</param>
        /// <returns>灰階化程度</returns>
        private double CalculateGrayishLevel(double[] rgbMean)
        {
            double rgDiff = Math.Abs(rgbMean[2] - rgbMean[1]); // R-G差異
            double rbDiff = Math.Abs(rgbMean[2] - rgbMean[0]); // R-B差異  
            double gbDiff = Math.Abs(rgbMean[1] - rgbMean[0]); // G-B差異

            return (rgDiff + rbDiff + gbDiff) / 3.0;
        }

        /// <summary>
        /// 根據色彩分析結果判斷是否為OTP瑕疵
        /// </summary>
        /// <param name="colorAnalysis">色彩分析結果</param>
        /// <param name="stop">站點編號</param>
        /// <param name="mode">檢測模式</param>
        /// <returns>是否為瑕疵</returns>
        private bool IsOtpDefectByColorAnalysis(OtpColorAnalysis colorAnalysis, int stop, OtpDetectionMode mode)
        {
            if (colorAnalysis.ValidPixelCount < 1000)
                return false;

            // 根據站點和檢測模式選擇不同的閾值
            var thresholds = GetOtpThresholds(stop, mode);

            // 計算缺陷信心度
            double defectConfidence = CalculateOtpDefectConfidence(colorAnalysis, thresholds);

            // 根據檢測模式設定不同的信心度閾值
            double confidenceThreshold;
            switch (mode)
            {
                case OtpDetectionMode.Strict:
                    confidenceThreshold = 0.8;    // 嚴格模式需要80%以上信心度
                    break;
                case OtpDetectionMode.Moderate:
                    confidenceThreshold = 0.6;    // 中等模式需要60%以上信心度
                    break;
                case OtpDetectionMode.Lenient:
                    confidenceThreshold = 0.4;    // 寬鬆模式需要40%以上信心度
                    break;
                default:
                    confidenceThreshold = 0.6;
                    break;
            }
            return defectConfidence >= confidenceThreshold;
        }

        /// <summary>
        /// 獲取OTP檢測閾值
        /// </summary>
        /// <param name="stop">站點編號</param>
        /// <param name="mode">檢測模式</param>
        /// <returns>閾值設定</returns>
        private OtpThresholds GetOtpThresholds(int stop, OtpDetectionMode mode)
        {
            // 基於最新統計數據設定閾值
            if (stop == 3)
            {
                // 第3站閾值設定（基於OK1的最新分析）
                // OK1: 飽和度=43.46±11.34, 亮度=163.91±7.32, 色彩豐富度=30.78±4.58
                // 主導色調: 均值=162.85°, 標準差=40.39°
                switch (mode)
                {
                    case OtpDetectionMode.Strict:
                        return new OtpThresholds
                        {
                            // 嚴格模式：使用1σ範圍
                            SaturationMin = 32.0,       // OK1均值-1σ = 43.46-11.34 = 32.12
                            BrightnessMin = 156.0,      // OK1均值-1σ = 163.91-7.32 = 156.59
                            ColorfulnessMin = 26.0,     // OK1均值-1σ = 30.78-4.58 = 26.20
                            GrayishLevelMax = 25.0,     // 灰階化程度上限
                            DominantHueMin = 120.0,     // 主導色調範圍 162.85±40 ≈ 120-200
                            DominantHueMax = 200.0      // 考慮色調的環形特性
                        };
                    case OtpDetectionMode.Moderate:
                        return new OtpThresholds
                        {
                            // 中等模式：使用1.5σ範圍
                            SaturationMin = 26.0,       // OK1均值-1.5σ ≈ 26
                            BrightnessMin = 153.0,      // OK1均值-1.5σ ≈ 153
                            ColorfulnessMin = 24.0,     // OK1均值-1.5σ ≈ 24
                            GrayishLevelMax = 30.0,     // 稍微放寬
                            DominantHueMin = 100.0,     // 放寬色調範圍
                            DominantHueMax = 220.0
                        };
                    case OtpDetectionMode.Lenient:
                        return new OtpThresholds
                        {
                            // 寬鬆模式：使用2σ範圍
                            SaturationMin = 21.0,       // OK1均值-2σ ≈ 21
                            BrightnessMin = 149.0,      // OK1均值-2σ ≈ 149
                            ColorfulnessMin = 22.0,     // OK1均值-2σ ≈ 22
                            GrayishLevelMax = 35.0,     // 更寬鬆
                            DominantHueMin = 80.0,      // 更寬的色調範圍
                            DominantHueMax = 240.0
                        };
                    default:
                        return new OtpThresholds();
                }
            }
            else if (stop == 4)
            {
                // 第4站閾值設定（基於OK2的最新分析）
                // OK2: 飽和度=46.83±9.03, 亮度=112.63±5.57, 色彩豐富度=37.01±3.88
                // 主導色調: 均值=179.21°, 標準差=30.73°

                switch (mode)
                {
                    case OtpDetectionMode.Strict:
                        return new OtpThresholds
                        {
                            // 嚴格模式：第4站整體數值較低
                            SaturationMin = 37.0,       // OK2均值-1σ = 46.83-9.03 = 37.80
                            BrightnessMin = 107.0,      // OK2均值-1σ = 112.63-5.57 = 107.06
                            ColorfulnessMin = 33.0,     // OK2均值-1σ = 37.01-3.88 = 33.13
                            GrayishLevelMax = 25.0,     // 灰階化程度上限
                            DominantHueMin = 148.0,     // 主導色調範圍 179.21±31 ≈ 148-210
                            DominantHueMax = 210.0
                        };
                    case OtpDetectionMode.Moderate:
                        return new OtpThresholds
                        {
                            // 中等模式
                            SaturationMin = 33.0,       // OK2均值-1.5σ ≈ 33
                            BrightnessMin = 104.0,      // OK2均值-1.5σ ≈ 104
                            ColorfulnessMin = 31.0,     // OK2均值-1.5σ ≈ 31
                            GrayishLevelMax = 30.0,
                            DominantHueMin = 130.0,     // 放寬色調範圍
                            DominantHueMax = 230.0
                        };
                    case OtpDetectionMode.Lenient:
                        return new OtpThresholds
                        {
                            // 寬鬆模式
                            SaturationMin = 28.0,       // OK2均值-2σ ≈ 28
                            BrightnessMin = 101.0,      // OK2均值-2σ ≈ 101
                            ColorfulnessMin = 29.0,     // OK2均值-2σ ≈ 29
                            GrayishLevelMax = 35.0,
                            DominantHueMin = 110.0,     // 更寬的色調範圍
                            DominantHueMax = 250.0
                        };
                    default:
                        return new OtpThresholds();
                }
            }

            return new OtpThresholds(); // 預設值
        }

        /// <summary>
        /// 計算OTP缺陷信心度
        /// </summary>
        /// <param name="analysis">色彩分析結果</param>
        /// <param name="thresholds">閾值設定</param>
        /// <returns>缺陷信心度 (0-1)</returns>
        private double CalculateOtpDefectConfidence(OtpColorAnalysis analysis, OtpThresholds thresholds)
        {
            double confidence = 0.0;

            // 飽和度檢測 (權重35%)
            if (analysis.Saturation < thresholds.SaturationMin * 0.7)
                confidence += 0.35;
            else if (analysis.Saturation < thresholds.SaturationMin)
                confidence += 0.20;

            // 亮度檢測 (權重30%)
            if (analysis.Brightness < thresholds.BrightnessMin * 0.9)
                confidence += 0.30;
            else if (analysis.Brightness < thresholds.BrightnessMin)
                confidence += 0.15;

            // 色彩豐富度檢測 (權重20%)
            if (analysis.ColorfulnessIndex < thresholds.ColorfulnessMin * 0.8)
                confidence += 0.20;
            else if (analysis.ColorfulnessIndex < thresholds.ColorfulnessMin)
                confidence += 0.10;

            // 灰階化檢測 (權重10%)
            if (analysis.GrayishLevel < thresholds.GrayishLevelMax * 0.5)
                confidence += 0.10;
            else if (analysis.GrayishLevel < thresholds.GrayishLevelMax)
                confidence += 0.05;

            // 色調檢測 (權重5%)
            if (analysis.DominantHue < thresholds.DominantHueMin ||
                analysis.DominantHue > thresholds.DominantHueMax)
                confidence += 0.05;

            return Math.Min(confidence, 1.0);
        }

        #endregion

        #region OTP相關數據結構

        /// <summary>
        /// OTP色彩分析結果
        /// </summary>
        public class OtpColorAnalysis
        {
            public double[] RgbMean { get; set; } = new double[3];
            public double[] RgbStdDev { get; set; } = new double[3];
            public double[] HsvMean { get; set; } = new double[3];
            public double[] HsvStdDev { get; set; } = new double[3];
            public double[] LabMean { get; set; } = new double[3];
            public double[] LabStdDev { get; set; } = new double[3];

            public double Saturation { get; set; }
            public double Brightness { get; set; }
            public double ColorfulnessIndex { get; set; }
            public double DominantHue { get; set; }
            public double ColorUniformity { get; set; }
            public double GrayishLevel { get; set; }
            public int ValidPixelCount { get; set; }
        }

        /// <summary>
        /// OTP檢測模式
        /// </summary>
        public enum OtpDetectionMode
        {
            Strict,    // 嚴格模式
            Moderate,  // 中等模式  
            Lenient    // 寬鬆模式
        }

        /// <summary>
        /// OTP檢測閾值
        /// </summary>
        public class OtpThresholds
        {
            // 未來可能需要調整的參數：
            // 1. 飽和度相關參數 (SaturationMin, SaturationWeight)
            // 2. 亮度相關參數 (BrightnessMin, BrightnessWeight)
            // 3. 色彩豐富度參數 (ColorfulnessMin, ColorfulnessWeight)
            // 4. 灰階化程度參數 (GrayishLevelMax, GrayishWeight)
            // 5. 色調範圍參數 (DominantHueMin, DominantHueMax, HueWeight)
            // 6. 各檢測模式的信心度閾值
            // 7. 各特徵的權重分配
            // 8. 最低有效像素數量要求
            // 9. 黑色遮罩的創建參數 (RGB總和閾值, HSV閾值)
            // 10. 形態學操作的kernel大小

            public double SaturationMin { get; set; }     // 飽和度最小值
            public double BrightnessMin { get; set; }     // 亮度最小值
            public double ColorfulnessMin { get; set; }   // 色彩豐富度最小值
            public double GrayishLevelMax { get; set; }   // 灰階化程度最大值
            public double DominantHueMin { get; set; }    // 主導色調最小值
            public double DominantHueMax { get; set; }    // 主導色調最大值
        }

        #endregion

        public void showResultMat(Mat img, int stop)
        {
            if (img == null) return;

            // 控制更新频率
            if (DateTime.Now - app.lastUpdateTime < app.minUpdateInterval)
                return;

            app.lastUpdateTime = DateTime.Now;

            // 深拷贝图像，避免线程安全问题
            Mat imgCopy = img.Clone();

            // 在后台线程处理图像
            BeginInvoke(new Action(() =>
            {
                try
                {
                    PictureBox targetPictureBox = GetPictureBoxByStop(stop);
                    if (targetPictureBox != null)
                    {
                        using (Mat resizedImg = new Mat())
                        {
                            Cv2.Resize(imgCopy, resizedImg, new Size(345, 345));

                            // 暫存舊圖像
                            var oldImage = targetPictureBox.Image;

                            // 設置新圖像
                            targetPictureBox.Image = resizedImg.ToBitmap();

                            // 舊圖像丟入待釋放佇列
                            if (oldImage != null && oldImage != targetPictureBox.Image)
                            {
                                _disposeQueue.Enqueue(oldImage);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"更新圖片時發生異常：{ex.Message}");
                }
                finally
                {
                    imgCopy.Dispose();
                }
            }));
        }

        // 根據 stop 值返回對應的 PictureBox
        private PictureBox GetPictureBoxByStop(int stop)
        {
            switch (stop)
            {
                case 1:
                    return cherngerPictureBox1;
                case 2:
                    return cherngerPictureBox2;
                case 3:
                    return cherngerPictureBox3;
                case 4:
                    return cherngerPictureBox4;
                default:
                    return null;
            }
        }
        private void WriteToLog(string content, string fileName)
        {
            try
            {
                // 建立logs資料夾（如果不存在）
                string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // 完整的日誌檔案路徑
                string logFile = Path.Combine(logDir, fileName);

                // 使用檔案鎖定確保多執行緒環境下的安全性
                lock (this)
                {
                    // 附加內容到檔案
                    File.AppendAllText(logFile, content + Environment.NewLine);
                }

                // 可選：同時使用 Serilog 記錄
                //Log.Debug($"已寫入日誌檔案 {fileName}");
            }
            catch (Exception ex)
            {
                // 如果寫入日誌檔案失敗，至少嘗試用 Serilog 記錄錯誤
                Log.Error($"寫入日誌檔案 {fileName} 時發生錯誤: {ex.Message}");
            }
        }

        public static class PerformanceProfiler
        {
            private static readonly ConcurrentDictionary<string, ConcurrentBag<double>> _timings =
                new ConcurrentDictionary<string, ConcurrentBag<double>>();

            private static readonly ConcurrentDictionary<int, ConcurrentDictionary<string, Stopwatch>> _activeTimers =
                new ConcurrentDictionary<int, ConcurrentDictionary<string, Stopwatch>>();

            private static readonly object _fileLock = new object();
            private static string _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "performance_log.txt");
            private static string _csvLogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "performance_log.csv");
            private static bool _isInitialized = false;
            private static System.Threading.Timer _autoSaveTimer;

            /// <summary>
            /// 初始化性能分析器
            /// </summary>
            /// <param name="logFilePath">日誌檔案路徑，如果為null則使用默認路徑</param>
            /// <param name="autoSaveInterval">自動保存間隔(毫秒)，設為0禁用自動保存</param>
            public static void Initialize(string logFilePath = null, int autoSaveInterval = 60000)
            {
                if (_isInitialized) return;

                // 產生包含日期時間的檔名
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string baseFileName = "performance_log_" + timestamp;

                if (!string.IsNullOrEmpty(logFilePath))
                {
                    string temp_directory = Path.GetDirectoryName(logFilePath);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(logFilePath) + "_" + timestamp;
                    string ext = Path.GetExtension(logFilePath);

                    _logFilePath = Path.Combine(temp_directory, fileNameWithoutExt + ext);
                    _csvLogFilePath = Path.Combine(temp_directory, fileNameWithoutExt + ".csv");
                }
                else
                {
                    _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, baseFileName + ".txt");
                    _csvLogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, baseFileName + ".csv");
                }

                // 確保目錄存在
                string directory = Path.GetDirectoryName(_logFilePath);
                if (!Directory.Exists(directory) && !string.IsNullOrEmpty(directory))
                    Directory.CreateDirectory(directory);

                // 寫入整齊格式的標題行 (文字檔)
                string header = string.Format(
                    "| {0,-25} | {1,-10} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-20} |\r\n",
                    "函數名稱", "調用次數", "平均時間(ms)", "最短時間(ms)", "最長時間(ms)", "總時間(ms)", "記錄日期時間"
                );

                string separator = new string('-', header.Length) + "\r\n";

                File.WriteAllText(_logFilePath,
                    $"性能分析報告 - 開始時間: {DateTime.Now}\r\n" +
                    separator +
                    header +
                    separator);

                // CSV 檔案 (用於匯入 Excel)
                File.WriteAllText(_csvLogFilePath,
                    "函數名稱,調用次數,平均執行時間(ms),最短時間(ms),最長時間(ms),總時間(ms),記錄日期時間\r\n");

                // 設置自動儲存
                if (autoSaveInterval > 0)
                {
                    _autoSaveTimer = new System.Threading.Timer(_ => FlushToDisk(), null, autoSaveInterval, autoSaveInterval);
                }

                _isInitialized = true;
            }

            /// <summary>
            /// 開始計時特定函數
            /// </summary>
            /// <param name="functionName">函數名稱</param>
            public static void StartMeasure(string functionName)
            {
                if (!_isInitialized)
                    Initialize();

                int threadId = Thread.CurrentThread.ManagedThreadId;
                var threadTimers = _activeTimers.GetOrAdd(threadId, _ => new ConcurrentDictionary<string, Stopwatch>());

                var sw = new Stopwatch();
                sw.Start();
                threadTimers[functionName] = sw;
            }

            /// <summary>
            /// 停止計時並記錄執行時間
            /// </summary>
            /// <param name="functionName">函數名稱</param>
            /// <returns>執行時間(毫秒)</returns>
            public static double StopMeasure(string functionName)
            {
                int threadId = Thread.CurrentThread.ManagedThreadId;

                if (!_activeTimers.TryGetValue(threadId, out var threadTimers) ||
                    !threadTimers.TryRemove(functionName, out var sw))
                {
                    Debug.WriteLine($"警告: 嘗試停止未啟動的計時器 '{functionName}'");
                    return 0;
                }

                sw.Stop();
                double elapsedMs = sw.Elapsed.TotalMilliseconds;

                // 記錄時間
                _timings.GetOrAdd(functionName, _ => new ConcurrentBag<double>())
                       .Add(elapsedMs);

                // 直接輸出到 Debug 視窗，方便即時查看
                Debug.WriteLine($"性能: {functionName} 執行時間 = {elapsedMs:F2} ms");

                return elapsedMs;
            }

            /// <summary>
            /// 停止計時並記錄執行時間，返回 TimeSpan (兼容舊版本)
            /// </summary>
            public static TimeSpan StopMeasureTimeSpan(string functionName)
            {
                double ms = StopMeasure(functionName);
                return TimeSpan.FromMilliseconds(ms);
            }

            /// <summary>
            /// 簡單包裝一個操作並測量其執行時間
            /// </summary>
            /// <param name="functionName">函數名稱</param>
            /// <param name="action">要執行的操作</param>
            /// <returns>執行時間(毫秒)</returns>
            public static double Measure(string functionName, Action action)
            {
                StartMeasure(functionName);
                double elapsed = 0;
                try
                {
                    action();
                }
                finally
                {
                    elapsed = StopMeasure(functionName);
                }
                return elapsed;
            }

            /// <summary>
            /// 將當前記錄的所有計時數據寫入磁盤
            /// </summary>
            public static List<double> GetTimings(string functionName)
            {
                if (_timings.TryGetValue(functionName, out var timingsBag))
                {
                    return timingsBag.ToList();
                }
                return null;
            }
            public static void FlushToDisk()
            {
                if (!_isInitialized) return;

                try
                {
                    StringBuilder sbText = new StringBuilder();  // 文字檔 (美觀格式)
                    StringBuilder sbCSV = new StringBuilder();   // CSV 檔案 (資料分析用)

                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    // 文字報表的表頭
                    string header = string.Format(
                        "| {0,-25} | {1,-10} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-20} |",
                        "函數名稱", "調用次數", "平均時間(ms)", "最短時間(ms)", "最長時間(ms)", "總時間(ms)", "記錄日期時間"
                    );
                    string separator = new string('-', header.Length);

                    sbText.AppendLine(separator);
                    sbText.AppendLine($"性能報告更新時間: {timestamp}");
                    sbText.AppendLine(separator);
                    sbText.AppendLine(header);
                    sbText.AppendLine(separator);

                    // 從 _timings 複製資料並按函數名稱排序
                    var sortedTimings = _timings.OrderBy(kvp => kvp.Key).ToList();

                    foreach (var kvp in sortedTimings)
                    {
                        string functionName = kvp.Key;
                        var timings = kvp.Value.ToArray(); // 複製當前值避免同時修改

                        if (timings.Length == 0)
                            continue;

                        double avgTime = timings.Average();
                        double minTime = timings.Min();
                        double maxTime = timings.Max();
                        double totalTime = timings.Sum();
                        int count = timings.Length;

                        // 文字報表格式 (整齊的表格)
                        sbText.AppendLine(string.Format(
                            "| {0,-25} | {1,-10} | {2,-15:F2} | {3,-15:F2} | {4,-15:F2} | {5,-15:F2} | {6,-20} |",
                            functionName.Length > 25 ? functionName.Substring(0, 22) + "..." : functionName,
                            count,
                            avgTime,
                            minTime,
                            maxTime,
                            totalTime,
                            timestamp
                        ));

                        // CSV 格式 (函數名,調用次數,平均時間,最短時間,最長時間,總時間,時間戳)
                        sbCSV.AppendLine($"{functionName},{count},{avgTime:F2},{minTime:F2},{maxTime:F2},{totalTime:F2},{timestamp}");
                    }

                    sbText.AppendLine(separator);

                    lock (_fileLock)
                    {
                        // 寫入文字報表
                        File.AppendAllText(_logFilePath, sbText.ToString() + Environment.NewLine);

                        // 寫入 CSV 檔案
                        File.AppendAllText(_csvLogFilePath, sbCSV.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"寫入性能報告失敗: {ex.Message}");
                }
                _timings.Clear();
            }

            /// <summary>
            /// 清除特定函數的計時記錄
            /// </summary>
            public static void ClearMeasurements(string functionName)
            {
                _timings.TryRemove(functionName, out _);
            }

            /// <summary>
            /// 清除所有函數的計時記錄
            /// </summary>
            public static void ClearAllMeasurements()
            {
                _timings.Clear();
            }

            /// <summary>
            /// 釋放資源
            /// </summary>
            public static void Shutdown()
            {
                _autoSaveTimer?.Dispose();
                FlushToDisk();
            }
        }


        private void button17_Click(object sender, EventArgs e)
        {
            #region 檢查狀態
            // **新增：檢查狀態，只有在更新後需要復歸狀態才能執行**
            /*
            if (app.currentState != app.SystemState.UpdatedNeedReset)
            {
                string message = app.currentState == app.SystemState.Running
                    ? "請先停止檢測。"
                    : app.currentState == app.SystemState.StoppedNeedUpdate
                    ? "請先執行更新計數。"
                    : app.currentState == app.SystemState.Stopped
                    ? "系統已在正常狀態，無需復歸。"
                    : "系統狀態不允許此操作。";

                CustomMessageBox.Show(
                    message,
                    "操作順序提醒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
                );
                return;
            }
            */
            #endregion
            // 防連點檢查 (必須在函數最開頭)
            lock (_button17Lock)
            {
                if (_isButton17Processing)
                {
                    return; // 直接忽略重複點擊
                }
                _isButton17Processing = true;
            }
            try
            {
                button17.Enabled = false; // 原有的UI視覺回饋
                if (!app.offline)
                {
                    if (!PLC_CheckM(21))
                    {
                        // 復歸開始前先鎖定UI
                        uiLock(false);

                        app.status = true;
                        PLC_SetM(3, true);
                        ShowCountdownDialog(15); //顯示倒數計時對話框
                        PLC_SetM(3, false);
                        ResultManager.ResetOkCounters();
                        #region 重置第一顆樣品狀態
                        lock (_initLock)
                        {
                            _systemInitialized = false;
                            Log.Information("異常復歸完成，重置第一顆樣品狀態，下次啟動將等待 D100 變化");
                        }
                        #endregion
                        var d803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803);
                        var d805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805);
                        var d807 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 807);
                        if (d803 == 0)
                        {
                            d807 = d807 - d805;
                        }
                        else if (d805 == 0)
                        {
                            d807 = d807 - d803;
                        }
                        PLC_SetD(803, 0);
                        PLC_SetD(805, 0);
                        PLC_SetD(807, d807);

                        _prevD803 = -1;
                        _prevD805 = -1;
                        _prevD807 = -1;

                        updateLabel();

                        //PLC_SetM(401, true); 改在停止檢測做 (button2)
                        PLC_SetM(401, true); //多做一次不是壞事

                        //軟體計數（sampleID）與PLC計數同步
                        PLC_SetD(97, app.counter["stop3"] % 40); //直接取餘數就好，因為PLC那邊會先+1
                        PLC_SetD(98, app.counter["stop3"] % 15);
                        PLC_SetD(99, app.counter["stop3"] % 15);
                        // === 復歸基準監控：寫LOG與初始化監控狀態 ===
                        try
                        {
                            // 讀回PLC基準值
                            int d97 = PLC_CheckD(97);
                            int d98 = PLC_CheckD(98);
                            int d99 = PLC_CheckD(99);
                            int d801 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 801); // NG
                            d803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803); // OK1
                            d805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805); // OK2
                            d807 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 807); // OK總
                            int d809 = PLC_CheckD(809);                                  // NULL

                            bool m6 = PLC_CheckM(6), m7 = PLC_CheckM(7), m8 = PLC_CheckM(8), m9 = PLC_CheckM(9);
                            bool emer = PLC_CheckM(21);

                            int sampleIdBase = app.counter["stop3"];
                            int stop0 = app.counter["stop0"];
                            int stop1 = app.counter["stop1"];
                            int stop2 = app.counter["stop2"];
                            int stop3 = app.counter["stop3"];

                            // 建立期望映射（下顆預期的M）
                            int idx15 = ((sampleIdBase % 15) + 15) % 15;
                            int idx40 = ((sampleIdBase % 40) + 40) % 40;
                            int expectedOK1 = 3000 + idx15;
                            int expectedOK2 = 4500 + idx15;
                            int expectedNG = 1500 + idx40;

                            // 重置監控基準，避免復歸後第一輪誤判/連跳
                            app.lastD98Accepted = d98;   // 站4 D98 比對基準
                            _prevD803 = d803;            // OK1 基準
                            _prevD805 = d805;            // OK2 基準
                            _prevD807 = d807;            // OK 總基準
                            _lastProductiveTime = null;   // 重置為null，等待第一個樣品推料

                            // 清空待配對佇列，避免殘留樣品ID誤配
                            while (app.pendingOK1.TryDequeue(out _)) { }
                            while (app.pendingOK2.TryDequeue(out _)) { }

                            // 總覽基準
                            Log.Information($"[Reset] 基準建立: SAMPLE_ID={sampleIdBase}, stop0/1/2/3={stop0}/{stop1}/{stop2}/{stop3} | D97={d97}, D98={d98}, D99={d99} | D803(OK1)={d803}, D805(OK2)={d805}, D807(OK)={d807}, D801(NG)={d801}, D809(NULL)={d809} | Gate M6/7/8/9={m6}/{m7}/{m8}/{m9}, EMER(M21)={emer}");

                            // 映射期望（對位基準）
                            Log.Information($"[Reset] 映射期望: OK1=>M{expectedOK1} (由D98={d98}), OK2=>M{expectedOK2} (由D99={d99}), NG=>M{expectedNG} (由D97={d97}), pack={app.pack}");
                        }
                        catch (Exception exLog)
                        {
                            Log.Warning($"[Reset] 基準監控記錄失敗: {exLog.Message}");
                        }
                        // === 復歸基準監控：結束 ===
                        PLC_SetM(0, false);
                        PLC_SetM(1, false);
                        PLC_SetM(2, false);
                        PLC_SetM(5, false);
                        PLC_SetY(20, false); //NG推料
                        PLC_SetY(21, false); //OK1推料
                        PLC_SetY(22, false); //OK2推料
                        PLC_SetM(333, false); //轉盤
                        #region 取像過程訊號清空 //M3 在PLC做一樣的事
                        for (int i = 150; i <= 159; i++) //NG出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 190; i <= 199; i++) //NG出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 335; i <= 354; i++) //NG出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 160; i <= 169; i++) //OK1出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 201; i <= 205; i++) //OK1出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 170; i <= 184; i++) //OK2出料
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 276; i <= 298; i++) //NULL佇列
                        {
                            if (i == 280 || i == 285 || i == 290) continue;
                            PLC_SetM(i, false);
                        }
                        for (int i = 360; i <= 379; i++) //NULL佇列
                        {
                            PLC_SetM(i, false);
                        }
                        /*
                        for (int i = 210; i <= 249; i++) //站4到NG佇列
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 301; i <= 315; i++) //站4到OK1佇列
                        {
                            PLC_SetM(i, false);
                        }
                        for (int i = 260; i <= 274; i++) //站4到OK2佇列
                        {
                            PLC_SetM(i, false);
                        }
                        */
                        #endregion
                        #region 閘門初始化
                        button26.Text = "ON";
                        button26.BackColor = Color.FromArgb(128, 255, 128); // 綠色
                        if (!app.offline)
                        {
                            PLC_SetM(6, true);
                        }
                        button12.Text = "OFF";
                        button12.BackColor = Color.FromArgb(255, 128, 128);
                        if (!app.offline)
                        {
                            PLC_SetM(9, true);
                        }
                        #endregion
                        ClearImageQueues();
                        app.reseting = true;
                        app.hasPerformedRecovery = true;

                        // 新增：顯示復歸成功通知
                        ShowRecoverySuccessNotification();

                        app.plc_stop = false; // 復歸完強制認為plc不可能在暫停狀態
                        app.status = false;
                        // 復歸完成後解鎖UI
                        uiLock(true);
                    }
                    else
                    {
                        alert alert = new alert();
                        alert.ShowDialog();
                    }
                }
                // **新增：狀態變更為可以開始檢測**
                app.currentState = app.SystemState.Stopped;
                UpdateButtonStates();
            }
            finally
            {
                // 確保旗標和按鈕狀態重置 (必須在 finally 區塊)
                button17.Enabled = true;

                lock (_button17Lock)
                {
                    _isButton17Processing = false;
                }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox6.Items.Clear();
            using (var db = new MydbDB())
            {
                var q =
                    from c in db.DefectCounts
                    where c.Type == comboBox3.Text
                    orderby c.LotId
                    select c;
                if (q.Count() > 0)
                {
                    foreach (var c in q)
                    {
                        if (!comboBox6.Items.Contains(c.LotId))
                        {
                            comboBox6.Items.Add(c.LotId);
                        }
                    }
                }
            }
            if (comboBox6.Items.Count > 0)
            {
                comboBox6.SelectedIndex = 0;
            }
        }


        private void button23_Click(object sender, EventArgs e)
        {
            save2excel(comboBox3.Text, comboBox6.Text, DateTime.FromOADate(dateTimePicker1.Value.AddMinutes(-1).ToOADate()), DateTime.FromOADate(dateTimePicker2.Value.AddMinutes(-1).ToOADate()), true);
        }

        private void button26_Click(object sender, EventArgs e) 
        {
            if (button26.Text == "ON") //關OK1
            {
                button26.Text = "OFF";
                button26.BackColor = Color.FromArgb(255, 128, 128);
                PLC_SetM(7, true);
            }
            else if (button26.Text == "OFF") //開OK1
            {
                button26.Text = "ON";
                button26.BackColor = Color.FromArgb(128, 255, 128);
                PLC_SetM(6, true);
            }
        }
        private void button12_Click(object sender, EventArgs e)
        {
            if (button12.Text == "ON") // 關OK2
            {
                button12.Text = "OFF";
                button12.BackColor = Color.FromArgb(255, 128, 128);
                PLC_SetM(9, true);
            }
            else if (button12.Text == "OFF") //開OK2
            {
                button12.Text = "ON";
                button12.BackColor = Color.FromArgb(128, 255, 128);
                PLC_SetM(8, true);
            }
        }

        

        private void button27_Click(object sender, EventArgs e) // 開門警示 待改 (ON代表 開門會急停?)
        {
            if (!app.offline)
            {
                if (app.param["dooralert"] == "true")
                {
                    button27.Text = "OFF";
                    button27.BackColor = button27.BackColor = Color.FromArgb(255, 128, 128);
                    using (var db = new MydbDB())
                    {
                        db.Parameters.Where(p => p.Name == "dooralert").Set(p => p.Value, "false").Update();
                    }
                    PLC_SetM(22, false);
                }
                else
                {
                    button27.Text = "ON";
                    button27.BackColor = button27.BackColor = Color.Lime;
                    using (var db = new MydbDB())
                    {
                        db.Parameters.Where(p => p.Name == "dooralert").Set(p => p.Value, "true").Update();
                    }
                    PLC_SetM(22, true);
                }
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            app.status = true;
            //PLC_SetM(0, true);
            PLC_SetM(3, true);
        }

        private void button33_Click(object sender, EventArgs e)
        {
            app.status = false;
            //PLC_SetM(0, true);
            PLC_SetM(3, false);
        }
        private void button34_Click_1(object sender, EventArgs e)
        {
            PLC_SetM(30, true);
        }

        private void button35_Click(object sender, EventArgs e)
        {
            try
            {
                // 初始化性能分析器，設定檔案路徑和自動保存間隔
                PerformanceProfiler.Initialize(@"C:\Workspace\anomalib\datasets\MVTec\newbush\performance_log.txt", 0);

                // 清除任何舊有的紀錄
                PerformanceProfiler.ClearAllMeasurements();

                // 指定要讀取的目錄
                string folderPath = @"C:\Users\User\Desktop\4";

                // 檢查目錄是否存在
                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show($"指定的目錄不存在：{folderPath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 獲取所有PNG檔案
                string[] pngFiles = Directory.GetFiles(folderPath, "*.jpg", SearchOption.TopDirectoryOnly);
                if (pngFiles.Length == 0)
                {
                    MessageBox.Show($"在目錄中找不到PNG檔案：{folderPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 創建目錄用於儲存處理後的影像
                string outputFolder = Path.Combine(folderPath, "processed");
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // 處理進度提示
                int totalFiles = pngFiles.Length;
                int processedFiles = 0;

                // 創建結果記錄檔
                StringBuilder resultLog = new StringBuilder();
                resultLog.AppendLine($"===== 圓形ROI檢測效能測試 - 開始時間: {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} =====");
                resultLog.AppendLine($"測試影像目錄: {folderPath}");
                resultLog.AppendLine($"測試影像數量: {totalFiles}");
                resultLog.AppendLine("---------------------------------------------------");

                // 新增: 創建數據集合來保存檢測到的圓心和半徑資訊
                List<Point> outerCenters = new List<Point>();
                List<Point> innerCenters = new List<Point>();
                List<int> outerRadii = new List<int>();
                List<int> innerRadii = new List<int>();
                List<double> centerDistances = new List<double>();

                // 新增: 創建CSV來記錄圓心和半徑信息
                StringBuilder circleDataCsv = new StringBuilder();
                circleDataCsv.AppendLine("檔案名稱,外圈圓心X,外圈圓心Y,外圈半徑,內圈圓心X,內圈圓心Y,內圈半徑," +
                    "外圈X偏差,外圈Y偏差,外圈半徑偏差,內圈X偏差,內圈Y偏差,內圈半徑偏差,圓心間距離");

                // 讀取預設值
                int stop = 4; // 測試站點
                
                int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
                int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
                int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);
                int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);
                
                // 處理每一個檔案
                foreach (string filePath in pngFiles)
                {
                    // 讀取影像
                    Mat inputImage = Cv2.ImRead(filePath);
                    if (inputImage.Empty())
                    {
                        resultLog.AppendLine($"無法讀取影像: {Path.GetFileName(filePath)}");
                        continue;
                    }

                    // 設定stop和count參數 (這裡假設我們測試站點1)
                    //int stop = 1;
                    int count = processedFiles;

                    // 測試處理時間 (整個方法計時)
                    PerformanceProfiler.StartMeasure($"DetectAndExtractROI_Total");
                    CircleDetectionResult detectionResult = testDetectAndExtractROI(inputImage, stop, count);
                    TimeSpan executionTime = PerformanceProfiler.StopMeasureTimeSpan($"DetectAndExtractROI_Total");

                    // 如果檢測成功，收集數據
                    
                    if (detectionResult.DetectedOuterRadius > 0)
                    {
                        // 收集數據
                        outerCenters.Add(detectionResult.DetectedOuterCenter);
                        innerCenters.Add(detectionResult.DetectedInnerCenter);
                        outerRadii.Add(detectionResult.DetectedOuterRadius);
                        innerRadii.Add(detectionResult.DetectedInnerRadius);
                        centerDistances.Add(detectionResult.CenterDistance);

                        // 計算與預設值的偏差
                        int outerCenterXDiff = detectionResult.DetectedOuterCenter.X - knownOuterCenterX;
                        int outerCenterYDiff = detectionResult.DetectedOuterCenter.Y - knownOuterCenterY;
                        int innerCenterXDiff = detectionResult.DetectedInnerCenter.X - knownInnerCenterX;
                        int innerCenterYDiff = detectionResult.DetectedInnerCenter.Y - knownInnerCenterY;
                        int outerRadiusDiff = detectionResult.DetectedOuterRadius - knownOuterRadius;
                        int innerRadiusDiff = detectionResult.DetectedInnerRadius - knownInnerRadius;

                        // 記錄到CSV文件
                        circleDataCsv.AppendLine($"{Path.GetFileName(filePath)}," +
                                              $"{detectionResult.DetectedOuterCenter.X},{detectionResult.DetectedOuterCenter.Y}," +
                                              $"{detectionResult.DetectedOuterRadius}," +
                                              $"{detectionResult.DetectedInnerCenter.X},{detectionResult.DetectedInnerCenter.Y}," +
                                              $"{detectionResult.DetectedInnerRadius}," +
                                              $"{outerCenterXDiff},{outerCenterYDiff},{outerRadiusDiff}," +
                                              $"{innerCenterXDiff},{innerCenterYDiff},{innerRadiusDiff}," +
                                              $"{detectionResult.CenterDistance:F2}");

                        // 繪製偵測結果到圖像上
                        Mat debugImage = inputImage.Clone();
                        Cv2.Circle(debugImage, detectionResult.DetectedOuterCenter, detectionResult.DetectedOuterRadius, Scalar.Red, 2);
                        Cv2.Circle(debugImage, detectionResult.DetectedInnerCenter, detectionResult.DetectedInnerRadius, Scalar.Blue, 2);
                        Cv2.Circle(debugImage, detectionResult.DetectedOuterCenter, 5, Scalar.Green, -1);
                        Cv2.Circle(debugImage, detectionResult.DetectedInnerCenter, 5, Scalar.Yellow, -1);

                        // 保存偵測結果圖像
                        string debugImagePath = Path.Combine(outputFolder, $"debug_{Path.GetFileName(filePath)}");
                        Cv2.ImWrite(debugImagePath, debugImage);
                        debugImage.Dispose();
                    }
                    
                    // 增加進度計數
                    processedFiles++;

                    // 顯示進度
                    if (processedFiles % 10 == 0 || processedFiles == totalFiles)
                    {
                        Console.WriteLine($"已處理: {processedFiles}/{totalFiles} 檔案");
                    }

                    // 記錄本次執行時間和偵測結果
                    resultLog.AppendLine($"檔案 {Path.GetFileName(filePath)} 處理時間: {executionTime.TotalMilliseconds:F2} ms");
                    if (detectionResult.DetectedOuterRadius > 0)
                    {
                        resultLog.AppendLine($"  外圈圓心: ({detectionResult.DetectedOuterCenter.X}, {detectionResult.DetectedOuterCenter.Y}), 半徑: {detectionResult.DetectedOuterRadius}");
                        resultLog.AppendLine($"  內圈圓心: ({detectionResult.DetectedInnerCenter.X}, {detectionResult.DetectedInnerCenter.Y}), 半徑: {detectionResult.DetectedInnerRadius}");
                        resultLog.AppendLine($"  圓心間距: {detectionResult.CenterDistance:F2} 像素");
                    }
                    else
                    {
                        resultLog.AppendLine("  未檢測到有效的圓形");
                    }

                    // 釋放資源
                    inputImage.Dispose();
                    if (detectionResult.ProcessedImage != null)
                        detectionResult.ProcessedImage.Dispose();
                }

                // 寫入圓形數據CSV
                string circleDataFilePath = Path.Combine(folderPath, "circle_detection_data.csv");
                File.WriteAllText(circleDataFilePath, circleDataCsv.ToString(), Encoding.UTF8);

                // 統計分析圓形數據
                resultLog.AppendLine("\n===== 圓心和半徑統計分析 =====");

                if (outerCenters.Count > 0)
                {
                    // 外圓圓心X座標分析
                    double outerCenterXMean = outerCenters.Average(p => p.X);
                    double outerCenterXStdDev = Math.Sqrt(outerCenters.Average(p => Math.Pow(p.X - outerCenterXMean, 2)));
                    var outerCenterXValues = outerCenters.Select(p => p.X).OrderBy(x => x).ToList();
                    int outerCenterXMedian = outerCenterXValues.Count > 0 ?
                        outerCenterXValues[outerCenterXValues.Count / 2] : 0;

                    resultLog.AppendLine($"外圓圓心X座標: 平均值={outerCenterXMean:F2}, 中位數={outerCenterXMedian}, 標準差={outerCenterXStdDev:F2}");

                    // 外圓圓心Y座標分析
                    double outerCenterYMean = outerCenters.Average(p => p.Y);
                    double outerCenterYStdDev = Math.Sqrt(outerCenters.Average(p => Math.Pow(p.Y - outerCenterYMean, 2)));
                    var outerCenterYValues = outerCenters.Select(p => p.Y).OrderBy(y => y).ToList();
                    int outerCenterYMedian = outerCenterYValues.Count > 0 ?
                        outerCenterYValues[outerCenterYValues.Count / 2] : 0;

                    resultLog.AppendLine($"外圓圓心Y座標: 平均值={outerCenterYMean:F2}, 中位數={outerCenterYMedian}, 標準差={outerCenterYStdDev:F2}");

                    // 外圓半徑分析
                    double outerRadiusMean = outerRadii.Average();
                    double outerRadiusStdDev = Math.Sqrt(outerRadii.Average(r => Math.Pow(r - outerRadiusMean, 2)));
                    var outerRadiusValues = outerRadii.OrderBy(r => r).ToList();
                    int outerRadiusMedian = outerRadiusValues.Count > 0 ?
                        outerRadiusValues[outerRadiusValues.Count / 2] : 0;

                    resultLog.AppendLine($"外圓半徑: 平均值={outerRadiusMean:F2}, 中位數={outerRadiusMedian}, 標準差={outerRadiusStdDev:F2}");

                    // 內圓圓心X座標分析
                    double innerCenterXMean = innerCenters.Average(p => p.X);
                    double innerCenterXStdDev = Math.Sqrt(innerCenters.Average(p => Math.Pow(p.X - innerCenterXMean, 2)));
                    var innerCenterXValues = innerCenters.Select(p => p.X).OrderBy(x => x).ToList();
                    int innerCenterXMedian = innerCenterXValues.Count > 0 ?
                        innerCenterXValues[innerCenterXValues.Count / 2] : 0;

                    resultLog.AppendLine($"內圓圓心X座標: 平均值={innerCenterXMean:F2}, 中位數={innerCenterXMedian}, 標準差={innerCenterXStdDev:F2}");

                    // 內圓圓心Y座標分析
                    double innerCenterYMean = innerCenters.Average(p => p.Y);
                    double innerCenterYStdDev = Math.Sqrt(innerCenters.Average(p => Math.Pow(p.Y - innerCenterYMean, 2)));
                    var innerCenterYValues = innerCenters.Select(p => p.Y).OrderBy(y => y).ToList();
                    int innerCenterYMedian = innerCenterYValues.Count > 0 ?
                        innerCenterYValues[innerCenterYValues.Count / 2] : 0;

                    resultLog.AppendLine($"內圓圓心Y座標: 平均值={innerCenterYMean:F2}, 中位數={innerCenterYMedian}, 標準差={innerCenterYStdDev:F2}");

                    // 內圓半徑分析
                    double innerRadiusMean = innerRadii.Average();
                    double innerRadiusStdDev = Math.Sqrt(innerRadii.Average(r => Math.Pow(r - innerRadiusMean, 2)));
                    var innerRadiusValues = innerRadii.OrderBy(r => r).ToList();
                    int innerRadiusMedian = innerRadiusValues.Count > 0 ?
                        innerRadiusValues[innerRadiusValues.Count / 2] : 0;

                    resultLog.AppendLine($"內圓半徑: 平均值={innerRadiusMean:F2}, 中位數={innerRadiusMedian}, 標準差={innerRadiusStdDev:F2}");

                    // 圓心距離分析
                    double centerDistanceMean = centerDistances.Average();
                    double centerDistanceStdDev = Math.Sqrt(centerDistances.Average(d => Math.Pow(d - centerDistanceMean, 2)));
                    var centerDistanceValues = centerDistances.OrderBy(d => d).ToList();
                    double centerDistanceMedian = centerDistanceValues.Count > 0 ?
                        centerDistanceValues[centerDistanceValues.Count / 2] : 0;

                    resultLog.AppendLine($"圓心距離: 平均值={centerDistanceMean:F2}, 中位數={centerDistanceMedian:F2}, 標準差={centerDistanceStdDev:F2}");

                    // 建議值
                    resultLog.AppendLine("\n===== 建議的固定參數 =====");
                    resultLog.AppendLine($"建議外圓圓心: X={outerCenterXMedian}, Y={outerCenterYMedian}");
                    resultLog.AppendLine($"建議內圓圓心: X={innerCenterXMedian}, Y={innerCenterYMedian}");
                    resultLog.AppendLine($"建議外圓半徑: {outerRadiusMedian}");
                    resultLog.AppendLine($"建議內圓半徑: {innerRadiusMedian}");

                    // 創建CSV統計文件
                    StringBuilder statsCsv = new StringBuilder();
                    statsCsv.AppendLine("參數,平均值,中位數,標準差,最小值,最大值");
                    statsCsv.AppendLine($"外圓圓心X,{outerCenterXMean:F2},{outerCenterXMedian},{outerCenterXStdDev:F2},{outerCenterXValues.First()},{outerCenterXValues.Last()}");
                    statsCsv.AppendLine($"外圓圓心Y,{outerCenterYMean:F2},{outerCenterYMedian},{outerCenterYStdDev:F2},{outerCenterYValues.First()},{outerCenterYValues.Last()}");
                    statsCsv.AppendLine($"外圓半徑,{outerRadiusMean:F2},{outerRadiusMedian},{outerRadiusStdDev:F2},{outerRadiusValues.First()},{outerRadiusValues.Last()}");
                    statsCsv.AppendLine($"內圓圓心X,{innerCenterXMean:F2},{innerCenterXMedian},{innerCenterXStdDev:F2},{innerCenterXValues.First()},{innerCenterXValues.Last()}");
                    statsCsv.AppendLine($"內圓圓心Y,{innerCenterYMean:F2},{innerCenterYMedian},{innerCenterYStdDev:F2},{innerCenterYValues.First()},{innerCenterYValues.Last()}");
                    statsCsv.AppendLine($"內圓半徑,{innerRadiusMean:F2},{innerRadiusMedian},{innerRadiusStdDev:F2},{innerRadiusValues.First()},{innerRadiusValues.Last()}");
                    statsCsv.AppendLine($"圓心距離,{centerDistanceMean:F2},{centerDistanceMedian:F2},{centerDistanceStdDev:F2},{centerDistanceValues.First():F2},{centerDistanceValues.Last():F2}");

                    // 保存統計CSV
                    string statsCsvPath = Path.Combine(folderPath, "circle_statistics.csv");
                    File.WriteAllText(statsCsvPath, statsCsv.ToString(), Encoding.UTF8);

                    // 生成直方圖
                    try
                    {
                        // 創建保存直方圖的目錄
                        string histogramPath = Path.Combine(folderPath, "histograms");
                        if (!Directory.Exists(histogramPath))
                        {
                            Directory.CreateDirectory(histogramPath);
                        }

                        // 創建直方圖的函數
                        Action<List<int>, string, string, string> createHistogramForInt = (data, title, filename, xAxisTitle) =>
                        {
                            Chart chart = new Chart();
                            chart.Size = new System.Drawing.Size(800, 600);

                            // 設置圖表區域
                            ChartArea chartArea = new ChartArea();
                            chartArea.AxisX.Title = xAxisTitle;
                            chartArea.AxisY.Title = "頻率";
                            chartArea.AxisX.MajorGrid.LineColor = System.Drawing.Color.LightGray;
                            chartArea.AxisY.MajorGrid.LineColor = System.Drawing.Color.LightGray;
                            chart.ChartAreas.Add(chartArea);

                            // 添加數據系列
                            Series series = new Series();
                            series.ChartType = SeriesChartType.Column;
                            series.Name = title;

                            // 計算直方圖的柱子數量 (Sturges' formula)
                            int numBins = (int)(Math.Log(data.Count, 2) + 1);
                            numBins = Math.Max(5, Math.Min(numBins, 20)); // 確保柱子數量在合理範圍

                            // 計算最小值、最大值和間距
                            int min = data.Min();
                            int max = data.Max();
                            double binWidth = (max - min) / (double)numBins;

                            // 填充直方圖數據
                            int[] counts = new int[numBins];
                            foreach (int value in data)
                            {
                                int binIndex = Math.Min((int)Math.Floor((value - min) / binWidth), numBins - 1);
                                counts[binIndex]++;
                            }

                            // 添加數據點
                            for (int i = 0; i < numBins; i++)
                            {
                                double binStart = min + i * binWidth;
                                double binCenter = binStart + binWidth / 2;
                                series.Points.AddXY(binCenter, counts[i]);
                            }

                            chart.Series.Add(series);

                            // 添加標題
                            Title chartTitle = new Title();
                            chartTitle.Text = title;
                            chartTitle.Font = new System.Drawing.Font("微軟正黑體", 14, System.Drawing.FontStyle.Bold);
                            chart.Titles.Add(chartTitle);

                            // 保存圖表為圖片
                            string imagePath = Path.Combine(histogramPath, filename + ".png");
                            chart.SaveImage(imagePath, ChartImageFormat.Png);
                        };

                        Action<List<double>, string, string, string> createHistogramForDouble = (data, title, filename, xAxisTitle) =>
                        {
                            Chart chart = new Chart();
                            chart.Size = new System.Drawing.Size(800, 600);

                            ChartArea chartArea = new ChartArea();
                            chartArea.AxisX.Title = xAxisTitle;
                            chartArea.AxisY.Title = "頻率";
                            chartArea.AxisX.MajorGrid.LineColor = System.Drawing.Color.LightGray;
                            chartArea.AxisY.MajorGrid.LineColor = System.Drawing.Color.LightGray;
                            chart.ChartAreas.Add(chartArea);

                            Series series = new Series();
                            series.ChartType = SeriesChartType.Column;
                            series.Name = title;

                            int numBins = (int)(Math.Log(data.Count, 2) + 1);
                            numBins = Math.Max(5, Math.Min(numBins, 20));

                            double min = data.Min();
                            double max = data.Max();
                            double binWidth = (max - min) / numBins;

                            int[] counts = new int[numBins];
                            foreach (double value in data)
                            {
                                int binIndex = Math.Min((int)Math.Floor((value - min) / binWidth), numBins - 1);
                                counts[binIndex]++;
                            }

                            for (int i = 0; i < numBins; i++)
                            {
                                double binStart = min + i * binWidth;
                                double binCenter = binStart + binWidth / 2;
                                series.Points.AddXY(binCenter, counts[i]);
                            }

                            chart.Series.Add(series);

                            Title chartTitle = new Title();
                            chartTitle.Text = title;
                            chartTitle.Font = new System.Drawing.Font("微軟正黑體", 14, System.Drawing.FontStyle.Bold);
                            chart.Titles.Add(chartTitle);

                            string imagePath = Path.Combine(histogramPath, filename + ".png");
                            chart.SaveImage(imagePath, ChartImageFormat.Png);
                        };

                        // 生成各項指標的直方圖
                        createHistogramForInt(outerCenterXValues, "外圓圓心X座標分佈", "outer_center_x_hist", "X座標");
                        createHistogramForInt(outerCenterYValues, "外圓圓心Y座標分佈", "outer_center_y_hist", "Y座標");
                        createHistogramForInt(outerRadiusValues, "外圓半徑分佈", "outer_radius_hist", "半徑");
                        createHistogramForInt(innerCenterXValues, "內圓圓心X座標分佈", "inner_center_x_hist", "X座標");
                        createHistogramForInt(innerCenterYValues, "內圓圓心Y座標分佈", "inner_center_y_hist", "Y座標");
                        createHistogramForInt(innerRadiusValues, "內圓半徑分佈", "inner_radius_hist", "半徑");
                        createHistogramForDouble(centerDistanceValues, "內外圓圓心距離分佈", "center_distance_hist", "距離");

                        resultLog.AppendLine($"\n直方圖已生成至: {histogramPath}");
                    }
                    catch (Exception ex)
                    {
                        resultLog.AppendLine($"\n錯誤: 生成直方圖時發生錯誤: {ex.Message}");
                        MessageBox.Show($"生成直方圖時發生錯誤: {ex.Message}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                // 寫入結果檔案
                string resultFilePath = Path.Combine(folderPath, "roi_detection_test_results.txt");
                File.WriteAllText(resultFilePath, resultLog.ToString());

                // 將性能分析結果寫入磁碟
                PerformanceProfiler.FlushToDisk();

                // 顯示完成訊息
                MessageBox.Show($"測試完成！共處理 {processedFiles} 個檔案。\n" +
                               $"圓形檢測數據已儲存至:\n{circleDataFilePath}\n" +
                               $"統計分析結果已儲存至:\n{resultFilePath}",
                               "測試完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"測試過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. 讓使用者選取來源資料夾 - 使用 CommonOpenFileDialog
                string folderPath = "";
                using (var dialog = new CommonOpenFileDialog())
                {
                    dialog.IsFolderPicker = true;
                    dialog.Title = "請選擇包含圖片的資料夾";
                    if (Directory.Exists(@"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release"))
                        dialog.InitialDirectory = @"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release\image\2025-05\0522(done)";

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        folderPath = dialog.FileName;
                    }
                    else
                    {
                        MessageBox.Show("未選取資料夾，操作取消。");
                        return;
                    }
                }

                // 2. 取得父資料夾路徑，用於建立ROI資料夾
                string parentFolderPath = Directory.GetParent(folderPath).FullName;

                // 3. 預先建立各站點的ROI資料夾 (ROI_1, ROI_2, ROI_3, ROI_4)
                Dictionary<int, string> roiFolderPaths = new Dictionary<int, string>();
                for (int station = 1; station <= 4; station++)
                {
                    string roiFolderPath = Path.Combine(parentFolderPath, $"ROI_{station}");
                    roiFolderPaths[station] = roiFolderPath;

                    // 使用既有的 FolderJob 方法建立資料夾
                    FolderJob(roiFolderPath, false, true);
                }

                // 4. 檢查來源目錄是否存在
                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show($"指定的目錄不存在：{folderPath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 5. 獲取所有jpg檔案
                string[] jpgFiles = Directory.GetFiles(folderPath, "*.jpg", SearchOption.TopDirectoryOnly);
                if (jpgFiles.Length == 0)
                {
                    MessageBox.Show($"在目錄中找不到JPG檔案：{folderPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // === 進度提示視窗 ===
                int processed = 0;
                using (var progressForm = new Form())
                using (var progressBar = new ProgressBar())
                using (var label = new Label())
                {
                    progressForm.Text = "處理進度";
                    progressForm.Width = 400;
                    progressForm.Height = 120;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.StartPosition = FormStartPosition.CenterParent;
                    progressForm.ControlBox = false;
                    progressForm.TopMost = true;

                    progressBar.Minimum = 0;
                    progressBar.Maximum = jpgFiles.Length;
                    progressBar.Dock = DockStyle.Top;
                    progressBar.Height = 30;

                    label.Dock = DockStyle.Fill;
                    label.TextAlign = ContentAlignment.MiddleCenter;
                    label.Text = "開始處理...";

                    progressForm.Controls.Add(label);
                    progressForm.Controls.Add(progressBar);

                    progressForm.Show(this);
                    progressForm.BringToFront();
                    Application.DoEvents();

                    foreach (string filePath in jpgFiles)
                    {
                        string ext = Path.GetExtension(filePath); // 取得原始副檔名（如 .jpg）
                        string originalFileName = Path.GetFileNameWithoutExtension(filePath); // 例如 0-1
                        string[] parts = originalFileName.Split('-');

                        if (parts.Length < 2)
                            continue; // 格式不符

                        // 取得序號 (a)
                        if (!int.TryParse(parts[0], out int serialNumber))
                            continue;

                        // 取得站號 (b)
                        if (!int.TryParse(parts[1], out int stationNumber))
                            continue;

                        // 檢查站號範圍 (1-4)
                        if (stationNumber < 1 || stationNumber > 4)
                            continue;

                        try
                        {
                            // 讀取影像
                            Mat inputImage = Cv2.ImRead(filePath);
                            if (inputImage.Empty())
                                continue;

                            // 直接調用 DetectAndExtractROI 方法
                            // 參數: inputImage, stop(站號), count(序號), chamfer=false(預設)
                            Mat roiResult = DetectAndExtractROI(inputImage, stationNumber, serialNumber, false);

                            if (roiResult != null && !roiResult.Empty())
                            {
                                // 根據站號選擇對應的ROI資料夾
                                string targetRoiFolderPath = roiFolderPaths[stationNumber];

                                // 輸出檔名：原始檔名 + "_roi"
                                string outputFileName = originalFileName + "_roi" + ext;
                                string outputFilePath = Path.Combine(targetRoiFolderPath, outputFileName);

                                // 保存ROI結果到對應站號的資料夾
                                Cv2.ImWrite(outputFilePath, roiResult);

                                // 釋放ROI結果圖像
                                roiResult.Dispose();
                            }

                            // 釋放原始圖像
                            inputImage.Dispose();

                            // 更新進度
                            processed++;
                            progressBar.Value = processed;
                            label.Text = $"處理中 ({processed}/{jpgFiles.Length})：{Path.GetFileName(filePath)} -> ROI_{stationNumber}";
                            Application.DoEvents();
                        }
                        catch (Exception fileEx)
                        {
                            // 記錄單個檔案處理錯誤，但繼續處理其他檔案
                            Log.Error($"處理檔案 {filePath} 時發生錯誤: {fileEx.Message}");
                            continue;
                        }
                    }
                    progressForm.Close();
                }

                // 6. 顯示完成訊息，列出所有建立的ROI資料夾
                string roiFoldersInfo = "";
                foreach (var kvp in roiFolderPaths)
                {
                    if (Directory.Exists(kvp.Value) && Directory.GetFiles(kvp.Value, "*.jpg").Length > 0)
                    {
                        int fileCount = Directory.GetFiles(kvp.Value, "*.jpg").Length;
                        roiFoldersInfo += $"ROI_{kvp.Key}: {fileCount} 個檔案\n";
                    }
                }

                MessageBox.Show($"處理完成！共處理 {processed} 個檔案。\n\n結果已儲存至：\n{roiFoldersInfo}",
                               "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"處理過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void button37_Click(object sender, EventArgs e)
        {
            if (!app.offline)
            {
                PLC_SetD(801, 0);
                PLC_SetD(807, 0);
                PLC_SetD(809, 0);
            }
            // 使用資料夾選擇對話框
            _systemInitialized = true;
            var folderDialog = new FolderBrowserDialog
            {
                Description = "請選擇包含測試圖片的資料夾",
                ShowNewFolderButton = true
            };
            if (Directory.Exists(@"C:\Workspace\bush\bin\x64\Release\image"))
                folderDialog.SelectedPath = @"C:\Workspace\bush\bin\x64\Release\image";
            // 顯示資料夾選擇對話框
            if (folderDialog.ShowDialog() != DialogResult.OK)
            {
                lbAdd("使用者取消了選擇資料夾操作", "inf", "");
                return;
            }

            // 獲取選擇的輸入資料夾路徑
            string testImagesFolder = folderDialog.SelectedPath;

            // 預先獲取所有圖片檔案路徑 
            string[] imageFiles = Directory.GetFiles(testImagesFolder, "*.png", SearchOption.TopDirectoryOnly)
                                  .Concat(Directory.GetFiles(testImagesFolder, "*.jpg", SearchOption.TopDirectoryOnly))
                                  .ToArray();

            if (imageFiles.Length == 0)
            {
                MessageBox.Show($"在目錄中找不到圖片檔案: {testImagesFolder}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 由 GitHub Copilot 產生 - 解析檔名並按樣品編號和站號排序 (a-b 格式，a是樣品編號，b是站號1-4)
            var sortedFiles = imageFiles.Select(file =>
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string[] parts = fileName.Split('-');
                int sampleId = 0;
                int stationId = 0;

                // 解析樣品編號和站數
                if (parts.Length >= 2)
                {
                    int.TryParse(parts[0], out sampleId);
                    int.TryParse(parts[1], out stationId);
                }

                return new { FilePath = file, SampleId = sampleId, StationId = stationId };
            })
            .OrderBy(item => item.SampleId)   // 由 GitHub Copilot 產生 - 先按樣品編號排序（數字排序）
            .ThenBy(item => item.StationId)   // 由 GitHub Copilot 產生 - 再按站號排序，確保同一樣品的4站按1234順序處理
            .ToArray();

            lbAdd($"找到 {sortedFiles.Length} 張測試圖片", "inf", "");

            // 由 GitHub Copilot 產生 - 檢查解析結果和圖片分布
            if (sortedFiles.All(item => item.StationId == 0))
            {
                MessageBox.Show("檔案命名格式不符預期 (應為'a-b'格式，其中b為1~4的站號)", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 由 GitHub Copilot 產生 - 統計圖片分布情況以便診斷
            var uniqueSamples = sortedFiles.Select(f => f.SampleId).Distinct().Count();
            var station1Count = sortedFiles.Count(f => f.StationId == 1);
            var station2Count = sortedFiles.Count(f => f.StationId == 2);
            var station3Count = sortedFiles.Count(f => f.StationId == 3);
            var station4Count = sortedFiles.Count(f => f.StationId == 4);

            lbAdd($"圖片分布: 樣品數={uniqueSamples}, 站1={station1Count}, 站2={station2Count}, 站3={station3Count}, 站4={station4Count}", "inf", "");

            // 由 GitHub Copilot 產生 - 顯示前10張和後10張圖片順序，用於驗證排序
            var first10 = sortedFiles.Take(10).Select(f => $"{f.SampleId}-{f.StationId}");
            var last10 = sortedFiles.Skip(Math.Max(0, sortedFiles.Length - 10)).Select(f => $"{f.SampleId}-{f.StationId}");
            lbAdd($"前10張: {string.Join(", ", first10)}", "inf", "");
            lbAdd($"後10張: {string.Join(", ", last10)}", "inf", "");

            // 由 GitHub Copilot 產生 - 檢查是否每個樣品都有完整的4個站點
            var sampleStationGroups = sortedFiles.GroupBy(f => f.SampleId)
                .Select(g => new { SampleId = g.Key, Stations = g.Select(f => f.StationId).OrderBy(s => s).ToList() })
                .ToList();

            var incompleteSamples = sampleStationGroups
                .Where(g => g.Stations.Count != 4 || !g.Stations.SequenceEqual(new[] { 1, 2, 3, 4 }))
                .ToList();

            if (incompleteSamples.Count > 0)
            {
                lbAdd($"警告: 發現 {incompleteSamples.Count} 個樣品的站點不完整！", "war", "");
                foreach (var sample in incompleteSamples.Take(10))  // 只顯示前10個
                {
                    string stationList = string.Join(",", sample.Stations);
                    lbAdd($"  樣品 {sample.SampleId}: 只有站點 [{stationList}]", "war", "");
                }
                if (incompleteSamples.Count > 10)
                {
                    lbAdd($"  ... 還有 {incompleteSamples.Count - 10} 個樣品有問題", "war", "");
                }
            }
            else
            {
                lbAdd($"檢查通過: 所有 {uniqueSamples} 個樣品都有完整的4個站點", "inf", "");
            }

            // 初始化記錄變數
            Stopwatch totalTimer = new Stopwatch();
            totalTimer.Start();

            // 設定測試模式
            app.testc = true;  // 開啟測試模式標記

            // 由 GitHub Copilot 產生 - 暫存原始狀態（包含系統狀態）
            bool originalStatus = app.status;
            bool originalSoftTriggerMode = app.SoftTriggerMode;
            app.SystemState originalSystemState = app.currentState;

            // 由 GitHub Copilot 產生 - 設定為運行狀態，並關閉軟體觸發模式
            // 必須設定 currentState 為 Running，否則停止按鈕會一直彈出警告
            app.status = true;
            app.currentState = app.SystemState.Running;
            app.SoftTriggerMode = false;
            app._reader.Set();

            // 由 GitHub Copilot 產生 - 初始化統計資料以避免匯出時發生 NullReferenceException
            ResultManager.InitializeStats();

            // 由 GitHub Copilot 產生 - 修復：只在開始時重設計數器一次
            // 重設計數器確保從0開始（只執行一次）
            app.counter["stop0"] = 0;
            app.counter["stop1"] = 0;
            app.counter["stop2"] = 0;
            app.counter["stop3"] = 0;

            int processedImageCount = 0;
            int successCount = 0;
            int skippedCount = 0;  // 由 GitHub Copilot 產生 - 記錄跳過的圖片數量
            int failedCount = 0;   // 由 GitHub Copilot 產生 - 記錄讀取失敗的圖片數量

            try
            {
                lbAdd($"開始離線測試模式，從 {testImagesFolder} 讀取圖片", "inf", "");

                // 依序處理每張圖片
                foreach (var fileInfo in sortedFiles)
                {
                    // 檢查站號是否有效
                    if (fileInfo.StationId < 1 || fileInfo.StationId > 4)
                    {
                        lbAdd($"圖片 {Path.GetFileName(fileInfo.FilePath)} 站號 {fileInfo.StationId} 不在有效範圍1-4內，已跳過", "war", "");
                        skippedCount++;  // 由 GitHub Copilot 產生
                        continue;
                    }

                    // 讀取圖片
                    Mat src = Cv2.ImRead(fileInfo.FilePath, ImreadModes.Color);
                    if (src.Empty())
                    {
                        lbAdd($"無法讀取圖片: {Path.GetFileName(fileInfo.FilePath)}", "war", "跳過此圖片");
                        // 由 GitHub Copilot 產生 - 即使 Mat 是空的也需要釋放資源
                        src.Dispose();
                        failedCount++;  // 由 GitHub Copilot 產生
                        continue;
                    }

                    // 使用 Receiver 方法模擬相機進圖 (站號從0開始，所以要減1)
                    int camID = fileInfo.StationId - 1;

                    // 由 GitHub Copilot 產生 - 移除每張圖的計數器重設（已在開頭統一重設）
                    // 計數器會在 Receiver 內部自動遞增，不應該在這裡重設

                    // 由 GitHub Copilot 產生 - 記錄當前計數器值用於除錯
                    int currentCounter = app.counter.TryGetValue("stop" + camID, out int val) ? val : 0;

                    // 使用 Receiver 方法將圖片送入處理流程
                    Receiver(camID, src, DateTime.Now);

                    // 由 GitHub Copilot 產生 - Receiver 會複製 Mat，原始 Mat 需要立即釋放以避免記憶體洩漏
                    src.Dispose();

                    processedImageCount++;
                    successCount++;

                    // 由 GitHub Copilot 產生 - 優化：動態調整延遲，確保系統穩定處理（防止 OK 進 NULL）
                    int totalQueueCount = app.Queue_Bitmap1.Count + app.Queue_Bitmap2.Count +
                                         app.Queue_Bitmap3.Count + app.Queue_Bitmap4.Count;

                    // 由 GitHub Copilot 產生 - 優化：提高基礎延遲，確保每張圖都有足夠處理時間
                    int delay = 200; // 基礎延遲提高到 200ms，防止 OK 樣品被判定為 NULL
                    
                    if (totalQueueCount > 60)
                    {
                        // 佇列超過 60 張，大幅減速
                        delay = 600;
                        if (processedImageCount % 20 == 0)
                        {
                            lbAdd($"佇列積壓嚴重 {totalQueueCount} 張，減速處理中...", "war", "");
                        }
                    }
                    else if (totalQueueCount > 40)
                    {
                        // 佇列超過 40 張，適度減速
                        delay = 400;
                    }
                    else if (totalQueueCount > 20)
                    {
                        // 佇列超過 20 張，微調減速
                        delay = 250;
                    }

                    await Task.Delay(delay);

                    // 由 GitHub Copilot 產生 - 優化：每 20 張圖更新一次進度
                    if (processedImageCount % 20 == 0)
                    {
                        lbAdd($"已處理 {processedImageCount}/{sortedFiles.Length} 張圖片 (佇列:{totalQueueCount}, 延遲:{delay}ms)", "inf", "");
                        Application.DoEvents(); // 更新UI
                    }

                    // 檢查程序是否需要停止
                    if (!app.status)
                    {
                        lbAdd("測試被使用者中斷", "war", "");
                        break;
                    }
                }

                // 由 GitHub Copilot 產生 - 優化：縮短等待間隔，加快隊列檢查頻率
                lbAdd("等待所有隊列處理完畢...", "inf", "");
                int timeout = 0;
                int maxTimeout = 4000;  // 由 GitHub Copilot 產生 - 增加到 4000 (4000 × 150ms = 10 分鐘)
                int lastQueueCount = -1;
                int noChangeCount = 0;  // 由 GitHub Copilot 產生 - 記錄佇列數量連續不變的次數

                while ((app.Queue_Bitmap1.Count + app.Queue_Bitmap2.Count +
                        app.Queue_Bitmap3.Count + app.Queue_Bitmap4.Count > 0) &&
                      timeout < maxTimeout && app.status)
                {
                    // 由 GitHub Copilot 產生 - 優化：將等待間隔從 300ms 減少到 150ms，加快反應速度
                    await Task.Delay(150);
                    timeout++;

                    int currentQueueCount = app.Queue_Bitmap1.Count + app.Queue_Bitmap2.Count +
                                           app.Queue_Bitmap3.Count + app.Queue_Bitmap4.Count;

                    // 由 GitHub Copilot 產生 - 優化：調整卡住檢測閾值（40 × 150ms = 6 秒）
                    if (currentQueueCount == lastQueueCount)
                    {
                        noChangeCount++;
                        if (noChangeCount >= 40)  // 40 × 150ms = 6 秒沒變化
                        {
                            lbAdd($"警告: 佇列處理似乎卡住了，剩餘 {currentQueueCount} 張圖片未處理", "war", "");
                            lbAdd("建議: 1) 檢查 AI 模型是否正常運作 2) 檢查日誌是否有錯誤訊息", "war", "");
                            break;
                        }
                    }
                    else
                    {
                        noChangeCount = 0;  // 重置計數器
                    }
                    lastQueueCount = currentQueueCount;

                    // 由 GitHub Copilot 產生 - 優化：降低日誌輸出頻率（每 20 次 = 3 秒更新一次）
                    if (timeout % 20 == 0)
                    {
                        int remainingImages = currentQueueCount;
                        int elapsedSeconds = timeout * 150 / 1000;
                        lbAdd($"隊列剩餘: 相機1={app.Queue_Bitmap1.Count}, 相機2={app.Queue_Bitmap2.Count}, " +
                              $"相機3={app.Queue_Bitmap3.Count}, 相機4={app.Queue_Bitmap4.Count} " +
                              $"(已等待 {elapsedSeconds} 秒)", "inf", "");
                        Application.DoEvents(); // 更新UI
                    }
                }

                // 由 GitHub Copilot 產生 - 檢查是否因為 timeout 而結束
                if (timeout >= maxTimeout)
                {
                    int remainingCount = app.Queue_Bitmap1.Count + app.Queue_Bitmap2.Count +
                                        app.Queue_Bitmap3.Count + app.Queue_Bitmap4.Count;
                    lbAdd($"警告: 等待超時（{maxTimeout * 150 / 1000} 秒），佇列中還有 {remainingCount} 張圖片未處理", "war", "");
                }

                // 由 GitHub Copilot 產生 - 完成測試，顯示詳細統計
                totalTimer.Stop();
                string resultMessage = $"離線測試完成\n" +
                                      $"總圖片數: {sortedFiles.Length}\n" +
                                      $"成功送入: {successCount} 張\n" +
                                      $"讀取失敗: {failedCount} 張\n" +
                                      $"站號錯誤: {skippedCount} 張\n" +
                                      $"耗時: {totalTimer.ElapsedMilliseconds / 1000.0:F1} 秒";

                lbAdd($"離線測試完成 - 總數:{sortedFiles.Length}, 成功:{successCount}, 失敗:{failedCount}, 跳過:{skippedCount}, 耗時:{totalTimer.ElapsedMilliseconds / 1000.0:F1}秒",
                      "inf", "");

                MessageBox.Show(resultMessage, "測試完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 由 GitHub Copilot 產生 - 離線測試後匯出檢測統計
                #region 匯出檢測統計
                DialogResult exportResult = MessageBox.Show(
                        "是否要匯出瑕疵統計資料？",
                        "匯出統計",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                if (exportResult == DialogResult.Yes)
                {
                    // 建立資料夾路徑
                    string statsFolder = System.IO.Path.Combine(Environment.CurrentDirectory, "Statistics");
                    if (!Directory.Exists(statsFolder))
                    {
                        Directory.CreateDirectory(statsFolder);
                    }

                    // 生成檔案名稱
                    string fileName = string.Format("瑕疵統計_離線測試_{0}_{1}_{2:yyyyMMdd_HHmmss}.csv",
                                                   app.produce_No,
                                                   app.LotID,
                                                   DateTime.Now);
                    string filePath = System.IO.Path.Combine(statsFolder, fileName);

                    // 匯出統計資料
                    bool success = ResultManager.ExportStatsToCsv(filePath);

                    if (success)
                    {
                        lbAdd($"瑕疵統計已成功匯出至: {filePath}", "inf", "");
                        MessageBox.Show($"瑕疵統計已成功匯出至:\n{filePath}", "匯出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        lbAdd("匯出瑕疵統計失敗", "war", "");
                        MessageBox.Show("匯出瑕疵統計失敗", "匯出失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                lbAdd("測試過程中發生錯誤", "err", ex.Message);
                MessageBox.Show($"測試過程中發生錯誤: {ex.Message}\n{ex.StackTrace}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 由 GitHub Copilot 產生 - 復原系統狀態（包含 currentState）
                app.status = originalStatus;
                app.currentState = originalSystemState;
                app.SoftTriggerMode = originalSoftTriggerMode;
                app.test = false;

                // 嘗試解放所有資源
                //GC.Collect();
            }
        }
        private void button38_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "選擇包含圖片的資料夾";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderDialog.SelectedPath;

                    // 支援的圖片格式
                    string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif" };

                    try
                    {
                        // 獲取資料夾內所有圖片檔案
                        var imageFiles = imageExtensions
                            .SelectMany(ext => Directory.GetFiles(selectedPath, ext, SearchOption.TopDirectoryOnly))
                            .ToList();

                        if (imageFiles.Count == 0)
                        {
                            MessageBox.Show("資料夾內沒有找到圖片檔案。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        int successCount = 0;
                        // 創建進度條對話框
                        using (Form progressForm = new Form())
                        {
                            progressForm.Text = "處理進度";
                            progressForm.Size = new System.Drawing.Size(400, 120);
                            progressForm.StartPosition = FormStartPosition.CenterParent;
                            progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                            progressForm.MaximizeBox = false;
                            progressForm.MinimizeBox = false;

                            ProgressBar progressBar = new ProgressBar();
                            progressBar.Location = new System.Drawing.Point(20, 20);
                            progressBar.Size = new System.Drawing.Size(340, 23);
                            progressBar.Maximum = imageFiles.Count;
                            progressBar.Value = 0;

                            Label statusLabel = new Label();
                            statusLabel.Location = new System.Drawing.Point(20, 50);
                            statusLabel.Size = new System.Drawing.Size(340, 20);
                            statusLabel.Text = "準備開始處理...";

                            progressForm.Controls.Add(progressBar);
                            progressForm.Controls.Add(statusLabel);

                            progressForm.Show();

                            // 處理每個圖片檔案
                            int processedCount = 0;

                            foreach (string imagePath in imageFiles)
                            {
                                try
                                {
                                    // 更新進度顯示
                                    statusLabel.Text = $"正在處理: {Path.GetFileName(imagePath)} ({processedCount + 1}/{imageFiles.Count})";
                                    Application.DoEvents(); // 刷新UI

                                    // 讀取圖片
                                    Mat src = Cv2.ImRead(imagePath);
                                    string nameWithoutExt = Path.GetFileNameWithoutExtension(Path.GetFileName(imagePath)); // "a-b"
                                                                                                                           // 分割並取得數字部分
                                    string[] parts = nameWithoutExt.Split('-');
                                    int stop = int.Parse(parts[1]); // 將 "b" 轉換為整數
                                    if (src.Empty())
                                    {
                                        Console.WriteLine($"無法讀取圖片: {imagePath}");
                                        processedCount++;
                                        progressBar.Value = processedCount;
                                        continue;
                                    }
                                    int contrastOffset = app.param.ContainsKey($"deepenContrast_{stop}") ?
                                        int.Parse(app.param[$"deepenContrast_{stop}"]) : 30;
                                    int brightnessOffset = app.param.ContainsKey($"deepenBrightness_{stop}") ?
                                                          int.Parse(app.param[$"deepenBrightness_{stop}"]) : 5;
                                    //轉灰階
                                    Cv2.CvtColor(src, src, ColorConversionCodes.BGR2GRAY);

                                    // 應用 ContrastAndClose 函數
                                    Mat result = ContrastAndClose(src, contrastOffset, brightnessOffset, 0);

                                    // 覆蓋原檔案
                                    Cv2.ImWrite(imagePath, result);

                                    // 釋放記憶體
                                    src.Dispose();
                                    result.Dispose();

                                    successCount++;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"處理圖片 {imagePath} 時發生錯誤: {ex.Message}");
                                }

                                processedCount++;
                                progressBar.Value = processedCount;
                                Application.DoEvents(); // 刷新UI
                            }

                            // 顯示完成狀態
                            statusLabel.Text = "處理完成!";
                            System.Threading.Thread.Sleep(1000); // 顯示完成狀態1秒
                        }

                        MessageBox.Show($"處理完成!\n成功處理: {successCount} 個檔案\n總檔案數: {imageFiles.Count}",
                                       "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"處理過程中發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void button39_Click(object sender, EventArgs e)
        {
            if (button39.Text == "設定")
            {
                if (textBox4.Text != "")
                {
                    if (int.Parse(textBox4.Text) > 100)
                    {
                        app.NGmax = int.Parse(textBox4.Text);
                        PLC_SetD(815, int.Parse(textBox4.Text)-100);
                        PLC_SetD(817, int.Parse(textBox4.Text));
                        lbAdd("設定NG籃滿料數量為:" + app.NGmax.ToString(), "inf", "");
                        button39.Text = "變更";
                        textBox4.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("請輸入有效數字。");
                    }
                }
                else
                {
                    MessageBox.Show("尚未輸入滿料數量。");
                }
            }
            else
            {
                app.NGmax = 0;
                button39.Text = "設定";
                textBox4.Enabled = true;
            }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            if (button40.Text == "設定")
            {
                if (textBox5.Text != "")
                {
                    if (int.Parse(textBox5.Text) > 100)
                    {
                        app.NULLmax = int.Parse(textBox5.Text);
                        PLC_SetD(819, int.Parse(textBox5.Text) - 100);
                        PLC_SetD(821, int.Parse(textBox5.Text));
                        lbAdd("設定NULL籃滿料數量為:" + app.NULLmax.ToString(), "inf", "");
                        button40.Text = "變更";
                        textBox5.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("請輸入有效數字。");
                    }
                }
                else
                {
                    MessageBox.Show("尚未輸入滿料數量。");
                }
            }
            else
            {
                app.NULLmax = 0;
                button40.Text = "設定";
                textBox5.Enabled = true;
            }
        }

        private void button44_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. 讓使用者選取要修正ROI的資料夾
                string folderPath = "";
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "請選擇要修正ROI的圖片資料夾";
                    if (Directory.Exists(@"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release"))
                        dialog.SelectedPath = @"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = dialog.SelectedPath;
                    }
                    else
                    {
                        MessageBox.Show("未選取資料夾，操作取消。");
                        return;
                    }
                }

                // 2. 檢查來源目錄是否存在
                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show($"指定的目錄不存在：{folderPath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. 讓使用者選擇要處理的站點
                int selectedStop = 1; // 預設值
                using (var stopSelectForm = new Form())
                {
                    stopSelectForm.Text = "選擇處理站點";
                    stopSelectForm.Size = new System.Drawing.Size(300, 150);
                    stopSelectForm.StartPosition = FormStartPosition.CenterScreen;
                    stopSelectForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    stopSelectForm.MaximizeBox = false;
                    stopSelectForm.MinimizeBox = false;

                    Label label = new Label();
                    label.Text = "請選擇要處理的站點:";
                    label.Location = new System.Drawing.Point(20, 20);
                    label.Size = new System.Drawing.Size(150, 20);

                    ComboBox comboBox = new ComboBox();
                    comboBox.Location = new System.Drawing.Point(180, 20);
                    comboBox.Size = new System.Drawing.Size(80, 20);
                    comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                    comboBox.Items.AddRange(new object[] { "站點 1", "站點 2" });
                    comboBox.SelectedIndex = 0; // 預設選擇站點1

                    Button okButton = new Button();
                    okButton.Text = "確定";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Location = new System.Drawing.Point(100, 60);
                    okButton.Size = new System.Drawing.Size(80, 30);

                    stopSelectForm.Controls.Add(label);
                    stopSelectForm.Controls.Add(comboBox);
                    stopSelectForm.Controls.Add(okButton);
                    stopSelectForm.AcceptButton = okButton;

                    if (stopSelectForm.ShowDialog() == DialogResult.OK)
                    {
                        selectedStop = comboBox.SelectedIndex + 1; // 轉換為1或2
                    }
                    else
                    {
                        return; // 用戶取消
                    }
                }

                // 4. 獲取所有圖片檔案
                string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif" };
                List<string> allImageFiles = new List<string>();

                foreach (string extension in imageExtensions)
                {
                    allImageFiles.AddRange(Directory.GetFiles(folderPath, extension, SearchOption.TopDirectoryOnly));
                }

                if (allImageFiles.Count == 0)
                {
                    MessageBox.Show($"在目錄中找不到圖片檔案：{folderPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 5. 確認對話框
                DialogResult confirmResult = MessageBox.Show(
                    $"確定要使用站點 {selectedStop} 的參數修正資料夾內的 {allImageFiles.Count} 個圖片檔案的ROI嗎？\n" +
                    "此操作會直接覆蓋原檔案，無法復原！\n\n" +
                    $"資料夾：{folderPath}\n" +
                    $"處理站點：{selectedStop}",
                    "確認修正ROI",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2
                );

                if (confirmResult != DialogResult.Yes)
                {
                    return;
                }

                // === 進度提示視窗 ===
                int processed = 0;
                int skipped = 0;
                using (var progressForm = new Form())
                using (var progressBar = new ProgressBar())
                using (var label = new Label())
                {
                    progressForm.Text = "修正ROI進度";
                    progressForm.Width = 450;
                    progressForm.Height = 120;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.StartPosition = FormStartPosition.CenterParent;
                    progressForm.ControlBox = false;
                    progressForm.TopMost = true;

                    progressBar.Minimum = 0;
                    progressBar.Maximum = allImageFiles.Count;
                    progressBar.Dock = DockStyle.Top;
                    progressBar.Height = 30;

                    label.Dock = DockStyle.Fill;
                    label.TextAlign = ContentAlignment.MiddleCenter;
                    label.Text = "開始修正ROI...";

                    progressForm.Controls.Add(label);
                    progressForm.Controls.Add(progressBar);

                    progressForm.Show(this);
                    progressForm.BringToFront();
                    Application.DoEvents();

                    foreach (string filePath in allImageFiles)
                    {
                        // 使用使用者選擇的站點參數
                        int stop = selectedStop;

                        // 讀取影像
                        Mat inputImage = Cv2.ImRead(filePath);
                        if (inputImage.Empty())
                        {
                            skipped++;
                            continue;
                        }

                        try
                        {
                            // 從資料庫讀取預設圓心和半徑 (與button36_Click相同邏輯)
                            int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
                            int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
                            int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);
                            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                            int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);

                            // 預設外圓圓心和內圓圓心
                            Point knownOuterCenter = new Point(knownOuterCenterX, knownOuterCenterY);
                            Point knownInnerCenter = new Point(knownInnerCenterX, knownInnerCenterY);

                            // 建立遮罩
                            Mat mask = new Mat(inputImage.Size(), MatType.CV_8UC1, Scalar.Black);

                            if (stop == 1 || stop == 2)
                            {
                                // 對於站點1和2，外圓白色，內圓也白色（填滿整個檢測區域）
                                Cv2.Circle(mask, knownOuterCenter, knownOuterRadius, Scalar.White, -1);
                                Cv2.Circle(mask, knownInnerCenter, knownInnerRadius, Scalar.White, -1);
                            }

                            // 應用遮罩
                            Mat roiResult = new Mat();
                            Cv2.BitwiseAnd(inputImage, inputImage, roiResult, mask);

                            // 如果是第1、2站，將內環區域填充為白色
                            if (stop == 1 || stop == 2)
                            {
                                Cv2.Circle(roiResult, knownInnerCenter, knownInnerRadius, Scalar.White, -1);
                            }

                            /*
                            // 對比度和亮度調整
                            if (stop == 1 || stop == 2)
                            {
                                // 轉為灰階進行對比度調整
                                Cv2.CvtColor(roiResult, roiResult, ColorConversionCodes.BGR2GRAY);

                                int contrastOffset = app.param.ContainsKey($"deepenContrast_{stop}") ?
                                                    int.Parse(app.param[$"deepenContrast_{stop}"]) : 30;
                                int brightnessOffset = app.param.ContainsKey($"deepenBrightness_{stop}") ?
                                                      int.Parse(app.param[$"deepenBrightness_{stop}"]) : 5;

                                roiResult = ContrastAndClose(roiResult, contrastOffset, brightnessOffset, stop);
                            }
                            */

                            // 直接覆蓋原檔案
                            Cv2.ImWrite(filePath, roiResult);

                            // 釋放資源
                            roiResult.Dispose();
                            mask.Dispose();
                            processed++;
                        }
                        catch (Exception ex)
                        {
                            // 記錄錯誤但繼續處理其他檔案
                            skipped++;
                            Console.WriteLine($"處理檔案 {filePath} 時發生錯誤: {ex.Message}");
                        }
                        finally
                        {
                            inputImage.Dispose();
                        }

                        // 更新進度
                        progressBar.Value = processed + skipped;
                        label.Text = $"修正中 ({processed + skipped}/{allImageFiles.Count})：{Path.GetFileName(filePath)}";
                        Application.DoEvents();
                    }
                    progressForm.Close();
                }

                string resultMessage = $"ROI修正完成！\n" +
                                     $"成功處理：{processed} 個檔案\n" +
                                     $"跳過：{skipped} 個檔案\n" +
                                     $"使用站點：{selectedStop}\n" +
                                     $"資料夾：{folderPath}";

                MessageBox.Show(resultMessage, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"修正ROI過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button45_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. 選擇要分析的資料夾
                string folderPath = "";
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "請選擇包含五彩鋅電鍍圖片的資料夾";
                    if (Directory.Exists(@"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release"))
                        dialog.SelectedPath = @"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = dialog.SelectedPath;
                    }
                    else
                    {
                        MessageBox.Show("未選取資料夾，操作取消。");
                        return;
                    }
                }

                // 2. 讓用戶輸入批次標籤
                string batchLabel = "";
                using (var labelForm = new Form())
                {
                    labelForm.Text = "批次標籤";
                    labelForm.Size = new System.Drawing.Size(350, 150);
                    labelForm.StartPosition = FormStartPosition.CenterScreen;
                    labelForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    labelForm.MaximizeBox = false;
                    labelForm.MinimizeBox = false;

                    Label label = new Label();
                    label.Text = "請輸入此批次的標籤 (例如: OK_batch1, NG_batch1):";
                    label.Location = new System.Drawing.Point(20, 20);
                    label.Size = new System.Drawing.Size(300, 20);

                    TextBox textBox = new TextBox();
                    textBox.Location = new System.Drawing.Point(20, 50);
                    textBox.Size = new System.Drawing.Size(280, 20);
                    textBox.Text = Path.GetFileName(folderPath); // 預設使用資料夾名稱

                    Button okButton = new Button();
                    okButton.Text = "確定";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Location = new System.Drawing.Point(120, 80);
                    okButton.Size = new System.Drawing.Size(80, 30);

                    labelForm.Controls.Add(label);
                    labelForm.Controls.Add(textBox);
                    labelForm.Controls.Add(okButton);
                    labelForm.AcceptButton = okButton;

                    if (labelForm.ShowDialog() == DialogResult.OK)
                    {
                        batchLabel = textBox.Text.Trim();
                        if (string.IsNullOrEmpty(batchLabel))
                        {
                            batchLabel = Path.GetFileName(folderPath);
                        }
                    }
                    else
                    {
                        return;
                    }
                }

                // 3. 獲取圖片文件
                string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif" };
                List<string> allImageFiles = new List<string>();

                foreach (string extension in imageExtensions)
                {
                    allImageFiles.AddRange(Directory.GetFiles(folderPath, extension, SearchOption.TopDirectoryOnly));
                }

                if (allImageFiles.Count == 0)
                {
                    MessageBox.Show($"在目錄中找不到圖片檔案：{folderPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 4. 創建輸出資料夾
                string outputDir = Path.Combine(folderPath, $"ColorAnalysis_{batchLabel}_{DateTime.Now:yyyyMMdd_HHmmss}");
                Directory.CreateDirectory(outputDir);

                // 5. 初始化色彩統計數據
                var colorStatistics = new List<ColorAnalysisResult>();

                // 6. 處理進度條
                using (var progressForm = new Form())
                using (var progressBar = new ProgressBar())
                using (var statusLabel = new Label())
                {
                    progressForm.Text = "色彩分析進度";
                    progressForm.Width = 450;
                    progressForm.Height = 120;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.StartPosition = FormStartPosition.CenterParent;
                    progressForm.ControlBox = false;
                    progressForm.TopMost = true;

                    progressBar.Minimum = 0;
                    progressBar.Maximum = allImageFiles.Count;
                    progressBar.Dock = DockStyle.Top;
                    progressBar.Height = 30;

                    statusLabel.Dock = DockStyle.Fill;
                    statusLabel.TextAlign = ContentAlignment.MiddleCenter;
                    statusLabel.Text = "開始色彩分析...";

                    progressForm.Controls.Add(statusLabel);
                    progressForm.Controls.Add(progressBar);

                    progressForm.Show(this);
                    progressForm.BringToFront();
                    Application.DoEvents();

                    // 7. 處理每張圖片
                    int processedCount = 0;
                    foreach (string imagePath in allImageFiles)
                    {
                        statusLabel.Text = $"分析中: {Path.GetFileName(imagePath)} ({processedCount + 1}/{allImageFiles.Count})";
                        Application.DoEvents();

                        try
                        {
                            Mat image = Cv2.ImRead(imagePath);
                            if (image.Empty())
                            {
                                processedCount++;
                                progressBar.Value = processedCount;
                                continue;
                            }

                            // 分析此圖片的色彩特徵
                            var result = AnalyzeImageColors(image, Path.GetFileName(imagePath), batchLabel);
                            if (result != null)
                            {
                                colorStatistics.Add(result);
                            }

                            image.Dispose();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"處理圖片 {imagePath} 時發生錯誤: {ex.Message}");
                        }

                        processedCount++;
                        progressBar.Value = processedCount;
                    }
                }

                // 8. 生成分析報告
                if (colorStatistics.Count > 0)
                {
                    GenerateColorAnalysisReport(colorStatistics, outputDir, batchLabel);
                    GenerateColorHistograms(colorStatistics, outputDir, batchLabel);
                    ExportColorStatisticsToCSV(colorStatistics, Path.Combine(outputDir, $"ColorStatistics_{batchLabel}.csv"));

                    MessageBox.Show($"色彩分析完成！\n" +
                                   $"分析圖片數量: {colorStatistics.Count}\n" +
                                   $"批次標籤: {batchLabel}\n" +
                                   $"結果已保存至: {outputDir}",
                                   "分析完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("沒有成功分析的圖片。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"色彩分析過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            // 創建檔案選擇對話框
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "選擇圖片進行白色像素占比分析";
            openFileDialog.Filter = "圖片檔案 (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|所有檔案 (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 讀取選擇的圖片
                    Mat selectedImage = Cv2.ImRead(openFileDialog.FileName);

                    if (selectedImage.Empty())
                    {
                        MessageBox.Show("無法讀取選擇的圖片！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 詢問用戶要使用哪個站點的參數進行分析
                    string stationInput = Microsoft.VisualBasic.Interaction.InputBox(
                        "請輸入站點編號 (1-4)：",
                        "選擇站點",
                        "1");

                    if (!int.TryParse(stationInput, out int station) || station < 1 || station > 4)
                    {
                        MessageBox.Show("請輸入有效的站點編號 (1-4)！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        selectedImage.Dispose();
                        return;
                    }

                    // 手動計算詳細資訊用於顯示
                    int minthresh = 0;
                    if (station == 1)
                    {
                        minthresh = 127;
                    }
                    else if (station == 2)
                    {
                        minthresh = 250;
                    }

                    Mat gray = new Mat();
                    Cv2.CvtColor(selectedImage, gray, ColorConversionCodes.BGR2GRAY);
                    Mat binary = new Mat();
                    Cv2.Threshold(gray, binary, minthresh, 255, ThresholdTypes.Binary);

                    int whitePixels = Cv2.CountNonZero(binary);
                    int totalPixels = binary.Rows * binary.Cols;
                    double ratio = (double)whitePixels / totalPixels * 100;

                    double standardRatio = 0;
                    double.TryParse(app.param[$"white_{station}"], out standardRatio);

                    // 呼叫現有的檢查方法
                    bool isValid = CheckWhitePixelRatio(selectedImage, station);

                    // 顯示結果
                    string resultMessage = $"圖片分析結果：\n\n" +
                                         $"檔案路徑：{openFileDialog.FileName}\n" +
                                         $"圖片尺寸：{selectedImage.Width} x {selectedImage.Height}\n" +
                                         $"總像素數：{totalPixels:N0}\n" +
                                         $"白色像素數：{whitePixels:N0}\n" +
                                         $"白色像素占比：{ratio:F4}%\n" +
                                         $"站點 {station} 標準值：{standardRatio:F2}%\n" +
                                         $"檢查結果：{(isValid ? "有效" : "無效")}";

                    MessageBox.Show(resultMessage, "白色像素分析結果", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 記錄到日誌
                    lbAdd($"手動分析圖片 - 站點{station}: {ratio:F4}% (標準值: {standardRatio:F2}%)",
                          isValid ? "inf" : "war",
                          openFileDialog.FileName);

                    // 清理資源
                    gray.Dispose();
                    binary.Dispose();
                    selectedImage.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"分析圖片時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lbAdd("手動分析圖片時發生錯誤", "err", ex.Message);
                }
            }
        }
        private void button47_Click(object sender, EventArgs e)
        {
            // **新增：檢查狀態，只有在停止後需要更新狀態才能執行**
            if (app.currentState != app.SystemState.StoppedNeedUpdate)
            {
                string message = app.currentState == app.SystemState.Running
                    ? "請先停止檢測後再執行更新計數。"
                    : app.currentState == app.SystemState.UpdatedNeedReset
                    ? "計數已經更新過，請執行異常復歸。"
                    : "系統狀態不允許此操作。";

                CustomMessageBox.Show(
                    message,
                    "操作順序提醒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
                );
                return;
            }

            var d803 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 803);
            var d805 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 805);
            var d807 = PLC_ModBus.GetValue_32bit(ValueUnit_32bt.D, 807);
            if (d803 == 0)
            {
                d807 = d807 - d805;
            }
            else if (d805 == 0)
            {
                d807 = d807 - d803;
            }
            PLC_SetD(803, 0);
            PLC_SetD(805, 0);
            PLC_SetD(807, d807);

            updateLabel();
            // **新增：狀態變更為更新後需要復歸**
            app.currentState = app.SystemState.UpdatedNeedReset;
            UpdateButtonStates();

            CustomMessageBox.Show(
                "計數更新完成，請執行異常復歸後方可開始下次檢測。",
                "更新完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
            );
        }
        private void button48_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. 選擇資料夾（使用 WindowsAPICodePack）
                string selectedFolder = "";
                using (var dialog = new CommonOpenFileDialog())
                {
                    dialog.IsFolderPicker = true;
                    dialog.Title = "請選擇包含 origin、OK、NG 資料夾的目錄";

                    // ✅ 修正：每次都重新設定預設路徑
                    string defaultPath = @".\image";
                    if (Directory.Exists(defaultPath))
                    {
                        dialog.InitialDirectory = Path.GetFullPath(defaultPath); // 使用絕對路徑
                    }

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        selectedFolder = dialog.FileName;
                    }
                    else
                    {
                        CustomMessageBox.Show("未選擇資料夾，操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            new Font("微軟正黑體", 14F));
                        return;
                    }
                }

                // 2. 驗證資料夾結構
                string originPath = Path.Combine(selectedFolder, "origin");
                string okPath = Path.Combine(selectedFolder, "OK");
                string ngPath = Path.Combine(selectedFolder, "NG");

                if (!Directory.Exists(originPath) || !Directory.Exists(okPath) || !Directory.Exists(ngPath))
                {
                    MessageBox.Show("選擇的資料夾必須包含 origin、OK、NG 三個子資料夾！",
                                  "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. 選擇站點
                int selectedStation = 0;
                using (var stationForm = new Form())
                {
                    stationForm.Text = "選擇站點";
                    stationForm.Size = new System.Drawing.Size(300, 150);
                    stationForm.StartPosition = FormStartPosition.CenterParent;
                    stationForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    stationForm.MaximizeBox = false;
                    stationForm.MinimizeBox = false;

                    Label lblStation = new Label
                    {
                        Text = "請選擇站點 (1-4):",
                        Location = new System.Drawing.Point(20, 20),
                        Size = new System.Drawing.Size(120, 25),
                        Font = new Font("微軟正黑體", 12F)
                    };

                    ComboBox cmbStation = new ComboBox
                    {
                        Location = new System.Drawing.Point(150, 20),
                        Size = new System.Drawing.Size(100, 25),
                        Font = new Font("微軟正黑體", 12F),
                        DropDownStyle = ComboBoxStyle.DropDownList
                    };
                    cmbStation.Items.AddRange(new object[] { "1", "2", "3", "4" });
                    cmbStation.SelectedIndex = 0;

                    Button btnOK = new Button
                    {
                        Text = "確定",
                        Location = new System.Drawing.Point(100, 70),
                        Size = new System.Drawing.Size(80, 35),
                        Font = new Font("微軟正黑體", 12F),
                        DialogResult = DialogResult.OK
                    };

                    stationForm.Controls.Add(lblStation);
                    stationForm.Controls.Add(cmbStation);
                    stationForm.Controls.Add(btnOK);
                    stationForm.AcceptButton = btnOK;

                    if (stationForm.ShowDialog() == DialogResult.OK)
                    {
                        selectedStation = int.Parse(cmbStation.SelectedItem.ToString());
                    }
                    else
                    {
                        MessageBox.Show("未選擇站點，操作取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                // 4. 建立輸出資料夾
                string trainOkPath = Path.Combine(selectedFolder, $"train_OK_{selectedStation}");
                string trainNgPath = Path.Combine(selectedFolder, $"train_NG_{selectedStation}");

                if (!Directory.Exists(trainOkPath))
                    Directory.CreateDirectory(trainOkPath);
                if (!Directory.Exists(trainNgPath))
                    Directory.CreateDirectory(trainNgPath);

                // 5. 解析檔名並提取編號
                // 檔名格式: a-b.jpg 或 a-b-瑕疵名稱-瑕疵分數.jpg
                var okNumbers = new HashSet<int>();
                var ngNumbers = new HashSet<int>();

                // 從 OK 資料夾提取編號
                foreach (var file in Directory.GetFiles(okPath, "*.jpg"))
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string[] parts = fileName.Split('-');

                    if (parts.Length >= 2)
                    {
                        if (int.TryParse(parts[0], out int number) &&
                            int.TryParse(parts[1], out int station) &&
                            station == selectedStation)
                        {
                            okNumbers.Add(number);
                        }
                    }
                }

                // 從 NG 資料夾提取編號
                foreach (var file in Directory.GetFiles(ngPath, "*.jpg"))
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string[] parts = fileName.Split('-');

                    if (parts.Length >= 2)
                    {
                        if (int.TryParse(parts[0], out int number) &&
                            int.TryParse(parts[1], out int station) &&
                            station == selectedStation)
                        {
                            ngNumbers.Add(number);
                        }
                    }
                }

                // 6. 複製檔案（使用進度條）
                int totalFiles = okNumbers.Count + ngNumbers.Count;
                int processedFiles = 0;

                using (var progressForm = new Form())
                using (var progressBar = new ProgressBar())
                using (var lblProgress = new Label())
                {
                    progressForm.Text = "複製檔案中...";
                    progressForm.Size = new System.Drawing.Size(450, 150);
                    progressForm.StartPosition = FormStartPosition.CenterParent;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.MaximizeBox = false;
                    progressForm.MinimizeBox = false;

                    progressBar.Location = new System.Drawing.Point(20, 40);
                    progressBar.Size = new System.Drawing.Size(400, 30);
                    progressBar.Maximum = totalFiles;

                    lblProgress.Location = new System.Drawing.Point(20, 10);
                    lblProgress.Size = new System.Drawing.Size(400, 25);
                    lblProgress.Font = new Font("微軟正黑體", 12F);
                    lblProgress.Text = $"準備複製 {totalFiles} 個檔案...";

                    progressForm.Controls.Add(progressBar);
                    progressForm.Controls.Add(lblProgress);
                    progressForm.Show();
                    progressForm.Refresh();

                    // 複製 OK 編號對應的 origin 圖片
                    foreach (int number in okNumbers)
                    {
                        string sourceFile = Path.Combine(originPath, $"{number}-{selectedStation}.jpg");
                        string destFile = Path.Combine(trainOkPath, $"{number}-{selectedStation}.jpg");

                        if (File.Exists(sourceFile))
                        {
                            File.Copy(sourceFile, destFile, true);
                            processedFiles++;
                            progressBar.Value = processedFiles;
                            lblProgress.Text = $"複製中... ({processedFiles}/{totalFiles}) - OK 編號: {number}";
                            progressForm.Refresh();
                        }
                    }

                    // 複製 NG 編號對應的 origin 圖片
                    foreach (int number in ngNumbers)
                    {
                        string sourceFile = Path.Combine(originPath, $"{number}-{selectedStation}.jpg");
                        string destFile = Path.Combine(trainNgPath, $"{number}-{selectedStation}.jpg");

                        if (File.Exists(sourceFile))
                        {
                            File.Copy(sourceFile, destFile, true);
                            processedFiles++;
                            progressBar.Value = processedFiles;
                            lblProgress.Text = $"複製中... ({processedFiles}/{totalFiles}) - NG 編號: {number}";
                            progressForm.Refresh();
                        }
                    }

                    lblProgress.Text = "複製完成！";
                    progressForm.Refresh();
                    System.Threading.Thread.Sleep(1000);
                }

                // 7. 顯示完成訊息
                string summaryMessage = $"分類完成！\n\n" +
                                      $"站點: {selectedStation}\n" +
                                      $"train_OK_{selectedStation}: {okNumbers.Count} 張圖片\n" +
                                      $"train_NG_{selectedStation}: {ngNumbers.Count} 張圖片\n\n" +
                                      $"儲存位置:\n{selectedFolder}";

                CustomMessageBox.Show(summaryMessage, "圖片分類完成",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information,
                                    new Font("微軟正黑體", 14F));

                // 8. 詢問是否開啟資料夾
                var openResult = CustomMessageBox.Show("是否要開啟輸出資料夾？", "完成",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                                      new Font("微軟正黑體", 14F));

                if (openResult == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", selectedFolder);
                }

                lbAdd($"圖片分類完成 - 站點{selectedStation}, OK:{okNumbers.Count}張, NG:{ngNumbers.Count}張", "inf", "");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"圖片分類時發生錯誤:\n{ex.Message}", "錯誤",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                lbAdd("圖片分類失敗", "err", ex.ToString());
            }
        }

        private void InitializeResourceDisposer()
        {
            Task.Run(async () => {
                while (!this.IsDisposed) // 表單未被釋放時持續運行
                {
                    IDisposable resource;
                    while (_disposeQueue.TryDequeue(out resource))
                    {
                        try
                        {
                            await _disposeSemaphore.WaitAsync();
                            try
                            {
                                resource.Dispose();
                            }
                            finally
                            {
                                _disposeSemaphore.Release();
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"釋放資源時出錯: {ex.Message}");
                        }
                    }
                    await Task.Delay(50); // 短暫休眠，減少CPU占用
                }
            });
        }
        private void UpdateDetectionRate()
        {
            // 記錄當前時間
            DateTime now = DateTime.Now;

            // 添加新的檢測時間到佇列
            sampleDetectionTimes.Enqueue(now);

            // 移除一分鐘以前的記錄
            DateTime oneMinuteAgo = now.AddMinutes(-1);
            while (sampleDetectionTimes.Count > 0 && sampleDetectionTimes.Peek() < oneMinuteAgo)
            {
                sampleDetectionTimes.Dequeue();
            }

            // 計算近一分鐘內的樣品數量
            recentMinuteSampleCount = sampleDetectionTimes.Count;
            // 計算每分鐘速率 (樣品/分鐘)
            double samplesPerMinute = recentMinuteSampleCount;
            // 安全地更新 UI
            BeginInvoke(new Action(() => {
                label51.Text = recentMinuteSampleCount.ToString(); // 近一分鐘總數
            }));
        }

        /// <summary>
        /// 檢查圖像中白色像素的占比是否在有效範圍內
        /// </summary>
        /// <param name="image">輸入圖像</param>
        /// <param name="stop">站點編號，用於獲取相應參數</param>
        /// <returns>如果像素占比在有效範圍內，返回true；否則返回false</returns>
        public bool CheckWhitePixelRatio(Mat image, int stop)
        {
            // 由 GitHub Copilot 產生 - 調機模式或離線測試模式下跳過白色像素檢查
            if (app.DetectMode == 1)
            {
                // 調機模式：不執行白色像素檢查，直接返回 true（視為有效圖像）
                return true;
            }

            // 由 GitHub Copilot 產生 - 離線測試模式也跳過白色像素檢查
            // 因為離線測試使用的是歷史圖片，白色像素比例可能與實時拍攝不同
            if (app.testc)
            {
                return true;
            }

            try
            {
                int minthresh = 0;
                if (stop == 1 || stop == 4)
                {
                    minthresh = 180;
                }
                else if (stop == 2)
                {
                    minthresh = 250;
                }
                else if (stop == 3)
                {
                    minthresh = 170;
                }

                minthresh = GetIntParam(app.param, $"whiteThresh_{stop}", 250);

                // 轉為灰階
                Mat gray = new Mat();
                Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                // 二值化 (200~255)
                Mat binary = new Mat();
                Cv2.Threshold(gray, binary, minthresh, 255, ThresholdTypes.Binary);

                // 計算白色像素數量
                int whitePixels = Cv2.CountNonZero(binary);

                // 計算占比 (百分比)
                double ratio = (double)whitePixels / (binary.Rows * binary.Cols) * 100;

                gray.Dispose();
                binary.Dispose();

                // 由 GitHub Copilot 產生 - 使用安全的參數讀取方式
                // 從參數中讀取標準值
                double tolerance = GetDoubleParam(app.param, $"whiteTolerance_{stop}", 3);
                double standardRatio;
                string whiteParamKey = $"white_{stop}";
                if (app.param != null && app.param.TryGetValue(whiteParamKey, out string whiteValue) &&
                    double.TryParse(whiteValue, out standardRatio))
                {
                    // 設定允許的誤差範圍 

                    bool a = ratio < standardRatio + tolerance;
                    bool b = ratio > standardRatio - tolerance;
                    bool isValid = isValid = (a && b);

                    // 記錄日誌
                    //lbAdd($"站點 {stop} 白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%, 允許範圍: [{lowerBound:F2}%, {upperBound:F2}%], 有效性: {(isValid ? "有效" : "無效")}", "inf", "PixelRatioCheck");
                    //Console.WriteLine($"站點 {stop} 白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%, 有效性: {(isValid ? "有效" : "無效")}", "inf", "PixelRatioCheck");
                    if (!isValid)
                    {
                        Log.Warning($"站點 {stop}白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%,  無效");
                    }

                    return isValid;
                }
                else
                {
                    // 若參數不存在，默認為有效
                    lbAdd($"站點 {stop} 的白色像素占比參數未設定，默認視為有效圖像", "war", "PixelRatioParamMissing");
                    return true;
                }
            }
            catch (Exception ex)
            {
                lbAdd($"檢查白色像素占比時發生錯誤: {ex.Message}", "err", ex.Message);
                return true; // 發生錯誤時，默認為有效
            }
        }
        public (bool isValid, double whiteRatio) CheckWhitePixelRatioWithValue(Mat image, int stop)
        {
            bool isValid = true;
            double ratio = 0.0;
            try
            {
                int minthresh = 0;
                if (stop == 1 || stop == 4)
                {
                    minthresh = 180;
                }
                else if (stop == 2)
                {
                    minthresh = 250;
                }
                else if (stop == 3)
                {
                    minthresh = 170;
                }

                minthresh = GetIntParam(app.param, $"whiteThresh_{stop}", 250);

                // 轉為灰階
                Mat gray = new Mat();
                Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                // 二值化 (200~255)
                Mat binary = new Mat();
                Cv2.Threshold(gray, binary, minthresh, 255, ThresholdTypes.Binary);

                // 計算白色像素數量
                int whitePixels = Cv2.CountNonZero(binary);

                // 計算占比 (百分比)
                ratio = (double)whitePixels / (binary.Rows * binary.Cols) * 100;

                gray.Dispose();
                binary.Dispose();

                // 從參數中讀取標準值
                double standardRatio;
                bool a = true;
                bool b = true;
                if (double.TryParse(app.param[$"white_{stop}"], out standardRatio))
                {
                    double tolerance = 0.0;
                    string toleranceKey = $"whiteTolerance_{stop}";

                    if (app.param.ContainsKey(toleranceKey) &&
                        double.TryParse(app.param[toleranceKey], out tolerance))
                    {
                        // 使用資料庫的容忍度
                        a = ratio < standardRatio + tolerance;
                        b = ratio > standardRatio - tolerance;
                        isValid = (a && b);
                    }
                    else
                    {
                        Log.Warning("white 或 whiteTolerance 參數有誤");
                        isValid = false;
                    }
                    // 記錄日誌
                    //lbAdd($"站點 {stop} 白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%, 允許範圍: [{lowerBound:F2}%, {upperBound:F2}%], 有效性: {(isValid ? "有效" : "無效")}", "inf", "PixelRatioCheck");
                    //Console.WriteLine($"站點 {stop} 白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%, 有效性: {(isValid ? "有效" : "無效")}", "inf", "PixelRatioCheck");
                    if (!isValid)
                    {
                        Log.Warning($"站點 {stop}白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%,  無效");
                    }

                    return (isValid, ratio);
                }
                else
                {
                    // 若參數不存在，默認為有效
                    lbAdd($"站點 {stop} 的白色像素占比參數未設定，默認視為有效圖像", "war", "PixelRatioParamMissing");
                    return (isValid, ratio); ;
                }
            }
            catch (Exception ex)
            {
                lbAdd($"檢查白色像素占比時發生錯誤: {ex.Message}", "err", ex.Message);
                return (isValid, ratio); ; // 發生錯誤時，默認為有效
            }
        }
        #region 取物體位置
        /// <summary>
        /// 檢查物體是否在預期位置上（預期位置就是圖像中心，容許偏差50像素）
        /// </summary>
        /// <param name="image">輸入圖像</param>
        /// <param name="stop">站點編號</param>
        /// <returns>物體位置檢測結果，包含是否有瑕疵(位置偏差)以及結果圖像</returns>
        /// 
        private List<Point> objectPositions = new List<Point>();
        /// <summary>
        /// 計算物體座標的統計數據並匯出成 CSV
        /// </summary>
        private void ExportObjectPositionStatistics(string filePath)
        {
            if (objectPositions.Count == 0)
            {
                MessageBox.Show("沒有可用的數據進行匯出。");
                return;
            }

            // 計算統計數據
            var xValues = objectPositions.Select(p => p.X).ToList();
            var yValues = objectPositions.Select(p => p.Y).ToList();

            double xMean = xValues.Average();
            double yMean = yValues.Average();

            double xMedian = GetMedian(xValues);
            double yMedian = GetMedian(yValues);

            int xMin = xValues.Min();
            int yMin = yValues.Min();

            int xMax = xValues.Max();
            int yMax = yValues.Max();

            double xStdDev = GetStandardDeviation(xValues, xMean);
            double yStdDev = GetStandardDeviation(yValues, yMean);

            // 匯出到 CSV
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine("Statistic,X,Y");
                writer.WriteLine($"Mean,{xMean:F2},{yMean:F2}");
                writer.WriteLine($"Median,{xMedian:F2},{yMedian:F2}");
                writer.WriteLine($"Min,{xMin},{yMin}");
                writer.WriteLine($"Max,{xMax},{yMax}");
                writer.WriteLine($"StdDev,{xStdDev:F2},{yStdDev:F2}");
                writer.WriteLine();
                writer.WriteLine("Index,ObjectCenterX,ObjectCenterY");
                for (int i = 0; i < objectPositions.Count; i++)
                {
                    writer.WriteLine($"{i + 1},{objectPositions[i].X},{objectPositions[i].Y}");
                }
            }

            MessageBox.Show($"檢測結果已成功匯出到: {filePath}");
        }

        /// <summary>
        /// 計算中位數
        /// </summary>
        private double GetMedian(List<int> values)
        {
            var sortedValues = values.OrderBy(v => v).ToList();
            int count = sortedValues.Count;

            if (count % 2 == 0)
            {
                return (sortedValues[count / 2 - 1] + sortedValues[count / 2]) / 2.0;
            }
            else
            {
                return sortedValues[count / 2];
            }
        }

        /// <summary>
        /// 計算標準差
        /// </summary>
        private double GetStandardDeviation(List<int> values, double mean)
        {
            double sumOfSquares = values.Select(v => Math.Pow(v - mean, 2)).Sum();
            return Math.Sqrt(sumOfSquares / values.Count);
        }
        #endregion

        public (bool hasDefect, Mat resultImg) CheckObjectPosition(Mat image, int stop)
        {
            
            try
            {
                int objBiasX = int.Parse(app.param[$"objBias_x_{stop}"]);
                int objBiasY = int.Parse(app.param[$"objBias_y_{stop}"]);

                // 1. 設定容許偏差範圍為50像素
                int tolerance = int.Parse(app.param[$"objTolerance_{stop}"]); ;

                // 2. 計算預期中心點  //預設物體中心!!
                Point defaultCenter = new Point(image.Width / 2 + objBiasX, image.Height / 2 + objBiasY);

                // 3. 轉為灰階
                Mat gray = new Mat();
                Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

                // 4. 固定閾值二值化 (210, 255)
                Mat binary = new Mat();
                Cv2.Threshold(gray, binary, 250, 255, ThresholdTypes.Binary);

                // 5. 尋找輪廓
                Point[][] contours;
                HierarchyIndex[] hierarchy;
                Cv2.FindContours(binary, out contours, out hierarchy,
                                RetrievalModes.External, ContourApproximationModes.ApproxSimple);

                // 結果圖像 (複製一份，不修改原圖)
                Mat resultImg = image.Clone();

                // 6. 篩選面積在20-100萬像素的最大輪廓
                Point[] largestContour = null;
                double maxArea = 0;

                foreach (var contour in contours)
                {
                    double area = Cv2.ContourArea(contour);
                    if (200000 <= area && area <= 1000000 && area > maxArea)
                    {
                        maxArea = area;
                        largestContour = contour;
                    }
                }

                // 如果沒找到合適大小的輪廓
                if (largestContour == null)
                {
                    lbAdd($"站點 {stop} - 未找到適當大小的物體輪廓", "war", "輪廓面積應在200000-1000000像素範圍");

                    // 在結果圖上標記預期位置和容許範圍
                    Cv2.Circle(resultImg, defaultCenter, 10, new Scalar(0, 0, 255), -1);  // 紅色-圖像中心
                    Cv2.Circle(resultImg, defaultCenter, tolerance, new Scalar(0, 255, 255), 2);  // 黃色圈-容許範圍
                    Cv2.PutText(resultImg, "No valid contour found",
                                new Point(50, 100), HersheyFonts.HersheySimplex, 1, new Scalar(0, 0, 255), 2);

                    return (true, resultImg);  // 未找到輪廓視為瑕疵
                }

                // 7. 計算輪廓質心
                Moments moments = Cv2.Moments(largestContour);
                if (moments.M00 == 0)
                {
                    lbAdd($"站點 {stop} - 無法計算物體質心", "err", "輪廓矩計算失敗");
                    Cv2.PutText(resultImg, "No moments found",
                                new Point(50, 100), HersheyFonts.HersheySimplex, 1, new Scalar(0, 0, 255), 2);
                    return (true, resultImg);  // 無法計算質心視為瑕疵
                }

                Point objectCenter = new Point(
                    (int)(moments.M10 / moments.M00),
                    (int)(moments.M01 / moments.M00)
                );

                // 8. 計算與圖像中心的偏差
                double offsetX = objectCenter.X - defaultCenter.X;
                double offsetY = objectCenter.Y - defaultCenter.Y;
                double offsetDistance = Math.Sqrt(offsetX * offsetX + offsetY * offsetY);
                // 將物體座標記錄到列表中
                objectPositions.Add(objectCenter);
                // 9. 判斷是否在容許範圍內
                bool isPositionedCorrectly = offsetDistance <= tolerance;

                // 10. 記錄結果
                lbAdd($"站點 {stop} - 物體中心: ({objectCenter.X}, {objectCenter.Y}), " +
                      $"圖像中心: ({defaultCenter.X}, {defaultCenter.Y}), " +
                      $"偏移: X={offsetX:F0}, Y={offsetY:F0}, 距離={offsetDistance:F1}, " +
                      $"容許偏差: {tolerance}, " +
                      $"位置狀態: {(isPositionedCorrectly ? "正常" : "偏離")}",
                      isPositionedCorrectly ? "Inf" : "war", "物體位置檢測");

                // 11. 視覺化結果

                // 繪製輪廓
                Cv2.DrawContours(resultImg, new[] { largestContour }, -1, new Scalar(0, 255, 0), 2);

                // 繪製圖像中心點（紅色）
                Cv2.Circle(resultImg, defaultCenter, 10, new Scalar(0, 0, 255), -1);
                Cv2.PutText(resultImg, "Image Center",
                           new Point(defaultCenter.X + 15, defaultCenter.Y - 15), HersheyFonts.HersheySimplex, 0.8, new Scalar(0, 0, 255), 2);

                // 繪製物體中心點（藍色）
                Cv2.Circle(resultImg, objectCenter, 10, new Scalar(255, 0, 0), -1);
                Cv2.PutText(resultImg, "Object Center",
                           new Point(objectCenter.X + 15, objectCenter.Y - 15), HersheyFonts.HersheySimplex, 0.8, new Scalar(255, 0, 0), 2);

                // 繪製從物體中心到圖像中心的連線
                Cv2.Line(resultImg, objectCenter, defaultCenter, new Scalar(255, 255, 0), 2);

                // 繪製容許範圍圓圈
                Cv2.Circle(resultImg, defaultCenter, tolerance,
                          isPositionedCorrectly ? new Scalar(0, 255, 0) : new Scalar(0, 0, 255), 2);

                // 顯示偏移資訊
                string offsetInfo = $"Offset: X={offsetX:F0}, Y={offsetY:F0}, Dist={offsetDistance:F1}";
                Cv2.PutText(resultImg, offsetInfo,
                            new Point(50, 50), HersheyFonts.HersheySimplex, 0.8, new Scalar(255, 255, 0), 2);

                // 顯示結果狀態
                string statusText = $"Status: {(isPositionedCorrectly ? "OK" : "NG")}";
                Cv2.PutText(resultImg, statusText,
                           new Point(50, 100), HersheyFonts.HersheySimplex, 1.0,
                           isPositionedCorrectly ? new Scalar(0, 255, 0) : new Scalar(0, 0, 255), 2);

                // 顯示結果圖像
                //showResultMat(resultImg, stop);

                //釋放
                 gray.Dispose();
                 binary.Dispose();

                // 回傳結果：位置不正確視為有瑕疵
                return (!isPositionedCorrectly, resultImg);
            }
            catch (Exception ex)
            {
                lbAdd($"檢查物體位置時發生錯誤: {ex.Message}", "err", ex.ToString());
                return (true, image.Clone());  // 發生錯誤時，返回有瑕疵
            }
        }

        private bool CheckChamfer(Mat image, int stop) //倒角AOI
        {
            // 轉換為灰階
            Mat gray = new Mat();
            Cv2.CvtColor(image, gray, ColorConversionCodes.BGR2GRAY);

            // 二值化處理
            Mat binary = new Mat();
            Cv2.Threshold(gray, binary, 127, 255, ThresholdTypes.Binary);

            // 獲取內圓中心和半徑參數
            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
            int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);

            // 創建掩碼，將外部區域填充為黑色
            Mat mask = new Mat(binary.Size(), MatType.CV_8UC1, Scalar.Black);
            Cv2.Circle(mask, new Point(knownInnerCenterX, knownInnerCenterY), knownInnerRadius, Scalar.White, -1);

            // 應用掩碼，僅保留內部圓區域
            Mat roi = new Mat();
            Cv2.BitwiseAnd(binary, binary, roi, mask);

            // 計算內部區域的白色像素數量
            int whitePixelCount = Cv2.CountNonZero(roi);

            // 計算內部區域的總像素數
            int totalPixels = (int)(Math.PI * knownInnerRadius * knownInnerRadius);

            // 計算白色像素占比
            double ratio = (double)whitePixelCount / totalPixels;

            double standardRatio;
            if (double.TryParse(app.param[$"chamfer_{stop}"], out standardRatio))
            {
                // 設定允許的誤差範圍 (±2%)
                double lowerBound = standardRatio - 2.0;
                double upperBound = standardRatio + 2.0;

                // 判斷是否在允許範圍內
                bool isValid = (ratio >= lowerBound && ratio <= upperBound);

                // 記錄日誌
                //lbAdd($"站點 {stop} 白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%, 允許範圍: [{lowerBound:F2}%, {upperBound:F2}%], 有效性: {(isValid ? "有效" : "無效")}", "inf", "PixelRatioCheck");
                Console.WriteLine($"站點 {stop} 白色像素占比: {ratio:F2}%, 標準值: {standardRatio}%, 允許範圍: [{lowerBound:F2}%, {upperBound:F2}%], 有效性: {(isValid ? "有效" : "無效")}", "inf", "PixelRatioCheck");

                return isValid;
            }
            else
            {
                // 若參數不存在，默認為有效
                lbAdd($"站點 {stop} 的白色像素占比參數未設定，默認視為有效圖像", "war", "PixelRatioParamMissing");
                return true;
            }
        }
        public double CalculateIoU(Rect rect1, Rect rect2)
        {
            // 計算交集矩形的左上角和右下角座標
            int interLeft = Math.Max(rect1.Left, rect2.Left);
            int interTop = Math.Max(rect1.Top, rect2.Top);
            int interRight = Math.Min(rect1.Right, rect2.Right);
            int interBottom = Math.Min(rect1.Bottom, rect2.Bottom);

            // 計算交集的寬度和高度
            int interWidth = Math.Max(0, interRight - interLeft);
            int interHeight = Math.Max(0, interBottom - interTop);

            // 計算交集面積
            int intersectionArea = interWidth * interHeight;

            // 計算兩個矩形的面積
            int area1 = rect1.Width * rect1.Height;
            int area2 = rect2.Width * rect2.Height;

            // 計算聯合面積
            int unionArea = area1 + area2 - intersectionArea;

            // 避免除以零的情況
            //if (unionArea == 0) return 0;
            if (area1 == 0) return 0.0;
            // 返回 IoU
            //return (double)intersectionArea / unionArea;
            return (double)intersectionArea / area1;
        }
        private bool IsPointInPolygon(Point point, Point[] polygon)
        {
            int n = polygon.Length;
            bool inside = false;

            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                if ((polygon[i].Y > point.Y) != (polygon[j].Y > point.Y) &&
                    point.X < (polygon[j].X - polygon[i].X) * (point.Y - polygon[i].Y) /
                            (polygon[j].Y - polygon[i].Y) + polygon[i].X)
                {
                    inside = !inside;
                }
            }

            return inside;
        }

        /// <summary>
        /// 判斷瑕疵是否位於非ROI區域，使用9點採樣法
        /// </summary>
        /// <param name="defectRect">瑕疵矩形框</param>
        /// <param name="nonRoiRect">非ROI矩形框</param>
        /// <param name="circleCenter">圓心坐標</param>
        /// <param name="innerRadius">內圓半徑</param>
        /// <param name="outerRadius">外圓半徑</param>
        /// <param name="innerExpandPixels">內圓擴展的像素數</param>
        /// <param name="outerExpandPixels">外圓擴展的像素數</param>
        /// <returns>瑕疵是否位於非ROI區域</returns>

        #region NROI01
        public bool IsDefectInNonRoiRegion_in(Rect defectRect, Rect nonRoiRect, /*Point circleCenter,*/
                                         /*double innerRadius, double outerRadius,*/ int stop,
                                         int innerExpandPixels, int outerExpandPixels)
        {
            // 1. 先計算非ROI區域的多邊形表示
            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
            int outerRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);
            int innerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);
            Point circleCenter = new Point(
                knownInnerCenterX,
                knownInnerCenterY
            );
            
            Point[] nonRoiPolygon = CalculateNonRoiPolygon(nonRoiRect, circleCenter, innerRadius, outerRadius,
                                                          innerExpandPixels, outerExpandPixels);
            
            // 計算瑕疵中心點
            Point defectCenter = new Point(
                defectRect.X + defectRect.Width / 2,
                defectRect.Y + defectRect.Height / 2
            );
            // 判斷瑕疵中心點是否在非ROI多邊形內

            bool isinpolygon = IsPointInPolygon(defectCenter, nonRoiPolygon);


            return isinpolygon;

            /*
            // 2. 在瑕疵矩形上均勻採樣9個點
            Point[] samplePoints = new Point[9];

            // 四個角點
            samplePoints[0] = new Point(defectRect.X, defectRect.Y); // 左上角
            samplePoints[1] = new Point(defectRect.X + defectRect.Width, defectRect.Y); // 右上角
            samplePoints[2] = new Point(defectRect.X, defectRect.Y + defectRect.Height); // 左下角
            samplePoints[3] = new Point(defectRect.X + defectRect.Width, defectRect.Y + defectRect.Height); // 右下角

            // 四邊中點
            samplePoints[4] = new Point(defectRect.X + defectRect.Width / 2, defectRect.Y); // 上邊中點
            samplePoints[5] = new Point(defectRect.X + defectRect.Width, defectRect.Y + defectRect.Height / 2); // 右邊中點
            samplePoints[6] = new Point(defectRect.X + defectRect.Width / 2, defectRect.Y + defectRect.Height); // 下邊中點
            samplePoints[7] = new Point(defectRect.X, defectRect.Y + defectRect.Height / 2); // 左邊中點

            // 中心點
            samplePoints[8] = new Point(defectRect.X + defectRect.Width / 2, defectRect.Y + defectRect.Height / 2);

            // 3. 計算有多少點落在非ROI區域內
            int pointsInNonRoi = 0;
            foreach (var point in samplePoints)
            {
                if (IsPointInPolygon(point, nonRoiPolygon))
                {
                    pointsInNonRoi++;
                }
            }

            // 4. 如果超過指定比例的點在非ROI區域內，則認為瑕疵在非ROI區域內
            // 這裡設定為20%，2個點就算
            double threshold = 0.2; // 20%         
            if(IsPointInPolygon(samplePoints[8], nonRoiPolygon))
            {
                return true;
            }
            else if ((double)pointsInNonRoi / samplePoints.Length >= threshold) //中心在裡面直接算
            {
                return true;
            } // 條件可以再加
            else
            {
                return false;
            }
            */
        }

        /// <summary>
        /// 計算非ROI區域的多邊形表示，使用固定像素擴展
        /// </summary>
        private Point[] CalculateNonRoiPolygon(Rect nonRoiRect, Point circleCenter, double innerRadius, double outerRadius,
                                             int innerExpandPixels, int outerExpandPixels)
        {
            // 1. 計算非ROI矩形的中心點
            Point nonRoiCenter = new Point(
                nonRoiRect.X + nonRoiRect.Width / 2,
                nonRoiRect.Y + nonRoiRect.Height / 2
            );

            // 2. 計算從圓心到非ROI中心的向量
            double vectorX = nonRoiCenter.X - circleCenter.X;
            double vectorY = nonRoiCenter.Y - circleCenter.Y;

            // 3. 計算向量長度
            double vectorLength = Math.Sqrt(vectorX * vectorX + vectorY * vectorY);

            // 4. 單位化向量
            double unitVectorX = vectorX / vectorLength;
            double unitVectorY = vectorY / vectorLength;

            // 5. 沿著向量方向，從圓心延伸到外圓邊緣的點
            Point outerEdgePoint = new Point(
                (int)(circleCenter.X + unitVectorX * outerRadius),
                (int)(circleCenter.Y + unitVectorY * outerRadius)
            );

            // 6. 計算垂直於該向量的單位向量
            double perpUnitVectorX = -unitVectorY;
            double perpUnitVectorY = unitVectorX;

            // 7. 構建多邊形的四個頂點，使用固定像素擴展
            Point[] nonRoiPolygon = new Point[4];

            // 在外圓邊緣點兩側各擴展指定像素
            nonRoiPolygon[0] = new Point(
                (int)(outerEdgePoint.X + perpUnitVectorX * outerExpandPixels),
                (int)(outerEdgePoint.Y + perpUnitVectorY * outerExpandPixels)
            );
            nonRoiPolygon[1] = new Point(
                (int)(outerEdgePoint.X - perpUnitVectorX * outerExpandPixels),
                (int)(outerEdgePoint.Y - perpUnitVectorY * outerExpandPixels)
            );

            // 內側邊界點
            Point innerEdgePoint = new Point(
                (int)(circleCenter.X + unitVectorX * innerRadius),
                (int)(circleCenter.Y + unitVectorY * innerRadius)
            );

            // 內側兩個頂點，也用固定像素擴展
            nonRoiPolygon[2] = new Point(
                (int)(innerEdgePoint.X - perpUnitVectorX * innerExpandPixels),
                (int)(innerEdgePoint.Y - perpUnitVectorY * innerExpandPixels)
            );
            nonRoiPolygon[3] = new Point(
                (int)(innerEdgePoint.X + perpUnitVectorX * innerExpandPixels),
                (int)(innerEdgePoint.Y + perpUnitVectorY * innerExpandPixels)
            );

            return nonRoiPolygon;
        }

        /// <summary>
        /// 繪製非ROI區域到影像上（用於視覺化調試）
        /// </summary>
        public Mat DrawNonRoiRegion_in(Mat inputImage, Rect nonRoiRect, /*Point circleCenter, double innerRadius, double outerRadius,*/
                                   int stop, int innerExpandPixels = 20, int outerExpandPixels = 40)
        {
            Mat result = inputImage.Clone();
            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
            int outerRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);
            int innerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);
            Point circleCenter = new Point(
                knownInnerCenterX,
                knownInnerCenterY
            );
            try
            {
                // 1. 計算非ROI多邊形
                Point[] nonRoiPolygon = CalculateNonRoiPolygon(
                    nonRoiRect, circleCenter, innerRadius, outerRadius, innerExpandPixels, outerExpandPixels);

                // 2. 繪製非ROI多邊形
                Cv2.FillConvexPoly(result, nonRoiPolygon, new Scalar(0, 255, 0));

                // 3. 繪製連接線（用於調試）
                Point nonRoiCenter = new Point(
                    nonRoiRect.X + nonRoiRect.Width / 2,
                    nonRoiRect.Y + nonRoiRect.Height / 2
                );
                Cv2.Line(result, circleCenter, nonRoiCenter, new Scalar(0, 0, 255), 1);

                // 4. 繪製原始非ROI矩形（用於比較）
                Cv2.Rectangle(result, nonRoiRect, new Scalar(255, 0, 0), 1);

                // 5. 顯示擴展參數信息
                Cv2.PutText(result,
                    $"in_ex: {innerExpandPixels}px, out_ex: {outerExpandPixels}px",
                    new Point(50, 100),
                    HersheyFonts.HersheyDuplex, 0.8, new Scalar(255, 255, 0), 1);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"繪製非ROI區域失敗: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 檢查點是否在多邊形內
        /// </summary>
        

        #endregion
        #region NROI02

        /// <summary>
        /// 判斷瑕疵是否位於非ROI區域
        /// </summary>
        /// <param name="defectRect">瑕疵矩形框</param>
        /// <param name="nonRoiRect">非ROI矩形框（YOLO檢測結果）</param>
        /// <param name="stop">站點編號，用於獲取參數</param>
        /// <param name="innerExpandPixels">內圓擴展像素（預設值）</param> 暫時不使用
        /// <param name="outerExpandPixels">外圓擴展像素（預設值）</param>
        /// <returns>瑕疵是否位於非ROI區域</returns>
        public bool IsDefectInNonRoiRegion_out(Rect defectRect, Rect nonRoiRect, int stop,
                                         int innerExpandPixels, int outerExpandPixels)
        {
            try
            {
                // 獲取圓心座標
                Point circleCenter = new Point();

                // 從app.param中讀取圓心座標
                if (app.param.ContainsKey($"known_inner_center_x_{stop}") && app.param.ContainsKey($"known_inner_center_y_{stop}"))
                {
                    circleCenter.X = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                    circleCenter.Y = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                }
                else
                {
                    // 如果沒有設定圓心，使用標準值
                    circleCenter.X = 1224; // 假設標準影像寬度一半
                    circleCenter.Y = 1024; // 假設標準影像高度一半
                }

                // 從app.param中獲取擴展參數
                if (app.param.ContainsKey($"expandNROI_out_{stop}"))
                {
                    outerExpandPixels = int.Parse(app.param[$"expandNROI_out_{stop}"]);
                }

                // 找出非ROI矩形框最靠近圓心的角點
                Point closestCorner = GetClosestCornerToCircle(nonRoiRect, circleCenter);

                // 計算瑕疵中心點
                Point defectCenter = new Point(
                    defectRect.X + defectRect.Width / 2,
                    defectRect.Y + defectRect.Height / 2
                );

                // 判斷最靠近的角點與圓心的X或Y座標相差是否小於100
                bool isClose = Math.Abs(closestCorner.X - circleCenter.X) < 100 ||
                               Math.Abs(closestCorner.Y - circleCenter.Y) < 100;
                // 如果角點與圓心很近，直接使用原檢測框
                if (isClose)
                {
                    // 使用原始框，計算瑕疵矩形與非ROI區域的IoU
                    //return CalculateIoU(defectRect, nonRoiRect) > 0;
                    // 如果角點與圓心很近，直接使用原檢測框
                    // 檢查瑕疵中心點是否在非ROI矩形內
                    return defectCenter.X >= nonRoiRect.X &&
                           defectCenter.X <= nonRoiRect.X + nonRoiRect.Width &&
                           defectCenter.Y >= nonRoiRect.Y &&
                           defectCenter.Y <= nonRoiRect.Y + nonRoiRect.Height;
                }
                else
                {
                    // 如果角點與圓心的距離較遠，則重新計算非ROI區域
                    // 找出對角點
                    Point farthestCorner = GetFarthestCornerToCorner(nonRoiRect, closestCorner);

                    // 計算A線段（從最近角點到最遠角點的向量）
                    Point aVector = new Point(
                        farthestCorner.X - closestCorner.X,
                        farthestCorner.Y - closestCorner.Y
                    );

                    // 計算垂直於A線段的方向向量
                    Point perpVector = new Point(-aVector.Y, aVector.X);

                    // 正規化垂直向量
                    double length = Math.Sqrt(perpVector.X * perpVector.X + perpVector.Y * perpVector.Y);
                    if (length > 0)
                    {
                        perpVector.X = (int)(perpVector.X / length * outerExpandPixels);
                        perpVector.Y = (int)(perpVector.Y / length * outerExpandPixels);
                    }

                    // 計算B線段和C線段的四個頂點
                    Point b1 = new Point(closestCorner.X + perpVector.X, closestCorner.Y + perpVector.Y);
                    Point b2 = new Point(farthestCorner.X + perpVector.X, farthestCorner.Y + perpVector.Y);
                    Point c1 = new Point(closestCorner.X - perpVector.X, closestCorner.Y - perpVector.Y);
                    Point c2 = new Point(farthestCorner.X - perpVector.X, farthestCorner.Y - perpVector.Y);

                    // 形成新的非ROI多邊形
                    Point[] nonRoiPolygon = new Point[] { b1, b2, c2, c1 };

                    // 檢查瑕疵中心點是否在非ROI多邊形內
                    return IsPointInPolygon(defectCenter, nonRoiPolygon);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"非ROI區域計算錯誤: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 繪製非ROI區域（用於視覺化調試）
        /// </summary>
        public Mat DrawNonRoiRegion_out(Mat inputImage, Rect nonRoiRect, int stop, int innerExpandPixels = 20, int outerExpandPixels = 40)
        {
            Mat result = inputImage.Clone();

            try
            {
                // 從app.param中獲取擴展參數
                if (app.param.ContainsKey($"expandNROI_out_{stop}"))
                {
                    outerExpandPixels = int.Parse(app.param[$"expandNROI_out_{stop}"]);
                }

                // 獲取圓心座標
                Point circleCenter = new Point();

                if (app.param.ContainsKey($"known_inner_center_x_{stop}") && app.param.ContainsKey($"known_inner_center_y_{stop}"))
                {
                    circleCenter.X = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                    circleCenter.Y = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                }
                else
                {
                    circleCenter.X = 1224;
                    circleCenter.Y = 1024;
                }

                // 找出非ROI矩形框最靠近圓心的角點
                Point closestCorner = GetClosestCornerToCircle(nonRoiRect, circleCenter);

                // 判斷與圓心的距離
                bool isClose = Math.Abs(closestCorner.X - circleCenter.X) < 100 ||
                               Math.Abs(closestCorner.Y - circleCenter.Y) < 100;

                if (isClose)
                {
                    // 直接使用原檢測框
                    Cv2.Rectangle(result, nonRoiRect, new Scalar(0, 255, 0), 2);
                    Cv2.PutText(result, "使用原始非ROI檢測框", new Point(nonRoiRect.X, nonRoiRect.Y - 10),
                            HersheyFonts.HersheyComplex, 0.5, new Scalar(0, 255, 0), 1);
                }
                else
                {
                    // 計算新的非ROI區域
                    Point farthestCorner = GetFarthestCornerToCorner(nonRoiRect, closestCorner);

                    // 計算A線段向量
                    Point aVector = new Point(
                        farthestCorner.X - closestCorner.X,
                        farthestCorner.Y - closestCorner.Y
                    );

                    // 計算垂直向量
                    Point perpVector = new Point(-aVector.Y, aVector.X);

                    // 正規化垂直向量
                    double length = Math.Sqrt(perpVector.X * perpVector.X + perpVector.Y * perpVector.Y);
                    if (length > 0)
                    {
                        perpVector.X = (int)(perpVector.X / length * outerExpandPixels);
                        perpVector.Y = (int)(perpVector.Y / length * outerExpandPixels);
                    }

                    // 計算四個頂點
                    Point b1 = new Point(closestCorner.X + perpVector.X, closestCorner.Y + perpVector.Y);
                    Point b2 = new Point(farthestCorner.X + perpVector.X, farthestCorner.Y + perpVector.Y);
                    Point c1 = new Point(closestCorner.X - perpVector.X, closestCorner.Y - perpVector.Y);
                    Point c2 = new Point(farthestCorner.X - perpVector.X, farthestCorner.Y - perpVector.Y);

                    // 繪製A線段
                    Cv2.Line(result, closestCorner, farthestCorner, new Scalar(0, 0, 255), 2);

                    // 繪製新的非ROI多邊形
                    Point[] nonRoiPolygon = new Point[] { b1, b2, c2, c1 };
                    Cv2.FillConvexPoly(result, nonRoiPolygon, new Scalar(0, 255, 0, 80));
                    Cv2.Polylines(result, new Point[][] { nonRoiPolygon }, true, new Scalar(0, 255, 0), 2);

                    // 標記最近角點和最遠角點
                    Cv2.Circle(result, closestCorner, 5, new Scalar(255, 0, 0), -1);
                    Cv2.Circle(result, farthestCorner, 5, new Scalar(0, 0, 255), -1);

                    Cv2.PutText(result, "重新計算非ROI區域", new Point(closestCorner.X, closestCorner.Y - 10),
                            HersheyFonts.HersheyComplex, 0.5, new Scalar(0, 255, 0), 1);
                }

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"繪製非ROI區域錯誤: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// 找出矩形中最靠近圓心的角點
        /// </summary>
        private Point GetClosestCornerToCircle(Rect rect, Point circleCenter)
        {
            // 矩形的四個角點
            Point[] corners = new Point[]
            {
                new Point(rect.X, rect.Y),                       // 左上
                new Point(rect.X + rect.Width, rect.Y),          // 右上
                new Point(rect.X, rect.Y + rect.Height),         // 左下
                new Point(rect.X + rect.Width, rect.Y + rect.Height) // 右下
            };

            // 找出最近的角點
            double minDistance = double.MaxValue;
            Point closestCorner = corners[0];

            foreach (var corner in corners)
            {
                double distance = Math.Sqrt(
                    Math.Pow(corner.X - circleCenter.X, 2) +
                    Math.Pow(corner.Y - circleCenter.Y, 2)
                );

                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestCorner = corner;
                }
            }

            return closestCorner;
        }

        /// <summary>
        /// 找出矩形中與給定角點距離最遠的角點（對角點）
        /// </summary>
        private Point GetFarthestCornerToCorner(Rect rect, Point corner)
        {
            // 矩形的四個角點
            Point[] corners = new Point[]
            {
                new Point(rect.X, rect.Y),                       // 左上
                new Point(rect.X + rect.Width, rect.Y),          // 右上
                new Point(rect.X, rect.Y + rect.Height),         // 左下
                new Point(rect.X + rect.Width, rect.Y + rect.Height) // 右下
            };

            // 找出最遠的角點
            double maxDistance = 0;
            Point farthestCorner = corners[0];

            foreach (var c in corners)
            {
                double distance = Math.Sqrt(
                    Math.Pow(c.X - corner.X, 2) +
                    Math.Pow(c.Y - corner.Y, 2)
                );

                if (distance > maxDistance)
                {
                    maxDistance = distance;
                    farthestCorner = c;
                }
            }

            return farthestCorner;
        }


        /// <summary>
        /// 檢查點是否在多邊形內 (射線算法)
        /// </summary>
        #endregion

        private void testchamfer_Click(object sender, EventArgs e)
        {
            // 此函數是測試倒角不均使用，採取白色像素計算，建議只能用在第四站
            // 觀察效果尚可，但須先找到常態參數

            // 創建資料夾選擇對話框
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "請選擇包含圖片的資料夾";
                if (Directory.Exists(@"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release"))
                    folderDialog.SelectedPath = @"C:\Users\User\Desktop\peilin2_TEST0303\bin\x64\Release";

                // 如果使用者選擇了資料夾並點擊確定
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderDialog.SelectedPath;

                    // 選擇要檢查的站點
                    int stop = 3; // 預設值
                    using (var stopSelectForm = new Form())
                    {
                        stopSelectForm.Text = "選擇站點";
                        stopSelectForm.Size = new System.Drawing.Size(300, 150);
                        stopSelectForm.StartPosition = FormStartPosition.CenterScreen;
                        stopSelectForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                        stopSelectForm.MaximizeBox = false;
                        stopSelectForm.MinimizeBox = false;

                        Label label = new Label();
                        label.Text = "請選擇站點:";
                        label.Location = new System.Drawing.Point(20, 20);
                        label.Size = new System.Drawing.Size(100, 20);

                        ComboBox comboBox = new ComboBox();
                        comboBox.Location = new System.Drawing.Point(130, 20);
                        comboBox.Size = new System.Drawing.Size(100, 20);
                        comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                        comboBox.Items.AddRange(new object[] { "1", "2", "3", "4" });
                        comboBox.SelectedIndex = 2; // 預設選擇站點3

                        Button okButton = new Button();
                        okButton.Text = "確定";
                        okButton.DialogResult = DialogResult.OK;
                        okButton.Location = new System.Drawing.Point(100, 60);
                        okButton.Size = new System.Drawing.Size(80, 30);

                        stopSelectForm.Controls.Add(label);
                        stopSelectForm.Controls.Add(comboBox);
                        stopSelectForm.Controls.Add(okButton);
                        stopSelectForm.AcceptButton = okButton;

                        if (stopSelectForm.ShowDialog() == DialogResult.OK)
                        {
                            stop = comboBox.SelectedIndex + 1;
                        }
                        else
                        {
                            return; // 用戶取消
                        }
                    }

                    // 建立統計數據的集合
                    List<ChamferStatistic> statistics = new List<ChamferStatistic>();

                    try
                    {
                        // 獲取資料夾中所有圖片檔案
                        string[] allImageFiles = Directory.GetFiles(folderPath, "*.png")
                            .Union(Directory.GetFiles(folderPath, "*.jpg"))
                            .Union(Directory.GetFiles(folderPath, "*.jpeg"))
                            .Union(Directory.GetFiles(folderPath, "*.bmp"))
                            .ToArray();

                        // 根據所選站點過濾圖片 (只保留檔名末尾為"-{stop}"的圖片)
                        string[] imageFiles = allImageFiles
                            .Where(filePath => Path.GetFileNameWithoutExtension(filePath).EndsWith($"-{stop}"))
                            .ToArray();

                        if (imageFiles.Length == 0)
                        {
                            MessageBox.Show($"所選資料夾中沒有站點 {stop} 的圖片檔案!");
                            return;
                        }

                        // 檢查必要參數是否存在
                        if (!app.param.ContainsKey($"known_inner_center_x_{stop}") ||
                            !app.param.ContainsKey($"known_inner_center_y_{stop}") ||
                            !app.param.ContainsKey($"known_inner_radius_{stop}"))
                        {
                            MessageBox.Show($"站點 {stop} 缺少必要的內圓參數設定，請先設定參數！");
                            return;
                        }

                        // 創建進度條表單
                        using (var progressForm = new Form())
                        {
                            progressForm.Text = "處理進度";
                            progressForm.Size = new System.Drawing.Size(400, 100);
                            progressForm.StartPosition = FormStartPosition.CenterScreen;
                            progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                            progressForm.MaximizeBox = false;
                            progressForm.MinimizeBox = false;

                            ProgressBar progressBar = new ProgressBar();
                            progressBar.Dock = DockStyle.Fill;
                            progressBar.Minimum = 0;
                            progressBar.Maximum = imageFiles.Length;
                            progressBar.Step = 1;

                            progressForm.Controls.Add(progressBar);

                            // 顯示進度表單但不阻塞主線程
                            progressForm.Show();

                            // 獲取內圓參數
                            int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                            int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                            int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);

                            // 縮小40像素的內圓半徑
                            int innerSmallRadius = Math.Max(knownInnerRadius - 40, 10); // 確保半徑不會小於10像素

                            // 處理每張圖片
                            foreach (string imagePath in imageFiles)
                            {
                                // 讀取圖片
                                Mat originalImage = Cv2.ImRead(imagePath);
                                if (originalImage.Empty())
                                {
                                    lbAdd($"無法讀取圖片: {imagePath}", "err", "ImageReadFailed");
                                    continue;
                                }

                                // 複製原始圖像以進行分析，保持原始圖像不變
                                Mat processImage = originalImage.Clone();

                                // 轉換為灰階
                                Mat gray = new Mat();
                                Cv2.CvtColor(processImage, gray, ColorConversionCodes.BGR2GRAY);

                                // 二值化處理
                                Mat binary = new Mat();
                                Cv2.Threshold(gray, binary, 127, 255, ThresholdTypes.Binary);

                                // 創建掩碼，將外部區域填充為黑色
                                Mat mask = new Mat(binary.Size(), MatType.CV_8UC1, Scalar.Black);

                                // 繪製內圓區域（原始內圓）
                                Cv2.Circle(mask, new Point(knownInnerCenterX, knownInnerCenterY), knownInnerRadius, Scalar.White, -1);

                                // 繪製縮小後的內圓區域為黑色（將最內部區域填黑）
                                Cv2.Circle(mask, new Point(knownInnerCenterX, knownInnerCenterY), innerSmallRadius, Scalar.Black, -1);

                                // 應用掩碼，僅保留環形區域
                                Mat roi = new Mat();
                                Cv2.BitwiseAnd(binary, binary, roi, mask);

                                // 計算環形區域的白色像素數量
                                int whitePixelCount = Cv2.CountNonZero(roi);

                                // 計算環形區域的總像素數
                                int totalPixels = (int)(Math.PI * (Math.Pow(knownInnerRadius, 2) - Math.Pow(innerSmallRadius, 2)));

                                // 計算白色像素占比
                                double ratio = (double)whitePixelCount / totalPixels * 100; // 轉為百分比

                                // 獲取標準比例（如果有）
                                double standardRatio = 0;
                                if (app.param.ContainsKey($"chamfer_whiteLower_{stop}") && double.TryParse(app.param[$"chamfer_whiteLower_{stop}"], out standardRatio))
                                {
                                    // 有設定標準比例
                                }

                                // 保存處理後的圖像用於檢查
                                string outputDir = Path.Combine(folderPath, "processed");
                                Directory.CreateDirectory(outputDir);

                                // 創建一個結果圖像（使用原始圖像）
                                Mat resultImage = originalImage.Clone();

                                // 在原始圖像上繪製檢測區域
                                Cv2.Circle(resultImage, new Point(knownInnerCenterX, knownInnerCenterY), knownInnerRadius, new Scalar(0, 255, 0), 2);
                                Cv2.Circle(resultImage, new Point(knownInnerCenterX, knownInnerCenterY), innerSmallRadius, new Scalar(0, 0, 255), 2);

                                // 在左上角添加檢測結果信息
                                // 增加黑色底框以提高文字可讀性
                                int infoBoxHeight = 130; // 根據顯示的行數調整
                                int infoBoxWidth = 400; // 根據文字長度調整

                                // 添加半透明背景
                                Mat overlay = resultImage.Clone();
                                Cv2.Rectangle(
                                    overlay,
                                    new Rect(10, 10, infoBoxWidth, infoBoxHeight),
                                    new Scalar(0, 0, 0),
                                    -1
                                );

                                // 應用半透明效果
                                double alpha = 0.7;
                                Cv2.AddWeighted(overlay, alpha, resultImage, 1 - alpha, 0, resultImage);

                                // 1. 檔名
                                string fileName = Path.GetFileName(imagePath);
                                Cv2.PutText(
                                    resultImage,
                                    $"文件: {fileName}",
                                    new Point(20, 30),
                                    HersheyFonts.HersheyDuplex,
                                    0.7,
                                    new Scalar(255, 255, 255),
                                    1,
                                    LineTypes.AntiAlias
                                );

                                // 2. 白色像素占比
                                Cv2.PutText(
                                    resultImage,
                                    $"白色像素占比: {ratio:F2}%",
                                    new Point(20, 60),
                                    HersheyFonts.HersheyDuplex,
                                    0.7,
                                    new Scalar(255, 255, 255),
                                    1,
                                    LineTypes.AntiAlias
                                );

                                // 3. 標準比例及差異（如果有）
                                if (standardRatio > 0)
                                {
                                    // 設定顏色基於比例差異
                                    Scalar textColor;
                                    double diff = Math.Abs(ratio - standardRatio);
                                    if (diff <= 2.0)
                                    {
                                        textColor = new Scalar(0, 255, 0); // 綠色 - 在允許範圍內
                                    }
                                    else
                                    {
                                        textColor = new Scalar(0, 0, 255); // 紅色 - 超出允許範圍
                                    }

                                    Cv2.PutText(
                                        resultImage,
                                        $"標準比例: {standardRatio:F2}% (差異: {diff:F2}%)",
                                        new Point(20, 90),
                                        HersheyFonts.HersheyDuplex,
                                        0.7,
                                        textColor,
                                        1,
                                        LineTypes.AntiAlias
                                    );
                                }

                                // 4. 生成時間
                                Cv2.PutText(
                                    resultImage,
                                    $"生成時間: {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}",
                                    new Point(20, 120),
                                    HersheyFonts.HersheyDuplex,
                                    0.7,
                                    new Scalar(255, 255, 255),
                                    1,
                                    LineTypes.AntiAlias
                                );

                                // 保存處理後的圖像（原始圖像+檢測結果）
                                string outputPath = Path.Combine(outputDir, Path.GetFileName(imagePath));
                                Cv2.ImWrite(outputPath, resultImage);

                                // 記錄統計數據
                                statistics.Add(new ChamferStatistic
                                {
                                    FileName = Path.GetFileName(imagePath),
                                    WhitePixelCount = whitePixelCount,
                                    TotalPixels = totalPixels,
                                    Ratio = ratio
                                });

                                // 更新進度條
                                progressBar.PerformStep();
                                Application.DoEvents();

                                // 釋放資源
                                originalImage.Dispose();
                                processImage.Dispose();
                                gray.Dispose();
                                binary.Dispose();
                                mask.Dispose();
                                roi.Dispose();
                                resultImage.Dispose();
                                if (overlay != null) overlay.Dispose();
                            }

                            progressForm.Close();
                        }

                        // 如果有統計數據，輸出CSV並生成直方圖
                        if (statistics.Count > 0)
                        {
                            // 計算統計值
                            List<double> ratioValues = statistics.Select(s => s.Ratio).ToList();
                            ratioValues.Sort(); // 排序以計算中位數

                            double minRatio = ratioValues.Min();
                            double maxRatio = ratioValues.Max();
                            double avgRatio = ratioValues.Average();
                            double stdDev = Math.Sqrt(ratioValues.Average(val => Math.Pow(val - avgRatio, 2)));

                            // 計算中位數
                            double medianRatio;
                            int middleIndex = ratioValues.Count / 2;
                            if (ratioValues.Count % 2 == 0)
                            {
                                medianRatio = (ratioValues[middleIndex - 1] + ratioValues[middleIndex]) / 2;
                            }
                            else
                            {
                                medianRatio = ratioValues[middleIndex];
                            }

                            // 創建直方圖
                            CreateHistogram(ratioValues, folderPath, stop);

                            // 建立CSV儲存對話框
                            using (SaveFileDialog saveDialog = new SaveFileDialog())
                            {
                                saveDialog.Filter = "CSV檔案|*.csv";
                                saveDialog.Title = "儲存統計結果";
                                saveDialog.FileName = $"Chamfer_Statistics_Station{stop}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.csv";

                                if (saveDialog.ShowDialog() == DialogResult.OK)
                                {
                                    using (StreamWriter writer = new StreamWriter(saveDialog.FileName, false, Encoding.UTF8))
                                    {
                                        // 寫入表頭
                                        writer.WriteLine("檔案名稱,白色像素數量,總像素數量,白色像素占比(%)");

                                        // 寫入每張圖片的數據
                                        foreach (var stat in statistics)
                                        {
                                            writer.WriteLine($"{stat.FileName},{stat.WhitePixelCount},{stat.TotalPixels},{stat.Ratio:F2}");
                                        }

                                        // 寫入統計摘要
                                        writer.WriteLine();
                                        writer.WriteLine("統計摘要");
                                        writer.WriteLine($"最小值(%),{minRatio:F2}");
                                        writer.WriteLine($"最大值(%),{maxRatio:F2}");
                                        writer.WriteLine($"平均值(%),{avgRatio:F2}");
                                        writer.WriteLine($"中位數(%),{medianRatio:F2}");
                                        writer.WriteLine($"標準差(%),{stdDev:F2}");
                                    }

                                    MessageBox.Show($"已成功處理 {statistics.Count} 張站點 {stop} 的圖片並儲存統計結果!\n" +
                                                   $"平均白色像素占比: {avgRatio:F2}%\n" +
                                                   $"中位數: {medianRatio:F2}%\n" +
                                                   $"標準差: {stdDev:F2}%\n" +
                                                   $"直方圖已儲存在 {folderPath} 資料夾\n" +
                                                   $"處理後的圖像已保存在 {Path.Combine(folderPath, "processed")} 資料夾",
                                                   "處理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show($"未找到站點 {stop} 的有效圖片，請確認檔名格式是否為 \"x-{stop}\"", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"處理過程中發生錯誤:\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        lbAdd($"處理Chamfer檢測時發生錯誤: {ex.Message}", "err", "ChamferProcessError");
                    }
                }
            }
        }
        public struct CircleDetectionResult
        {
            public Mat ProcessedImage;
            public Point DetectedOuterCenter;
            public Point DetectedInnerCenter;
            public int DetectedOuterRadius;
            public int DetectedInnerRadius;
            public double CenterDistance;
        }
        private CircleDetectionResult testDetectAndExtractROI(Mat inputImage, int stop, int count)
        {
            CircleDetectionResult result = new CircleDetectionResult();
            try
            {
                // 開始計時整個方法的執行時間
                PerformanceProfiler.StartMeasure("DetectAndExtractROI_Total");

                // 記錄執行資訊的 StringBuilder
                StringBuilder logInfo = new StringBuilder();
                logInfo.AppendLine($"===== 圓形ROI偵測記錄 - 時間: {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")} =====");
                logInfo.AppendLine($"站點: {stop}, 樣品編號: {count}");

                // 區段1: 參數讀取和準備工作
                PerformanceProfiler.StartMeasure("Preprocessing");

                // 從 app.param 讀取霍夫圓參數 (加入防呆處理)
                int outerMinRadius = int.Parse(app.param[$"outer_minRadius_{stop}"]);
                int outerMaxRadius = int.Parse(app.param[$"outer_maxRadius_{stop}"]);
                int innerMinRadius = int.Parse(app.param[$"inner_minRadius_{stop}"]);
                int innerMaxRadius = int.Parse(app.param[$"inner_maxRadius_{stop}"]);
                int outerP1 = app.param.ContainsKey($"outer_p1_{stop}") ? int.Parse(app.param[$"outer_p1_{stop}"]) : 120;
                int outerP2 = app.param.ContainsKey($"outer_p2_{stop}") ? int.Parse(app.param[$"outer_p2_{stop}"]) : 20;
                int innerP1 = app.param.ContainsKey($"inner_p1_{stop}") ? int.Parse(app.param[$"inner_p1_{stop}"]) : 120;
                int innerP2 = app.param.ContainsKey($"inner_p2_{stop}") ? int.Parse(app.param[$"inner_p2_{stop}"]) : 20;
                int outerMinDist = app.param.ContainsKey($"outer_minDist_{stop}") ? int.Parse(app.param[$"outer_minDist_{stop}"]) : 50;
                int innerMinDist = app.param.ContainsKey($"inner_minDist_{stop}") ? int.Parse(app.param[$"inner_minDist_{stop}"]) : 50;
                //int shrink = app.param.ContainsKey($"shrink_{stop}") ? int.Parse(app.param[$"shrink_{stop}"]) : 0;

                // 從資料庫讀取預設圓心和半徑

                int knownOuterCenterX = int.Parse(app.param[$"known_outer_center_x_{stop}"]);
                int knownOuterCenterY = int.Parse(app.param[$"known_outer_center_y_{stop}"]);
                int knownInnerCenterX = int.Parse(app.param[$"known_inner_center_x_{stop}"]);
                int knownInnerCenterY = int.Parse(app.param[$"known_inner_center_y_{stop}"]);
                int knownOuterRadius = int.Parse(app.param[$"known_outer_radius_{stop}"]);
                int knownInnerRadius = int.Parse(app.param[$"known_inner_radius_{stop}"]);

                /*
                int knownOuterCenterX = 0;
                int knownOuterCenterY = 0;
                int knownInnerCenterX = 0;
                int knownInnerCenterY = 0;
                int knownOuterRadius = 0;
                int knownInnerRadius = 0;
                */
                logInfo.AppendLine($"參數設置:");
                logInfo.AppendLine($"  外圈半徑範圍: {outerMinRadius}-{outerMaxRadius}, 內圈半徑範圍: {innerMinRadius}-{innerMaxRadius}");
                logInfo.AppendLine($"  預設外圓圓心: ({knownOuterCenterX}, {knownOuterCenterY})");
                logInfo.AppendLine($"  預設內圓圓心: ({knownInnerCenterX}, {knownInnerCenterY})");
                logInfo.AppendLine($"  預設半徑: 外圈={knownOuterRadius}, 內圈={knownInnerRadius}");

                int centerTolerance = 50; // 圓心容許偏差
                Mat roi_gray = new Mat();

                // 區段1.1: 預處理 - 灰階和模糊化
                Mat gray = inputImage.CvtColor(ColorConversionCodes.BGR2GRAY);
                Mat blurred = gray.GaussianBlur(new Size(5, 5), 1);

                Cv2.Threshold(blurred, blurred, 115, 255, ThresholdTypes.Binary);

                PerformanceProfiler.StopMeasure("Preprocessing");

                // 區段2: 霍夫圓檢測 - 外圈
                PerformanceProfiler.StartMeasure("HoughCircles_Outer");

                var outerCircles = Cv2.HoughCircles(
                    blurred,
                    HoughModes.Gradient,
                    dp: 1,
                    minDist: outerMinDist,
                    param1: outerP1,
                    param2: outerP2,
                    minRadius: outerMinRadius,
                    maxRadius: outerMaxRadius);

                PerformanceProfiler.StopMeasure("HoughCircles_Outer");

                // 區段3: 霍夫圓檢測 - 內圈
                PerformanceProfiler.StartMeasure("HoughCircles_Inner");

                var innerCircles = Cv2.HoughCircles(
                    blurred,
                    HoughModes.Gradient,
                    dp: 1,
                    minDist: innerMinDist,
                    param1: innerP1,
                    param2: innerP2,
                    minRadius: innerMinRadius,
                    maxRadius: innerMaxRadius);

                PerformanceProfiler.StopMeasure("HoughCircles_Inner");

                // 區段4: 最佳圓選擇
                PerformanceProfiler.StartMeasure("FindBestCircles");

                // 篩選和驗證檢測到的圓
                CircleSegment? bestOuterCircle = null;
                CircleSegment? bestInnerCircle = null;
                double minCenterDistance = double.MaxValue;

                foreach (var outerCircle in outerCircles)
                {
                    foreach (var innerCircle in innerCircles)
                    {
                        // 計算圓心距離
                        double centerDistance = Math.Sqrt(Math.Pow(outerCircle.Center.X - innerCircle.Center.X, 2) +
                                                       Math.Pow(outerCircle.Center.Y - innerCircle.Center.Y, 2));

                        // 檢查圓心距離和半徑比
                        if (centerDistance <= centerTolerance && centerDistance < minCenterDistance)
                        {
                            double radiusRatio = outerCircle.Radius / innerCircle.Radius;
                            // 如果需要額外比較半徑比例: if (radiusRatio > 1.5 && radiusRatio < 3.0)
                            {
                                bestOuterCircle = outerCircle;
                                bestInnerCircle = innerCircle;
                                minCenterDistance = centerDistance;
                            }
                        }
                    }
                }

                if (bestOuterCircle == null || bestInnerCircle == null)
                {
                    throw new Exception("找不到有效的內外圓！");
                }

                // 檢測到的圓心和半徑
                Point detectedOuterCenter = new Point((int)bestOuterCircle.Value.Center.X, (int)bestOuterCircle.Value.Center.Y);
                int detectedOuterRadius = (int)bestOuterCircle.Value.Radius;
                Point detectedInnerCenter = new Point((int)bestInnerCircle.Value.Center.X, (int)bestInnerCircle.Value.Center.Y);
                int detectedInnerRadius = (int)bestInnerCircle.Value.Radius;

                // 設置結果結構的相關字段

                result.DetectedOuterCenter = detectedOuterCenter;
                result.DetectedInnerCenter = detectedInnerCenter;
                result.DetectedOuterRadius = detectedOuterRadius;
                result.DetectedInnerRadius = detectedInnerRadius;

                // 記錄檢測到的圓形信息
                logInfo.AppendLine("檢測到的圓形信息:");
                logInfo.AppendLine($"  外圈圓心: ({detectedOuterCenter.X}, {detectedOuterCenter.Y}), 半徑: {detectedOuterRadius}");
                logInfo.AppendLine($"  內圈圓心: ({detectedInnerCenter.X}, {detectedInnerCenter.Y}), 半徑: {detectedInnerRadius}");

                // 計算與預設值的偏差
                int outerCenterXDiff = Math.Abs(detectedOuterCenter.X - knownOuterCenterX);
                int outerCenterYDiff = Math.Abs(detectedOuterCenter.Y - knownOuterCenterY);
                int innerCenterXDiff = Math.Abs(detectedInnerCenter.X - knownInnerCenterX);
                int innerCenterYDiff = Math.Abs(detectedInnerCenter.Y - knownInnerCenterY);
                int outerRadiusDiff = Math.Abs(detectedOuterRadius - knownOuterRadius);
                int innerRadiusDiff = Math.Abs(detectedInnerRadius - knownInnerRadius);

                logInfo.AppendLine("與預設值的偏差:");
                logInfo.AppendLine($"  外圈圓心偏差: X={outerCenterXDiff}, Y={outerCenterYDiff}, 半徑偏差: {outerRadiusDiff}");
                logInfo.AppendLine($"  內圈圓心偏差: X={innerCenterXDiff}, Y={innerCenterYDiff}, 半徑偏差: {innerRadiusDiff}");

                // 圓心間距離
                double centerDistance0 = Math.Sqrt(Math.Pow(detectedOuterCenter.X - detectedInnerCenter.X, 2) +
                                                 Math.Pow(detectedOuterCenter.Y - detectedInnerCenter.Y, 2));
                logInfo.AppendLine($"  內外圓圓心距離: {centerDistance0:F2} 像素");
                result.CenterDistance = centerDistance0;

                PerformanceProfiler.StopMeasure("FindBestCircles");

                // 區段5: 平移矩陣建立
                PerformanceProfiler.StartMeasure("MatrixCreation");

                // 預設外圓圓心 (用於整體圖像平移)
                Point knownOuterCenter = new Point(knownOuterCenterX, knownOuterCenterY);

                // 預設內圓圓心 (用於特殊內環檢測模式)
                Point knownInnerCenter = new Point(knownInnerCenterX, knownInnerCenterY);

                // 計算平移向量 (從檢測圓心到預設圓心)
                int shiftX = knownOuterCenterX - detectedOuterCenter.X;
                int shiftY = knownOuterCenterY - detectedOuterCenter.Y;
                logInfo.AppendLine($"應用平移向量: ({shiftX}, {shiftY})");

                // 建立平移矩陣
                Mat translationMatrix = new Mat(2, 3, MatType.CV_64FC1);
                double[] translationData = new double[] { 1, 0, shiftX, 0, 1, shiftY };
                Marshal.Copy(translationData, 0, translationMatrix.Data, translationData.Length);

                PerformanceProfiler.StopMeasure("MatrixCreation");

                // 區段6: 圖像平移
                PerformanceProfiler.StartMeasure("WarpAffine");

                // 使用平移矩陣進行整個圖像平移
                Mat shiftedImage = new Mat();
                Cv2.WarpAffine(inputImage, shiftedImage, translationMatrix, inputImage.Size());

                PerformanceProfiler.StopMeasure("WarpAffine");

                // 區段7: 遮罩建立
                PerformanceProfiler.StartMeasure("MaskCreation");

                // 建立遮罩
                Mat mask = new Mat(shiftedImage.Size(), MatType.CV_8UC1, Scalar.Black);

                // 取實際檢測半徑
                int finalOuterRadius = detectedOuterRadius;
                int finalInnerRadius = detectedInnerRadius;
                int calculatedOuterRadius = 0;
                logInfo.AppendLine($"最終使用的半徑: 外圈={finalOuterRadius}, 內圈={finalInnerRadius}");

                // 判斷是否為特殊內環檢測模式
                bool isInnerHoughMode = false;

                if (app.param.ContainsKey($"innerHough_{stop}"))
                {
                    isInnerHoughMode = (int.Parse(app.param[$"innerHough_{stop}"]) == 1);
                }

                if (isInnerHoughMode)
                {
                    // 特殊內環檢測模式 - 內圓變成白色遮罩，且外援用內圓的延伸
                    int offsetY = app.param.ContainsKey($"innerHoughOffsetY_{stop}") ?
                                  int.Parse(app.param[$"innerHoughOffsetY_{stop}"]) : 0;
                    int roiRadius = app.param.ContainsKey($"innerHoughRoiRadius_{stop}") ?
                                   int.Parse(app.param[$"innerHoughRoiRadius_{stop}"]) : 0;
                    //int roiRadius = 170;

                    calculatedOuterRadius = finalInnerRadius + roiRadius;
                    logInfo.AppendLine($"內環檢測模式: offsetY={offsetY}, roiRadius={roiRadius}, calculatedOuterRadius={calculatedOuterRadius}");

                    //外圓算法不同，進行覆蓋，內圓跟其他一樣
                    result.DetectedOuterCenter = detectedInnerCenter;
                    result.DetectedOuterCenter.Y = detectedInnerCenter.Y + offsetY;
                    result.DetectedOuterRadius = calculatedOuterRadius;
                    // 繪製白色圓形遮罩
                    Cv2.Circle(mask, knownInnerCenter, calculatedOuterRadius, Scalar.White, -1);

                    //Cv2.Circle(mask, knownInnerCenter, calculatedOuterRadius, Scalar.Red, 2);
                    //Cv2.Circle(mask, knownInnerCenter, finalInnerRadius, Scalar.Red, 2);
                }
                else

                {
                    // 標準模式 - 外圓白色，內圓黑色
                    Cv2.Circle(mask, knownOuterCenter, finalOuterRadius, Scalar.White, -1);
                    Cv2.Circle(mask, knownInnerCenter, finalInnerRadius, Scalar.Black, -1);
                }


                PerformanceProfiler.StopMeasure("MaskCreation");

                // 區段8: 遮罩應用
                PerformanceProfiler.StartMeasure("MaskApplication");

                // 應用遮罩
                Mat roi_full = new Mat();
                Cv2.BitwiseAnd(shiftedImage, shiftedImage, roi_full, mask);

                // 轉灰階
                Cv2.CvtColor(roi_full, roi_gray, ColorConversionCodes.BGR2GRAY);

                PerformanceProfiler.StopMeasure("MaskApplication");

                // 區段9: 後處理
                PerformanceProfiler.StartMeasure("Postprocessing");

                // 根據站點進行不同的後處理

                if (stop == 1 || stop == 2)
                {
                    int contrastOffset = app.param.ContainsKey($"deepenContrast_{stop}") ?
                                        int.Parse(app.param[$"deepenContrast_{stop}"]) : 30;
                    int brightnessOffset = app.param.ContainsKey($"deepenBrightness_{stop}") ?
                                          int.Parse(app.param[$"deepenBrightness_{stop}"]) : 30;

                    // 調整對比度和亮度
                    roi_gray = AdjustContrast(roi_gray, contrastOffset, brightnessOffset);
                }

                PerformanceProfiler.StopMeasure("Postprocessing");

                // 區段10: 顏色轉換
                PerformanceProfiler.StartMeasure("ColorConversion");

                // 將灰階圖轉換回彩色圖
                Mat roi_color = new Mat();
                Cv2.CvtColor(roi_gray, roi_color, ColorConversionCodes.GRAY2BGR);

                PerformanceProfiler.StopMeasure("ColorConversion");

                // 區段11: 儲存影像
                //PerformanceProfiler.StartMeasure("SaveImage");

                // 儲存處理後的影像
                string outputPath = Path.Combine(@"C:\Users\User\Desktop\4\ROI", $"ROI_{count}.png");
                Cv2.ImWrite(outputPath, roi_color);

                //PerformanceProfiler.StopMeasure("SaveImage");

                // 停止計時整個方法的執行時間
                PerformanceProfiler.StopMeasure("DetectAndExtractROI_Total");

                // 返回處理後的影像
                result.ProcessedImage = roi_color;

                return result;
            }
            catch (Exception ex)
            {
                Log.Error($"testDetectAndExtractROI 錯誤: {ex.Message}");
                return result;
            }
        }
        /// <summary>
        /// 創建並儲存白色像素占比直方圖
        /// </summary>
        /// <param name="ratioValues">白色像素占比列表</param>
        /// <param name="folderPath">儲存資料夾路徑</param>
        /// <param name="stop">站點編號</param>
        private void CreateHistogram(List<double> ratioValues, string folderPath, int stop)
        {
            try
            {
                // 確定直方圖的參數
                double min = Math.Floor(ratioValues.Min());
                double max = Math.Ceiling(ratioValues.Max());
                int binCount = 20; // 直方圖的柱子數量
                double binWidth = (max - min) / binCount;

                // 創建柱子計數陣列
                int[] binCounts = new int[binCount];

                // 統計每個柱子中的值的數量
                foreach (var ratio in ratioValues)
                {
                    int binIndex = (int)((ratio - min) / binWidth);
                    if (binIndex >= binCount) binIndex = binCount - 1; // 防止越界
                    if (binIndex < 0) binIndex = 0;
                    binCounts[binIndex]++;
                }

                // 找出最大計數以確定直方圖高度
                int maxCount = binCounts.Max();

                // 創建直方圖圖像
                int width = 800;
                int height = 600;
                int margin = 60;
                int graphHeight = height - 2 * margin;
                int graphWidth = width - 2 * margin;

                // 創建白色背景
                Mat histImage = new Mat(height, width, MatType.CV_8UC3, new Scalar(255, 255, 255));

                // 繪製座標軸
                Cv2.Line(histImage, new Point(margin, height - margin), new Point(width - margin, height - margin), new Scalar(0, 0, 0), 2); // X軸
                Cv2.Line(histImage, new Point(margin, margin), new Point(margin, height - margin), new Scalar(0, 0, 0), 2); // Y軸

                // 繪製標題
                Cv2.PutText(
                    histImage,
                    $"stop {stop} white pixels ratio histogram",
                    new Point(width / 2 - 150, 30),
                    HersheyFonts.HersheyDuplex,
                    0.8,
                    new Scalar(0, 0, 0),
                    2,
                    LineTypes.AntiAlias
                );

                // 繪製 X 軸標籤
                Cv2.PutText(
                    histImage,
                    "white pixels ratio (%)",
                    new Point(width / 2 - 80, height - 10),
                    HersheyFonts.HersheyDuplex,
                    0.7,
                    new Scalar(0, 0, 0),
                    1,
                    LineTypes.AntiAlias
                );

                // 繪製 Y 軸標籤
                Cv2.PutText(
                    histImage,
                    "count",
                    new Point(10, height / 2),
                    HersheyFonts.HersheyDuplex,
                    0.7,
                    new Scalar(0, 0, 0),
                    1,
                    LineTypes.AntiAlias
                );

                // 繪製 X 軸刻度和標籤
                int xTickCount = 10;
                double xTickStep = (max - min) / xTickCount;
                for (int i = 0; i <= xTickCount; i++)
                {
                    double value = min + i * xTickStep;
                    int x = margin + (int)(i * graphWidth / xTickCount);

                    // 繪製刻度線
                    Cv2.Line(histImage, new Point(x, height - margin), new Point(x, height - margin + 5), new Scalar(0, 0, 0), 1);

                    // 繪製標籤
                    Cv2.PutText(
                        histImage,
                        $"{value:F1}",
                        new Point(x - 15, height - margin + 25),
                        HersheyFonts.HersheyDuplex,
                        0.5,
                        new Scalar(0, 0, 0),
                        1,
                        LineTypes.AntiAlias
                    );
                }

                // 繪製 Y 軸刻度和標籤
                int yTickCount = 10;
                for (int i = 0; i <= yTickCount; i++)
                {
                    int value = i * maxCount / yTickCount;
                    int y = height - margin - (int)(i * graphHeight / yTickCount);

                    // 繪製刻度線
                    Cv2.Line(histImage, new Point(margin - 5, y), new Point(margin, y), new Scalar(0, 0, 0), 1);

                    // 繪製標籤
                    Cv2.PutText(
                        histImage,
                        $"{value}",
                        new Point(margin - 40, y + 5),
                        HersheyFonts.HersheyDuplex,
                        0.5,
                        new Scalar(0, 0, 0),
                        1,
                        LineTypes.AntiAlias
                    );

                    // 繪製網格線 (虛線)
                    if (i > 0 && i < yTickCount)
                    {
                        for (int dashX = margin + 5; dashX < width - margin; dashX += 10)
                        {
                            Cv2.Line(histImage, new Point(dashX, y), new Point(dashX + 5, y), new Scalar(200, 200, 200), 1);
                        }
                    }
                }

                // 繪製柱子
                for (int i = 0; i < binCount; i++)
                {
                    double binLeft = min + i * binWidth;
                    int x1 = margin + (int)(graphWidth * (binLeft - min) / (max - min));
                    int x2 = margin + (int)(graphWidth * (binLeft + binWidth - min) / (max - min));
                    int y = height - margin - (int)((double)binCounts[i] / maxCount * graphHeight);

                    // 隨機顏色的柱子，但使用藍色漸變
                    Scalar barColor = new Scalar(200, 100 + i * 5, 50 + i * 10);

                    // 填充柱子
                    Cv2.Rectangle(histImage, new Point(x1, height - margin), new Point(x2, y), barColor, -1);

                    // 繪製柱子邊框
                    Cv2.Rectangle(histImage, new Point(x1, height - margin), new Point(x2, y), new Scalar(0, 0, 0), 1);

                    // 如果柱子夠寬，顯示計數
                    if (x2 - x1 > 20 && binCounts[i] > 0)
                    {
                        Cv2.PutText(
                            histImage,
                            $"{binCounts[i]}",
                            new Point(x1 + (x2 - x1) / 2 - 10, y - 10),
                            HersheyFonts.HersheyDuplex,
                            0.5,
                            new Scalar(0, 0, 0),
                            1,
                            LineTypes.AntiAlias
                        );
                    }
                }

                // 繪製統計信息
                double avg = ratioValues.Average();
                double median = ratioValues.Count % 2 == 0
                    ? (ratioValues[ratioValues.Count / 2 - 1] + ratioValues[ratioValues.Count / 2]) / 2
                    : ratioValues[ratioValues.Count / 2];
                double stdDev = Math.Sqrt(ratioValues.Average(v => Math.Pow(v - avg, 2)));

                Cv2.PutText(
                    histImage,
                    $"sample count: {ratioValues.Count}, average: {avg:F2}%, median: {median:F2}%, standard deviation: {stdDev:F2}%",
                    new Point(margin, 55),
                    HersheyFonts.HersheyDuplex,
                    0.6,
                    new Scalar(0, 0, 0),
                    1,
                    LineTypes.AntiAlias
                );

                // 畫出平均線和中位數線
                int avgX = margin + (int)(graphWidth * (avg - min) / (max - min));
                int medianX = margin + (int)(graphWidth * (median - min) / (max - min));

                // 平均線 (紅色虛線)
                for (int y = margin; y < height - margin; y += 5)
                {
                    Cv2.Line(histImage, new Point(avgX, y), new Point(avgX, y + 3), new Scalar(0, 0, 255), 2);
                }

                // 中位數線 (綠色虛線)
                for (int y = margin; y < height - margin; y += 5)
                {
                    Cv2.Line(histImage, new Point(medianX, y), new Point(medianX, y + 3), new Scalar(0, 255, 0), 2);
                }

                // 添加平均值和中位數標籤
                Cv2.PutText(
                    histImage,
                    "average",
                    new Point(avgX + 5, margin + 15),
                    HersheyFonts.HersheyDuplex,
                    0.5,
                    new Scalar(0, 0, 255),
                    1,
                    LineTypes.AntiAlias
                );

                Cv2.PutText(
                    histImage,
                    "median",
                    new Point(medianX + 5, margin + 30),
                    HersheyFonts.HersheyDuplex,
                    0.5,
                    new Scalar(0, 255, 0),
                    1,
                    LineTypes.AntiAlias
                );

                // 保存直方圖
                string histogramPath = Path.Combine(folderPath, $"Histogram_Station{stop}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.png");
                Cv2.ImWrite(histogramPath, histImage);

                // 釋放資源
                histImage.Dispose();
            }
            catch (Exception ex)
            {
                lbAdd($"創建直方圖時發生錯誤: {ex.Message}", "err", "HistogramError");
            }
        }
        // 統計數據的類別
        private class ChamferStatistic
        {
            public string FileName { get; set; }
            public int WhitePixelCount { get; set; }
            public int TotalPixels { get; set; }
            public double Ratio { get; set; }
        }
        public string GetAppFolderName()
        {
            return app.foldername;
        }
        // 由 GitHub Copilot 產生
        // 修正: 內部 Clone 確保 Queue_Save 擁有獨立副本，避免呼叫方提早釋放導致 ObjectDisposedException
        public void SaveImageAsync(Mat image, string path)
        {
            app.Queue_Save.Enqueue(new ImageSave(image.Clone(), path)); // ✅ Clone 在這裡
            app._sv.Set();
        }

        private int CalculateTotalAreaWithoutOverlap(List<Rect> rects)
        {
            if (rects == null || rects.Count == 0)
                return 0;

            if (rects.Count == 1)
                return rects[0].Width * rects[0].Height;

            // 使用遮罩來計算精確的聯合面積
            var allRects = rects.ToList();
            var boundingBox = GetBoundingBox(allRects);

            // 創建遮罩
            Mat mask = new Mat(boundingBox.Height, boundingBox.Width, MatType.CV_8UC1, Scalar.Black);

            foreach (var rect in allRects)
            {
                // 調整矩形相對於邊界框的位置
                Rect adjustedRect = new Rect(
                    rect.X - boundingBox.X,
                    rect.Y - boundingBox.Y,
                    rect.Width,
                    rect.Height
                );

                // 在遮罩上畫白色矩形
                Cv2.Rectangle(mask, adjustedRect, Scalar.White, -1);
            }

            // 計算白色像素數量（即總面積）
            int totalArea = Cv2.CountNonZero(mask);

            mask.Dispose();
            return totalArea;
        }

        private Rect GetBoundingBox(List<Rect> rects)
        {
            if (rects.Count == 0)
                return new Rect();

            int minX = rects.Min(r => r.X);
            int minY = rects.Min(r => r.Y);
            int maxX = rects.Max(r => r.X + r.Width);
            int maxY = rects.Max(r => r.Y + r.Height);

            return new Rect(minX, minY, maxX - minX, maxY - minY);
        }
        private void ShowAlert()
        {
            if (alertForm == null || alertForm.IsDisposed)
            {
                this.Enabled = false;
                app.plc_stop = true;

                alertForm = new alert();
                alertForm.FormClosed += (s, e) => {
                    alertForm = null;
                    app.plc_stop = false;
                    this.Enabled = true;
                    app.alertTriggered = false; // 視窗關閉時重置標記
                    // 自動呼叫停止檢測
                    //button2_Click(button2, EventArgs.Empty);
                };
                alertForm.Show();
                alertForm.BringToFront();
            }
        }

        #region 色彩分析
        // 色彩分析結果類別
        public class ColorAnalysisResult
        {
            public string FileName { get; set; }
            public string BatchLabel { get; set; }

            // RGB統計
            public double[] RgbMean { get; set; } = new double[3];
            public double[] RgbStdDev { get; set; } = new double[3];
            public int[,] RgbHistogram { get; set; } = new int[3, 256];

            // HSV統計
            public double[] HsvMean { get; set; } = new double[3];
            public double[] HsvStdDev { get; set; } = new double[3];
            public int[,] HsvHistogram { get; set; } = new int[3, 256];

            // Lab統計
            public double[] LabMean { get; set; } = new double[3];
            public double[] LabStdDev { get; set; } = new double[3];

            // 自定義特徵
            public double ColorfulnessIndex { get; set; }  // 色彩豐富度指數
            public double DominantHue { get; set; }        // 主導色調
            public double Saturation { get; set; }        // 平均飽和度
            public double Brightness { get; set; }        // 平均亮度
            public int ValidPixelCount { get; set; }      // 有效像素數量（非黑色）
        }

        // 分析單張圖片的色彩特徵
        private ColorAnalysisResult AnalyzeImageColors(Mat image, string fileName, string batchLabel)
        {
            try
            {
                var result = new ColorAnalysisResult
                {
                    FileName = fileName,
                    BatchLabel = batchLabel
                };

                // 創建遮罩，排除黑色背景像素 (RGB值都小於10的像素)
                Mat mask = new Mat(image.Size(), MatType.CV_8UC1);
                Mat grayTemp = new Mat();
                Cv2.CvtColor(image, grayTemp, ColorConversionCodes.BGR2GRAY);
                Cv2.Threshold(grayTemp, mask, 10, 255, ThresholdTypes.Binary);

                // 計算有效像素數量
                result.ValidPixelCount = Cv2.CountNonZero(mask);

                if (result.ValidPixelCount < 100) // 如果有效像素太少，跳過
                {
                    mask.Dispose();
                    grayTemp.Dispose();
                    return null;
                }

                // 轉換到不同色彩空間
                Mat hsvImage = new Mat();
                Mat labImage = new Mat();
                Cv2.CvtColor(image, hsvImage, ColorConversionCodes.BGR2HSV);
                Cv2.CvtColor(image, labImage, ColorConversionCodes.BGR2Lab);

                // 分析RGB
                AnalyzeColorSpace(image, mask, result.RgbMean, result.RgbStdDev, result.RgbHistogram);

                // 分析HSV
                AnalyzeColorSpace(hsvImage, mask, result.HsvMean, result.HsvStdDev, result.HsvHistogram);

                // 分析Lab
                AnalyzeLabColorSpace(labImage, mask, result.LabMean, result.LabStdDev);

                // 計算自定義特徵
                CalculateCustomFeatures(image, hsvImage, mask, result);

                // 釋放資源
                mask.Dispose();
                grayTemp.Dispose();
                hsvImage.Dispose();
                labImage.Dispose();

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"分析圖片 {fileName} 色彩時發生錯誤: {ex.Message}");
                return null;
            }
        }

        // 分析色彩空間統計
        private void AnalyzeColorSpace(Mat image, Mat mask, double[] means, double[] stdDevs, int[,] histograms)
        {
            Mat[] channels = image.Split();

            for (int c = 0; c < 3; c++)
            {
                // 計算均值和標準差
                Scalar meanScalar, stdDevScalar;
                Cv2.MeanStdDev(channels[c], out meanScalar, out stdDevScalar, mask);
                means[c] = meanScalar.Val0;
                stdDevs[c] = stdDevScalar.Val0;

                // 計算直方圖
                Mat hist = new Mat();
                int[] histSize = { 256 };
                Rangef[] ranges = { new Rangef(0, 256) };
                Cv2.CalcHist(new[] { channels[c] }, new[] { 0 }, mask, hist, 1, histSize, ranges);

                for (int i = 0; i < 256; i++)
                {
                    histograms[c, i] = (int)hist.Get<float>(i);
                }

                hist.Dispose();
                channels[c].Dispose();
            }
        }

        // 分析Lab色彩空間
        private void AnalyzeLabColorSpace2(Mat labImage, Mat mask, double[] means, double[] stdDevs)
        {
            Mat[] channels = labImage.Split();

            for (int c = 0; c < 3; c++)
            {
                Scalar meanScalar, stdDevScalar;
                Cv2.MeanStdDev(channels[c], out meanScalar, out stdDevScalar, mask);
                means[c] = meanScalar.Val0;
                stdDevs[c] = stdDevScalar.Val0;
                channels[c].Dispose();
            }
        }

        // 計算自定義特徵
        private void CalculateCustomFeatures(Mat bgrImage, Mat hsvImage, Mat mask, ColorAnalysisResult result)
        {
            Mat[] hsvChannels = hsvImage.Split();

            // 計算平均飽和度和亮度
            Scalar satMean, valMean, satStdDev, valStdDev;
            Cv2.MeanStdDev(hsvChannels[1], out satMean, out satStdDev, mask);
            Cv2.MeanStdDev(hsvChannels[2], out valMean, out valStdDev, mask);

            result.Saturation = satMean.Val0;
            result.Brightness = valMean.Val0;

            // 計算色彩豐富度指數 (基於飽和度和亮度的變化)
            result.ColorfulnessIndex = satStdDev.Val0 + valStdDev.Val0;

            // 計算主導色調 (Hue直方圖的峰值)
            Mat hueHist = new Mat();
            int[] histSize = { 180 };
            Rangef[] ranges = { new Rangef(0, 180) };
            Cv2.CalcHist(new[] { hsvChannels[0] }, new[] { 0 }, mask, hueHist, 1, histSize, ranges);

            Point minLoc, maxLoc;
            double minVal, maxVal;
            Cv2.MinMaxLoc(hueHist, out minVal, out maxVal, out minLoc, out maxLoc);
            result.DominantHue = maxLoc.Y * 2; // 轉換回0-360度

            // 釋放資源
            foreach (var channel in hsvChannels)
                channel.Dispose();
            hueHist.Dispose();
        }

        // 生成色彩分析報告
        private void GenerateColorAnalysisReport(List<ColorAnalysisResult> results, string outputDir, string batchLabel)
        {
            try
            {
                string reportPath = Path.Combine(outputDir, $"ColorAnalysisReport_{batchLabel}.txt");

                using (StreamWriter writer = new StreamWriter(reportPath, false, Encoding.UTF8))
                {
                    writer.WriteLine($"=== 五彩鋅電鍍色彩分析報告 ===");
                    writer.WriteLine($"批次標籤: {batchLabel}");
                    writer.WriteLine($"分析時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    writer.WriteLine($"圖片數量: {results.Count}");
                    writer.WriteLine();

                    // 統計摘要
                    writer.WriteLine("=== 統計摘要 ===");

                    // RGB統計
                    var rgbRMeans = results.Select(r => r.RgbMean[2]).ToList(); // Red
                    var rgbGMeans = results.Select(r => r.RgbMean[1]).ToList(); // Green  
                    var rgbBMeans = results.Select(r => r.RgbMean[0]).ToList(); // Blue

                    writer.WriteLine($"RGB平均值:");
                    writer.WriteLine($"  R通道: 均值={rgbRMeans.Average():F2}, 標準差={CalculateStdDev(rgbRMeans):F2}");
                    writer.WriteLine($"  G通道: 均值={rgbGMeans.Average():F2}, 標準差={CalculateStdDev(rgbGMeans):F2}");
                    writer.WriteLine($"  B通道: 均值={rgbBMeans.Average():F2}, 標準差={CalculateStdDev(rgbBMeans):F2}");
                    writer.WriteLine();

                    // HSV統計
                    var saturations = results.Select(r => r.Saturation).ToList();
                    var brightnesses = results.Select(r => r.Brightness).ToList();
                    var dominantHues = results.Select(r => r.DominantHue).ToList();

                    writer.WriteLine($"HSV特徵:");
                    writer.WriteLine($"  飽和度: 均值={saturations.Average():F2}, 標準差={CalculateStdDev(saturations):F2}");
                    writer.WriteLine($"  亮度: 均值={brightnesses.Average():F2}, 標準差={CalculateStdDev(brightnesses):F2}");
                    writer.WriteLine($"  主導色調: 均值={dominantHues.Average():F2}°, 標準差={CalculateStdDev(dominantHues):F2}°");
                    writer.WriteLine();

                    // 自定義特徵
                    var colorfulnessIndices = results.Select(r => r.ColorfulnessIndex).ToList();
                    writer.WriteLine($"色彩豐富度指數: 均值={colorfulnessIndices.Average():F2}, 標準差={CalculateStdDev(colorfulnessIndices):F2}");
                    writer.WriteLine();

                    // 詳細數據
                    writer.WriteLine("=== 詳細數據 ===");
                    writer.WriteLine("檔名,R均值,G均值,B均值,飽和度,亮度,主導色調,色彩豐富度,有效像素數");

                    foreach (var result in results)
                    {
                        writer.WriteLine($"{result.FileName}," +
                                       $"{result.RgbMean[2]:F2}," +
                                       $"{result.RgbMean[1]:F2}," +
                                       $"{result.RgbMean[0]:F2}," +
                                       $"{result.Saturation:F2}," +
                                       $"{result.Brightness:F2}," +
                                       $"{result.DominantHue:F2}," +
                                       $"{result.ColorfulnessIndex:F2}," +
                                       $"{result.ValidPixelCount}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成色彩分析報告時發生錯誤: {ex.Message}");
            }
        }

        // 生成色彩直方圖
        private void GenerateColorHistograms(List<ColorAnalysisResult> results, string outputDir, string batchLabel)
        {
            try
            {
                // 合併所有圖片的直方圖數據
                int[,] combinedRgbHist = new int[3, 256];
                int[,] combinedHsvHist = new int[3, 256];

                foreach (var result in results)
                {
                    for (int c = 0; c < 3; c++)
                    {
                        for (int i = 0; i < 256; i++)
                        {
                            combinedRgbHist[c, i] += result.RgbHistogram[c, i];
                            combinedHsvHist[c, i] += result.HsvHistogram[c, i];
                        }
                    }
                }

                // 創建RGB直方圖
                CreateHistogramImage(combinedRgbHist,
                                   new string[] { "Blue", "Green", "Red" },
                                   new Scalar[] { new Scalar(255, 0, 0), new Scalar(0, 255, 0), new Scalar(0, 0, 255) },
                                   Path.Combine(outputDir, $"RGB_Histogram_{batchLabel}.png"),
                                   $"RGB Color Histogram - {batchLabel}");

                // 創建HSV直方圖
                CreateHistogramImage(combinedHsvHist,
                                   new string[] { "Hue", "Saturation", "Value" },
                                   new Scalar[] { new Scalar(255, 255, 0), new Scalar(255, 0, 255), new Scalar(0, 255, 255) },
                                   Path.Combine(outputDir, $"HSV_Histogram_{batchLabel}.png"),
                                   $"HSV Color Histogram - {batchLabel}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成直方圖時發生錯誤: {ex.Message}");
            }
        }

        // 創建直方圖圖像
        private void CreateHistogramImage(int[,] histData, string[] channelNames, Scalar[] colors,
                                        string outputPath, string title)
        {
            int width = 800;
            int height = 600;
            int margin = 80;
            int graphHeight = height - 2 * margin;
            int graphWidth = width - 2 * margin;

            Mat histImage = new Mat(height, width, MatType.CV_8UC3, new Scalar(255, 255, 255));

            // 找出最大值用於歸一化
            int maxVal = 0;
            for (int c = 0; c < 3; c++)
            {
                for (int i = 0; i < 256; i++)
                {
                    maxVal = Math.Max(maxVal, histData[c, i]);
                }
            }

            if (maxVal == 0) maxVal = 1; // 避免除以零

            // 繪製座標軸
            Cv2.Line(histImage, new Point(margin, height - margin),
                     new Point(width - margin, height - margin), new Scalar(0, 0, 0), 2);
            Cv2.Line(histImage, new Point(margin, margin),
                     new Point(margin, height - margin), new Scalar(0, 0, 0), 2);

            // 繪製標題
            Cv2.PutText(histImage, title, new Point(width / 2 - 150, 30),
                        HersheyFonts.HersheyDuplex, 0.8, new Scalar(0, 0, 0), 2);

            // 繪製直方圖
            for (int c = 0; c < 3; c++)
            {
                for (int i = 0; i < 256; i++)
                {
                    int x = margin + (i * graphWidth / 256);
                    int y = height - margin - (int)((double)histData[c, i] / maxVal * graphHeight);

                    if (i > 0)
                    {
                        int prevX = margin + ((i - 1) * graphWidth / 256);
                        int prevY = height - margin - (int)((double)histData[c, i - 1] / maxVal * graphHeight);
                        Cv2.Line(histImage, new Point(prevX, prevY), new Point(x, y), colors[c], 2);
                    }
                }

                // 添加圖例
                Cv2.Line(histImage, new Point(width - 150, 60 + c * 25),
                         new Point(width - 120, 60 + c * 25), colors[c], 3);
                Cv2.PutText(histImage, channelNames[c], new Point(width - 110, 65 + c * 25),
                            HersheyFonts.HersheyDuplex, 0.6, new Scalar(0, 0, 0), 1);
            }

            Cv2.ImWrite(outputPath, histImage);
            histImage.Dispose();
        }
        #endregion

        // 匯出統計數據到CSV
        private void ExportColorStatisticsToCSV(List<ColorAnalysisResult> results, string csvPath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    // 寫入標題行
                    writer.WriteLine("檔名,批次標籤,R均值,G均值,B均值,R標準差,G標準差,B標準差," +
                                   "H均值,S均值,V均值,H標準差,S標準差,V標準差," +
                                   "L均值,A均值,B均值(Lab),L標準差,A標準差,B標準差(Lab)," +
                                   "飽和度,亮度,主導色調,色彩豐富度指數,有效像素數");

                    // 寫入數據行
                    foreach (var result in results)
                    {
                        writer.WriteLine($"{result.FileName}," +
                                       $"{result.BatchLabel}," +
                                       $"{result.RgbMean[2]:F3}," + // R
                                       $"{result.RgbMean[1]:F3}," + // G
                                       $"{result.RgbMean[0]:F3}," + // B
                                       $"{result.RgbStdDev[2]:F3}," +
                                       $"{result.RgbStdDev[1]:F3}," +
                                       $"{result.RgbStdDev[0]:F3}," +
                                       $"{result.HsvMean[0]:F3}," + // H
                                       $"{result.HsvMean[1]:F3}," + // S
                                       $"{result.HsvMean[2]:F3}," + // V
                                       $"{result.HsvStdDev[0]:F3}," +
                                       $"{result.HsvStdDev[1]:F3}," +
                                       $"{result.HsvStdDev[2]:F3}," +
                                       $"{result.LabMean[0]:F3}," + // L
                                       $"{result.LabMean[1]:F3}," + // A
                                       $"{result.LabMean[2]:F3}," + // B
                                       $"{result.LabStdDev[0]:F3}," +
                                       $"{result.LabStdDev[1]:F3}," +
                                       $"{result.LabStdDev[2]:F3}," +
                                       $"{result.Saturation:F3}," +
                                       $"{result.Brightness:F3}," +
                                       $"{result.DominantHue:F3}," +
                                       $"{result.ColorfulnessIndex:F3}," +
                                       $"{result.ValidPixelCount}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"匯出CSV時發生錯誤: {ex.Message}");
            }
        }

        // 計算標準差的輔助函數
        private double CalculateStdDev(List<double> values)
        {
            if (values.Count == 0) return 0;
            double mean = values.Average();
            return Math.Sqrt(values.Average(v => Math.Pow(v - mean, 2)));
        }

        /// <summary>
        /// 擴展矩形框，確保不超出影像邊界
        /// </summary>
        /// <param name="rect">原始矩形</param>
        /// <param name="expandPixels">擴展像素數</param>
        /// <param name="imageWidth">影像寬度</param>
        /// <param name="imageHeight">影像高度</param>
        /// <returns>擴展後的矩形</returns>
        private Rect ExpandRect(Rect rect, int expandPixels, int imageWidth, int imageHeight)
        {
            int newX = Math.Max(0, rect.X - expandPixels);
            int newY = Math.Max(0, rect.Y - expandPixels);
            int newWidth = Math.Min(imageWidth - newX, rect.Width + 2 * expandPixels);
            int newHeight = Math.Min(imageHeight - newY, rect.Height + 2 * expandPixels);

            return new Rect(newX, newY, newWidth, newHeight);
        }

        #region 圓心校正工具
        private void ShowCircleCalibrationTool()
        {
            if (string.IsNullOrEmpty(app.produce_No))
            {
                MessageBox.Show("請先設定料號", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var calibrationTool = new CircleCalibrationForm())
            {
                calibrationTool.ShowDialog();
            }
        }
        #endregion
        #region 提示視窗
        private void ShowStopDetectionNotification()
        {
            try
            {
                // 使用新的 CustomMessageBox 替代原本的 MessageBox
                CustomMessageBox.Show(
                    "檢測已停止，請把當前使用中的包裝箱內的樣品倒回入料口，並點擊\"更新計數\"。",
                    "停止通知",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
                );
            }
            catch (Exception ex)
            {
                Log.Error($"顯示停止檢測通知時發生錯誤: {ex.Message}");
                // 備用方案：如果 CustomMessageBox 失敗，使用原本的 MessageBox
                MessageBox.Show("檢測已停止", "停止通知", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void ShowRecoverySuccessNotification()
        {
            try
            {
                CustomMessageBox.Show(
                    "復歸成功，可點擊\"開始檢測\"。",
                    "異常復歸完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    new System.Drawing.Font("微軟正黑體", 16F, System.Drawing.FontStyle.Bold)
                );

                // 記錄到日誌
                lbAdd("已顯示復歸成功通知", "inf", "");
            }
            catch (Exception ex)
            {
                lbAdd("顯示復歸成功通知時發生錯誤", "err", ex.Message);
            }
        }
        // 由 GitHub Copilot 產生
        private void ShowCountdownDialog(int seconds)
        {
            // 計算初始文字所需的尺寸
            string initialText = $"復歸中... {seconds}";
            Font countdownFont = new Font("微軟正黑體", 20F, FontStyle.Bold);

            // 創建倒數對話框
            var countdownForm = new Form()
            {
                Text = "復歸中",
                StartPosition = FormStartPosition.CenterScreen, // 改為螢幕中央
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                AutoSize = false
            };

            // 使用 Graphics 計算文字所需空間
            using (Graphics g = countdownForm.CreateGraphics())
            {
                System.Drawing.SizeF textSize = g.MeasureString(initialText, countdownFont);

                // 設定適當的視窗大小，包含邊距
                int margin = 40;
                int formWidth = Math.Max(350, (int)Math.Ceiling(textSize.Width) + margin * 2);
                int formHeight = Math.Max(120, (int)Math.Ceiling(textSize.Height) + margin * 2);

                countdownForm.Size = new System.Drawing.Size(formWidth, formHeight);

                // 創建 Label
                var lblCountdown = new Label()
                {
                    Location = new System.Drawing.Point(margin, margin),
                    Size = new System.Drawing.Size(formWidth - margin * 2, formHeight - margin * 2),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = countdownFont,
                    ForeColor = Color.Red,
                    AutoSize = false
                };

                countdownForm.Controls.Add(lblCountdown);
                countdownForm.Show();
                countdownForm.BringToFront();

                try
                {
                    // 倒數計時邏輯
                    for (int i = seconds; i >= 0; i--)
                    {
                        if (countdownForm.IsDisposed) break;

                        // 更新文字內容
                        string currentText = i > 0 ? $"復歸中... {i}" : "復歸完成";
                        lblCountdown.Text = currentText;

                        // 重新計算文字大小，動態調整視窗大小
                        System.Drawing.SizeF currentTextSize = g.MeasureString(currentText, countdownFont);

                        int newFormWidth = Math.Max(350, (int)Math.Ceiling(currentTextSize.Width) + margin * 2);
                        int newFormHeight = Math.Max(120, (int)Math.Ceiling(currentTextSize.Height) + margin * 2);

                        // 只有當尺寸需要變化時才調整，並保持中央位置
                        if (countdownForm.Width != newFormWidth || countdownForm.Height != newFormHeight)
                        {
                            // 計算新的中央位置
                            Screen screen = Screen.PrimaryScreen;
                            int newX = (screen.WorkingArea.Width - newFormWidth) / 2;
                            int newY = (screen.WorkingArea.Height - newFormHeight) / 2;

                            countdownForm.Size = new System.Drawing.Size(newFormWidth, newFormHeight);
                            countdownForm.Location = new System.Drawing.Point(newX, newY);
                            lblCountdown.Size = new System.Drawing.Size(newFormWidth - margin * 2, newFormHeight - margin * 2);
                        }

                        countdownForm.Refresh();
                        Application.DoEvents(); // 讓UI有機會更新

                        if (i > 0)
                        {
                            Thread.Sleep(1000); // 等待1秒
                        }
                    }

                    // 顯示完成訊息1秒後關閉
                    Thread.Sleep(1000);
                }
                finally
                {
                    if (!countdownForm.IsDisposed)
                    {
                        countdownForm.Close();
                        countdownForm.Dispose();
                    }
                }
            }
        }
        // 由 GitHub Copilot 產生
        private void UpdateButtonStates()
        {
            BeginInvoke(new Action(() =>
            {
                // 由 GitHub Copilot 產生
                // 調機模式下不限制按鈕，所有按鈕都可正常使用
                if (app.DetectMode == 1)
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button47.Enabled = true;
                    button17.Enabled = true;

                    // 保持調機模式的文字顯示
                    button1.Text = "開始調機";
                    button2.Text = "停止調機";

                    return; // 調機模式下直接返回，不執行後續的狀態限制
                }

                switch (app.currentState)
                {
                    case app.SystemState.Stopped: //復歸完
                        button1.Enabled = true;   // 可以開始檢測
                        button2.Enabled = false;  // 不能停止（沒在運行）
                        button47.Enabled = true; // 不能更新計數 -> 改成也可以按
                        button17.Enabled = true; // 不能復歸（沒需要）->改成可以一直復歸

                        button1.Text = "開始檢測";
                        button1.BackColor = System.Drawing.Color.FromArgb(128, 255, 128); // 綠色
                        break;

                    case app.SystemState.Running: //開始完
                        button1.Enabled = false;  // 不能開始（已在運行）
                        button2.Enabled = true;   // 可以停止檢測
                        button47.Enabled = false; // 不能更新計數（在運行中）
                        button17.Enabled = false; // 不能復歸（在運行中）

                        button2.Text = "停止檢測";
                        button2.BackColor = System.Drawing.Color.FromArgb(255, 128, 128); // 紅色
                        break;

                    case app.SystemState.StoppedNeedUpdate: //停止完
                        button1.Enabled = false;  // 不能開始（需先更新計數）
                        button2.Enabled = false;  // 不能停止（已停止）
                        button47.Enabled = true;  // 可以更新計數
                        button17.Enabled = false; // 不能復歸（需先更新計數）

                        button47.BackColor = System.Drawing.Color.FromArgb(255, 255, 128); // 黃色提醒
                        break;

                    case app.SystemState.UpdatedNeedReset: //更新完
                        button1.Enabled = false;  // 不能開始（需先復歸）
                        button2.Enabled = false;  // 不能停止（已停止）
                        button47.Enabled = false; // 不能更新計數（已更新）
                        button17.Enabled = true;  // 可以復歸

                        button17.BackColor = System.Drawing.Color.FromArgb(255, 165, 0); // 橘色提醒
                        break;
                }
            }));
        }
        #endregion
        #region 取資料庫參數
        /// <summary>
        /// 由 GitHub Copilot 產生
        /// 安全取得 int 參數，避免格式錯誤造成例外
        /// </summary>
        public static int GetIntParam(IDictionary<string, string> param, string key, int defaultValue)
        {
            if (param != null && param.TryGetValue(key, out var s) && int.TryParse(s, out var v))
                return v;
            return defaultValue;
        }

        /// <summary>
        /// 由 GitHub Copilot 產生
        /// 安全取得 double 參數，避免格式錯誤造成例外
        /// </summary>
        public static double GetDoubleParam(IDictionary<string, string> param, string key, double defaultValue)
        {
            if (param != null && param.TryGetValue(key, out var s) && double.TryParse(s, out var v))
                return v;
            return defaultValue;
        }
        #endregion
        private void ClearImageQueues()
        {
            // 清空佇列並釋放 Mat 物件
            while (app.Queue_Bitmap1.TryDequeue(out var img1)) img1?.image?.Dispose();
            while (app.Queue_Bitmap2.TryDequeue(out var img2)) img2?.image?.Dispose();
            while (app.Queue_Bitmap3.TryDequeue(out var img3)) img3?.image?.Dispose();
            while (app.Queue_Bitmap4.TryDequeue(out var img4)) img4?.image?.Dispose();
            while (app.Queue_Save.TryDequeue(out var imgSave)) imgSave?.image?.Dispose();

            Log.Information("復歸前已清空所有影像佇列");
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. 讓使用者選取資料夾
                string folderPath = "";
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "請選擇包含圖片的資料夾";
                    if (Directory.Exists(@".\image"))
                        dialog.SelectedPath = @".\image";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = dialog.SelectedPath;
                    }
                    else
                    {
                        MessageBox.Show("未選取資料夾，操作取消。");
                        return;
                    }
                }

                // 2. 檢查目錄是否存在
                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show($"指定的目錄不存在：{folderPath}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. 獲取所有支援的圖片檔案
                string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif" };
                List<string> allImageFiles = new List<string>();

                foreach (string extension in imageExtensions)
                {
                    allImageFiles.AddRange(Directory.GetFiles(folderPath, extension, SearchOption.TopDirectoryOnly));
                }

                if (allImageFiles.Count == 0)
                {
                    MessageBox.Show($"在目錄中找不到圖片檔案：{folderPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 4. 確認對話框
                DialogResult confirmResult = MessageBox.Show(
                    $"確定要對資料夾內的 {allImageFiles.Count} 個圖片檔案進行黑點檢測嗎？\n\n" +
                    $"資料夾：{folderPath}\n" +
                    $"站點：第3站",
                    "確認批次檢測",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2
                );

                if (confirmResult != DialogResult.Yes)
                {
                    return;
                }

                // 5. 創建輸出資料夾
                string outputFolder = Path.Combine(folderPath, "BlackDotResults");
                FolderJob(outputFolder, false, true);

                // 6. 顯示進度視窗並處理圖片
                int processed = 0;
                int detectedCount = 0;
                int okCount = 0;

                using (var progressForm = new Form())
                using (var progressBar = new ProgressBar())
                using (var label = new Label())
                {
                    progressForm.Text = "黑點檢測進度";
                    progressForm.Width = 450;
                    progressForm.Height = 120;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.StartPosition = FormStartPosition.CenterParent;
                    progressForm.ControlBox = false;
                    progressForm.TopMost = true;

                    progressBar.Minimum = 0;
                    progressBar.Maximum = allImageFiles.Count;
                    progressBar.Dock = DockStyle.Top;
                    progressBar.Height = 30;

                    label.Dock = DockStyle.Fill;
                    label.TextAlign = ContentAlignment.MiddleCenter;
                    label.Text = "開始檢測黑點...";

                    progressForm.Controls.Add(label);
                    progressForm.Controls.Add(progressBar);

                    progressForm.Show(this);
                    progressForm.BringToFront();
                    Application.DoEvents();

                    // 處理每個圖片檔案
                    foreach (string filePath in allImageFiles)
                    {
                        try
                        {
                            // 更新進度顯示
                            label.Text = $"正在處理: {Path.GetFileName(filePath)}\n已處理: {processed}/{allImageFiles.Count}";
                            progressBar.Value = processed;
                            Application.DoEvents();

                            // 讀取圖片
                            Mat inputImage = Cv2.ImRead(filePath);
                            if (inputImage == null || inputImage.Empty())
                            {
                                Log.Warning($"無法讀取圖片: {filePath}");
                                continue;
                            }

                            // 呼叫 DetectBlackDots 函數
                            int stop = 3; // 第3站
                            int count = processed; // 使用處理順序作為 count

                            var (hasBlackDots, resultImage, detectedContours) = DetectBlackDots(inputImage, stop, count);

                            // 統計結果
                            if (hasBlackDots)
                            {
                                detectedCount++;
                            }
                            else
                            {
                                okCount++;
                            }

                            // 儲存結果圖像
                            if (resultImage != null && !resultImage.Empty())
                            {
                                string resultFileName = hasBlackDots
                                    ? $"{Path.GetFileNameWithoutExtension(filePath)}_NG.png"
                                    : $"{Path.GetFileNameWithoutExtension(filePath)}_OK.png";

                                string resultPath = Path.Combine(outputFolder, resultFileName);
                                Cv2.ImWrite(resultPath, resultImage);

                                resultImage.Dispose();
                            }

                            // 釋放資源
                            inputImage.Dispose();

                            processed++;
                        }
                        catch (Exception ex)
                        {
                            Log.Error($"處理圖片時發生錯誤: {filePath}, 錯誤: {ex.Message}");
                        }
                    }

                    progressForm.Close();
                }

                // 7. 生成檢測報告
                string reportPath = Path.Combine(outputFolder, "detection_report.txt");
                StringBuilder report = new StringBuilder();
                report.AppendLine("===== 黑點檢測報告 =====");
                report.AppendLine($"檢測時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                report.AppendLine($"資料夾: {folderPath}");
                report.AppendLine($"站點: 第3站");
                report.AppendLine($"總處理數量: {processed}");
                report.AppendLine($"檢測到黑點: {detectedCount}");
                report.AppendLine($"無黑點: {okCount}");
                report.AppendLine($"檢出率: {(processed > 0 ? (detectedCount * 100.0 / processed).ToString("F2") : "0")}%");
                report.AppendLine($"結果儲存位置: {outputFolder}");
                report.AppendLine("========================");

                File.WriteAllText(reportPath, report.ToString(), Encoding.UTF8);

                // 8. 顯示完成訊息
                string resultMessage = $"黑點檢測完成！\n\n" +
                                      $"總處理數量：{processed} 個檔案\n" +
                                      $"檢測到黑點：{detectedCount} 個\n" +
                                      $"無黑點：{okCount} 個\n" +
                                      $"檢出率：{(processed > 0 ? (detectedCount * 100.0 / processed).ToString("F2") : "0")}%\n\n" +
                                      $"結果已儲存至：\n{outputFolder}\n\n" +
                                      $"報告檔案：{reportPath}";

                MessageBox.Show(resultMessage, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 9. 詢問是否開啟結果資料夾
                DialogResult openFolder = MessageBox.Show(
                    "是否要開啟結果資料夾？",
                    "開啟資料夾",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (openFolder == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("explorer.exe", outputFolder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"批次黑點檢測過程發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log.Error($"批次黑點檢測錯誤: {ex.Message}");
            }
        }

        private void button49_Click(object sender, EventArgs e)
        {
            if (button49.Text == "正常檢測")
            {
                button49.Text = "全進OK";
                button49.BackColor = Color.FromArgb(255, 128, 128);
                app.allOK = true;
                Console.WriteLine("true");
            }
            else if (button49.Text == "全進OK")
            {
                button49.Text = "正常檢測";
                button49.BackColor = Color.FromArgb(128, 255, 128);
                app.allOK = false;
                Console.WriteLine("false");
            }
        }

    }
    /// <summary>
    /// 管理瑕疵計數的寫入、讀取和更新
    /// </summary>
    // 在 #region 介面事件 的最後面，新增一個專門的 region

    #region 自訂訊息框
    public static class CustomMessageBox
    {
        public static DialogResult Show(string message, string title = "",
            MessageBoxButtons buttons = MessageBoxButtons.OK,
            MessageBoxIcon icon = MessageBoxIcon.None,
            Font font = null,
            bool topMost = false,
            Form owner = null) // ← 加上 owner 參數
        {
            Form messageBoxForm = new Form();
            Label messageLabel = new Label();
            Button button1 = new Button();
            Button button2 = new Button();
            Button button3 = new Button();

            // 設定字體
            if (font == null)
            {
                font = new Font("微軟正黑體", 16F, FontStyle.Regular);
            }

            try
            {
                // 計算文字所需的尺寸
                using (Graphics g = messageBoxForm.CreateGraphics())
                {
                    // 設定最大寬度限制
                    int maxTextWidth = 500; // 增加寬度
                    SizeF textSize = g.MeasureString(message, font, maxTextWidth);

                    // 計算實際需要的寬度和高度，並增加額外空間
                    int textWidth = (int)Math.Ceiling(textSize.Width) + 20; // 增加寬度緩衝
                    int textHeight = (int)Math.Ceiling(textSize.Height) + 30; // 增加高度緩衝

                    // 設定邊距
                    int margin = 40;
                    int buttonHeight = 50;
                    int buttonMargin = 30; // 增加按鈕邊距

                    // 計算視窗尺寸
                    int formWidth = Math.Max(400, textWidth + margin * 2); // 最小寬度400
                    int formHeight = textHeight + buttonHeight + margin * 2 + buttonMargin;

                    // 設定Form屬性
                    messageBoxForm.Text = title;
                    messageBoxForm.Size = new System.Drawing.Size(formWidth, formHeight);
                    messageBoxForm.StartPosition = FormStartPosition.CenterParent;
                    messageBoxForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    messageBoxForm.MaximizeBox = false;
                    messageBoxForm.MinimizeBox = false;
                    messageBoxForm.TopMost = topMost; // ← 設定置頂

                    // ← 根據是否有 owner 決定 StartPosition
                    if (owner != null)
                    {
                        messageBoxForm.StartPosition = FormStartPosition.CenterParent;
                    }
                    else
                    {
                        messageBoxForm.StartPosition = FormStartPosition.CenterScreen;
                    }

                    // 設定Label屬性
                    messageLabel.Text = message;
                    messageLabel.Font = font;
                    messageLabel.Location = new System.Drawing.Point(margin, margin);
                    messageLabel.Size = new System.Drawing.Size(formWidth - margin * 2, textHeight);
                    messageLabel.TextAlign = ContentAlignment.TopCenter; // 改為頂部置中，避免文字被截斷
                    messageLabel.AutoSize = false;

                    // 計算按鈕位置
                    int buttonY = margin + textHeight + buttonMargin;

                    // 根據按鈕類型設定按鈕
                    switch (buttons)
                    {
                        case MessageBoxButtons.OK:
                            button1.Text = "確定";
                            button1.DialogResult = DialogResult.OK;
                            button1.Size = new System.Drawing.Size(100, 40);
                            button1.Font = new System.Drawing.Font("微軟正黑體", 12F, FontStyle.Regular);
                            button1.Location = new System.Drawing.Point(
                                (formWidth - button1.Width) / 2,
                                buttonY
                            );
                            messageBoxForm.Controls.Add(button1);
                            messageBoxForm.AcceptButton = button1;
                            break;

                        case MessageBoxButtons.YesNo:
                            button1.Text = "是";
                            button1.DialogResult = DialogResult.Yes;
                            button1.Size = new System.Drawing.Size(80, 40);
                            button1.Font = new Font("微軟正黑體", 12F, FontStyle.Regular);

                            button2.Text = "否";
                            button2.DialogResult = DialogResult.No;
                            button2.Size = new System.Drawing.Size(80, 40);
                            button2.Font = new Font("微軟正黑體", 12F, FontStyle.Regular);

                            // 計算兩個按鈕的位置（置中排列）
                            int totalButtonWidth = button1.Width + button2.Width + 20;
                            int startX = (formWidth - totalButtonWidth) / 2;

                            button1.Location = new System.Drawing.Point(startX, buttonY);
                            button2.Location = new System.Drawing.Point(startX + button1.Width + 20, buttonY);

                            messageBoxForm.Controls.Add(button1);
                            messageBoxForm.Controls.Add(button2);
                            messageBoxForm.AcceptButton = button1;
                            break;

                        case MessageBoxButtons.YesNoCancel:
                            button1.Text = "是";
                            button1.DialogResult = DialogResult.Yes;
                            button1.Size = new System.Drawing.Size(80, 40);
                            button1.Font = new Font("微軟正黑體", 12F, FontStyle.Regular);

                            button2.Text = "否";
                            button2.DialogResult = DialogResult.No;
                            button2.Size = new System.Drawing.Size(80, 40);
                            button2.Font = new Font("微軟正黑體", 12F, FontStyle.Regular);

                            button3.Text = "取消";
                            button3.DialogResult = DialogResult.Cancel;
                            button3.Size = new System.Drawing.Size(80, 40);
                            button3.Font = new Font("微軟正黑體", 12F, FontStyle.Regular);

                            // 計算三個按鈕的位置（置中排列）
                            int totalButton3Width = button1.Width + button2.Width + button3.Width + 40;
                            int start3X = (formWidth - totalButton3Width) / 2;

                            button1.Location = new System.Drawing.Point(start3X, buttonY);
                            button2.Location = new System.Drawing.Point(start3X + button1.Width + 20, buttonY);
                            button3.Location = new System.Drawing.Point(start3X + button1.Width + button2.Width + 40, buttonY);

                            messageBoxForm.Controls.Add(button1);
                            messageBoxForm.Controls.Add(button2);
                            messageBoxForm.Controls.Add(button3);
                            messageBoxForm.AcceptButton = button1;
                            messageBoxForm.CancelButton = button3;
                            break;
                    }

                    messageBoxForm.Controls.Add(messageLabel);
                }
            }
            catch (Exception ex)
            {
                // 如果發生錯誤，使用預設大小
                messageBoxForm.Size = new System.Drawing.Size(450, 200); // 增加預設大小
                messageLabel.Location = new System.Drawing.Point(20, 20);
                messageLabel.Size = new System.Drawing.Size(400, 100); // 增加Label高度
                messageLabel.TextAlign = ContentAlignment.TopCenter;
                Console.WriteLine($"CustomMessageBox 設定時發生錯誤: {ex.Message}");
            }

            return messageBoxForm.ShowDialog();
        }
    }
    #endregion
    #region 計數
    public static class DefectCountManager
    {
        private static readonly object _dbLock = new object();

        /// <summary>
        /// 寫入單筆瑕疵計數記錄到資料庫
        /// </summary>
        /// <param name="type">產品料號</param>
        /// <param name="name">瑕疵名稱</param>
        /// <param name="count">瑕疵數量</param>
        /// <param name="lotId">批號 (可選)</param>
        /// <returns>寫入是否成功</returns>
        public static bool WriteDefectCount(string type, string name, int count, string lotId = null)
        {
            try
            {
                lock (_dbLock)
                {
                    using (var db = new MydbDB())
                    {
                        db.DefectCounts
                          .Value(p => p.Type, type)
                          .Value(p => p.Name, name)
                          .Value(p => p.Count, count)
                          .Value(p => p.Time, DateTime.Now)
                          .Value(p => p.LotId, lotId)
                          .Insert();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] 寫入 DefectCount 失敗: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 批量寫入多筆瑕疵計數記錄，使用同一個時間戳
        /// </summary>
        /// <param name="items">要寫入的記錄清單，每筆包含 (type, name, count, lotId)</param>
        /// <param name="forceWrite">是否強制寫入（不檢查重複），用於停止時的最終寫入</param>
        /// <returns>寫入是否成功</returns>
        public static bool WriteDefectCounts(List<(string Type, string Name, int Count, string LotId)> items, bool forceWrite = false)
        {
            if (items == null || items.Count == 0)
            {
                Log.Warning("WriteDefectCounts: items 為 null 或空，無法寫入");
                return false;
            }

            try
            {
                DateTime now = DateTime.Now;

                // 由 GitHub Copilot 產生
                // 診斷：記錄寫入開始
                Log.Information("===== WriteDefectCounts 開始 =====");
                Log.Information($"準備寫入 {items.Count} 筆資料，forceWrite={forceWrite}");

                lock (_dbLock)
                {
                    using (var db = new MydbDB())
                    {
                        int insertedCount = 0;
                        int skippedCount = 0;

                        foreach (var item in items)
                        {
                            // 由 GitHub Copilot 產生
                            Log.Debug($"處理項目: Type='{item.Type}', Name='{item.Name}', Count={item.Count}, LotId='{item.LotId}'");

                            if (forceWrite)
                            {
                                // 強制寫入模式（停止時使用）：直接寫入，不檢查重複
                                db.DefectCounts
                                  .Value(p => p.Type, item.Type)
                                  .Value(p => p.Name, item.Name)
                                  .Value(p => p.Count, item.Count)
                                  .Value(p => p.Time, now)
                                  .Value(p => p.LotId, item.LotId)
                                  .Insert();

                                insertedCount++;
                                Log.Information($"  [強制寫入] {item.Name} = {item.Count}");
                            }
                            else
                            {
                                // 【修正】檢查是否已存在相同 Type, Name, LotId 且時間在 10 秒內的記錄
                                var existingCount = db.DefectCounts
                                    .Count(p => p.Type == item.Type &&
                                               p.Name == item.Name &&
                                               p.LotId == item.LotId &&
                                               p.Time >= now.AddSeconds(-10));

                                if (existingCount == 0)
                                {
                                    // 不存在，才寫入
                                    db.DefectCounts
                                      .Value(p => p.Type, item.Type)
                                      .Value(p => p.Name, item.Name)
                                      .Value(p => p.Count, item.Count)
                                      .Value(p => p.Time, now)
                                      .Value(p => p.LotId, item.LotId)
                                      .Insert();

                                    insertedCount++;
                                    Log.Information($"  [新增] {item.Name} = {item.Count}");
                                }
                                else
                                {
                                    skippedCount++;
                                    Log.Debug($"  [跳過] {item.Name} = {item.Count} (10秒內已存在 {existingCount} 筆相同記錄)");
                                }
                            }
                        }

                        // 由 GitHub Copilot 產生
                        // 診斷：記錄寫入統計
                        Log.Information($"寫入完成: 新增 {insertedCount} 筆, 跳過 {skippedCount} 筆");
                        Log.Information("===== WriteDefectCounts 結束 =====");
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error($"[Error] 批量寫入 DefectCount 失敗: {ex.Message}");
                Log.Error($"例外堆疊: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 從 app.dc 字典寫入所有瑕疵計數資料和統計資料 (OK/NG/NULL)
        /// </summary>
        /// <param name="type">產品料號</param>
        /// <param name="lotId">批號</param>
        /// <param name="defectCounts">瑕疵計數字典</param>
        /// <param name="generalCounts">一般計數字典 (OK/NG/NULL)</param>
        /// <param name="forceWrite">是否強制寫入（不檢查重複），用於停止時的最終寫入</param>
        /// <returns>寫入是否成功</returns>
        public static bool WriteAllDefectCounts(string type, string lotId,
    IDictionary<string, int> defectCounts,
    IDictionary<string, int> generalCounts = null,
    bool forceWrite = false)
        {
            try
            {
                // 由 GitHub Copilot 產生
                // 診斷：記錄函數呼叫參數
                Log.Information("===== WriteAllDefectCounts 開始 =====");
                Log.Information($"參數: type='{type}', lotId='{lotId}', forceWrite={forceWrite}");

                var itemsToWrite = new List<(string Type, string Name, int Count, string LotId)>();

                // 由 GitHub Copilot 產生
                // 診斷：記錄瑕疵計數處理
                Log.Information($"處理 defectCounts (共 {defectCounts?.Count ?? 0} 筆)：");

                // 加入各類瑕疵計數（排除 SAMPLE_ID，因為會單獨處理）
                if (defectCounts != null)
                {
                    foreach (var item in defectCounts)
                    {
                        // 【修正】跳過 SAMPLE_ID，避免重複寫入
                        if (item.Key != "SAMPLE_ID")
                        {
                            // 由 GitHub Copilot 產生
                            // 診斷：檢查是否誤將 OK/NG/NULL 加入 app.dc
                            if (item.Key == "OK" || item.Key == "NG" || item.Key == "NULL")
                            {
                                Log.Warning($"  [警告] 在 defectCounts 中發現一般計數項目: '{item.Key}' = {item.Value}");
                                Log.Warning($"  [警告] 此項目應該在 generalCounts 中，而非 defectCounts (app.dc)");
                                // 由 GitHub Copilot 產生
                                // 決策：跳過此項目，不寫入（因為會在 generalCounts 處理）
                                continue;
                            }

                            itemsToWrite.Add((type, item.Key, item.Value, lotId));
                            Log.Information($"  加入瑕疵: '{item.Key}' = {item.Value}");
                        }
                        else
                        {
                            Log.Debug($"  跳過 SAMPLE_ID (單獨處理)");
                        }
                    }
                }

                // 由 GitHub Copilot 產生
                // 診斷：記錄一般計數處理
                Log.Information($"處理 generalCounts (共 {generalCounts?.Count ?? 0} 筆)：");

                // 如果有一般計數，也加入
                if (generalCounts != null)
                {
                    foreach (var item in generalCounts)
                    {
                        if (item.Key == "OK" || item.Key == "NG" || item.Key == "NULL")
                        {
                            itemsToWrite.Add((type, item.Key, item.Value, lotId));
                            Log.Information($"  加入一般計數: '{item.Key}' = {item.Value}");
                        }
                        else
                        {
                            Log.Warning($"  [警告] generalCounts 中發現非一般計數項目: '{item.Key}' = {item.Value}");
                        }
                    }
                }

                // 由 GitHub Copilot 產生
                // 診斷：記錄最終要寫入的項目
                Log.Information($"準備寫入 {itemsToWrite.Count} 筆資料到資料庫");

                bool result = WriteDefectCounts(itemsToWrite, forceWrite);

                Log.Information($"WriteDefectCounts 執行結果: {(result ? "成功" : "失敗")}");
                Log.Information("===== WriteAllDefectCounts 結束 =====");

                return result;
            }
            catch (Exception ex)
            {
                Log.Error($"[Error] 批量寫入 DefectCount 失敗: {ex.Message}");
                Log.Error($"例外堆疊: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 從資料庫讀取特定料號和批號的最新瑕疵計數
        /// </summary>
        /// <param name="type">產品料號</param>
        /// <param name="lotId">批號 (可選)</param>
        /// <returns>瑕疵計數字典，若無資料則回傳空字典</returns>
        public static Dictionary<string, int> ReadLatestDefectCounts(string type, string lotId = null)
        {
            var result = new Dictionary<string, int>();

            try
            {
                using (var db = new MydbDB())
                {
                    // 可以根據需要調整查詢邏輯，例如「每種瑕疵的最後一筆記錄」
                    var query = db.DefectCounts.AsQueryable()
                                  .Where(dc => dc.Type == type);

                    if (!string.IsNullOrEmpty(lotId))
                    {
                        query = query.Where(dc => dc.LotId == lotId);
                    }

                    // 按名稱分組，取每組中時間最晚的一筆
                    var latestRecords = query
                                        .AsEnumerable()
                                        .GroupBy(dc => dc.Name)
                                        .Select(g => g.OrderByDescending(dc => dc.Time).First());

                    foreach (var record in latestRecords)
                    {
                        result[record.Name] = record.Count;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] 讀取 DefectCount 失敗: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 定期寫入 DefectCount 到資料庫，可呼叫於定時任務中
        /// </summary>
        /// <param name="forceWrite">是否強制寫入（按停止時使用）</param>
        /// <returns>寫入是否成功</returns>
        public static bool PerformPeriodicWrite(bool forceWrite = false)
        {
            try
            {
                // 由 GitHub Copilot 產生
                // 診斷：記錄 app.dc 的完整內容
                Log.Information("===== PerformPeriodicWrite 開始 =====");
                Log.Information($"forceWrite = {forceWrite}");
                Log.Information($"app.produce_No = {app.produce_No}");
                Log.Information($"app.LotID = {app.LotID}");

                if (app.dc == null)
                {
                    Log.Warning("app.dc 為 null，無法寫入");
                    return false;
                }

                if (app.dc.Count == 0)
                {
                    Log.Warning("app.dc 為空，無法寫入");
                    return false;
                }

                if (string.IsNullOrEmpty(app.produce_No))
                {
                    Log.Warning("app.produce_No 為空，無法寫入");
                    return false;
                }

                // 由 GitHub Copilot 產生
                // 診斷：逐項列出 app.dc 的內容
                Log.Information($"app.dc 共有 {app.dc.Count} 筆資料：");
                foreach (var item in app.dc.OrderBy(x => x.Key))
                {
                    Log.Information($"  瑕疵名稱: '{item.Key}', 數量: {item.Value}");
                }

                // 從 PLC 讀取計數
                int ngCount = Form1.PLC_CheckD(801);
                int okCount = Form1.PLC_CheckD(807);
                int nullCount = Form1.PLC_CheckD(809);

                // 由 GitHub Copilot 產生
                // 診斷：記錄從 PLC 讀取的數值
                Log.Information($"從 PLC 讀取的計數：");
                Log.Information($"  D801 (NG) = {ngCount}");
                Log.Information($"  D807 (OK原始) = {okCount}");
                Log.Information($"  D809 (NULL) = {nullCount}");

                // 修正：將 OK 數量取整到 50 的倍數（向下取整）
                int q = okCount / 50;
                int adjustedOkCount = q * 50;

                // 由 GitHub Copilot 產生
                // 診斷：記錄 OK 數量調整
                if (okCount != adjustedOkCount)
                {
                    Log.Information($"OK 數量調整: {okCount} -> {adjustedOkCount} (取整到50的倍數)");
                }

                // 準備一般計數字典(OK/NG/NULL)
                Dictionary<string, int> generalCounts = new Dictionary<string, int>
                {
                    { "OK", adjustedOkCount },
                    { "NG", ngCount },
                    { "NULL", nullCount }
                };

                // 由 GitHub Copilot 產生
                // 診斷：記錄 generalCounts 內容
                Log.Information($"generalCounts 內容：");
                foreach (var item in generalCounts)
                {
                    Log.Information($"  {item.Key} = {item.Value}");
                }

                // 【新增】寫入最終 SAMPLE_ID（停止時）
                if (forceWrite)
                {
                    Log.Information("執行強制寫入 (停止時)");

                    int currentSampleId = 0;
                    if (ResultManager.counter.TryGetValue("SAMPLE_ID", out currentSampleId))
                    {
                        // 由 GitHub Copilot 產生
                        Log.Information($"當前 SAMPLE_ID = {currentSampleId}");

                        try
                        {
                            using (var db = new MydbDB())
                            {
                                int updatedRows = db.DefectCounts
                                    .Where(p => p.Type == app.produce_No &&
                                               p.Name == "SAMPLE_ID" &&
                                               p.LotId == app.LotID)
                                    .Set(p => p.Count, currentSampleId)
                                    .Set(p => p.Time, DateTime.Now)
                                    .Update();

                                if (updatedRows == 0)
                                {
                                    db.DefectCounts
                                        .Value(p => p.Type, app.produce_No)
                                        .Value(p => p.Name, "SAMPLE_ID")
                                        .Value(p => p.Count, currentSampleId)
                                        .Value(p => p.Time, DateTime.Now)
                                        .Value(p => p.LotId, app.LotID)
                                        .Insert();

                                    Log.Information($"停止時插入最終 SAMPLE_ID = {currentSampleId}");
                                }
                                else
                                {
                                    Log.Information($"停止時更新最終 SAMPLE_ID = {currentSampleId} (更新了 {updatedRows} 筆)");
                                }
                            }
                        }
                        catch (Exception exSample)
                        {
                            Log.Error($"寫入最終 SAMPLE_ID 失敗: {exSample.Message}");
                        }
                    }
                    else
                    {
                        Log.Warning("無法從 ResultManager.counter 取得 SAMPLE_ID");
                    }
                }

                // 由 GitHub Copilot 產生
                // 呼叫寫入函數前記錄
                Log.Information("準備呼叫 WriteAllDefectCounts...");
                bool result = WriteAllDefectCounts(app.produce_No, app.LotID, app.dc, generalCounts, forceWrite);

                Log.Information($"WriteAllDefectCounts 執行結果: {(result ? "成功" : "失敗")}");
                Log.Information("===== PerformPeriodicWrite 結束 =====");

                return result;
            }
            catch (Exception ex)
            {
                Log.Error($"[Error] 定期寫入 DefectCount 失敗: {ex.Message}");
                Log.Error($"例外堆疊: {ex.StackTrace}");
                return false;
            }
        }
    }
    #endregion

}

# region ClientApp
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

public class DetectionResponse
{
    public List<DetectionResult> detections { get; set; }
    public string error { get; set; }
}


#endregion

#region 檢測結果管理
public class ResultManager
{
    //using peilin;

    public static int totalStations = 0; //這邊依模型數量動態加 LINE1820
    private static ConcurrentDictionary<int, SampleResult> results = new ConcurrentDictionary<int, SampleResult>();

    public static ConcurrentDictionary<string, int> counter = new ConcurrentDictionary<string, int>();

    // 擴展 sampleResults 結構，增加 timeout 欄位
    private static ConcurrentDictionary<int, (bool isNG, string defectName, float? score, double? timeout)> sampleResults =
        new ConcurrentDictionary<int, (bool isNG, string defectName, float? score, double? timeout)>();

    #region 統計用
    // 新增：站點與瑕疵統計字典
    private static ConcurrentDictionary<int, ConcurrentDictionary<string, int>> stationDefectStats =
        new ConcurrentDictionary<int, ConcurrentDictionary<string, int>>();

    // 新增：站點的 OK/NG 總計
    private static ConcurrentDictionary<int, (int ok, int ng)> stationOkNgStats =
        new ConcurrentDictionary<int, (int ok, int ng)>();

    // 樣品ID -> 站點ID -> StationResult
    private static ConcurrentDictionary<int, Dictionary<int, StationResult>> sampleStationResults =
    new ConcurrentDictionary<int, Dictionary<int, StationResult>>();

    // 初始化統計資料
    public static void InitializeStats()
    {
        stationDefectStats.Clear();
        stationOkNgStats.Clear();

        // 預設四個站點
        for (int i = 1; i <= 4; i++)
        {
            stationDefectStats[i] = new ConcurrentDictionary<string, int>();
            stationOkNgStats[i] = (0, 0);
        }
    }
    #endregion

    static ResultManager()
    {
        // 初始化 counter 字典
        counter.TryAdd("OK", 0);
        counter.TryAdd("OK1", 0);
        counter.TryAdd("OK2", 0);
        counter.TryAdd("SAMPLE_ID", 0);
        counter.TryAdd("NG", 0);
        counter.TryAdd("NULL", 0);
        counter.TryAdd("stop0", 0);
        counter.TryAdd("stop1", 0);
        counter.TryAdd("stop2", 0);
        counter.TryAdd("stop3", 0);
    }
    private static readonly object okLock = new object();
    private static string activeOkCounter = "OK1"; //一開始先從OK1 之後都在函數內變動

    public void AddResult(int sampleId, StationResult stationResult) //stationResult是一整個結構，包含所有結果
    {
        // 獲取或創建 SampleResult
        var sampleResult = results.GetOrAdd(sampleId, new SampleResult(sampleId));

        // 添加站點结果
        sampleResult.AddStationResult(stationResult);

        // 檢查是否收集到所有站點的结果
        if (sampleResult.IsComplete(totalStations))
        {
            // 結果統合
            CombineResults(sampleResult);

            // 從字典中移除已處理的樣品
            results.TryRemove(sampleId, out _);
        }
    }

    private void CombineResults(SampleResult sampleResult)
    {
        // 檢查是否有任何站點結果被標記為 NULL
        bool hasNullResult = sampleResult.StationResults.Values.Any(r => r.DefectName == "NULL_invalid");

        if (hasNullResult)
        {
            // 如果有 NULL 結果，整個樣品視為 NULL
            // 但仍然繼續處理並儲存結果
            Log.Warning($"樣品 {sampleResult.SampleId} 有無效取像，將標記為 NULL_invalid");

            CalculateAndSendPLCSignal(sampleResult, false, isNull: true);

            return;
        }

        // 如果沒有 NULL 結果，按原來邏輯處理
        bool finalIsNG = sampleResult.StationResults.Values.Any(r => r.IsNG);

        //Console.WriteLine($"樣品 {sampleResult.SampleId} 的最終結果：{(finalIsNG ? "NG" : "OK")}");

        // 發送 PLC 信號
        CalculateAndSendPLCSignal(sampleResult, finalIsNG, isNull: false);
    }

    private void SaveFinalResult(SampleResult sampleResult, bool isNG, bool isNull, double timeout = 0.0)
    {
        try
        {
            // 儲存影像和紀錄 - 以下邏輯需要根據 NULL 狀態調整

            // 新增部分 - 將所有站點結果添加到統計管理器
            if (!isNull) // NULL 樣品可能不需要統計分數
            {
                foreach (var stationResult in sampleResult.StationResults.Values)
                {
                    // 只對有分數的結果進行統計
                    if (stationResult.OkNgScore.HasValue)
                    {
                        ScoreStatisticsManager.AddScore(
                            stationResult.Stop,
                            stationResult.OkNgScore.Value,
                            stationResult.IsNG,
                            stationResult.DefectName
                        );
                    }
                }
            }

            var timestamp = DateTime.Now;
            string basePath = Path.Combine(
                @".\image",
                timestamp.ToString("yyyy-MM"),
                timestamp.ToString("MMdd"),
                app.foldername
            );

            // 決定結果資料夾
            string resultFolder;
            if (isNull)
                resultFolder = "NULL";
            else if (isNG)
                resultFolder = "NG";
            else
                resultFolder = "OK";

            string saveDir = Path.Combine(basePath, resultFolder);
            if (!Directory.Exists(saveDir)) Directory.CreateDirectory(saveDir);

            // 新增功能：依據站點存放圖片到各自對應資料夾
            // 檢查是否需要儲存 Stations 結果
            bool shouldSaveStations = false;
            using (var db = new MydbDB())
            {
                var param = db.Parameters.Where(p => p.Name == "saveStations").FirstOrDefault();
                shouldSaveStations = param?.Value == "true";
            }

            if (shouldSaveStations)
            {
                string stationBasePath = Path.Combine(basePath, "Stations");
                if (!Directory.Exists(stationBasePath)) Directory.CreateDirectory(stationBasePath);

                foreach (var stationResult in sampleResult.StationResults.Values)
                {
                    string fname = "";
                    if (isNull)
                    {
                        fname = $"{sampleResult.SampleId}-{stationResult.Stop}-{stationResult.DefectName}-{stationResult.OkNgScore:F2}.jpg"; //DefectName 是  NULL_invalid
                    }
                    else if (!stationResult.IsNG) //OK
                    {
                        fname = $"{sampleResult.SampleId}-{stationResult.Stop}-{stationResult.DefectName}-{stationResult.OkNgScore:F2}.jpg";
                    }
                    else //NG
                    {
                        if (stationResult.DefectScore.HasValue)
                        {
                            fname = $"{sampleResult.SampleId}-{stationResult.Stop}-{stationResult.DefectName}-{stationResult.DefectScore:F2}.jpg";
                        }
                        else
                        {
                            fname = $"{sampleResult.SampleId}-{stationResult.Stop}-{stationResult.DefectName}.jpg";
                        }
                    }

                    string savePath = Path.Combine(saveDir, fname);

                    string markText = "";
                    if (isNull)
                    {
                        markText = $"NULL | {stationResult.DefectName} | Score={stationResult.DefectScore:F2}";
                    }
                    else if (!stationResult.IsNG)
                    {
                        markText = $"OK | {stationResult.DefectName} | Score={stationResult.DefectScore:F2}";
                    }
                    else
                    {
                        if (stationResult.IsNG)
                        {
                            markText = $"NG | {stationResult.DefectName} | Score={stationResult.DefectScore:F2}";
                        }
                        else
                        {
                            markText = $"NG - {stationResult.DefectName}";
                        }
                    }

                    // 在圖上標記文字
                    using (Mat markedImage = stationResult.FinalMap.Clone())
                    {
                        stationResult.FinalMap.Dispose();
                        stationResult.FinalMap = null;

                        Scalar textColor = isNull ? Scalar.Yellow : (stationResult.IsNG ? Scalar.Red : Scalar.Green);

                        Cv2.PutText(
                            markedImage,
                            markText,
                            new Point(30, 60),
                            HersheyFonts.HersheySimplex,
                            2.0,
                            textColor,
                            5
                        );

                        // 保存原有路徑圖像（需要 Clone 因為 using 會釋放）
                        app.Queue_Save.Enqueue(new ImageSave(markedImage.Clone(), savePath));

                        // 新增功能：依據站點存放
                        string stationFolder = Path.Combine(stationBasePath, $"Station{stationResult.Stop}");
                        if (!Directory.Exists(stationFolder)) Directory.CreateDirectory(stationFolder);

                        // 站點路徑下的檔名可以與原來相同
                        string stationSavePath = Path.Combine(stationFolder, fname);
                        app.Queue_Save.Enqueue(new ImageSave(markedImage.Clone(), stationSavePath));
                    }

                    // 觸發儲存
                    app._sv.Set();
                }
            }
            // 更新統計資料
            foreach (var stationResult in sampleResult.StationResults.Values)
            {
                int station = stationResult.Stop;

                // 確保站點字典存在
                var defectDict = stationDefectStats.GetOrAdd(station,
                    _ => new ConcurrentDictionary<string, int>());

                // 更新 OK/NG 計數
                var currentStats = stationOkNgStats.GetOrAdd(station, (0, 0));
                if (stationResult.IsNG)
                {
                    // 更新瑕疵類型計數
                    if (!string.IsNullOrEmpty(stationResult.DefectName))
                    {
                        defectDict.AddOrUpdate(stationResult.DefectName, 1, (k, v) => v + 1);
                    }

                    // 更新 NG 計數
                    stationOkNgStats[station] = (currentStats.ok, currentStats.ng + 1);
                }
                else
                {
                    // 更新 OK 計數
                    stationOkNgStats[station] = (currentStats.ok + 1, currentStats.ng);
                }
            }
            // 收集樣品級別結果 - 紀錄這個樣品的最終狀態
            string finalDefectName = "";
            float? finalScore = null;

            // app.dc (defectcount 計數)
            if (isNG)
            {
                // 如果是 NG，我們找出分數最高(最糟糕)的瑕疵作為最終結果
                foreach (var result in sampleResult.StationResults.Values.Where(r => r.IsNG))
                {
                    if (!string.IsNullOrEmpty(result.DefectName) &&
                        (finalScore == null || (result.DefectScore.HasValue && result.DefectScore.Value > finalScore.Value)))
                    {
                        finalDefectName = result.DefectName;
                        finalScore = result.DefectScore;
                    }
                }
                // 只有發現了有效的最嚴重缺陷，才計數一次
                if (!string.IsNullOrEmpty(finalDefectName))
                {
                    // 由 GitHub Copilot 產生 // 修改: 使用 AddOrUpdate 原子遞增
                    app.dc.AddOrUpdate(finalDefectName, 1, (key, oldValue) => oldValue + 1);
                }
            }
            else
            {
                foreach (var result in sampleResult.StationResults.Values)
                {
                    if (!string.IsNullOrEmpty(result.DefectName) &&
                        (finalScore == null || (result.DefectScore.HasValue && result.DefectScore.Value > finalScore.Value)))
                    {
                        finalDefectName = result.DefectName;
                        finalScore = result.DefectScore;
                    }
                }
                // 存入樣品結果收集器
                //sampleResults[sampleResult.SampleId] = (isNG, finalDefectName, finalScore);
            }

            // 存入樣品結果收集器
            sampleResults[sampleResult.SampleId] = (isNG, finalDefectName, finalScore, timeout);

            // 複製每個樣品的所有站點結果
            if (sampleResult.StationResults != null && sampleResult.StationResults.Count > 0)
            {
                // 轉成普通 Dictionary 以避免外部被移除
                sampleStationResults[sampleResult.SampleId] = sampleResult.StationResults
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"保存樣品 {sampleResult.SampleId} 的最終結果時發生錯誤：{ex.Message}");
        }
    }
    
    private void CalculateAndSendPLCSignal(SampleResult sampleResult, bool isNG, bool isNull)
    {
        //Log.Debug($"樣品 {sampleResult.SampleId} 進入CalculateAndSendPLCSignal時間記錄: {DateTime.Now.ToString("HH:mm:ss.fff")} ms");
        if (app.allOK == true)
        {
            isNG = false;
        }
        int sampleId = sampleResult.SampleId;
        try
        {
            // ✅ **核心邏輯：檢查是否在同步模式**
            if (app.isInSyncMode)
            {
                TimeSpan elapsed = DateTime.Now - app.syncStartTime;

                if (elapsed.TotalMilliseconds < app.syncWaitTimeMs)
                {
                    // ✅ 還在等待期間，**所有樣品（包含 NG）都送 NULL**
                    isNull = true;
                    Log.Warning($"[同步] 樣品 {sampleId} 在同步等待期間（已過 {elapsed.TotalMilliseconds:F0}/{app.syncWaitTimeMs} ms），改送 NULL");
                }
                else
                {
                    // ✅ 等待結束，讀取 PLC 並覆蓋計數
                    try
                    {
                        // 由 GitHub Copilot 產生
                        // 修正: 讀取 PLC 並覆蓋計數
                        string lane = ResultManager.activeOkCounter;
                        int plcCount;

                        if (lane == "OK1")
                        {
                            plcCount = Form1.PLC_CheckD(803);
                        }
                        else // OK2
                        {
                            plcCount = Form1.PLC_CheckD(805);
                        }

                        // ✅ 檢查讀取是否成功
                        if (plcCount < 0)
                        {
                            Log.Error($"[同步失敗] PLC 讀取失敗，觸發停機");
                            app.plc_stop = true;
                            return;
                        }

                        // 計算實際包數

                        int softwareCount = ResultManager.counter[lane];
                        // 覆蓋軟體計數
                        if (softwareCount != plcCount)
                        {
                            lock (ResultManager.okLock)
                            {
                                ResultManager.counter[lane] = plcCount;
                            }
                        }
                        else // softwareCount == plcCount == app.pack
                        {
                            // 重置當前計數器
                            ResultManager.counter[lane] = 0;

                            // 切換到另一個計數器
                            ResultManager.activeOkCounter = (lane == "OK1") ? "OK2" : "OK1";

                            Log.Information($"[同步] {lane} 達到 {plcCount} 顆（== {app.pack}），重置為 0 並切換至 {ResultManager.activeOkCounter}");

                        }

                        Log.Information($"[同步完成] {lane}: 軟體計數 {softwareCount} → PLC實際值 {plcCount}");

                        // ✅ 關閉同步模式
                        app.isInSyncMode = false;

                        // 這個樣品恢復正常處理（NG 送 NG，OK 送 OK）
                        isNull = false;
                    }
                    catch (Exception plcEx)
                    {
                        Log.Error($"[同步失敗] PLC 通訊錯誤: {plcEx.Message}");
                        app.plc_stop = true;
                        return;
                    }
                }
            }
            // 由 GitHub Copilot 產生
            // 修改: 使用 TryGetValue 避免 Check-Then-Act 反模式
            if (!app.samplePhotoTimes.TryGetValue(sampleId, out DateTime photoTime))
            {
                Log.Warning($"樣品 {sampleId} 的拍照時間記錄不存在，無法計算精確推料時間");
                return; // 沒有時間記錄，不繼續處理 (不計數也不推料)
            }
        
            // 獲取當前時間
            DateTime resultTime = DateTime.Now;
            TimeSpan timeDifference = resultTime - photoTime;
            double left = 0.0;
            bool timeout = false;

            // 計算預期推料時間點
            DateTime expectedPushTime = photoTime;

            //拍照-送訊號的時間長度 若大於拍照到站口的時間長度 則超時
            if (ResultManager.activeOkCounter == "OK1" && isNG == false)
            {
                left = (double.Parse(app.param["fourToOK1_time_ms_4"]) - double.Parse(app.param["delay4"])) - timeDifference.TotalMilliseconds;
                expectedPushTime = photoTime.AddMilliseconds(double.Parse(app.param["fourToOK1_time_ms_4"]) - double.Parse(app.param["delay4"]));
            }
            else if (ResultManager.activeOkCounter == "OK2" && isNG == false)
            {
                left = (double.Parse(app.param["fourToOK2_time_ms_4"]) - double.Parse(app.param["delay4"])) - timeDifference.TotalMilliseconds;
                expectedPushTime = photoTime.AddMilliseconds(double.Parse(app.param["fourToOK2_time_ms_4"]) - double.Parse(app.param["delay4"]));
            }
            else if (isNG == true)
            {
                left = (double.Parse(app.param["fourToNG_time_ms_4"]) - double.Parse(app.param["delay4"])) - timeDifference.TotalMilliseconds;
                expectedPushTime = photoTime.AddMilliseconds(double.Parse(app.param["fourToNG_time_ms_4"]) - double.Parse(app.param["delay4"]));
            }
            if (left < 150)
            {
                isNull = true;
                timeout = true;                
            }


            int sigM = 0;
            if (isNull)
            {
                sigM = 13;
            }
            else if(isNG)
            {
                sigM = 1500 + sampleId % 40; //1500~1539
                //sigM = 10;
            }
            else
            {
                if (ResultManager.activeOkCounter == "OK1")
                {
                    sigM = 3000 + sampleId % 15; // 3000~3014
                    //sigM = 11;
                }
                else // OK2
                {
                    sigM = 4500 + sampleId % 15; // 4500~4514
                    //sigM = 12;
                }
            }
            
            bool plcSignalSuccess = false;
            string result = "";

            try
            {
                DateTime sendTime = DateTime.Now;

                

                // 發送 PLC 信號

                if (isNull)
                {
                    IncrementCounter(counter, "NULL");
                    Form1.PLC_SetM(sigM, true, true); //emer設為false 讓OK/NG優先
                    Thread.Sleep(50);
                    Form1.PLC_SetM(sigM, false, true);
                    // 【優化】移除高頻寫入 SAMPLE_ID，改在停止時寫入
                    // WriteSampleToDefectCounts(sampleId, sendTime);
                    result = "NULL";
                    plcSignalSuccess = true;

                    // 統一Log格式 - NULL
                    //Log.Information($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:NULL");
                }
                else if (isNG)
                {
                    IncrementCounter(counter, "NG");
                    Form1.PLC_SetM(sigM, true, true); //emer設為true 
                    Thread.Sleep(50);
                    Form1.PLC_SetM(sigM, false, true); //看這邊能否讓PLC自動設false
                    //移除高頻寫入 SAMPLE_ID，改在停止時寫入
                    // WriteSampleToDefectCounts(sampleId, sendTime);
                    result = "NG";
                    plcSignalSuccess = true;

                    // 統一Log格式 - NG
                    //Log.Information($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:NG");
                }
                else
                {
                    // OK 處理
                    Form1.PLC_SetM(sigM, true, true); //emer設為true 
                    Thread.Sleep(50);
                    Form1.PLC_SetM(sigM, false, true); //看這邊能否讓PLC自動設false

                    if (ResultManager.activeOkCounter == "OK1")
                    {
                        app.pendingOK1.Enqueue(sampleId);
                        result = "OK1";
                    }
                    else
                    {
                        app.pendingOK2.Enqueue(sampleId);
                        result = "OK2";
                    }

                    IncrementOkCounters(counter);
                    // 【優化】移除高頻寫入 SAMPLE_ID，改在停止時寫入
                    // WriteSampleToDefectCounts(sampleId, sendTime);
                    plcSignalSuccess = true;

                    //Log.Information($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:{result}");
                }

                // PLC 信號發送成功，才更新計數和保存結果
                if (plcSignalSuccess)
                {
                    // 只有在 PLC 信號發送成功後才計數
                    if (timeout)
                    {
                        SaveFinalResult(sampleResult, isNG, isNull, left);
                        Log.Warning($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 預期推料:無 | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:NULL (超時)");
                    }
                    else
                    {
                        if (isNull)
                        {
                            Log.Information($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 預期推料:無 | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:NULL");
                        }
                        else if (isNG)
                        {
                            Log.Information($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 預期推料:{expectedPushTime:HH:mm:ss.fff} | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:NG");
                        }
                        else
                        {
                            Log.Information($"[樣品 {sampleId}] 拍攝:{photoTime:HH:mm:ss.fff} | 預期推料:{expectedPushTime:HH:mm:ss.fff} | 送訊號:{sendTime:HH:mm:ss.fff} | 結果:{result}");
                        }
                        // 更新 ResultManager.counter 中的 SAMPLE_ID
                        ResultManager.counter.AddOrUpdate("SAMPLE_ID", sampleId, (key, oldValue) => sampleId); //停止按鈕，寫入時需要
                        SaveFinalResult(sampleResult, isNG, isNull);
                    }
                }
                else
                {
                    // 這個分支在目前邏輯中不會執行到，但保留以防未來邏輯變更
                    Log.Warning($"樣品 {sampleId} PLC 信號發送失敗，未更新計數");
                }

            }
            catch (Exception plcEx)
            {
                // 紀錄 PLC 通訊錯誤
                Log.Error($"PLC 信號發送失敗: {plcEx.Message}，樣品 {sampleId} 未計數");

                // 可以在這裡增加警報或其他機制通知操作員
                //lbAdd("PLC通訊失敗", "err", $"無法發送推料信號，請檢查PLC連線: {plcEx.Message}");

                // 可選: 將失敗的樣品ID加入重試隊列，等到PLC恢復後再處理
                // AddToPlcRetryQueue(sampleId, isNG);
            }

            Form1.PerformanceProfiler.FlushToDisk();
            // 用 TryRemove 保證執行緒安全
            // 完成後從字典中移除時間記錄，無論成功或失敗
            app.samplePhotoTimes.TryRemove(sampleId, out _);
        }
        catch (Exception ex)
        {
            Log.Error($"計算延遲時間時發生錯誤: {ex.Message}");
        }
    }

    public static void IncrementCounter(ConcurrentDictionary<string, int> dict, string key)
    {
        // AddOrUpdate 方法：
        // 如果键不存在，则添加键并将值初始化为 1
        // 如果键存在，则更新键的值，为其当前值加 1
        dict.AddOrUpdate(
            key,
            1, // 键不存在时的初始值
            (existingKey, existingValue) => existingValue + 1 // 键存在时的更新函数
        );
    }
    public static void IncrementOkCounters(ConcurrentDictionary<string, int> dict, int increment = 1)
    {
        string logMessage = "";
        lock (okLock)
        {
            // 先增加總數
            dict.AddOrUpdate("OK", increment, (key, oldValue) => oldValue + increment);

            // 再增加目前活躍計數器的數量
            dict.AddOrUpdate(activeOkCounter, increment, (key, oldValue) => oldValue + increment);

            // 檢查是否達到包裝數量
            if (dict[activeOkCounter] % app.pack == 0 && dict[activeOkCounter] > 0) // 已經包裝完一包
            {
                // ✅ 設定同步旗標
                app.isInSyncMode = true;

                // 記錄同步開始時間
                app.syncStartTime = DateTime.Now;

                // 計算等待時間（根據當前活躍計數器決定）
                if (activeOkCounter == "OK1")
                {
                    app.syncWaitTimeMs = int.Parse(app.param["fourToOK1_time_ms_4"])
                                       - int.Parse(app.param["delay4"])
                                       + 100;
                }
                else // OK2
                {
                    app.syncWaitTimeMs = int.Parse(app.param["fourToOK2_time_ms_4"])
                                       - int.Parse(app.param["delay4"])
                                       + 100;
                }

                logMessage = $"[同步] {activeOkCounter} 達到 {dict[activeOkCounter]} 顆，" +
                        $"從 {app.syncStartTime:HH:mm:ss.fff} 開始同步，持續 {app.syncWaitTimeMs} ms";
            }
            /*
            // 4. ✅ 檢查是否達到包裝數量（但**不在同步期間**切換）
            // ⚠️ 關鍵：只有在「非同步模式」且「確實達到 app.pack」時才切換
            if (!app.isInSyncMode && dict[activeOkCounter] == app.pack)
            {
                dict[activeOkCounter] = 0;
                activeOkCounter = (activeOkCounter == "OK1") ? "OK2" : "OK1";

                logMessage += $"\n[切換] 計數器切換至 {activeOkCounter}";
            }
            */
        }
        // 準備資訊日誌訊息(在 lock 外寫入)
        string infoMessage = $"[OKN監控] {DateTime.Now:HH:mm:ss.fff},          active:{activeOkCounter}, count:{dict[activeOkCounter]}";

        // 在 lock 外寫入 Log
        if (!string.IsNullOrEmpty(logMessage))
        {
            Log.Warning(logMessage);
        }
        Log.Information(infoMessage);
    }
    public static void ResetOkCounters()
    {
        lock (okLock)
        {
            try
            {
                // 重置活躍計數器為 OK1
                activeOkCounter = "OK1";

                // 重置計數器數值
                counter["OK1"] = 0;
                counter["OK2"] = 0;

                // 如果需要重置總計數,可以取消以下註解:
                // counter["OK"] = 0;

                Log.Information($"OK 計數器已重置，當前活躍計數器: {activeOkCounter}");
            }
            catch (Exception ex)
            {
                Log.Error($"重置 OK 計數器時發生錯誤: {ex.Message}");
                throw; // 重新拋出例外,讓上層處理
            }
        }
    }
    // ...existing code...
    private void WriteSampleToDefectCounts(int sampleId, DateTime timestamp)
    {
        const int maxRetries = 3;
        int retryCount = 0;

        while (retryCount < maxRetries)
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 使用 LinqToDB 的正確 UPSERT 語法
                    // 方案1: 使用 InsertOrUpdate (LinqToDB 推薦)
                    db.DefectCounts
                        .InsertOrUpdate(
                            // Insert 條件
                            () => new DefectCount
                            {
                                Type = app.produce_No,
                                Name = "SAMPLE_ID",
                                Count = sampleId,
                                Time = timestamp,
                                LotId = app.LotID
                            },
                            // Update 條件 (當記錄已存在時)
                            dc => new DefectCount
                            {
                                Count = sampleId,
                                Time = timestamp
                            },
                            // 判斷是否存在的條件
                            () => new DefectCount
                            {
                                Type = app.produce_No,
                                Name = "SAMPLE_ID",
                                LotId = app.LotID
                            }
                        );

                    Log.Debug($"樣品 {sampleId} 已儲存至 DefectCounts");
                }

                // 成功執行，跳出重試循環
                break;
            }
            catch (System.Data.SQLite.SQLiteException ex) when (ex.ErrorCode == 5) // SQLITE_BUSY
            {
                retryCount++;
                if (retryCount >= maxRetries)
                {
                    Log.Error($"寫入樣品 {sampleId} 到 DefectCounts 失敗 (資料庫忙碌，已重試 {maxRetries} 次): {ex.Message}");
                    throw;
                }

                // 遞增延遲時間後重試
                Thread.Sleep(50 * retryCount);
                Log.Warning($"資料庫忙碌，重試寫入樣品 {sampleId} (第 {retryCount} 次)");
            }
            catch (Exception ex)
            {
                Log.Error($"寫入樣品 {sampleId} 到 DefectCounts 失敗: {ex.Message}");
                throw;
            }
        }
    }
    // ...existing code...
    private void WriteSampleToDefectCounts_OLD(int sampleId, DateTime timestamp)
    {
        try
        {
            using (var db = new MydbDB())
            {
                // 嘗試更新現有記錄
                int updatedRows = db.DefectCounts
                    .Where(p => p.Type == app.produce_No &&
                               p.Name == "SAMPLE_ID" &&
                               p.LotId == app.LotID)
                               //p.Count == sampleId)
                    .Set(p => p.Count, sampleId)
                    .Set(p => p.Time, timestamp)
                    .Update();

                // 如果沒有更新任何記錄，表示記錄不存在，需要插入
                if (updatedRows == 0)
                {
                    db.DefectCounts
                        .Value(p => p.Type, app.produce_No)           // 料號
                        .Value(p => p.Name, "SAMPLE_ID")              // 固定名稱
                        .Value(p => p.Count, sampleId)                // 樣品ID作為計數
                        .Value(p => p.Time, timestamp)                // 完成時間
                        .Value(p => p.LotId, app.LotID)               // LotID
                        .Insert();

                    Log.Debug($"樣品 {sampleId} 已新增至 DefectCounts");
                }
                else
                {
                    //Log.Debug($"樣品 {sampleId} 已更新 DefectCounts 時間戳");
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error($"寫入樣品 {sampleId} 到 DefectCounts 失敗: {ex.Message}");
        }
    }
    public static class ScoreAnalyzer
    {
        /// <summary>
        /// 收集按瑕疵類型和結果(OK/NG)分類的分數
        /// </summary>
        public static Dictionary<string, List<(float score, bool isOK)>> CollectDefectScores()
        {
            var defectScores = new Dictionary<string, List<(float score, bool isOK)>>();

            foreach (var sample in sampleResults)
            {
                if (!sample.Value.score.HasValue) continue;

                //string defectName = sample.Value.isNG ? sample.Value.defectName ?? "未分類" : "OK";

                // 不論 isNG，全部都用 defectName（沒有就 "未分類"）
                string defectName = sample.Value.defectName ?? "未分類";
                bool isOK = !sample.Value.isNG;
                float score = sample.Value.score.Value;

                if (!defectScores.ContainsKey(defectName))
                    defectScores[defectName] = new List<(float, bool)>();

                defectScores[defectName].Add((score, isOK));
            }

            return defectScores;
        }

        /// <summary>
        /// 計算統計數據
        /// </summary>
        public static Dictionary<string, (float min, float max, float avg, float median, float p5, float p25, float p75, float p95, float p99)>
            CalculateStatistics(Dictionary<string, List<(float score, bool isOK)>> defectScores)
        {
            var stats = new Dictionary<string, (float min, float max, float avg, float median, float p5, float p25, float p75, float p95, float p99)>();

            foreach (var kvp in defectScores)
            {
                string defectName = kvp.Key;
                var scores = kvp.Value.Select(x => x.score).OrderBy(s => s).ToList();

                if (scores.Count == 0) continue;

                float min = scores.First();
                float max = scores.Last();
                float avg = scores.Average();
                float median = scores.Count % 2 == 0
                    ? (scores[scores.Count / 2 - 1] + scores[scores.Count / 2]) / 2
                    : scores[scores.Count / 2];

                // 新增：計算 5%、25% 和 75% 分位數
                float p5 = scores.Count > 20 ? scores[(int)(scores.Count * 0.05)] : min;
                float p25 = scores.Count > 4 ? scores[(int)(scores.Count * 0.25)] : min;
                float p75 = scores.Count > 4 ? scores[(int)(scores.Count * 0.75)] : max;
                float p95 = scores.Count > 20 ? scores[(int)(scores.Count * 0.95)] : max;
                float p99 = scores.Count > 100 ? scores[(int)(scores.Count * 0.99)] : max;

                stats[defectName] = (min, max, avg, median, p5, p25, p75, p95, p99);
            }

            return stats;
        }

        /// <summary>
        /// 找尋最佳閾值
        /// </summary>
        public static Dictionary<string, float> FindOptimalThresholds(
            Dictionary<string, List<(float score, bool isOK)>> defectScores,
            float targetFalsePositiveRate = 0.05f)
        {
            var thresholds = new Dictionary<string, float>();

            foreach (var kvp in defectScores)
            {
                string defectName = kvp.Key;
                var scores = kvp.Value;

                // 若全是OK品，用P95作為閾值
                if (scores.All(s => s.isOK))
                {
                    var sortedScores = scores.Select(s => s.score).OrderBy(s => s).ToList();
                    int index = Math.Min((int)(sortedScores.Count * 0.95), sortedScores.Count - 1);
                    thresholds[defectName] = sortedScores[index];
                    continue;
                }

                // 如果有OK/NG混合，尋找最佳分類閾值
                float bestThreshold = 0;
                float bestMetric = -1;

                // 測試不同閾值（步進0.01）
                for (float t = 0; t <= 1.0f; t += 0.01f)
                {
                    int tp = scores.Count(s => !s.isOK && s.score >= t);  // 真陽性
                    int fp = scores.Count(s => s.isOK && s.score >= t);   // 假陽性
                    int tn = scores.Count(s => s.isOK && s.score < t);    // 真陰性
                    int fn = scores.Count(s => !s.isOK && s.score < t);   // 假陰性

                    float tpRate = tp / (float)Math.Max(1, tp + fn);
                    float fpRate = fp / (float)Math.Max(1, fp + tn);

                    // 使用F1分數或其他指標
                    float precision = tp / (float)Math.Max(1, tp + fp);
                    float recall = tpRate;
                    float f1 = 2 * (precision * recall) / Math.Max(0.001f, precision + recall);

                    // 降低誤判率的同時保持較高召回率的指標
                    float metric = f1 - (fpRate > targetFalsePositiveRate ? fpRate - targetFalsePositiveRate : 0) * 2;

                    if (metric > bestMetric)
                    {
                        bestMetric = metric;
                        bestThreshold = t;
                    }
                }

                thresholds[defectName] = bestThreshold;
            }

            return thresholds;
        }
    }


    #region 結果統計
    // 新增：獲取特定站點的瑕疵統計
    public static Dictionary<string, int> GetDefectStatsByStation(int station)
    {
        if (stationDefectStats.TryGetValue(station, out var stats))
        {
            return new Dictionary<string, int>(stats);
        }
        return new Dictionary<string, int>();
    }

    // 新增：獲取特定站點的 OK/NG 統計
    public static (int ok, int ng) GetOkNgStatsByStation(int station)
    {
        if (stationOkNgStats.TryGetValue(station, out var stats))
        {
            return stats;
        }
        return (0, 0);
    }

    // 新增：獲取所有站點的瑕疵統計總覽
    public static Dictionary<int, Dictionary<string, int>> GetAllStationDefectStats()
    {
        var result = new Dictionary<int, Dictionary<string, int>>();
        foreach (var pair in stationDefectStats)
        {
            result[pair.Key] = new Dictionary<string, int>(pair.Value);
        }
        return result;
    }

    // 新增：將統計結果輸出到 CSV 檔案
    public static bool ExportStatsToCsv(string filePath)
    {
        try
        {
            using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                #region 總體統計
                // 由 GitHub Copilot 產生 - 總體統計摘要（增加 null 安全檢查）
                writer.WriteLine();
                writer.WriteLine($"LOT ID:,{app.LotID}");
                writer.WriteLine($"工單數量:,{app.order}");
                writer.WriteLine("==== 總體統計 ====");

                // 由 GitHub Copilot 產生 - 以樣品為單位計算總數（不是站點單位）
                // 從sampleResults中統計唯一的樣品ID數量
                int totalSamples = 0;
                int totalOK = 0;
                int totalNG = 0;

                // 由 GitHub Copilot 產生 - 統計所有樣品的OK/NG狀態（一個樣品有任一站NG就算NG）
                var sampleStatusGroup = sampleResults.GroupBy(r => r.Key).Select(g => new
                {
                    SampleId = g.Key,
                    IsNG = g.Any(r => r.Value.isNG)
                }).ToList();

                totalSamples = sampleStatusGroup.Count;
                totalOK = sampleStatusGroup.Count(s => !s.IsNG);
                totalNG = sampleStatusGroup.Count(s => s.IsNG);

                // 由 GitHub Copilot 產生 - 如果sampleResults為空，從站點統計推算
                if (totalSamples == 0)
                {
                    // 使用站點統計推算：取最小站點的樣品數作為總數（代表完成所有站點檢測的樣品數）
                    var stationTotals = stationOkNgStats.Values.Select(s => s.ok + s.ng).ToList();
                    totalSamples = stationTotals.Count > 0 ? stationTotals.Min() : 0;

                    // 從站點瑕疵統計推算OK/NG（取平均值）
                    if (stationTotals.Count > 0 && totalSamples > 0)
                    {
                        int avgOK = (int)stationOkNgStats.Values.Average(s => s.ok);
                        int avgNG = (int)stationOkNgStats.Values.Average(s => s.ng);
                        totalOK = avgOK;
                        totalNG = avgNG;
                    }
                }
                if (!app.offline)
                {
                    int ng = Form1.PLC_CheckD(801);
                    int ok = Form1.PLC_CheckD(807);
                    if (ok != 0 && ng != 0)
                    {
                        totalOK = ok;
                        totalNG = ng;
                        totalSamples = ok + ng;
                    }
                }
                // 由 GitHub Copilot 產生 - 計算 OK 率和不良率
                double okPercentage = totalSamples > 0 ? (totalOK * 100.0 / totalSamples) : 0;
                double ngPercentage = totalSamples > 0 ? (totalNG * 100.0 / totalSamples) : 0;
                
                writer.WriteLine($"總樣品數,{totalSamples}");
                writer.WriteLine($"OK數量,{totalOK},OK率(%),{okPercentage:F2}");
                writer.WriteLine($"NG數量,{totalNG},不良率(%),{ngPercentage:F2}");
                #endregion
                #region 瑕疵類型統計
                // 瑕疵類型統計
                var defectTypeSummary = new Dictionary<string, int>();
                foreach (var sample in sampleResults.Where(s => s.Value.isNG))
                {
                    string defectName = sample.Value.defectName ?? "未知瑕疵";
                    if (defectTypeSummary.ContainsKey(defectName))
                        defectTypeSummary[defectName]++;
                    else
                        defectTypeSummary[defectName] = 1;
                }

                writer.WriteLine();
                writer.WriteLine("==== 瑕疵統計（只統計 NG 樣品） ====");
                writer.WriteLine("瑕疵類型,數量,比例");
                foreach (var defect in defectTypeSummary.OrderByDescending(d => d.Value))
                {
                    // 計算此瑕疵在所有NG樣本中的比例
                    float defectPercentage = totalNG > 0 ? ((float)defect.Value / totalNG) : 0;
                    writer.WriteLine($"{defect.Key},{defect.Value},{defectPercentage:F2}");
                    //writer.WriteLine($"{defect.Key},{defect.Value}");
                }
                #endregion
                #region 分數分布分析
                // 由 GitHub Copilot 產生 - 添加新的分數分布分析部分（加入 null 檢查）
                writer.WriteLine();
                writer.WriteLine("==== 所有檢出瑕疵分布（包含 OK 樣品中的低分瑕疵） ====");

                // 由 GitHub Copilot 產生 - 收集分數數據，加入 try-catch 避免 NullReferenceException
                try
                {
                    var defectScores = ScoreAnalyzer.CollectDefectScores();
                    if (defectScores != null && defectScores.Count > 0)
                    {
                        var scoreStats = ScoreAnalyzer.CalculateStatistics(defectScores);
                        var optimalThresholds = ScoreAnalyzer.FindOptimalThresholds(defectScores);

                        writer.WriteLine("瑕疵類型,樣本數量,最小值,最大值,平均值,5%分位數,25%分位數,中位數,75%分位數,95%分位數,99%分位數,推薦閾值");

                        foreach (var defect in scoreStats)
                        {
                            string defectName = defect.Key;
                            var stats = defect.Value;
                            int count = defectScores.ContainsKey(defectName) ? defectScores[defectName].Count : 0;
                            float recommendedThreshold = optimalThresholds.ContainsKey(defectName) ?
                                                         optimalThresholds[defectName] : 0.5f;

                            writer.WriteLine($"{defectName},{count}," +
                                             $"{stats.min:F2},{stats.max:F2},{stats.avg:F2}," +
                                             $"{stats.p5:F2},{stats.p25:F2}," +
                                             $"{stats.median:F2},{stats.p75:F2}," +
                                             $"{stats.p95:F2},{stats.p99:F2}," +
                                             $"{recommendedThreshold:F2}");
                        }
                    }
                    else
                    {
                        writer.WriteLine("無分數數據");
                    }
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"分數分析失敗: {ex.Message}");
                }
                #endregion
                #region 12/34 合併統計OK/NG
                // 由 GitHub Copilot 產生 - 修正: 以樣品為單位統計，任一站 NG 則該樣品算 NG
                int group12OK = 0, group12NG = 0;
                int group34OK = 0, group34NG = 0;

                // 由 GitHub Copilot 產生 - 遍歷所有樣品，檢查站點 1 和 2 的結果
                foreach (var samplePair in sampleStationResults)
                {
                    int sampleId = samplePair.Key;
                    var stations = samplePair.Value;

                    // 檢查站點 1 和 2（只要有其中一站的資料）
                    bool hasStat1or2 = stations.ContainsKey(1) || stations.ContainsKey(2);
                    if (hasStat1or2)
                    {
                        // 只要站點 1 或 2 任一個是 NG，該樣品就算 NG
                        bool isNG12 = (stations.ContainsKey(1) && stations[1].IsNG) ||
                                      (stations.ContainsKey(2) && stations[2].IsNG);
                        
                        if (isNG12)
                            group12NG++;
                        else
                            group12OK++;
                    }

                    // 檢查站點 3 和 4（只要有其中一站的資料）
                    bool hasStat3or4 = stations.ContainsKey(3) || stations.ContainsKey(4);
                    if (hasStat3or4)
                    {
                        // 只要站點 3 或 4 任一個是 NG，該樣品就算 NG
                        bool isNG34 = (stations.ContainsKey(3) && stations[3].IsNG) ||
                                      (stations.ContainsKey(4) && stations[4].IsNG);
                        
                        if (isNG34)
                            group34NG++;
                        else
                            group34OK++;
                    }
                }
                
                writer.WriteLine();
                writer.WriteLine("==== 合併組別OK/NG統計 ====");
                writer.WriteLine("組別,OK數量,NG數量");
                writer.WriteLine($"1+2,{group12OK},{group12NG}");
                writer.WriteLine($"3+4,{group34OK},{group34NG}");
                #endregion
                #region 站點統計
                writer.WriteLine();
                writer.WriteLine("==== 站點統計 ====");
                writer.WriteLine("站點,OK數量,NG數量,瑕疵類型,瑕疵數量");
                foreach (var stationPair in stationDefectStats)
                {
                    int station = stationPair.Key;
                    var defectCounts = stationPair.Value;

                    // 由 GitHub Copilot 產生 - 使用 TryGetValue 避免 KeyNotFoundException
                    if (!stationOkNgStats.TryGetValue(station, out var okNgStats))
                    {
                        okNgStats = (0, 0);
                    }

                    // 如果沒有瑕疵，輸出站點基本資訊
                    if (defectCounts.Count == 0)
                    {
                        writer.WriteLine($"{station},{okNgStats.ok},{okNgStats.ng},無,0");
                    }
                    else
                    {
                        foreach (var defectPair in defectCounts)
                        {
                            // 每一行都填充站點和 OK/NG 數量
                            writer.WriteLine($"{station},{okNgStats.ok},{okNgStats.ng},{defectPair.Key},{defectPair.Value}");
                        }
                    }
                }
                #endregion

            }
            return true;
        }
        catch (Exception ex)
        {
            Log.Error($"匯出統計資料時發生錯誤: {ex.Message}");
            return false;   
        }
    }

    // 新增：清除統計資料
    public static void ClearStats()
    {
        stationDefectStats.Clear();
        stationOkNgStats.Clear();
        sampleResults.Clear();

        // 預設四個站點
        for (int i = 1; i <= 4; i++)
        {
            stationDefectStats[i] = new ConcurrentDictionary<string, int>();
            stationOkNgStats[i] = (0, 0);
        }
    }
    #endregion
}

public class SampleResult
{
    public int SampleId { get; }
    public ConcurrentDictionary<int, StationResult> StationResults { get; }

    public SampleResult(int sampleId)
    {
        SampleId = sampleId;
        StationResults = new ConcurrentDictionary<int, StationResult>();
    }

    public void AddStationResult(StationResult stationResult)
    {
        StationResults[stationResult.Stop] = stationResult; //一整個結構傳進去，字典的型別本來就是這樣寫的
    }

    public bool IsComplete(int totalStations) //有bug的話可以看一下add的時機點
    {
        return StationResults.Count == totalStations;
    }
}

public class StationResult
{
    public int Stop { get; set; } // 站点编号
    public bool IsNG { get; set; }
    public float? OkNgScore { get; set; }
    public Mat FinalMap { get; set; }
    public string DefectName { get; set; }
    public float? DefectScore { get; set; }
    public string OriName { get; set; }
}
#endregion
public class ScoreStatisticsManager
{
    // 每個站點的分數紀錄
    private static readonly ConcurrentDictionary<int, ConcurrentBag<float>> stationScores = new ConcurrentDictionary<int, ConcurrentBag<float>>();

    // 每個站點的瑕疵類型統計
    private static readonly ConcurrentDictionary<int, ConcurrentDictionary<string, int>> defectCounts =
        new ConcurrentDictionary<int, ConcurrentDictionary<string, int>>();

    // 每個站點的OK/NG計數
    private static readonly ConcurrentDictionary<int, (int ok, int ng)> okNgCounts =
        new ConcurrentDictionary<int, (int ok, int ng)>();

    // 檔案鎖
    private static readonly object csvLock = new object();

    /// <summary>
    /// 初始化統計管理器
    /// </summary>
    public static void Initialize()
    {
        // 清空所有資料
        stationScores.Clear();
        defectCounts.Clear();
        okNgCounts.Clear();

        // 初始化4個站點的資料結構
        for (int i = 1; i <= 4; i++)
        {
            stationScores[i] = new ConcurrentBag<float>();
            defectCounts[i] = new ConcurrentDictionary<string, int>();
            okNgCounts[i] = (0, 0);
        }
    }

    /// <summary>
    /// 新增一筆檢測結果數據
    /// </summary>
    public static void AddScore(int station, float score, bool isNG, string defectName = null)
    {
        // 確保站點索引有效
        if (station < 1 || station > 4)
            return;

        // 添加分數記錄
        if (!stationScores.ContainsKey(station))
            stationScores[station] = new ConcurrentBag<float>();

        stationScores[station].Add(score);

        // 更新OK/NG計數
        var current = okNgCounts.GetOrAdd(station, (0, 0));
        if (isNG)
            okNgCounts[station] = (current.ok, current.ng + 1);
        else
            okNgCounts[station] = (current.ok + 1, current.ng);

        // 若為NG，統計瑕疵類型
        if (isNG && !string.IsNullOrEmpty(defectName))
        {
            if (!defectCounts.ContainsKey(station))
                defectCounts[station] = new ConcurrentDictionary<string, int>();

            defectCounts[station].AddOrUpdate(defectName, 1, (key, count) => count + 1);
        }
    }

    /// <summary>
    /// 生成直方圖並顯示
    /// </summary>
    public static void GenerateHistogram(Chart chart, int station, int binCount = 10)
    {
        if (chart == null || !stationScores.ContainsKey(station) || stationScores[station].IsEmpty)
            return;

        try
        {
            // 清空圖表
            chart.Series.Clear();
            chart.ChartAreas.Clear();
            chart.Titles.Clear();

            // 建立圖表區域
            ChartArea chartArea = new ChartArea($"Station{station}Area");
            chart.ChartAreas.Add(chartArea);

            // 建立標題
            Title title = new Title($"站點 {station} 分數分佈");
            chart.Titles.Add(title);

            // 建立系列
            Series histSeries = new Series($"Station{station}");
            histSeries.ChartType = SeriesChartType.Column;
            histSeries.Color = Color.SteelBlue;
            chart.Series.Add(histSeries);

            // 獲取分數數據
            var scores = stationScores[station].ToArray();

            // 如果沒有數據，直接返回
            if (scores.Length == 0)
                return;

            // 計算直方圖數據
            double min = scores.Min();
            double max = scores.Max();
            double range = max - min;
            double binWidth = range / binCount;

            // 建立柱狀圖數據
            int[] bins = new int[binCount];

            foreach (var score in scores)
            {
                int binIndex = (int)((score - min) / binWidth);
                if (binIndex == binCount) // 處理最大值
                    binIndex--;
                bins[binIndex]++;
            }

            // 添加數據點
            for (int i = 0; i < binCount; i++)
            {
                double binStart = min + i * binWidth;
                double binEnd = binStart + binWidth;
                string label = $"{binStart:F2} - {binEnd:F2}";
                histSeries.Points.AddXY(label, bins[i]);
            }

            // 設定軸標籤
            chartArea.AxisX.Title = "分數範圍";
            chartArea.AxisY.Title = "頻率";
            chartArea.AxisX.LabelStyle.Angle = -45;
            chartArea.AxisX.LabelStyle.Font = new Font("Arial", 8);

            // 增加OK/NG百分比標籤
            var okNgStat = okNgCounts.GetOrAdd(station, (0, 0));
            int total = okNgStat.ok + okNgStat.ng;
            double okPercent = total > 0 ? (double)okNgStat.ok / total * 100 : 0;
            double ngPercent = total > 0 ? (double)okNgStat.ng / total * 100 : 0;

            Title statTitle = new Title($"OK: {okNgStat.ok}({okPercent:F1}%), NG: {okNgStat.ng}({ngPercent:F1}%)");
            statTitle.DockedToChartArea = chartArea.Name;
            statTitle.Docking = Docking.Bottom;
            statTitle.Font = new Font("Arial", 9, FontStyle.Bold);
            chart.Titles.Add(statTitle);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成直方圖時發生錯誤: {ex.Message}");
        }
    }

    /// <summary>
    /// 導出所有站點的數據到CSV文件
    /// </summary>
    public static bool ExportToCsv(string folderPath, string prefix = "")
    {
        try
        {
            // 確保目錄存在
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            // 檔案名稱採用當前時間
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"{prefix}分數統計_{timestamp}.csv";
            string fullPath = Path.Combine(folderPath, fileName);

            lock (csvLock)
            {
                using (StreamWriter sw = new StreamWriter(fullPath, false, System.Text.Encoding.UTF8))
                {
                    // 寫入標題行
                    sw.WriteLine("日期,站點,總樣本數,OK數量,NG數量,OK率(%),NG率(%),平均分數,最高分數,最低分數,標準差");

                    // 寫入每個站點的統計數據
                    foreach (var stationEntry in stationScores)
                    {
                        int station = stationEntry.Key;
                        var scores = stationEntry.Value.ToArray();

                        if (scores.Length == 0)
                            continue;

                        var okNgStat = okNgCounts.GetOrAdd(station, (0, 0));
                        int totalSamples = okNgStat.ok + okNgStat.ng;
                        double okRate = totalSamples > 0 ? (double)okNgStat.ok / totalSamples * 100 : 0;
                        double ngRate = totalSamples > 0 ? (double)okNgStat.ng / totalSamples * 100 : 0;

                        double avgScore = scores.Average();
                        double maxScore = scores.Max();
                        double minScore = scores.Min();

                        // 計算標準差
                        double variance = scores.Sum(x => Math.Pow(x - avgScore, 2)) / scores.Length;
                        double stdDev = Math.Sqrt(variance);

                        sw.WriteLine($"{DateTime.Now.ToString("yyyy-MM-dd")},{station},{totalSamples},{okNgStat.ok},{okNgStat.ng},{okRate:F2},{ngRate:F2},{avgScore:F4},{maxScore:F4},{minScore:F4},{stdDev:F4}");
                    }

                    // 添加分隔行
                    sw.WriteLine();

                    // 添加瑕疵類型統計
                    sw.WriteLine("站點,瑕疵類型,出現次數,百分比(%)");
                    foreach (var stationEntry in defectCounts)
                    {
                        int station = stationEntry.Key;
                        var defects = stationEntry.Value;
                        var okNgStat = okNgCounts.GetOrAdd(station, (0, 0));
                        int totalNG = okNgStat.ng;

                        foreach (var defect in defects)
                        {
                            string defectName = defect.Key;
                            int count = defect.Value;
                            double percentage = totalNG > 0 ? (double)count / totalNG * 100 : 0;

                            sw.WriteLine($"{station},{defectName},{count},{percentage:F2}");
                        }
                    }
                }
            }

            Console.WriteLine($"CSV檔案已成功導出到: {fullPath}");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"導出CSV檔案時發生錯誤: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 生成瑕疵類型圓餅圖
    /// </summary>
    public static void GenerateDefectPieChart(Chart chart, int station)
    {
        if (chart == null || !defectCounts.ContainsKey(station))
            return;

        try
        {
            // 清空圖表
            chart.Series.Clear();
            chart.ChartAreas.Clear();
            chart.Titles.Clear();
            chart.Legends.Clear();

            // 建立圖表區域
            ChartArea chartArea = new ChartArea($"Defects{station}Area");
            chart.ChartAreas.Add(chartArea);

            // 建立標題
            Title title = new Title($"站點 {station} 瑕疵類型統計");
            chart.Titles.Add(title);

            // 建立系列
            Series pieSeries = new Series($"Defects{station}");
            pieSeries.ChartType = SeriesChartType.Pie;
            chart.Series.Add(pieSeries);

            // 建立圖例
            Legend legend = new Legend($"Legend{station}");
            chart.Legends.Add(legend);
            pieSeries.Legend = legend.Name;

            // 獲取瑕疵數據
            var defects = defectCounts[station];
            if (defects.Count == 0)
                return;

            // 添加數據點
            foreach (var defect in defects)
            {
                string label = $"{defect.Key} ({defect.Value})";
                pieSeries.Points.AddXY(label, defect.Value);
            }

            // 設定餅圖顯示樣式
            foreach (var point in pieSeries.Points)
            {
                point.LabelFormat = "{0} #PERCENT{P1}";
                point.Label = "#VALX: #PERCENT{P1}";
                point.IsValueShownAsLabel = true;
                point.LabelForeColor = Color.White;
                point.Font = new Font("Arial", 9, FontStyle.Bold);
            }

            pieSeries.BorderColor = Color.White;
            pieSeries.BorderWidth = 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成瑕疵圓餅圖時發生錯誤: {ex.Message}");
        }
    }
}

#region PositionClass
public class ImagePosition
{
    public Mat image;
    public int count;
    public int stop;
    public DateTime time;
    public string name;
    public ImagePosition() { }
    public ImagePosition(Mat image, int count)
    {
        this.image = image;
        this.count = count;
    }
    public ImagePosition(Mat image, int stop, int count)
    {
        this.image = image;
        this.stop = stop;
        this.count = count;
    }
    public ImagePosition(Mat image, int count, DateTime time)
    {
        this.image = image;
        this.count = count;
        this.time = time;
    }
    public ImagePosition(Mat image, int stop, int count, DateTime time)
    {
        this.image = image;
        this.stop = stop;
        this.count = count;
        this.time = time;
    }
    public ImagePosition(Mat image, int count, int stop, string name)
    {
        this.image = image;
        this.count = count;
        this.stop = stop;
        this.name = name;
    }
    public ImagePosition(Mat image, int count, string name)
    {
        this.image = image;
        this.count = count;
        this.name = name;
    }
    public ImagePosition(Mat image, int count, int stop, string name, DateTime time)
    {
        this.image = image;
        this.count = count;
        this.stop = stop;
        this.name = name;
        this.time = time;
    }
}
public class ImageSave
{
    public Mat image;
    public string path;
    public ImageSave(Mat image, string path)
    {
        this.image = image;
        this.path = path;
    }
}

#endregion
#region DetectData
public class DetectData
{
    public int SerialNum;
    public string OKNG;
    public double roundness;
    public double max;
    public double min;
    public double r;
    public DateTime time;
    public bool[] defect_Type;
    public float score;
    public float score2;
    public Mat register;
    public double gapTime;
    public DateTime time2;
    public DetectData()
    {
    }
    public DetectData(int SerialNum, string OKNG, double roundness, double max, double min, double r, DateTime time, bool[] defect_Type, float score, float score2, double gapTime, DateTime time2)
    {
        this.SerialNum = SerialNum;
        this.OKNG = OKNG;
        this.roundness = roundness;
        this.max = max;
        this.min = min;
        this.r = r;
        this.defect_Type = defect_Type;
        this.time = time;
        this.score = score;
        this.gapTime = gapTime;
        this.time2 = time2;
    }
}
#endregion
public class PendingPushInfo
{
    // 推料對時所需的完整資訊
    public int SampleId { get; set; }
    public string Lane { get; set; } // "OK1" / "OK2"
    public int SigM { get; set; } // 實際送出的 M
    public int ExpectedM { get; set; } // 依送出當下 D98/D99 推算
    public int D98AtSend { get; set; }
    public int D99AtSend { get; set; }
    public DateTime EnqueuedAt { get; set; }
}
public class app
{
    // 由 GitHub Copilot 產生
    // 修改: 使用執行緒安全的 ConcurrentDictionary
    public static ConcurrentDictionary<int, string> detect_result = new ConcurrentDictionary<int, string>();
    public static ConcurrentDictionary<int, bool[]> detect_result_check = new ConcurrentDictionary<int, bool[]>();
    public static EventWaitHandle _wh1 = new AutoResetEvent(false),
                                  _wh2 = new AutoResetEvent(false),
                                  _wh3 = new AutoResetEvent(false),
                                  _wh4 = new AutoResetEvent(false),
                                  _sv = new AutoResetEvent(false),
                                  _reader = new AutoResetEvent(false),
                                  _AI = new AutoResetEvent(false),
                                  _show = new AutoResetEvent(false);
    public static Task T1, T2, T3, T4, Tsv, Treader, TAI, Tshow;
    public static ConcurrentQueue<ImagePosition> Queue_Bitmap1 = new ConcurrentQueue<ImagePosition>();
    public static ConcurrentQueue<ImagePosition> Queue_Bitmap2 = new ConcurrentQueue<ImagePosition>();
    public static ConcurrentQueue<ImagePosition> Queue_Bitmap3 = new ConcurrentQueue<ImagePosition>();
    public static ConcurrentQueue<ImagePosition> Queue_Bitmap4 = new ConcurrentQueue<ImagePosition>();
    public static ConcurrentQueue<ImageSave> Queue_Save = new ConcurrentQueue<ImageSave>();
    public static ConcurrentQueue<ImagePosition> Queue_Send = new ConcurrentQueue<ImagePosition>();
    public static ConcurrentQueue<ImagePosition> Queue_Show = new ConcurrentQueue<ImagePosition>();
    // 等待 D803/D805 跳變時要配對的樣品ID佇列（只拿來 Log，不寫檔）
    public static System.Collections.Concurrent.ConcurrentQueue<int> pendingOK1 =
        new System.Collections.Concurrent.ConcurrentQueue<int>();
    public static System.Collections.Concurrent.ConcurrentQueue<int> pendingOK2 =
        new System.Collections.Concurrent.ConcurrentQueue<int>();
    public static System.Collections.Concurrent.ConcurrentQueue<PendingPushInfo> pendingPushOK1 =
        new System.Collections.Concurrent.ConcurrentQueue<PendingPushInfo>();
    public static System.Collections.Concurrent.ConcurrentQueue<PendingPushInfo> pendingPushOK2 =
        new System.Collections.Concurrent.ConcurrentQueue<PendingPushInfo>();
    //public static Dictionary<string, InferenceSession> onnxSessions = new Dictionary<string, InferenceSession>();
    // 由 GitHub Copilot 產生
    // 修改: 使用執行緒安全的 ConcurrentDictionary
    public static ConcurrentDictionary<string, int> counter = new ConcurrentDictionary<string, int>();
    public static ConcurrentDictionary<string, int> dc = new ConcurrentDictionary<string, int>();
    public static ConcurrentDictionary<string, string> param = new ConcurrentDictionary<string, string>();
    public static ConcurrentDictionary<string, string> models = new ConcurrentDictionary<string, string>();
    public static ConcurrentDictionary<string, string> metas = new ConcurrentDictionary<string, string>();
    public static ConcurrentDictionary<string, int> pos = new ConcurrentDictionary<string, int>();
    public static Dictionary<string, Mat> img;
    // 由 GitHub Copilot 產生
    // 修改: 使用執行緒安全的 ConcurrentDictionary
    public static ConcurrentDictionary<int, DateTime> samplePhotoTimes = new ConcurrentDictionary<int, DateTime>();
    public static List<string> defectLabels_in = new List<string>();
    public static List<string> defectLabels_out = new List<string>();
    public static List<string> defect_name = new List<string>();
    // 由 GitHub Copilot 產生
    // 新增: 每個站點的瑕疵名稱快取,避免重複查詢資料庫
    public static ConcurrentDictionary<int, List<string>> defectNamesPerStop = new ConcurrentDictionary<int, List<string>>();
    // 由 GitHub Copilot 產生
    // 新增: 倒角檢測快取 (料號_站點 -> 是否需要檢測)
    public static ConcurrentDictionary<string, bool> chamferDetectionCache = new ConcurrentDictionary<string, bool>();
    public static List<bool> ng_rec = new List<bool>();
    public static List<string> p2label_name = new List<string>();
    public static List<Label> p2label = new List<Label>();
    public static List<DateTime> finishtime = new List<DateTime>();
    public static DateTime rec_time;
    public static DateTime piece_count;
    public static DateTime lastIn = new DateTime();
    public static int null_count = 0;
    public static int ng_count = 0;
    public static int ok_count = 0;

    // 新增系統狀態枚舉
    public enum SystemState
    {
        Stopped,           // 停止狀態 - 可以開始檢測
        Running,           // 運行中 - 只能停止檢測
        StoppedNeedUpdate, // 停止後需要更新計數
        UpdatedNeedReset,  // 更新計數後需要異常復歸
    }
    public static SystemState currentState = SystemState.Stopped;

    public static string username = "", type = "";
    public static int DetectMode = 0, user = 3; //user:(0=工程師、1=管理者、2=作業員、3=未登入)
    public static bool status = false, live = false, paramUpdate = false; //機台運行狀態
    public static bool SoftTriggerMode = false;
    public static bool Mode = false;
    
    // 由 GitHub Copilot 產生 - 新增調機模式旗標
    public static bool isAdjustmentMode = false; // 調機模式旗標（parameter_info 使用）

    public static float okNgThreshold = 0.5f;   // OK/NG 判定閾值
    public static int margin = 45;             // 給 boundingRect 擴張的 margin
    public static int size = 256;
    public static Size targetSize = new Size(size, size); // 最終要給分類模型的輸入大小 (256×256)
    public static double iouThreshold = 0.3;   // 用於合併Rect(重疊度)
    //public static int shrink = 0;

    public static bool setComplete = true;
    public static bool start_error = false;
    public static bool okswitch = false;

    public static bool offline = true;
    public static bool usekey = false;
    public static bool alertTriggered = false; // 新增標記
    public static bool reseting = false;
    public static bool hasPerformedRecovery = false;
    public static bool enableProfiling = false; // 是否啟用效能測量
    public static HashSet<int> profilingStations = new HashSet<int>(); // 要測量的站點
    public static int postPackNullQuota = 0; // 滿箱後強制NULL的剩餘配額
    public static int lastD98Accepted = -1;
    public static int ProductiveTimeoutSec = 10;

    public static bool testc = false;
    public static bool testd = false;
    public static bool test = true;
    public static int mode = 0;
    public static bool allOK = false;
    // 由 GitHub Copilot 產生
    // 修正: 簡化同步機制，使用簡單旗標
    public static bool isInSyncMode = false;              // 是否在同步模式
    public static DateTime syncStartTime = DateTime.MinValue; // 同步開始時間
    public static int syncWaitTimeMs = 0;                 // 同步等待時間（毫秒）

    public static int NGmax = 0, NULLmax = 0;
    public static int pack = 0, order = 0;
    public static string LotID = "";
    public static string produce_No = "";
    public static string produce_innerModelPath = "";
    public static string produce_outerModelPath = "";
    public static string produce_innerServerUrl = "";
    public static string produce_inner_NROI_ServerUrl = "";
    public static string produce_outerServerUrl = "";
    public static string produce_outer_NROI_ServerUrl = "";
    public static string produce_chamferModelPath = "";
    public static string produce_chamferServerUrl = "";
    public static string produce_station1ServerUrl = "";
    public static string produce_station2ServerUrl = "";
    public static bool has_NROI_InnerModel = false;
    public static bool has_NROI_OuterModel = false;

    public static int continue_NG = 0, continue_NULL = 0;
    public static string foldername = "";
    public static string cur_pic_name = "";
    public static int empty = 0;

    public static InferenceSession classifyNet;
    public static string classifyNetName = "";
    public static ResultManager resultManager = new ResultManager();

    public static int output_result = 0;
    public static bool plc_stop = false;
    public static int okr = 0;

    public static DateTime lastUpdateTime = DateTime.MinValue;
    // 最小更新间隔，单位为毫秒（这里设置为 33 毫秒，大约 30 帧每秒）
    public static readonly TimeSpan minUpdateInterval = TimeSpan.FromMilliseconds(33);
}

