================================================================================
                  peilin AOI 系統  (bushingAOI / 培林光學檢測)
                          專案完整說明文件 (readme.txt)
================================================================================

最後更新日期 : 2026-05-14
目標框架     : .NET Framework 4.8 (x64)
應用類型     : C# Windows Forms (WinExe)
組件名稱     : peilin.exe
專案 GUID    : {9E92D6B9-AB70-496D-82BC-D11D8EC96ACC}
程式進入點   : Program.cs -> peilin.Form1
建置輸出     : bin\x64\Release\peilin.exe


================================================================================
[1] 專案總覽 (Project Overview)
================================================================================

本系統為工業用「襯套 (Bushing)」自動光學檢測 (AOI, Automated Optical
Inspection) 應用程式。透過 Basler 工業相機擷取產品影像、以 Python YOLO HTTP
伺服器執行 AI 推論進行物件 / 瑕疵偵測，配合 PLC (Modbus RTU) 訊號控制氣動
排出模組，於產線上即時分類良品 / 不良品 (NG) / 空白 (NULL)，並回寫統計資料
至 SQLite 資料庫。

註: 程式碼中雖保留 anomalyTensorRT.cs 與 AD_TRT_dll1~4.dll 之 P/Invoke 介
    面，但目前主流程「未實際呼叫」這些 TensorRT 介面 (Inference1~4 /
    CreateModel1~4 全程無外部呼叫點)，實際 AI 推論完全透過 YoloDetection.cs
    經由 HTTP 送至本機 Python 伺服器執行。TensorRT 相關檔案為歷史遺留
    (legacy)，CLAUDE.md 亦載明「now using .pt models with dynamic loading」。

系統可同時管理「最多 4 個檢測站」(Stations)，每站獨立運作但共用同一
條 PLC 串列埠、同一個資料庫與全域 `app` 狀態。

主要應用場景:
  - PTFE / 金屬複合襯套品檢
  - 內外徑 (ID / OD)、高度 (H)、溝槽 (Groove) 規格驗證
  - 表面瑕疵 (污漬、皺褶、開口、缺料、變形) 即時偵測


================================================================================
[2] 系統架構 (System Architecture)
================================================================================

------------------------------------ 處理管線 ----------------------------------

  [PLC 觸發訊號]
        |
        v
  [相機擷取 (Basler Pylon)]  --(GrabbedEventArgs)-->  Camera0.Receiver()
        |
        v
  [影像佇列 Queue_Bitmap1~4]  ------- 4 站獨立 -------
        |
        v
  [位置偵測 / ROI 擷取]  (algorithm.SafeROI, 圓心/半徑/倒角)
        |
        v
  [AI 推論]   (實際使用)
     +-- YOLO 物件 / 瑕疵偵測 (HTTP 5001~5005, Python 伺服器)
           -> JSON 回傳 bounding boxes + class id + 信心度

     [遺留未用]
     -- anomalyTensorRT.cs + AD_TRT_dll1~4 (P/Invoke 宣告保留, 主流程未呼叫)
        |
        v
  [缺陷判定]  (DefectCheck 表逐站閾值比對)
        |
        v
  [Queue_Save / Queue_Send / Queue_Show]
        |        |          |
        |        |          +--> UI 顯示 (CherngerPictureBox)
        |        +-------------> PLC 氣吹排出指令 (Modbus 寫入)
        +----------------------> 影像存檔 (report\{料號}\) + DB 統計

------------------------------------ 模組層級 ----------------------------------

  UI 層       : Form1, ParameterConfigForm, 各校正 Form
  業務邏輯層  : ParameterSetupManager, algorithm, YoloDetection, PytorchClient
  通訊層      : Camera0 (Basler), PLC_Test (Modbus)
  資料層      : LinqToDB + SQLite (mydb.sqlite)
  外部服務    : Python YOLO 伺服器 (透過 .bat 啟動)  -- 唯一實際運作的 AI 後端
  原生 DLL    : SLUG2DLL.dll (智慧卡認證)
  遺留 (未使用): anomalyTensorRT.cs / AD_TRT_dll1~4.dll  (歷史 P/Invoke, 無呼叫)


================================================================================
[3] 核心模組說明 (Core Modules)
================================================================================

[Form1.cs]               17,000+ 行, 系統總協調者
                         包含全域 static `app` 類別:
                           - Queue_Bitmap1~4 / Queue_Save / Queue_Send / Queue_Show
                           - counter / dc 計數
                           - param / models / metas / pos 參數快取
                           - _wh1~4 / _sv / _reader / _AI / _show 等待控制
                           - pendingOK1, pendingOK2, pendingPushOK1, pendingPushOK2
                             (PLC 計數對應的待推樣品 ID 佇列)
                           - currentState 系統狀態列舉

[Camera0.cs]             Basler Pylon 相機封裝 (basler 命名空間)
                         - m_BaslerCameras[4]  4 台相機陣列
                         - InitCamera / StartGrabbing / StopGrabbing / GrabImage
                         - 事件驅動式 ImageGrabbedEventArgs
                         - 觸發時間、最小間隔檢查、誤觸防護
                         - CameraLog_yyyymmdd.txt / CameraWarning 獨立日誌

[anomalyTensorRT.cs]     [遺留] TensorRT 異常偵測封裝 (AnomalyTensorRT 命名空間)
                         - 含 AD_TRT_dll1~4.dll 之 P/Invoke 宣告
                           (single_inference / create_model / close)
                         - 對應的 TensorRT.SingleTest1~4 方法 -- 主程式從未呼叫
                         - Form1.cs 第 72-103 行雖宣告對應 DllImport, 但
                           Inference1~4 / CreateModel1~4 同樣零呼叫
                         - 實際 AI 推論已改由 YoloDetection.cs (HTTP)取代

[YoloDetection.cs]       YOLO HTTP 客戶端
                         - 與 Python 伺服器 (localhost) 透過 JSON / Base64 通訊
                         - LoadYoloModel(modelPath, serverBaseUrl)
                         - PerformObjectDetection(Mat image, serverUrl)
                         - 自動暖機 (warmup) 與重新載入機制
                         - 連接埠分配:
                             5001  內環檢測 (inner contour)
                             5002  外環檢測 (outer contour)
                             5003  內環異常檢測 (inner NROI)
                             5004  外環異常檢測 (outer NROI)
                             5005  倒角檢測 (chamfer)

[PytorchClient.cs]       通用 PyTorch 伺服器啟動 / 健康檢查工具
                         - StartPythonServer(batFilePath)
                         - IsServerAvailable(serverBaseUrl)

[PLC_Test.cs]            PLC Modbus RTU 序列埠通訊
                         - 支援暫存器類型: X / Y / M / D / T / C / S / ECnX / ECnY
                         - 關鍵暫存器:
                             D803  站 1 OK 計數 (32-bit)
                             D805  站 2 OK 計數 (32-bit)
                             D807  站 3 OK 計數
                             D809  站 4 OK 計數
                         - 優先佇列 (priority queue) 機制
                             氣吹排出指令最高優先 → 避免 PLC 受阻
                         - 1 秒監控 timer 偵測計數上升並回寫 log

[algorithm.cs]           - SafeROI()  邊界安全的 ROI 裁切, 避免溢出當機

[ParameterConfigForm.cs] 參數管理三區介面
[ParameterSetupManager.cs]
                         - Reference Zone        (來源料號參數, 唯讀)
                         - Added-Unmodified Zone (已複製尚未編輯)
                         - Added-Modified Zone   (已編輯待儲存)
                         - 不同顏色標示狀態:  白 / 淺綠 / 淺藍
                         - 參數分類: Camera / Position / Detection / Timing /
                                     Testing

[校正工具 (6 種)]
   CircleCalibrationForm    圓心 / 半徑 / 倒角  (內圓 / 外圓 / 倒角)
   ContrastCalibrationForm  對比度 / 亮度  即時滑桿預覽
   PixelCalibrationForm     像素 - 公釐換算 (PixelToMM)
   WhiteCalibrationForm     白色像素閾值 (含 IQR 離群分析)
   ObjectBiasCalibrationForm  物體中心偏移 (objBias 中位數 / 標準差)
   gapThreshCalibrationForm 開口閾值, 內彎 / 外凸像素檢測

[Program.cs]             主進入點, STAThread, Application.Run(new Form1())


================================================================================
[4] 資料庫結構 (SQLite Schema)
================================================================================

資料庫檔案 : .\setting\mydb.sqlite
ORM        : LinqToDB 4.1 (主要) + EntityFramework 6.4.4 (備援)
連線字串   : Data Source=.\setting\mydb.sqlite  (見 App.config)
連線模式   : 支援 WAL 並發

主要資料表:
  Type           料號規格表  (TypeColumn, ID, OD, H, hasgroove, PTFEColor,
                              material, boxorpack, thick, hasYZP, package)
  Camera         各料號相機參數  (exposure / gain / delay)
  Param          演算法參數     (4 站獨立或共享, 由 type / stop 欄位區分)
  Parameter      影像存檔等全域選項 (boolean)
  DefectType     瑕疵種類總表
  DefectCheck    料號 x 站 x 瑕疵 之啟用旗標與閾值
  DefectCount    即時統計 (lot_id, name, count, time)
  Blow           氣吹排出位置 / 時序設定
  Total          總抽樣計數
  User           登入帳號 (user_name / password / level)
  LastSetting    上次工作階段值 (last_no_index, last_user_index,
                                  last_lotid, last_order, last_pack,
                                  last_NGLimit, last_NULLLimit)

資料存取規範 (重要):
  必須使用「Update-First, Insert-If-None」模式以避免 PK 衝突:
    int updated = db.table.Where(...).Set(...).Update();
    if (updated == 0) db.table.Value(...).Insert();
  不建立 entity 物件做 Insert; 全用 .Value() 串接。


================================================================================
[5] 設定檔與外部資源
================================================================================

App.config
  - connectionStrings  mydb (SQLite)
  - userSettings       ExposureClock / Gain
  - assemblyBinding    Newtonsoft.Json 13.0, OnnxRuntime 1.20.1,
                       System.Memory 4.0.1.2, ...

bin\x64\Release\setting\
  mydb.sqlite                 主資料庫 (含多份備援: 1mydb.sqlite, 2mydb.sqlite,
                              「(AI參數動刀前)」「(砍param前)」等)
  camera.txt / param.txt      文字匯出
  type.txt / mydb.csv         手動匯出
  param_session_*.json        每個料號的工作階段快照
                              (000 = 預設 / 各料號各一份)
  temp\                       暫存

bin\x64\Release\models\
  {料號}_{in|out|1|2}.pt       PyTorch 模型 (供 Python YOLO 伺服器載入)
  {料號}_*.json                模型中繼資料 (正規化參數)

bin\x64\Release\
  peilin.exe                  主程式
  peilin.exe.config           執行期設定
  SLUG2DLL.dll / SLUG2DLL64.dll  智慧卡 (SmartKey) 認證
  opencv_world455.dll         OpenCV 4.5.5 核心
  onnxruntime.dll             ONNX Runtime 1.20.1 (NuGet 自動 copy, 主流程未直接使用)
  AD_TRT_dll1~4.dll           [遺留] TensorRT 推論 DLL -- 程式不會載入
  yolov5.dll                  [遺留] YOLOv5 原生支援
  zlibwapi.dll                [遺留] TensorRT 相依壓縮庫
  openvino_dldt.*             [遺留] OpenVINO 後端 (備援推論, 未啟用)
  start_server.bat            主 YOLO 伺服器啟動腳本
  yolo-server.py              Python YOLO 伺服器主程式
  yolo_{料號}_{站號}.bat       各料號各站啟動腳本 (目前 3 料號 x 4 站 = 12 支)
                              料號: 10201120714TP, 10214090420T, 10215112360T
  best.pt / best0.pt          通用 YOLO 模型
  report\                     檢測影像 / 直方圖 / OK / NG / NULL 分桶
  logs\                       Serilog 日誌
  Statistics\                 統計報表 (Excel)


================================================================================
[6] 建置 (Build)
================================================================================

開發工具      : Visual Studio 2019/2022 (Solution Format Version 12.00)
方案檔        : peilin.sln    (peilin.csproj)
語言版本      : C# 7.3
建置設定      : Release | x64  (建議)

命令列建置:
  msbuild peilin.sln /p:Configuration=Release /p:Platform=x64

IDE 建置:
  1. 開啟 peilin.sln
  2. 設定 Configuration = Release, Platform = x64
  3. 還原 NuGet 套件 (約 30+ 套件, 見 packages.config)
  4. 建置 → 輸出至 bin\x64\Release\peilin.exe

NuGet 主要依賴:
  - OpenCvSharp4 4.5.3              影像處理
  - LijsDev.Basler.Pylon.NET.x64    7.1.0.25066  工業相機
  - Basler.Pylon / Basler.PylonC.NET            相容性
  - linq2db / linq2db.SQLite 4.1.0  ORM
  - EntityFramework 6.4.4           ORM 備援
  - Stub.System.Data.SQLite.Core    1.0.116.0    SQLite Provider
  - Microsoft.ML.OnnxRuntime 1.20.1 ONNX 推論
  - Newtonsoft.Json 13.0.1          JSON
  - NPOI 2.5.5                      Excel 匯出
  - Serilog 2.11 + Sinks.File/Console   結構化日誌
  - Portable.BouncyCastle 1.8.9     加密
  - System.Text.Json 9.0.0
  - Microsoft-WindowsAPICodePack    Windows Shell

其他相依 (本機 / 內部 DLL, 不在 NuGet):
  - CcvLib.dll, CherngerControls.dll, CherngerTools.dll  自製控制項與工具
  - ConcurrentList.dll                                  自製並行集合
  - AS.Video.Core.dll                                   影像播放


================================================================================
[7] 執行 (Run)
================================================================================

前置需求:
  1. Windows 10 / 11 x64
  2. .NET Framework 4.8 Runtime
  3. Basler Pylon Runtime (相容於 SDK 7.x) 或對應相機驅動
  4. Python 3.x + Ultralytics YOLO + PyTorch  (yolo-server.py 啟動所需)
     -- 主程式的 AI 推論完全依賴此 Python 伺服器, 沒有它系統等同失明
     -- 若 Python 端使用 GPU 加速, 需配置對應 NVIDIA 驅動 + CUDA
  5. 智慧卡 (SmartKey) Dongle —— 預設啟用; 透過 app.usekey 可關閉
  6. SQLite 資料庫位於 .\setting\mydb.sqlite

(註: AD_TRT_dll1~4.dll 雖存在於 Release 目錄, 但程式不會載入,
     因此「不」需要 TensorRT 或對應 CUDA 版本)

啟動順序:
  a) 執行 peilin.exe
  b) Serilog 初始化, 寫入 .\logs\peilin_log-YYYYMMDD.txt
  c) SQLite 連線開啟  (App.config 之 mydb 連線字串)
  d) 智慧卡 (SmartKey) 認證 (若 app.usekey = true)
  e) 載入料號清單, 顯示登入畫面 (login.cs)
  f) Form1_Load 初始化 4 站佇列 / 等待控制
  g) 依當前料號透過 yolo_{料號}_{站號}.bat 啟動 4 個 Python 伺服器
  h) Basler 相機初始化與軟體觸發
  i) PLC 監控 Timer 啟動 (1 秒, 監看 D803/D805/D807/D809)
  j) 系統狀態為 UpdatedNeedReset — 操作員需先執行「異常復歸」

啟動失敗常見原因:
  - mydb.sqlite 缺失或路徑錯誤
  - SmartKey 未插或被防毒軟體攔截
  - 連接埠 5001~5005 被佔用  →  YOLO 伺服器啟動超時
  - Python 環境缺失 / PyTorch 版本不符  →  yolo-server.py 啟動失敗
  - Basler 相機未供電或 IP 未配置


================================================================================
[8] 多站台同步運作機制
================================================================================

站台 ID 規則:
  - camID = 0 ~ 3, 4 個檢測站
  - 對應 m_BaslerCameras[camID]
  - 對應 Queue_Bitmap{camID+1}
  - 對應 _wh{camID+1} 等待控制
  - 對應 YOLO 連接埠  5001 ~ 5004 (+ 5005 倒角)
  - 對應 PLC 計數暫存器  D803 / D805 / D807 / D809
  - (anomalyTensorRT 雖然也按站區分 Inference1~4, 但主流程未呼叫)

共用資源 (需注意執行緒安全):
  - PLC 串列埠 (單一 SerialPort, 由優先佇列序列化)
  - SQLite 連線池 (LinqToDB 自動管理)
  - 全域 `app` 靜態狀態 (Queue_Save / counter / param)
  - 共享拋棄佇列 (_disposeQueue) 與 _disposeSemaphore (上限 4)


================================================================================
[9] 重要程式碼規範 (Code Style)
================================================================================

(必要規則, 已寫於 CLAUDE.md / CLAUDE-zh-tw.md)

語言:
  - 所有註解 / AI 助理回覆統一使用「繁體中文」
  - AI 產生的程式碼區段起始須加註:  // 由 GitHub Copilot 產生
  - 禁止在原始碼中使用 emoji

命名空間衝突:
  - System.Drawing.* 與 OpenCvSharp 皆有 Point / Size / Rect
  - 必須完整指定:  System.Drawing.Point p = new System.Drawing.Point(x, y);

記憶體管理 (關鍵):
  - OpenCvSharp Mat 使用完畢必須 Dispose() — 高速產線下 Mat 累積 = 必當機
  - System.Drawing.Bitmap 也必須 .Dispose()
  - 影像處理盡量使用 using 語句
  - 大型 / 長壽物件使用壓縮儲存
  - 透過 _disposeQueue + _disposeSemaphore 非同步釋放

資料庫:
  - 永遠先 Update, 0 筆才 Insert
  - 用 .Value() 串接, 不建立 entity 物件
  - 全部 using 包覆 DB 連線

PLC 通訊:
  - 氣吹排出指令最高優先
  - 不可在序列埠處理中 GC.Collect() (歷史教訓: 造成佇列堆積與計數失準)


================================================================================
[10] 日誌 (Logging)
================================================================================

框架      : Serilog 2.11
輸出格式  : "{Timestamp:HH:mm:ss} [{Level:u3}] {Message}{NewLine}{Exception}"
分割策略  : 每日 1 檔  (Sinks.File RollingInterval = Day)

主要日誌:
  .\logs\peilin_log-YYYYMMDD.txt          系統主日誌
  CameraLog_YYYYMMDD.txt                  相機事件 (擷取 / 觸發)
  CameraWarning_YYYYMMDD.txt              相機警告 (誤觸 / 超時)
  camera_trigger.log                      觸發訊號歷史


================================================================================
[11] 操作工作流程 (Operational Workflows)
================================================================================

(A) 新增料號:
   1. type_info 表單新增料號 (TypeColumn, ID, OD, H, ...)
   2. ParameterConfigForm 選擇來源料號 → 複製參數
   3. 各分類調整: Camera / Position / Detection / Timing / Testing
   4. 依需要執行對應校正工具 (Circle / Contrast / Pixel / White /
      ObjectBias / GapThresh)
   5. 儲存至資料庫 (只寫入 Added-Modified 區, 保留 Reference 區)
   6. 放置 AI 模型: bin\x64\Release\models\{料號}_*.pt
   7. 如需 YOLO 物件偵測, 建立 yolo_{料號}_{站號}.bat
   8. 用樣本影像驗證, 觀察 NG / OK / NULL 分佈

(B) 修改檢測參數:
   - 開啟 ParameterConfigForm → 載入目標料號
   - Reference 可選自己 (作為基準快照)
   - 在 Added 區編輯, 必要時開啟校正工具
   - 儲存後系統自動回寫 mydb.sqlite 與 param_session_*.json

(C) 除錯影像處理問題:
   - 啟用 Parameter.存圖
   - 檢視 .\report\{料號}\ 內的影像
   - 對照 .\logs\peilin_log-*.txt 時間軸
   - 檢視 CameraLog_*.txt 確認觸發節奏
   - PLC 通訊問題: 檢查 D803/D805/D807/D809 累計與軟體 counter 是否一致


================================================================================
[12] 已知限制與注意事項
================================================================================

- 緊急停止後可能造成 PLC 推桿計數與軟體 counter 相差 1~2 筆 (硬體特性)
- AI 推論完全依賴 Python YOLO 伺服器, 主程式本身沒有內嵌推論能力
- AD_TRT_dll1~4.dll / yolov5.dll / openvino_dldt.* 等檔案為歷史遺留,
  雖然仍隨 Release 一起打包, 但實際執行階段不會被載入
- 智慧卡認證若被防毒軟體攔截, 主程式無法啟動
- YOLO 伺服器以 Python 子行程模式啟動, 主程式關閉後須手動殺掉殘留 python.exe
- bin\x64\Release\setting\ 內有多份歷史備份 mydb.sqlite, 部署前應整理


================================================================================
[13] 文件索引 (Documentation Index)
================================================================================

[根目錄]
  CLAUDE.md             英文版開發指引 (AI 助理規範 + 架構摘要)
  CLAUDE-zh-tw.md       繁體中文版開發指引 (與 CLAUDE.md 內容同步)

[docs\]
  readme.txt                                    本檔
  IMPLEMENTATION_SUMMARY.md                     參數三區重構詳記
  ConcurrentDictionary_Fix_Summary.md           並行字典修補摘要
  code-review.md / code-review-zh-tw.md         程式碼審查紀錄
  copilot-instructions.md                       Copilot 指令範本
  dataset_management_recommendations.md         資料集管理建議
  取像定位流程.txt                              取像 / 定位流程說明
  memory-fixes\                                 歷次記憶體修復紀錄

[調整參數手冊\]    (18 份, MD/TXT 雙版本)
  01.基本說明
  02.檢測參數設定完整操作手冊
  03.完整操作檢查清單
  10.相機參數
  11.位置參數
  12.檢測參數
  13.時間參數
  20.檢測精度校正(使用原始圖像)
  21.白色像素校正(使用原始圖像)
  22.對比檢測校正(使用ROI圖像)
  23.物體偏移校正(使用ROI圖像)
  24.開口閾值校正(使用ROI圖像)
  30.小技巧
  31.常見問題
  新增料號流程
  更換料號流程

[參數手冊doc\]    同上之 Word (.docx) 版本


================================================================================
[14] 目錄結構 (Top-level Tree)
================================================================================

bushingAOI\
  ├─ peilin.sln                     方案檔
  ├─ peilin.csproj                  專案檔
  ├─ App.config                     應用組態
  ├─ packages.config                NuGet 清單
  ├─ Program.cs                     主進入點
  ├─ Form1.cs / Form1.Designer.cs   主視窗 (17,000+ 行)
  ├─ Camera0.cs                     相機封裝
  ├─ PLC_Test.cs / .designer.cs     PLC Modbus
  ├─ anomalyTensorRT.cs             [遺留, 未呼叫] TensorRT P/Invoke 包裝
  ├─ YoloDetection.cs               YOLO HTTP 客戶端  (實際 AI 入口)
  ├─ PytorchClient.cs               Python 伺服器啟動 / 健康檢查工具
  ├─ algorithm.cs                   工具方法 (SafeROI)
  ├─ ParameterConfigForm.cs         三區參數 UI
  ├─ ParameterSetupManager.cs       三區參數邏輯
  ├─ ParameterModels.cs             參數模型類別
  ├─ CircleCalibrationForm.cs       校正 - 圓
  ├─ ContrastCalibrationForm.cs     校正 - 對比
  ├─ PixelCalibrationForm.cs        校正 - 像素
  ├─ WhiteCalibrationForm.cs        校正 - 白平衡
  ├─ ObjectBiasCalibrationForm.cs   校正 - 物體偏移
  ├─ gapThreshCalibrationForm.cs    校正 - 開口閾值
  ├─ SaveConfirmDialog.cs           儲存確認對話框
  ├─ SourceSelectionDialog.cs       來源料號選擇
  ├─ login.cs                       登入畫面
  ├─ mbForm.cs                      訊息框
  ├─ alert.cs                       警告對話框
  ├─ blow_info.cs                   氣吹參數編輯
  ├─ camera_info.cs                 相機參數編輯
  ├─ defect_check_info.cs           瑕疵檢測對應編輯
  ├─ defect_type_info.cs            瑕疵類型維護
  ├─ keepday.cs                     資料保留期設定
  ├─ parameter_info.cs              通用參數編輯
  ├─ type_info.cs                   料號維護
  ├─ user_info.cs                   使用者維護
  ├─ delaybutton.cs                 自訂延遲按鈕控制項
  ├─ DummyForm.cs                   佔位用
  ├─ MemoryLeakTest.cs              記憶體洩漏測試
  ├─ VectorOfString2.cs             相容包裝
  ├─ onnxTest.cs / onnx_Test.cs     舊版 ONNX 測試
  ├─ testAOI.cs / testAOI2.cs       舊版 AOI 測試
  ├─ testPerPixel.cs                舊版逐像素測試
  ├─ testroi.cs                     舊版 ROI 測試
  ├─ mydb\                          LinqToDB T4 樣板 + 產生程式碼
  ├─ Properties\                    AssemblyInfo / Resources / Settings
  ├─ Resources\                     圖檔 (應用程式 Logo 等)
  ├─ LinqToDB.Templates\            T4 樣板組
  ├─ archive\                       舊版備份 (legacy-cs 等)
  ├─ packages\                      NuGet 套件
  ├─ obj\                           建置中繼物
  ├─ bin\
  │   └─ x64\Release\               執行目錄 (見 §5)
  ├─ docs\                          技術文件 (含本檔)
  ├─ 調整參數手冊\                  使用者手冊 (Markdown / TXT)
  ├─ 參數手冊doc\                   使用者手冊 (Word)
  ├─ logs\                          開發期日誌
  └─ mydb\                          T4 樣板


================================================================================
[15] 版權與作者
================================================================================

組件版本    : 1.0.0
圖示檔      : alhrw-pnvie-001.ico
Logo        : 內嵌於 Properties\Resources.resx

開發備註    : Git 主分支為 main, 開發過程請遵守 CLAUDE.md 之 commit 格式
              (type(scope): subject + body 條列每檔變更與理由)。

================================================================================
                                  - END -
================================================================================
