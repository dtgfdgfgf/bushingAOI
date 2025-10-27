# CLAUDE.md (繁體中文版)

此檔案為 Claude Code (claude.ai/code) 在本專案中工作時的指導文件。

## 專案概述

這是一個用於工業軸承製造品質控制的 C# Windows Forms AOI (自動光學檢測) 應用程式。系統整合了 Basler 相機、PLC 通訊、YOLO 物件偵測,以及使用 TensorRT 的異常偵測來即時檢測產品。

**目標框架:** .NET Framework 4.8 (x64)
**主要入口:** `Program.cs` → `Form1`
**建置輸出:** `bin\x64\Release\peilin.exe`

## 建置與執行

### 建置專案
```bash
# 在 Visual Studio 中建置
msbuild peilin.sln /p:Configuration=Release /p:Platform=x64

# 或使用 Visual Studio IDE (建議)
# 將設定設為 "Release" 平台設為 "x64"
```

### 執行應用程式
- 執行 `bin\x64\Release\peilin.exe`
- 應用程式需要 SQLite 資料庫位於 `.\setting\mydb.sqlite`
- 智慧卡加密狗認證預設啟用 (可透過 `app.usekey` 停用)

### 外部相依性
- **YOLO 的 Python 伺服器**: 透過 `bin\x64\Release\yolo_*.bat` 批次檔啟動
- **TensorRT DLL**: `AD_TRT_dll1.dll` 至 `AD_TRT_dll4.dll` 用於異常偵測
- **Basler Pylon SDK**: 相機操作所需

## 系統架構

### 多站設計
系統同時支援最多 4 個檢測站:
- 每個站有自己的 Basler 相機 (透過 `Camera0.cs` 設定)
- 每個站執行獨立的 AI 推論 (TensorRT + YOLO)
- 共享的 PLC 通訊透過序列埠協調所有站 (Modbus 協定)

### 核心處理管線
```
相機觸發 → 影像擷取 → 位置偵測 →
ROI 提取 → AI 推論 (異常 + YOLO) →
缺陷分類 → 結果儲存 → PLC 訊號
```

### 關鍵架構元件

**Form1.cs** (17,000+ 行)
- 主應用程式表單與系統協調
- 包含 `app` 靜態類別與全域狀態 (佇列、計數器、參數)
- 影像處理管線協調
- PLC 通訊處理器

**Camera0.cs**
- Basler Pylon 相機包裝器
- 多相機管理 (`m_BaslerCameras` 陣列)
- 事件驅動的影像擷取與監控/記錄

**anomalyTensorRT.cs**
- 4 個 TensorRT DLL 實例的包裝器 (`AD_TRT_dll1` 至 `AD_TRT_dll4`)
- 異常偵測模型載入與推論
- 回傳異常熱圖 (Mat) 與分數 (float)

**YoloDetection.cs**
- 基於 Python 的 YOLO 伺服器的 HTTP 客戶端
- 模型載入、預熱與批次偵測
- 與 localhost 伺服器在埠 5001-5004 通訊

**PLC_Test.cs**
- 與 PLC 的 Modbus 序列通訊
- 處理觸發訊號、計數器 (D803, D805, D807) 與吹氣指令
- 優先佇列系統防止 PLC 指令擁塞

**ParameterConfigForm.cs / ParameterSetupManager.cs**
- 參數管理 UI 採用 3 區系統:
  - 參考區 (來源料號參數,唯讀)
  - 新增未修改區 (已複製但尚未編輯)
  - 新增已修改區 (已編輯參數準備儲存)
- 位置/圓形/像素參數的校正工具整合

**algorithm.cs**
- SafeROI 輔助方法,用於邊界安全的影像區域提取

### 資料庫結構 (SQLite)
位於 `.\setting\mydb.sqlite` 包含以下表格:
- **Cameras**: 每個料號的相機參數 (曝光、增益、延遲)
- **params**: 每個料號的演算法參數
- **Types**: 料號資訊
- **DefectTypes**: 全域缺陷類型定義
- **DefectChecks**: 特定料號的缺陷偵測對應
- **DefectCounts**: 缺陷統計報表
- **Totals**: 生產計數器
- **Blows**: 吹氣分選參數
- **Users**: 使用者認證

## 程式碼風格要求

### 最小變更原則
**關鍵:** 修改或新增程式碼時，遵循最小變更原則。

**理念:**
- 以最小的變更達成所需功能
- 盡可能保留現有的程式碼結構與模式
- 僅修改必要的部分以修復錯誤或新增功能
- 尊重現有的架構與設計決策

**這並不代表可以犧牲程式碼品質或架構:**
- 維持適當的錯誤處理、記憶體管理與執行緒安全
- 遵循本文件中的所有其他指導方針
- 僅在對正確性或關鍵可維護性絕對必要時才重構
- 從一開始就以適當的架構設計新功能

**實際應用:**
- **錯誤修復**: 僅變更造成問題的特定程式碼行
- **新功能**: 整合到現有模式而非重寫周圍程式碼
- **參數新增**: 新增欄位而不重組現有資料流
- **效能改善**: 優化特定瓶頸,而非整個子系統
- **重構**: 僅在程式碼損壞、無法維護或造成生產風險時

**範例 - 新增參數:**
```csharp
// 良好: 最小變更 - 在現有參數旁新增參數
if (updatedRows == 0)
{
    db.params.Value(p => p.PartNumber, partNum)
             .Value(p => p.ExistingParam, existingVal)
             .Value(p => p.NewParam, newVal)  // 僅新增此行
             .Insert();
}

// 避免: 不必要地重組運作正常的程式碼
// 不要為了新增一個欄位就重寫整個參數載入系統
```

**平衡:** 目標是在生產系統中保持穩定性與可預測性。進行外科手術式的變更來解決問題而不引入新風險。然而,絕不為了最小變更而犧牲關鍵要求(記憶體安全、執行緒安全、資料完整性)。

### 程式碼變更溝通
**提議或進行程式碼變更時，必須提供:**

1. **精確位置**: 檔案路徑、類別名稱、方法名稱、行號/範圍
2. **明確理由**: 為何需要此變更 (錯誤修復、新功能、效能、安全性)
3. **變更範圍**: 將修改什麼、什麼會保持不變
4. **影響評估**: 哪些元件/執行緒/工作流程會受影響

**範例格式:**
```
位置: Form1.cs:1234 在 ProcessImage() 方法中
理由: 記憶體洩漏 - YOLO 推論後 Mat 物件未釋放
變更: 在 Mat 建立處新增 using 陳述式 (第 1234-1240 行)
影響: 僅影響 YOLO 處理路徑;不影響異常偵測或 PLC 通訊
```

**這確保了:**
- 實作前清楚了解將變更什麼
- 最小化非預期的修改
- 更好的審查與除錯能力
- 變更歷史的文件化

### 語言與註解
- **所有註解必須使用繁體中文**
- 在 AI 生成程式碼段落開頭加上註解: `// 由 GitHub Copilot 產生`

### 資料庫操作 (LinqToDB)
始終使用「先更新,若無則插入」模式避免主鍵衝突:
```csharp
// 嘗試更新現有記錄
int updatedRows = db.table.Where(條件).Set(欄位, 值).Update();

// 若無更新記錄則插入新記錄
if (updatedRows == 0)
{
    db.table.Value(欄位, 值).Insert();
}
```
- 避免為 Insert 建立實體物件;使用 `.Value()` 語法
- 資料庫連線始終使用 `using` 陳述式
- 所有資料庫操作都包含適當的錯誤處理

### 記憶體管理
**關鍵:** 此應用程式以工業速度處理高解析度影像。不當的記憶體管理會導致當機。

#### 嚴格的記憶體規範
- **OpenCV Mat 物件**: 使用後必須立即釋放
- **System.Drawing.Bitmap**: 必須明確呼叫 `.Dispose()`
- 影像處理操作使用 `using` 陳述式
- 長期存在的大型物件應使用壓縮儲存
- 應用程式使用 `_disposeQueue` 與 `_disposeSemaphore` 進行非同步釋放

#### 記憶體超載預防
**此系統由於以工業速度處理高解析度影像而在極端記憶體限制下運作。記憶體超載會導致應用程式當機與生產線故障。**

**強制性做法:**
1. **立即釋放**: 盡可能在相同方法範圍內釋放 Mat/Bitmap 物件
2. **佇列管理**: 監控佇列深度 (`Queue_Bitmap1-4`, `Queue_Save`, `Queue_Send`, `Queue_Show`) 以防止記憶體累積
3. **大型物件生命週期**: 追蹤並限制 >85KB 物件的生命週期 (大型物件堆積閾值)
4. **避免 GC.Collect()**: 除非絕對必要且有文件記載,否則絕不明確呼叫 `GC.Collect()`;讓執行時期管理回收
5. **影像複製最小化**: 避免不必要的影像複製;安全時使用參考,複製時釋放原始物件
6. **記憶體剖析**: 定期在生產負載條件下剖析記憶體使用

### 多執行緒管理
**關鍵:** 此應用程式執行 4 個以上具有共享資源的並行處理管線。執行緒安全違規會導致資料損壞與當機。

#### 執行緒安全要求
**系統協調多個執行緒處理相機饋送、AI 推論、PLC 通訊與 UI 更新。不當的同步化會導致:**
- 共享狀態的資料競爭
- 資源洩漏 (未釋放的鎖、未釋放的物件)
- 跨 CPU 核心的快取不一致
- PLC 通訊的死鎖

**強制性做法:**

1. **同步化原語**:
   - 所有 `app` 類別靜態狀態修改使用 `lock` 陳述式
   - 正確使用等待控制代碼 (`_wh1-4`, `_sv`, `_reader`, `_AI`, `_show`) 配合逾時模式
   - 非同步操作使用 `SemaphoreSlim` (例如 `_disposeSemaphore` 限制並行釋放)

2. **執行緒安全集合**:
   - 所有跨執行緒佇列使用 `ConcurrentQueue<T>` (`Queue_Bitmap1-4` 等)
   - 絕不在不了解快照語義的情況下迭代並行集合
   - 使用 `TryDequeue`/`TryPeek` 模式,絕不直接索引

3. **共享資源存取**:
   - **PLC SerialPort**: 單一寫入者模式配合優先佇列;絕不阻塞埠
   - **資料庫連線**: 每個執行緒始終透過 `using` 區塊使用獨立連線;絕不共享連線物件
   - **相機物件**: 每個相機實例僅由其專用執行緒存取;絕不交叉參考
   - **TensorRT DLL 實例**: 每個站使用專用實例 (1-4);絕不共享推論會話

4. **資料競爭預防**:
   - **讀取-修改-寫入操作**: 如果值決定後續寫入,讀取共享計數器/旗標前始終加鎖
   - **事件處理器**: 假設事件在任意執行緒上觸發;透過 `Invoke`/`BeginInvoke` 封送至 UI 執行緒進行控制項更新
   - **靜態快取**: 共享查找表使用 `ConcurrentDictionary`;跨執行緒存取絕不使用一般 `Dictionary`

5. **記憶體可見性**:
   - 跨執行緒讀寫的簡單旗標使用 `volatile`
   - 原子計數器更新使用 `Interlocked` 操作
   - 理解鎖提供記憶體屏障;未同步化的讀取可能看到過時的快取值

6. **死鎖預防**:
   - **鎖排序**: 始終以一致順序取得鎖 (如需多個鎖則記錄順序)
   - **逾時模式**: 使用 `Monitor.TryEnter` 或等待控制代碼逾時;絕不無限期阻塞
   - **避免巢狀鎖**: 最小化鎖範圍;持有鎖時絕不呼叫外部程式碼

7. **多執行緒環境中的資源清理**:
   - 確保關閉時排空釋放佇列 (`_disposeQueue`)
   - 使用取消權杖優雅終止執行緒
   - 釋放共享資源前等待工作執行緒

**應避免的常見執行緒安全違規:**
- 不加鎖修改 `app.counter`, `app.param`, `app.models`
- 跨執行緒共享 `DataConnection` 物件
- 從多個執行緒呼叫 `SerialPort.Write` 而不序列化
- 從非 UI 執行緒存取 Windows Forms 控制項
- 釋放仍被工作執行緒參考的物件
- 在 PLC 觸發事件與影像處理狀態間建立競爭條件

### 命名空間消歧義
**即使存在 `using System.Drawing` 也始終完全限定 `System.Drawing.*` 類型:**
```csharp
// 正確
System.Drawing.Point point = new System.Drawing.Point(x, y);
System.Drawing.Bitmap bitmap = ...;

// 錯誤 (與 OpenCvSharp 衝突)
Point point = new Point(x, y);
```
這是因為 OpenCvSharp 也定義了 `Point`, `Size`, `Rect` 等。

## 重要的全域狀態 (`app` 類別)

Form1.cs 中的 `app` 靜態類別包含關鍵的共享狀態:
- **佇列**: `Queue_Bitmap1-4`, `Queue_Save`, `Queue_Send`, `Queue_Show` 用於管線階段
- **計數器**: `app.counter`, `app.dc` 用於生產追蹤
- **參數**: `app.param`, `app.models`, `app.metas`, `app.pos` 從資料庫載入
- **等待控制代碼**: `_wh1-4`, `_sv`, `_reader`, `_AI`, `_show` 用於執行緒同步化
- **系統狀態**: `app.currentState` 列舉 (追蹤系統操作狀態)
- **PLC 佇列**: `pendingOK1`, `pendingOK2`, `pendingPushOK1`, `pendingPushOK2` 用於推桿追蹤

**所有 `app` 狀態的修改必須同步化。複合操作使用鎖。**

## PLC 通訊注意事項

### 計數機制
- 推桿計數由 PLC 管理 (暫存器 D803, D805, D807)
- 軟體顯示從 PLC 讀取的值
- `app.counter["stop" + camID]` 追蹤軟體端計數
- `SAMPLE_ID` 衍生自 `app.counter`
- 已知問題: 緊急停止可能導致軟體報表與實際計數間 +1-2 的計數不匹配

### 指令優先級
- 吹氣指令具有最高優先級以防止佇列擁塞
- 最近的提交改進了 PLC SerialPort 接收處理並移除過多的 `GC.Collect()` 呼叫

## 最近的重大變更

### ParameterConfigForm 重構
詳見 `IMPLEMENTATION_SUMMARY.md`:
- 新的 3 區 UI 架構 (參考 → 新增未修改 → 新增已修改)
- 修復 `objBias` 參數錯誤分類 (移至偵測類別)
- 修復校正後位置參數消失
- 透過背景顏色的視覺狀態指示器
- 校正工具 (Circle, Contrast, Pixel, White, ObjectBias, GapThresh) 現在保留參考區

### 記憶體與效能
- 移除過多的 `GC.Collect()` 呼叫 (見提交: "避免PLC指令堆積/修復記憶體錯誤")
- 改進 PLC SerialPort 事件處理
- 預設將離線模式設為 false

## 會話與設定檔

### 參數會話檔
- 位置: `bin\x64\Release\setting\param_session_*.json`
- 格式: JSON 以料號為鍵
- 包含所有參數類別以快速恢復會話
- 範例: `param_session_10201120714TP.json`

### 模型檔案結構
- 異常模型: `bin\x64\Release\models\{PartNumber}_{in|out|1|2}.pt`
- YOLO 模型: 透過 HTTP 提供的 Python .pt 檔案
- TensorRT 引擎: 先前儲存,現在使用 .pt 模型動態載入
- 元資料: 每個模型附帶的 `.json` 檔案用於正規化參數

## 測試與校正

### 校正工具 (可從 ParameterConfigForm 存取)
- **CircleCalibrationForm**: 圓形位置偵測校正
- **ContrastCalibrationForm**: 對比度閾值調整
- **PixelCalibrationForm**: 逐像素異常閾值
- **WhiteCalibrationForm**: 白平衡參考
- **ObjectBiasCalibrationForm**: 物件位置偏移修正
- **gapThreshCalibrationForm**: 間隙偵測閾值

### 測試檔案 (多為舊版)
- `testAOI.cs`, `testAOI2.cs`, `testPerPixel.cs`, `testroi.cs`, `onnxTest.cs`, `onnx_Test.cs`
- 這些是開發測試工具,不是主應用程式流程的一部分

## 日誌記錄

### Serilog 設定
- 主日誌: `.\logs\peilin_log-{Date}.txt`
- 相機特定日誌: `CameraLog_{Date}.txt`, `CameraWarning_{Date}.txt`
- 範本: `{Timestamp:HH:mm:ss} [{Level:u3}] {Message}{NewLine}{Exception}`

## 多相機協調

4 個相機各自獨立運作但共享:
1. **PLC 通訊通道** (序列埠, Modbus 協定)
2. **資料庫連線池** (SQLite 配合 LinqToDB)
3. **全域 `app` 狀態** 用於跨站協調
4. **共享釋放佇列** (`_disposeQueue`) 配合信號量限制 4 個並行釋放

系統使用:
- `_wh1` 至 `_wh4` 等待控制代碼用於各站影像處理
- 並行佇列隔離站處理管線
- 影像元資料中嵌入站 ID (`ImagePosition` 類別)

**所有共享資源存取必須適當同步化以防止競爭條件。**

## 關鍵工作流程

### 新增新料號
1. 使用 `type_info.cs` 表單新增料號至資料庫
2. 透過 `ParameterConfigForm` 設定參數:
   - 選擇來源料號作為參考
   - 複製與修改參數
   - 依需要執行校正工具
   - 儲存至資料庫
3. 新增 AI 模型至 `bin\x64\Release\models\{PartNumber}_*.pt`
4. 如需物件偵測則建立 YOLO 批次檔
5. 生產前用樣本影像測試

### 修改偵測參數
1. 開啟 `ParameterConfigForm`
2. 載入目標料號
3. 選擇來源參考 (可為相同料號)
4. 參數分類:
   - 相機: 曝光、增益、延遲
   - 位置: 圓形偵測、ROI 定義
   - 偵測: 異常閾值、objBias、對比度
   - 時序: 延遲與逾時
   - 測試: 驗證參數
5. 位置/視覺參數使用校正工具
6. 儲存至資料庫 (僅儲存新增區,保留參考區)

### 偵錯影像處理問題
1. 在 Parameters 表格中啟用影像儲存
2. 檢查 `bin\x64\Release\report\{PartNumber}\` 中儲存的影像
3. 檢視 `.\logs\` 中的日誌以了解時序與錯誤
4. 使用相機日誌診斷擷取問題
5. 檢查 PLC 通訊日誌以了解觸發訊號問題

## 效能考量

- **影像釋放**: 使用非同步釋放佇列防止 UI 阻塞
- **AI 推論**: 每個站可使用不同的 TensorRT DLL 實例 (1-4)
- **資料庫寫入**: 盡可能批次處理,多列更新使用交易
- **PLC 指令**: 優先佇列防止吹氣指令延遲
- **YOLO 預熱**: 系統在閒置期間後自動預熱 YOLO 模型
- **Mat 物件**: 絕不持有參考超過需要的時間;釋放至關重要
- **記憶體壓力**: 持續負載下監控佇列深度與影像物件生命週期
- **執行緒競爭**: 在峰值吞吐量期間剖析 `app` 狀態的鎖競爭

## 常見陷阱

1. **不釋放 OpenCV Mat**: 導致快速記憶體增長與當機
2. **使用不限定的 `Point`/`Size`**: System.Drawing 與 OpenCvSharp 之間模糊
3. **未檢查存在性直接 Insert**: 導致主鍵約束違規
4. **不同步化修改 `app` 狀態**: 跨站競爭條件
5. **阻塞 PLC 序列埠**: 吹氣指令必須最高優先級
6. **推論前忘記載入模型**: CreateModel 呼叫前檢查模型路徑是否存在
7. **跨執行緒共享資料庫連線**: 每個執行緒始終建立新連線
8. **過度呼叫 GC.Collect()**: 降低效能;在最近的提交中已移除
9. **呼叫外部程式碼時持有鎖**: 死鎖風險;最小化鎖範圍
10. **忽略佇列深度監控**: 不受控的佇列增長導致記憶體耗盡

## YOLO 伺服器管理

系統使用外部 Python 程序進行 YOLO 偵測:
- 透過 `bin\x64\Release\yolo_{PartNumber}_{StationNum}.bat` 啟動
- 與 localhost:5001-5004 上的 HTTP 通訊
- 支援透過 `/load_model` 端點熱切換模型
- 預熱機制保持模型在 GPU 記憶體中
- 批次偵測端點: `/detect_batch`

範例工作流程 (由 `YoloDetection.cs` 處理):
1. 透過批次檔啟動伺服器
2. 等待伺服器可用
3. 為當前料號載入模型
4. 用虛擬影像預熱
5. 發送包含 Base64 編碼影像的偵測請求
6. 解析包含邊界框與類別 ID 的 JSON 回應
