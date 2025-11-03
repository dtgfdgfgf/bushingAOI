# CLAUDE-zh-tw.md

此檔案為 Claude Code（claude.ai/code）在處理此專案時提供指導。

## 專案概述

這是一個用於工業軸套製造品質控制的 C# Windows Forms AOI（自動光學檢測）應用程式。系統整合 Basler 攝影機、PLC 通訊、YOLO 物件偵測及使用 TensorRT 的異常偵測，以即時檢測產品。

**目標框架：** .NET Framework 4.8（x64）
**主要進入點：** `Program.cs` → `Form1`
**建置輸出：** `bin\x64\Release\peilin.exe`

## 建置與執行

### 建置專案
```bash
# 在 Visual Studio 中建置
msbuild peilin.sln /p:Configuration=Release /p:Platform=x64

# 或使用 Visual Studio IDE（建議）
# 將組態設定為 "Release"，平台設定為 "x64"
```

### 執行應用程式
- 執行 `bin\x64\Release\peilin.exe`
- 應用程式需要位於 `.\setting\mydb.sqlite` 的 SQLite 資料庫
- 預設啟用智慧卡加密狗驗證（可透過 `app.usekey` 停用）

### 外部相依性
- **YOLO Python 伺服器**：透過 `bin\x64\Release\yolo_*.bat` 批次檔啟動
- **TensorRT DLL**：`AD_TRT_dll1.dll` 到 `AD_TRT_dll4.dll` 用於異常偵測
- **Basler Pylon SDK**：相機操作所需

## 系統架構

### 多站台設計
系統支援最多同時 4 個檢測站台：
- 每個站台擁有自己的 Basler 攝影機（透過 `Camera0.cs` 配置）
- 每個站台執行獨立的 AI 推論（TensorRT + YOLO）
- 共用的 PLC 通訊透過序列埠（Modbus 協定）協調所有站台

### 核心處理流程
```
相機觸發 → 影像擷取 → 位置偵測 →
ROI 提取 → AI 推論（異常 + YOLO）→
瑕疵分類 → 結果儲存 → PLC 訊號
```

### 關鍵架構元件

**Form1.cs**（17,000+ 行）
- 主要應用程式表單與系統編排
- 包含具有全域狀態的 `app` 靜態類別（佇列、計數器、參數）
- 影像處理管線協調
- PLC 通訊處理程式

**Camera0.cs**
- Basler Pylon 攝影機包裝器
- 多攝影機管理（`m_BaslerCameras` 陣列）
- 事件驅動的影像抓取，附帶監控/日誌記錄

**anomalyTensorRT.cs**
- 4 個 TensorRT DLL 實例（`AD_TRT_dll1` 到 `AD_TRT_dll4`）的包裝器
- 異常偵測模型載入與推論
- 回傳異常熱力圖（Mat）與分數（float）

**YoloDetection.cs**
- 基於 Python 的 YOLO 伺服器 HTTP 客戶端
- 模型載入、預熱與批次偵測
- 與本機伺服器通訊，埠號 5001-5004

**PLC_Test.cs**
- PLC 的 Modbus 序列通訊
- 處理觸發訊號、計數器（D803、D805、D807）與吹氣指令
- 優先佇列系統用於防止 PLC 指令堆積

**ParameterConfigForm.cs / ParameterSetupManager.cs**
- 參數管理 UI，採用 3 區系統：
  - 參考區（來源料號參數，唯讀）
  - 已新增-未修改區（已複製但尚未編輯）
  - 已新增-已修改區（已編輯參數，準備儲存）
- 校正工具整合，用於位置/圓形/像素參數

**algorithm.cs**
- SafeROI 輔助方法，用於邊界安全的影像區域提取

### 資料庫結構（SQLite）
位於 `.\setting\mydb.sqlite`，包含以下資料表：
- **Cameras**：每個料號的攝影機參數（曝光、增益、延遲）
- **params**：每個料號的演算法參數
- **Types**：料號資訊
- **DefectTypes**：全域瑕疵類型定義
- **DefectChecks**：料號特定的瑕疵偵測映射
- **DefectCounts**：用於報告的瑕疵統計
- **Totals**：生產計數器
- **Blows**：吹氣分類參數
- **Users**：使用者驗證

## 程式碼風格要求

### 語言與註解
- **所有註解必須使用繁體中文**
- 在 AI 產生的程式碼段落開頭加上註解：`// 由 GitHub Copilot 產生`

### 資料庫操作（LinqToDB）
務必使用「先更新，若無則插入」模式來避免主鍵衝突：
```csharp
// 嘗試更新現有記錄
int updatedRows = db.table.Where(條件).Set(欄位, 值).Update();

// 若無更新記錄則插入新記錄
if (updatedRows == 0)
{
    db.table.Value(欄位, 值).Insert();
}
```
- 避免為 Insert 建立實體物件；使用 `.Value()` 語法
- 永遠對資料庫連線使用 `using` 陳述式
- 對所有資料庫操作包含適當的錯誤處理

### 記憶體管理
**關鍵：** 此應用程式以工業速度處理高解析度影像。不良的記憶體管理會導致當機。

- **OpenCV Mat 物件**：使用後必須立即釋放
- **System.Drawing.Bitmap**：必須明確呼叫 `.Dispose()`
- 對影像處理操作使用 `using` 陳述式
- 長期存活的大型物件應使用壓縮儲存
- 應用程式使用 `_disposeQueue` 與 `_disposeSemaphore` 進行非同步釋放

### 命名空間消歧
**即使存在 `using System.Drawing`，也務必完整限定 `System.Drawing.*` 類型：**
```csharp
// 正確
System.Drawing.Point point = new System.Drawing.Point(x, y);
System.Drawing.Bitmap bitmap = ...;

// 錯誤（與 OpenCvSharp 衝突）
Point point = new Point(x, y);
```
這是因為 OpenCvSharp 也定義了 `Point`、`Size`、`Rect` 等。

## 重要的全域狀態（`app` 類別）

Form1.cs 中的 `app` 靜態類別包含關鍵的共用狀態：
- **佇列**：`Queue_Bitmap1-4`、`Queue_Save`、`Queue_Send`、`Queue_Show` 用於管線階段
- **計數器**：`app.counter`、`app.dc` 用於生產追蹤
- **參數**：`app.param`、`app.models`、`app.metas`、`app.pos` 從資料庫載入
- **等待控制代碼**：`_wh1-4`、`_sv`、`_reader`、`_AI`、`_show` 用於執行緒同步
- **系統狀態**：`app.currentState` 列舉（追蹤系統操作狀態）
- **PLC 佇列**：`pendingOK1`、`pendingOK2`、`pendingPushOK1`、`pendingPushOK2` 用於推桿追蹤

## PLC 通訊注意事項

### 計數機制
- 推桿計數由 PLC 管理（暫存器 D803、D805、D807）
- 軟體顯示從 PLC 讀取的值
- `app.counter["stop" + camID]` 追蹤軟體端計數
- `SAMPLE_ID` 衍生自 `app.counter`
- 已知問題：緊急停止可能導致軟體報告與實際計數之間相差 +1-2 個

### 指令優先順序
- 吹氣指令具有最高優先順序以防止佇列堆積
- 最近的提交改進了 PLC SerialPort 接收處理，並移除過多的 `GC.Collect()` 呼叫

## 最近的重大變更

### ParameterConfigForm 重構
詳見 `IMPLEMENTATION_SUMMARY.md` 詳細文件：
- 新的 3 區 UI 架構（參考 → 已新增-未修改 → 已新增-已修改）
- 修正 `objBias` 參數錯誤分類（移至 Detection 類別）
- 修正校正後位置參數消失的問題
- 透過背景顏色提供視覺狀態指示器
- 校正工具（Circle、Contrast、Pixel、White、ObjectBias、GapThresh）現在保留參考區

### 記憶體與效能
- 移除過多的 `GC.Collect()` 呼叫（參見提交：「避免PLC指令堆積/修復記憶體錯誤」）
- 改進 PLC SerialPort 事件處理
- 將離線模式預設設為 false

## 工作階段與配置檔案

### 參數工作階段檔案
- 位置：`bin\x64\Release\setting\param_session_*.json`
- 格式：JSON，以料號作為鍵值
- 包含所有參數類別，用於快速工作階段還原
- 範例：`param_session_10201120714TP.json`

### 模型檔案結構
- 異常模型：`bin\x64\Release\models\{PartNumber}_{in|out|1|2}.pt`
- YOLO 模型：透過 HTTP 提供的 Python .pt 檔案
- TensorRT 引擎：先前儲存，現在使用 .pt 模型進行動態載入
- 中繼資料：每個模型附帶的 `.json` 檔案，用於標準化參數

## 測試與校正

### 校正工具（可從 ParameterConfigForm 存取）
- **CircleCalibrationForm**：圓形位置偵測校正
- **ContrastCalibrationForm**：對比度閾值調整
- **PixelCalibrationForm**：每像素異常閾值
- **WhiteCalibrationForm**：白平衡參考
- **ObjectBiasCalibrationForm**：物體位置偏差修正
- **gapThreshCalibrationForm**：間隙偵測閾值

### 測試檔案（大多為舊版）
- `testAOI.cs`、`testAOI2.cs`、`testPerPixel.cs`、`testroi.cs`、`onnxTest.cs`、`onnx_Test.cs`
- 這些是開發測試工具，不屬於主要應用程式流程

## 日誌記錄

### Serilog 配置
- 主要日誌：`.\logs\peilin_log-{Date}.txt`
- 攝影機特定日誌：`CameraLog_{Date}.txt`、`CameraWarning_{Date}.txt`
- 範本：`{Timestamp:HH:mm:ss} [{Level:u3}] {Message}{NewLine}{Exception}`

## 多攝影機協調

4 個攝影機各自獨立運作，但共用：
1. **PLC 通訊通道**（序列埠，Modbus 協定）
2. **資料庫連線池**（SQLite 搭配 LinqToDB）
3. **全域 `app` 狀態**，用於跨站台協調
4. **共用釋放佇列**（`_disposeQueue`），使用旗號限制 4 個並行釋放

系統使用：
- `_wh1` 到 `_wh4` 等待控制代碼，用於每站台影像處理
- 並行佇列來隔離站台處理管線
- 站台 ID 嵌入在影像中繼資料中（`ImagePosition` 類別）

## 關鍵工作流程

### 新增料號
1. 使用 `type_info.cs` 表單將料號新增至資料庫
2. 透過 `ParameterConfigForm` 配置參數：
   - 選擇來源料號作為參考
   - 複製並修改參數
   - 視需要執行校正工具
   - 儲存至資料庫
3. 將 AI 模型新增至 `bin\x64\Release\models\{PartNumber}_*.pt`
4. 如需物件偵測，建立 YOLO 批次檔
5. 生產前使用範例影像測試

### 修改偵測參數
1. 開啟 `ParameterConfigForm`
2. 載入目標料號
3. 選擇來源參考（可以是相同料號）
4. 參數分類為：
   - Camera：曝光、增益、延遲
   - Position：圓形偵測、ROI 定義
   - Detection：異常閾值、objBias、對比度
   - Timing：延遲與逾時
   - Testing：驗證參數
5. 對位置/視覺參數使用校正工具
6. 儲存至資料庫（僅儲存已新增區，保留參考區）

### 除錯影像處理問題
1. 在 Parameters 資料表中啟用影像儲存
2. 檢查 `bin\x64\Release\report\{PartNumber}\` 中儲存的影像
3. 查看 `.\logs\` 中的日誌，了解時間與錯誤
4. 使用 Camera 日誌診斷抓取問題
5. 檢查 PLC 通訊日誌，了解觸發訊號問題

## 效能考量

- **影像釋放**：使用非同步釋放佇列以防止 UI 阻塞
- **AI 推論**：每個站台可使用不同的 TensorRT DLL 實例（1-4）
- **資料庫寫入**：盡可能批次處理，對多列更新使用交易
- **PLC 指令**：優先佇列防止吹氣指令延遲
- **YOLO 預熱**：系統在閒置期間後自動預熱 YOLO 模型
- **Mat 物件**：永遠不要持有超過必要時間的參考；釋放至關重要

## 常見陷阱

1. **未釋放 OpenCV Mat**：導致記憶體快速增長與當機
2. **使用未限定的 `Point`/`Size`**：System.Drawing 與 OpenCvSharp 之間模稜兩可
3. **未檢查存在性就直接 Insert**：導致主鍵約束違規
4. **未同步就修改 `app` 狀態**：跨站台競爭條件
5. **阻塞 PLC 序列埠**：吹氣指令必須是最高優先順序
6. **推論前忘記載入模型**：在 CreateModel 呼叫前檢查模型路徑是否存在

## YOLO 伺服器管理

系統使用外部 Python 處理程序進行 YOLO 偵測：
- 透過 `bin\x64\Release\yolo_{PartNumber}_{StationNum}.bat` 啟動
- 透過 HTTP 與本機主機通訊，埠號 5001-5004
- 支援透過 `/load_model` 端點進行模型熱交換
- 預熱機制以將模型保持在 GPU 記憶體中
- 批次偵測端點：`/detect_batch`

範例工作流程（由 `YoloDetection.cs` 處理）：
1. 透過批次檔啟動伺服器
2. 等待伺服器可用
3. 載入目前料號的模型
4. 使用虛擬影像進行預熱
5. 使用 Base64 編碼影像傳送偵測請求
6. 解析包含邊界框與類別 ID 的 JSON 回應

## 文檔管理規則

### 雙語文檔要求
**關鍵：** 此專案維護英文與繁體中文兩種語言的文檔。

當生成或更新 `CLAUDE.md`（英文版）時，你**必須**同時生成或更新 `CLAUDE-zh-tw.md`（繁體中文版）：
- 兩個檔案必須包含相同資訊，僅語言不同
- `CLAUDE.md` 必須使用英文
- `CLAUDE-zh-tw.md` 必須使用繁體中文（CLAUDE.md 的完整翻譯）
- 任何結構變更、新章節或內容更新都必須在兩個檔案間同步

### Git 推送要求
在進行任何文檔或程式碼變更後，遵循此工作流程：

1. **暫存並提交變更：**
   ```bash
   git add .
   git commit -m "..."
   git push
   ```

2. **驗證推送成功：**
   ```bash
   git status
   # 或
   git log --oneline -3
   ```

### 結構化 Commit 訊息格式
每個檔案變更都必須在 commit 訊息中明確記錄，使用此結構：

```
<類型>(<範圍>): <主題>

<內容>
- 檔案1：變更描述與理由
- 檔案2：變更描述與理由
- 檔案3：變更描述與理由

<註腳>
```

**Commit 訊息要求：**

- **類型：** feat、fix、docs、refactor、chore、test、style 等
- **範圍：** 受影響的模組或元件（例如：ai、notion、webhook、docs）
- **主題：** 簡短摘要（50 字元以內）
- **內容：** 每個變更檔案的詳細說明：
  - 每個檔案變更了什麼
  - 為什麼需要這個變更
  - 變更的影響
- **註腳：** 共同作者、參考資料、重大變更

**範例：**

```
docs(project): 新增雙語文檔管理規則

- CLAUDE.md: 新增文檔管理章節，包含雙語要求、git 工作流程與 commit 訊息標準
- CLAUDE-zh-tw.md: 建立完整的 CLAUDE.md 繁體中文翻譯

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>
```

```
feat(parameters): 實作 3 區參數管理系統

- ParameterConfigForm.cs: 重構 UI 以支援參考/已新增-未修改/已新增-已修改區
- ParameterSetupManager.cs: 新增區域追蹤邏輯與儲存驗證
- CLAUDE.md: 更新文檔以反映新的參數管理工作流程
- CLAUDE-zh-tw.md: 同步繁體中文文檔

此變更透過明確區分來源參數與已編輯副本來改進參數編輯工作流程，防止意外覆寫。

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>
```

### 文檔同步
當進行任何程式碼變更時，必須檢視並視需要更新 `CLAUDE.md`。

**需要更新文檔的關鍵領域：**
- 架構變更（新模組、修改的流程）
- 新功能或指令
- 資料庫 schema 修改
- API 整合或外部服務變更
- 配置或環境變數變更
- 工作流程或商業邏輯的變更

**同步工作流程：**
1. 進行程式碼變更
2. 評估變更是否影響文檔
3. 如果是，更新 `CLAUDE.md`（英文）
4. 立即更新 `CLAUDE-zh-tw.md`（繁體中文翻譯）
5. 將兩個文檔檔案與程式碼變更一起提交

**文檔必須始終反映程式碼庫的當前狀態，以防止資訊偏移。**
