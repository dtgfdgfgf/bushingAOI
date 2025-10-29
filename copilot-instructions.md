# Copilot Instructions - AOI Bushing Inspection System

此文件為 GitHub Copilot 和 AI 程式碼助手提供程式碼審查和產生指引。

## 核心原則

1. **所有註解和文件必須使用繁體中文**
2. **AI 產生的程式碼段落開頭必須標註**: `// 由 GitHub Copilot 產生`
3. **記憶體管理是最高優先級** - 圖像處理不當會導致系統崩潰
4. **執行緒安全至關重要** - 多站點並行處理需要嚴格同步
5. **PLC 通訊的即時性不可妥協** - 氣吹指令必須優先處理

---

## When reviewing code, focus on:

### 🔴 Critical - 記憶體管理 (Memory Management)

#### OpenCV Mat 物件
```csharp
// ✅ 正確：立即 Dispose
using (var mat = new Mat())
{
    // 處理影像
} // 自動釋放

// ✅ 正確：手動 Dispose
Mat mat = ProcessImage();
try
{
    // 使用 mat
}
finally
{
    mat?.Dispose();
}

// ❌ 錯誤：未釋放
Mat mat = new Mat();
// ... 使用後未 Dispose，造成記憶體洩漏
```

**檢查要點：**
- [ ] 每個 `new Mat()` 都有對應的 `.Dispose()` 或 `using` 語句
- [ ] 從方法回傳的 Mat 物件在呼叫端被正確處理
- [ ] 長期持有的 Mat 物件使用壓縮儲存或定期釋放
- [ ] 使用 `_disposeQueue` 和 `_disposeSemaphore` 進行非同步釋放
- [ ] 避免在 LINQ 或迭代中建立未追蹤的 Mat 物件

#### System.Drawing.Bitmap
```csharp
// ✅ 正確
using (var bitmap = new System.Drawing.Bitmap(path))
{
    // 處理圖片
}

// ❌ 錯誤：未 Dispose
System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(path);
// ... 使用後未釋放
```

**檢查要點：**
- [ ] 所有 Bitmap 物件都使用 `using` 或明確呼叫 `.Dispose()`
- [ ] Bitmap 與 Mat 之間的轉換後兩者都被正確釋放

#### 記憶體管理最佳實踐
- [ ] 避免在高頻迴圈中分配大型物件
- [ ] 使用物件池重用常用資源
- [ ] 移除不必要的 `GC.Collect()` 呼叫（會影響效能）
- [ ] 監控 `app.Queue_*` 佇列大小，防止記憶體累積

---

### 🔴 Critical - 型別明確性 (Type Disambiguation)

由於 `OpenCvSharp` 和 `System.Drawing` 都定義了 `Point`, `Size`, `Rectangle` 等型別，**必須完全限定 System.Drawing 的型別**。

```csharp
// ✅ 正確：完全限定
System.Drawing.Point point = new System.Drawing.Point(x, y);
System.Drawing.Size size = new System.Drawing.Size(width, height);
System.Drawing.Rectangle rect = new System.Drawing.Rectangle(x, y, w, h);
System.Drawing.Bitmap bitmap = ...;

// ❌ 錯誤：模糊參考（即使有 using System.Drawing）
Point point = new Point(x, y);  // 編譯器無法確定是哪個 Point
Size size = new Size(w, h);     // 可能是 OpenCvSharp.Size
```

**檢查要點：**
- [ ] 所有 `System.Drawing.*` 型別使用完全限定名稱
- [ ] 避免使用 `using Point = System.Drawing.Point` 這類別名
- [ ] 檢查方法簽章中的參數型別是否明確
- [ ] 在型別推斷（var）時確認右側型別正確

---

### 🔴 Critical - 資料庫操作 (Database Operations)

使用 **LinqToDB** 進行 SQLite 操作，必須遵循 "update-first, insert-if-none" 模式避免主鍵衝突。

```csharp
// ✅ 正確：Update-first, Insert-if-none 模式
using (var db = new PeiLin.Data.PeilinDB())
{
    // 嘗試更新現有記錄
    int updatedRows = db.Cameras
        .Where(c => c.PartNumber == partNumber && c.CamID == camID)
        .Set(c => c.Exposure, exposure)
        .Set(c => c.Gain, gain)
        .Update();

    // 若無更新記錄則插入新記錄
    if (updatedRows == 0)
    {
        db.Cameras
            .Value(c => c.PartNumber, partNumber)
            .Value(c => c.CamID, camID)
            .Value(c => c.Exposure, exposure)
            .Value(c => c.Gain, gain)
            .Insert();
    }
}

// ❌ 錯誤：直接 Insert（可能造成主鍵衝突）
db.Cameras.Insert(() => new Camera {
    PartNumber = partNumber,
    CamID = camID
});

// ❌ 錯誤：建立實體物件插入（不符合專案慣例）
var camera = new Camera { PartNumber = partNumber, ... };
db.Insert(camera);
```

**檢查要點：**
- [ ] 所有資料庫連線使用 `using` 語句確保釋放
- [ ] Insert 操作前先檢查記錄是否存在（或使用 Update-Insert 模式）
- [ ] 使用 `.Value()` 語法而非建立實體物件
- [ ] 包含適當的錯誤處理（try-catch）
- [ ] 多筆更新使用 Transaction 包裝
- [ ] 避免在迴圈中重複開啟連線（應在外層開啟）
- [ ] 查詢結果記得呼叫 `.ToList()` 或 `.FirstOrDefault()` 執行

---

### 🟠 Important - 多執行緒安全 (Thread Safety)

系統有 4 個檢測站點並行處理，共享 `app` 全域狀態需要同步機制。

```csharp
// ✅ 正確：使用 lock 保護共享狀態
lock (app.counterLock)
{
    app.counter["stop" + camID]++;
}

// ✅ 正確：使用 ConcurrentQueue
app.Queue_Bitmap1.Enqueue(imageData);

// ✅ 正確：使用 WaitHandle 同步
app._wh1.Set();  // 觸發站點 1 處理
app._wh2.WaitOne();  // 等待站點 2 完成

// ❌ 錯誤：直接修改共享狀態
app.counter["total"]++;  // 可能發生競態條件

// ❌ 錯誤：未同步存取 Dictionary
var value = app.param[key];  // 在多執行緒環境中不安全
```

**檢查要點：**
- [ ] 修改 `app.counter`, `app.param`, `app.metas` 等共享字典時使用 lock
- [ ] 使用 `ConcurrentQueue` 而非普通 Queue 處理跨執行緒資料
- [ ] 正確使用 `_wh1-4`, `_sv`, `_reader`, `_AI`, `_show` 等 WaitHandle
- [ ] 避免長時間持有鎖（死鎖風險）
- [ ] 檢查是否有競態條件（race condition）
- [ ] 確認 `pendingOK1/2`, `pendingPushOK1/2` 等 PLC 佇列的執行緒安全
- [ ] 使用 `Interlocked` 類別進行原子操作

---

### 🟠 Important - PLC 通訊 (PLC Communication)

PLC 透過 Modbus 串列埠通訊，必須遵循優先權機制防止指令堆積。

```csharp
// ✅ 正確：氣吹指令最高優先權
PLC_Test.SendBlowCommand(station, duration);  // 立即執行

// ✅ 正確：一般指令加入佇列
lock (pendingOK1)
{
    pendingOK1.Enqueue(new OKCommand { StationID = 1, Value = result });
}

// ❌ 錯誤：阻塞串列埠
serialPort.Write(data);
Thread.Sleep(1000);  // 阻塞其他站點通訊

// ❌ 錯誤：忽略指令優先權
SendPLCCommand(normalCommand);  // 可能延遲氣吹指令
```

**檢查要點：**
- [ ] 氣吹指令（air blow）不進入一般佇列，直接發送
- [ ] 使用 `pendingOK1/2`, `pendingPushOK1/2` 佇列管理非緊急指令
- [ ] 避免在 SerialPort 事件處理中執行耗時操作
- [ ] 正確處理 PLC 計數器（D803, D805, D807）讀取
- [ ] 錯誤處理：串列埠斷線、逾時、CRC 錯誤
- [ ] 避免過於頻繁的 PLC 讀寫（遵循時序要求）
- [ ] `SAMPLE_ID` 與 `app.counter` 的一致性

---

### 🟠 Important - AI 推論管理 (AI Inference)

系統使用 TensorRT（異常檢測）和 YOLO（物件偵測），需要正確的模型生命週期管理。

```csharp
// ✅ 正確：檢查模型檔案存在
string modelPath = $@".\models\{partNumber}_in.pt";
if (!File.Exists(modelPath))
{
    Log.Error($"模型檔案不存在: {modelPath}");
    return null;
}

// ✅ 正確：使用對應的 TensorRT DLL 實例
anomalyTensorRT ad = camID switch
{
    1 => app.AD1,
    2 => app.AD2,
    3 => app.AD3,
    4 => app.AD4,
    _ => throw new ArgumentException("無效的 CamID")
};

// ✅ 正確：YOLO 伺服器連線檢查
if (!yoloDetection.IsServerAlive())
{
    Log.Warning("YOLO 伺服器未回應，嘗試重新啟動");
    StartYoloServer(partNumber, camID);
}

// ❌ 錯誤：未檢查模型是否載入
var result = ad.RunInference(image);  // 可能拋出異常

// ❌ 錯誤：在錯誤的執行緒呼叫 AI 推論
Task.Run(() => yolo.Detect(image));  // 可能造成資源競爭
```

**檢查要點：**
- [ ] 模型檔案路徑存在性驗證
- [ ] TensorRT DLL 實例（AD1-4）的正確分配
- [ ] YOLO 伺服器健康檢查和自動重啟機制
- [ ] 模型預熱（warmup）在長時間閒置後執行
- [ ] 正確處理推論失敗（timeout, 記憶體不足）
- [ ] 批次處理時的影像編碼/解碼效率
- [ ] 推論結果的記憶體釋放（特別是 Mat 物件）
- [ ] 模型切換時的資源清理

---

### 🟡 Recommended - 程式碼風格 (Code Style)

```csharp
// ✅ 正確：繁體中文註解
// 計算 ROI 區域並提取影像特徵
public Mat ExtractROI(Mat source, System.Drawing.Rectangle roi)
{
    // 使用安全 ROI 避免越界
    var safeROI = algorithm.SafeROI(source, roi);
    return source.SubMat(safeROI);
}

// ✅ 正確：標註 AI 產生的程式碼
// 由 GitHub Copilot 產生
private void CalculateAnomalyScore(Mat heatmap, float threshold)
{
    // ... AI 產生的實作
}

// ❌ 錯誤：英文註解
// Calculate ROI and extract features
public Mat ExtractROI(...)

// ❌ 錯誤：未標註 AI 產生
private void CalculateAnomalyScore(...)  // AI 產生但未標註
```

**檢查要點：**
- [ ] 所有註解使用繁體中文
- [ ] AI 產生的程式碼段落標註 `// 由 GitHub Copilot 產生`
- [ ] 變數命名：C# 慣例（camelCase 區域變數，PascalCase 方法/屬性）
- [ ] 使用有意義的變數名稱（避免 `temp`, `data1`, `x1` 等模糊命名）
- [ ] 遵循專案現有的程式碼格式（縮排、括號位置）
- [ ] 公開方法包含 XML 文件註解
- [ ] 複雜演算法包含說明註解

---

### 🟡 Recommended - 錯誤處理與日誌 (Error Handling & Logging)

使用 **Serilog** 進行結構化日誌記錄。

```csharp
// ✅ 正確：結構化日誌
Log.Information("開始處理影像 - 站點: {CamID}, 料號: {PartNumber}", camID, partNumber);
Log.Error(ex, "AI 推論失敗 - 站點: {CamID}, 模型: {ModelPath}", camID, modelPath);

// ✅ 正確：適當的錯誤處理
try
{
    var result = ProcessImage(mat);
}
catch (OpenCVException ex)
{
    Log.Error(ex, "OpenCV 處理錯誤: {Message}", ex.Message);
    // 使用預設值或跳過此影像
    return defaultResult;
}
catch (Exception ex)
{
    Log.Fatal(ex, "未預期的錯誤");
    // 通知使用者並記錄詳細資訊
    MessageBox.Show($"系統錯誤: {ex.Message}");
    throw;
}

// ❌ 錯誤：吞掉例外
try
{
    ProcessImage(mat);
}
catch { }  // 沒有記錄或處理

// ❌ 錯誤：非結構化日誌
Log.Information($"處理影像 - {camID} - {partNumber}");  // 使用字串插值
```

**檢查要點：**
- [ ] 使用 Serilog 的結構化日誌（`{參數名稱}` 而非字串插值）
- [ ] 關鍵操作記錄 Information 層級
- [ ] 錯誤記錄包含完整例外堆疊（`Log.Error(ex, ...)`）
- [ ] 不吞掉例外（除非有明確理由）
- [ ] 錯誤訊息對使用者友善（避免技術細節）
- [ ] 相機相關錯誤記錄到 `CameraWarning_{Date}.txt`
- [ ] 效能敏感區域使用 `Log.Debug` 而非 `Information`

---

### 🟡 Recommended - 效能考量 (Performance)

```csharp
// ✅ 正確：批次資料庫操作
using (var db = new PeilinDB())
{
    db.BeginTransaction();
    foreach (var defect in defects)
    {
        // 批次插入
        db.DefectCounts.Value(...).Insert();
    }
    db.CommitTransaction();
}

// ✅ 正確：避免不必要的轉換
Mat mat = ...; // 直接使用 Mat
// 避免 Mat -> Bitmap -> Mat 的往返轉換

// ✅ 正確：使用 SafeROI 避免越界檢查
var safeROI = algorithm.SafeROI(mat, requestedROI);
var cropped = mat.SubMat(safeROI);

// ❌ 錯誤：高頻呼叫 GC
for (int i = 0; i < 1000; i++)
{
    ProcessImage();
    GC.Collect();  // 嚴重影響效能
}

// ❌ 錯誤：在迴圈中開啟資料庫連線
foreach (var item in items)
{
    using (var db = new PeilinDB())  // 應在外層開啟
    {
        db.Update(item);
    }
}
```

**檢查要點：**
- [ ] 移除不必要的 `GC.Collect()` 呼叫
- [ ] 批次處理資料庫操作使用 Transaction
- [ ] 避免重複的影像格式轉換
- [ ] 使用 `SafeROI` 而非手動邊界檢查
- [ ] 非同步處理使用 `async/await` 而非 `Task.Run`
- [ ] 佇列大小監控防止記憶體累積（`Queue_Bitmap1-4`）
- [ ] 大型影像使用 ROI 而非複製整張
- [ ] YOLO 批次偵測優於單張偵測

---

### 🟡 Recommended - 參數與配置管理 (Parameter Management)

```csharp
// ✅ 正確：從 app 全域狀態讀取參數
var threshold = app.param[$"{partNumber}_anomalyThreshold_{camID}"];
var exposure = app.param[$"{partNumber}_exposure_{camID}"];

// ✅ 正確：參數變更後更新資料庫
using (var db = new PeilinDB())
{
    int updated = db.params
        .Where(p => p.PartNumber == partNumber && p.ParamName == "threshold")
        .Set(p => p.ParamValue, newValue)
        .Update();

    if (updated == 0)
    {
        db.params.Value(...).Insert();
    }
}

// ✅ 正確：校準工具整合
var calibrationForm = new CircleCalibrationForm(partNumber, camID);
if (calibrationForm.ShowDialog() == DialogResult.OK)
{
    // 從校準表單取得新參數
    UpdatePositionParameters(calibrationForm.CalibratedParams);
}

// ❌ 錯誤：硬編碼參數
var threshold = 0.85f;  // 應從資料庫或 app.param 讀取

// ❌ 錯誤：參數變更未同步到資料庫
app.param[key] = newValue;  // 僅記憶體，重啟後遺失
```

**檢查要點：**
- [ ] 參數從 `app.param` 字典讀取而非硬編碼
- [ ] 參數變更同步到資料庫（params 表）
- [ ] 校準工具（Circle, Contrast, Pixel, White, ObjectBias, GapThresh）的結果正確保存
- [ ] Reference Zone 參數只讀，不可修改
- [ ] Added-Modified Zone 的參數正確標記和儲存
- [ ] 參數 Session 檔案（`param_session_*.json`）格式正確
- [ ] 料號切換時參數正確載入

---

### 🟢 Optional - 測試與維護性 (Testing & Maintainability)

```csharp
// ✅ 正確：方法職責單一
private Mat ExtractROI(Mat source, System.Drawing.Rectangle roi)
{
    return source.SubMat(algorithm.SafeROI(source, roi));
}

private float CalculateAnomalyScore(Mat heatmap)
{
    return heatmap.Mean().Val0;
}

// ✅ 正確：可測試的設計
public class ImageProcessor
{
    private readonly IAnomalyDetector _detector;

    public ImageProcessor(IAnomalyDetector detector)
    {
        _detector = detector;  // 依賴注入便於測試
    }
}

// ❌ 錯誤：方法過長（God Method）
private void ProcessImage(Mat image)
{
    // 300 行混合影像處理、AI 推論、資料庫操作、UI 更新...
}

// ❌ 錯誤：難以測試
private void ProcessData()
{
    var db = new PeilinDB();  // 硬編碼依賴
    var result = db.Query(...);
}
```

**檢查要點：**
- [ ] 方法長度控制在 50 行以內（複雜邏輯除外）
- [ ] 避免 God Class/Method（如 Form1.cs 的部分方法）
- [ ] 可測試性：避免硬編碼依賴
- [ ] 相關功能群組為內聚模組
- [ ] 魔術數字提取為常數或參數
- [ ] 複雜條件提取為有意義的布林變數
- [ ] 公開 API 的向後相容性

---

## 專案特定檢查清單 (Project-Specific Checklist)

### 新增料號 (Adding New Part Number)
- [ ] 使用 `type_info.cs` 新增料號到資料庫
- [ ] 透過 `ParameterConfigForm` 設定參數
- [ ] 執行所有必要的校準工具
- [ ] 新增 AI 模型檔案到 `.\models\{PartNumber}_*.pt`
- [ ] 建立 YOLO batch 檔案（如需物件偵測）
- [ ] 使用測試影像驗證

### 修改檢測參數 (Modifying Detection Parameters)
- [ ] 在 `ParameterConfigForm` 中選擇正確的參考來源
- [ ] 確認參數分類正確（Camera/Position/Detection/Timing/Testing）
- [ ] 使用校準工具調整視覺參數
- [ ] 儲存到資料庫（僅 Added Zone）
- [ ] 驗證 Reference Zone 未被修改

### 除錯影像處理問題 (Debugging Image Processing)
- [ ] 啟用影像儲存（Parameters 表）
- [ ] 檢查 `.\report\{PartNumber}\` 目錄
- [ ] 查看 `.\logs\` 中的時序和錯誤
- [ ] 檢查 Camera logs 診斷擷取問題
- [ ] 檢查 PLC logs 診斷觸發信號

---

## 常見陷阱速查表 (Common Pitfalls Quick Reference)

| 問題 | 症狀 | 解決方案 |
|------|------|----------|
| **未 Dispose Mat** | 記憶體快速增長、OutOfMemoryException | 使用 `using` 或 `.Dispose()` |
| **型別模糊** | Point/Size 編譯錯誤 | 完全限定 `System.Drawing.Point` |
| **主鍵衝突** | Insert 拋出 SQLite 錯誤 | 使用 Update-Insert 模式 |
| **競態條件** | 間歇性錯誤、資料不一致 | 使用 lock 或 Concurrent 集合 |
| **PLC 指令堆積** | 氣吹延遲、系統卡頓 | 氣吹指令優先、避免阻塞 |
| **模型未載入** | AI 推論失敗 | 檢查檔案存在性、預熱 |
| **過度 GC** | 效能低落、UI 卡頓 | 移除不必要的 `GC.Collect()` |
| **跨執行緒 UI** | InvalidOperationException | 使用 `Invoke` 或 `BeginInvoke` |

---

## 自動化檢查建議 (Automated Checks)

建議搭配以下工具進行自動化程式碼審查：

1. **Roslyn Analyzers**
   - `Microsoft.CodeAnalysis.NetAnalyzers`
   - 自訂規則檢查 Mat/Bitmap Dispose

2. **SonarLint / SonarQube**
   - 程式碼異味檢測
   - 安全漏洞掃描

3. **FxCop**
   - .NET Framework 特定規則

4. **Custom Pre-commit Hook**
   - 檢查繁體中文註解
   - 驗證 AI 標註

---

## 參考資源 (References)

- **專案文件**: `CLAUDE.md` - 完整架構與編碼規範
- **實作摘要**: `IMPLEMENTATION_SUMMARY.md` - ParameterConfigForm 重構
- **資料庫**: `.\setting\mydb.sqlite` - Schema 參考
- **日誌**: `.\logs\peilin_log-{Date}.txt` - 執行時期行為
- **模型**: `.\models\` - AI 模型檔案

---

**最後更新**: 2025-10-29
**維護者**: 請保持此文件與 CLAUDE.md 同步更新
