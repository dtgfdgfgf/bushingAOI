# FinalMap 記憶體洩漏根本原因完整分析

## 執行摘要

**記憶體洩漏確診**：
- **主因 1**: `sampleStationResults` 永久保留所有 `FinalMap`，從未釋放
- **主因 2**: FinalMap 賦值邏輯不一致，部分情況下共享引用導致 double-free 風險
- **累積洩漏**: 60 MB/樣品 × 3600 樣品 = **216 GB**
- **崩潰時間**: 36 分鐘（符合實測）

---

## 問題 1：FinalMap 賦值的三種模式（不一致性）

### 模式分類

#### ✅ **模式 A：Clone（正確但過度）**
```csharp
// Line 828, 1057, 1306, 1686, 1902, 2263, 2610, 2895, 3126
FinalMap = input.image.Clone()
FinalMap = resultImage.Clone()
FinalMap = chamfer_resultImage.Clone()
FinalMap = blackDotResultImage.Clone()
```
- **特點**: 建立獨立副本
- **優點**: 生命週期獨立，安全
- **缺點**: 記憶體開銷大（每個 15 MB）

#### ❌ **模式 B：直接引用（危險）**
```csharp
// Line 876, 1526, 2107, 2786 - gapResult
FinalMap = gapResult  // ❌ 來自 findGapWidth()

// Line 1449, 2041, 2072, 2750 - input.image
FinalMap = input.image  // ❌ 共享 input.image

// Line 1478 - CheckObjectResultMat
FinalMap = CheckObjectResultMat  // ❌ 來自 CheckObjectPosition()
```
- **特點**: 直接賦值引用
- **風險 1**: `input.image` 被 finally 釋放時，FinalMap 變成懸空指針
- **風險 2**: `gapResult` 和 `CheckObjectResultMat` 的所有權不明確

---

### 問題 1.1：`gapResult` 的所有權問題

**來源追蹤**：
```csharp
// Line 858, 1508
(gapIsNG, gapResult, non) = findGapWidth(gapInputImage, input.stop);

// findGapWidth 函數定義 (Line 9021)
public (bool isNG, Mat img, List<Point> gapPositions) findGapWidth(Mat img, int stop)
```

**問題**：
- `findGapWidth()` 返回的 `Mat img` 是否是新建立的？
- 如果是在函數內 Clone 的，誰負責釋放？
- 如果是傳入 `img` 的引用，會與 `gapInputImage` 衝突

**潛在 double-free 情境**：
```csharp
try {
    gapInputImage = input.image.Clone();
    (gapIsNG, gapResult, non) = findGapWidth(gapInputImage, input.stop);
    // 如果 gapResult == gapInputImage (同一個引用)
} finally {
    gapInputImage?.Dispose();  // 釋放
}

// 稍後
StationResult GapResult = new StationResult {
    FinalMap = gapResult,  // 已被釋放的引用！
};
```

---

### 問題 1.2：`input.image` 的共享引用問題

**危險賦值**：
```csharp
// Line 1449 (getMat2)
StationResult WhitePixelResult = new StationResult {
    FinalMap = input.image,  // ❌ 直接引用
    ...
};

// getMat2 的 finally 區塊
finally {
    input.image?.Dispose();  // ❌ 釋放 input.image
}
```

**記憶體錯誤**：
1. `FinalMap` 指向 `input.image`
2. finally 釋放 `input.image`
3. `FinalMap` 變成懸空指針
4. SaveFinalResult 中訪問 `FinalMap` → **ObjectDisposedException** 或 **Access Violation**

---

## 問題 2：sampleStationResults 永久保留記憶體

### 洩漏路徑

**Line 16603 (SaveFinalResult)**：
```csharp
// 複製每個樣品的所有站點結果
if (sampleResult.StationResults != null && sampleResult.StationResults.Count > 0)
{
    sampleStationResults[sampleResult.SampleId] = sampleResult.StationResults
        .ToDictionary(kv => kv.Key, kv => kv.Value);
}
```

**問題**：
- `sampleStationResults` 是 **static** ConcurrentDictionary
- **從未清除**，永久保留所有樣品
- 每個樣品的 4 個 StationResult，每個包含 15 MB FinalMap
- **永久駐留記憶體**

### 洩漏計算

```
每個樣品:
- 4 個 StationResult (4 個站點)
- 每個 FinalMap = 15 MB
- 總計 = 60 MB/樣品

累積 3600 樣品:
60 MB × 3600 = 216,000 MB = 211 GB

實際觀察:
- 36 分鐘後崩潰
- failed to allocate 15040512 bytes (14.3 MB)
- 符合記憶體耗盡症狀
```

---

### 問題 2.1：sampleStationResults 的使用場景

**唯一用途**：ExportStatsToCsv (Line 17189, 17226)

**分析**：
```csharp
public static bool ExportStatsToCsv(string filePath)
{
    // 讀取 sampleStationResults 生成報表
    foreach (var samplePair in sampleStationResults.OrderBy(r => r.Key))
    {
        // 遍歷所有樣品的所有站點結果
        // 但 **不需要 FinalMap**！只需要分數和瑕疵名稱
    }
}
```

**關鍵發現**：
- ExportStatsToCsv **根本不需要 FinalMap**
- 只需要 `IsNG`, `DefectName`, `DefectScore`
- **保留整個 StationResult (含 FinalMap) 完全沒必要**

---

## 問題 2.2：FinalMap 在 SaveFinalResult 後的命運

### 檢查 FinalMap 是否被後續使用

**Line 16499 (SaveFinalResult 內部)**：
```csharp
foreach (var stationResult in sampleResult.StationResults.Values)
{
    // ...
    using (Mat markedImage = stationResult.FinalMap.Clone())
    {
        // 在 FinalMap 上標記文字
        Cv2.PutText(markedImage, markText, ...);
        
        // 存檔（又 Clone 一次）
        app.Queue_Save.Enqueue(new ImageSave(markedImage.Clone(), savePath));
        app.Queue_Save.Enqueue(new ImageSave(markedImage.Clone(), stationSavePath));
    }
}
```

**關鍵**：
- SaveFinalResult **內部已經處理完 FinalMap**
- 已經 Clone 到 Queue_Save
- **SaveFinalResult 結束後，FinalMap 不再被需要**

**Line 16603 之後**：
```csharp
sampleStationResults[sampleResult.SampleId] = sampleResult.StationResults
    .ToDictionary(kv => kv.Key, kv => kv.Value);
// ❌ 這裡保留了 FinalMap，但已經沒用了
```

---

## 根本原因總結

### 洩漏點 1：sampleStationResults 永久保留（最嚴重）

| 項目 | 詳情 |
|------|------|
| 位置 | Line 16603 (SaveFinalResult) |
| 問題 | static ConcurrentDictionary 永久保留所有樣品的 StationResult |
| 影響 | 60 MB/樣品 × 3600 = 216 GB 累積洩漏 |
| 必要性 | ❌ ExportStatsToCsv 不需要 FinalMap |
| 優先級 | **P0 - 立即修正** |

### 洩漏點 2：FinalMap 賦值不一致（潛在 crash）

| 模式 | 位置 | 問題 | 風險 |
|------|------|------|------|
| Clone | Line 828, 1057... | 記憶體開銷大但安全 | 低 |
| 直接引用 input.image | Line 1449, 2041... | finally 釋放後變懸空 | **高（crash）** |
| 引用函數返回值 | Line 876, 1478... | 所有權不明確 | **中（洩漏或crash）** |

---

## 修正方案

### 方案 A：廢棄 sampleStationResults 的 FinalMap（推薦）

**目標**：只保留統計數據，不保留圖像

**修改 Line 16603**：
```csharp
// 原始（錯誤）
sampleStationResults[sampleResult.SampleId] = sampleResult.StationResults
    .ToDictionary(kv => kv.Key, kv => kv.Value);

// 修正：建立不含 FinalMap 的副本
sampleStationResults[sampleResult.SampleId] = sampleResult.StationResults
    .ToDictionary(
        kv => kv.Key, 
        kv => new StationResult {
            Stop = kv.Value.Stop,
            IsNG = kv.Value.IsNG,
            OkNgScore = kv.Value.OkNgScore,
            FinalMap = null,  // ✅ 不保留圖像
            DefectName = kv.Value.DefectName,
            DefectScore = kv.Value.DefectScore,
            OriName = kv.Value.OriName
        }
    );
```

**優點**：
- 記憶體從 60 MB/樣品降至 < 1 KB/樣品（**99.998% 減少**）
- ExportStatsToCsv 功能不受影響
- 不需要定期清理

---

### 方案 B：在 CombineResults 後釋放 FinalMap

**位置**：Line 16332 (CombineResults)

```csharp
private void CombineResults(SampleResult sampleResult)
{
    // 現有邏輯...
    bool finalIsNG = sampleResult.StationResults.Values.Any(r => r.IsNG);
    SaveFinalResult(sampleResult, finalIsNG, isNull: false);
    CalculateAndSendPLCSignal(sampleResult, finalIsNG, isNull: false);

    // ✅ 新增：SaveFinalResult 完成後立即釋放 FinalMap
    foreach (var stationResult in sampleResult.StationResults.Values)
    {
        stationResult.FinalMap?.Dispose();
        stationResult.FinalMap = null;  // 避免 double-free
    }

    results.TryRemove(sampleResult.SampleId, out _);
}
```

**問題**：
- 如果 Line 16603 保留了引用，這裡釋放會導致懸空指針
- **必須配合方案 A 一起使用**

---

### 方案 C：定期清理 sampleStationResults（次要）

**位置**：SaveFinalResult 結尾

```csharp
// Line 16610 後添加
// 只保留最近 100 個樣品，避免無限增長
if (sampleStationResults.Count > 100)
{
    var oldSamples = sampleStationResults.Keys
        .OrderBy(k => k)
        .Take(sampleStationResults.Count - 100)
        .ToList();
    
    foreach (var oldId in oldSamples)
    {
        if (sampleStationResults.TryRemove(oldId, out var oldResults))
        {
            foreach (var sr in oldResults.Values)
            {
                sr.FinalMap?.Dispose();
            }
        }
    }
}
```

**缺點**：
- 治標不治本
- 仍然保留 100 × 60 MB = 6 GB 記憶體
- 如果 ExportStatsToCsv 需要完整歷史會有問題

---

### 方案 D：統一 FinalMap 賦值邏輯（修正不一致性）

**規則**：
1. **如果 FinalMap 生命週期 > getMat1-4 函數**：必須 Clone
2. **如果來源會被 finally 釋放**：必須 Clone
3. **如果來源所有權不明確**：必須 Clone

**需要修正的位置**：

```csharp
// Line 876, 1526, 2107, 2786
FinalMap = gapResult,  // ❌
// 改為
FinalMap = gapResult?.Clone(),  // ✅

// Line 1449, 2041, 2072, 2750
FinalMap = input.image,  // ❌
// 改為
FinalMap = input.image?.Clone(),  // ✅

// Line 1478
FinalMap = CheckObjectResultMat,  // ❌
// 改為
FinalMap = CheckObjectResultMat?.Clone(),  // ✅
```

**但注意**：
- 如果採用方案 A，SaveFinalResult 後立即丟棄 FinalMap
- 這些 Clone 仍然必要（避免 SaveFinalResult 內部訪問已釋放的 Mat）

---

## 推薦實施順序

### 階段 1：立即修正（P0）

1. **方案 A**：Line 16603 不保留 FinalMap（優先級最高）
2. **方案 B**：Line 16332 釋放 FinalMap
3. **方案 D**：統一所有 FinalMap 賦值為 Clone

**預期效果**：
- 記憶體流量從 60 MB/樣品降至 15 MB/樣品（Receiver Clone）
- 累積洩漏從 216 GB 降至 0
- 可穩定運行 24/7

---

### 階段 2：優化（P1）

4. **方案 C**：定期清理 sampleStationResults（保險措施）
5. 移除 getMat1-4 內部不必要的 Clone（如 gapInputImage, roiInputImage）
6. 評估 Receiver Clone 是否可優化

**預期效果**：
- 記憶體流量降至 6 GB/分鐘
- 峰值記憶體 < 2 GB

---

## 驗證方法

### 記憶體監控

```csharp
// 在 SaveFinalResult 開頭添加
var process = System.Diagnostics.Process.GetCurrentProcess();
long memoryUsed = process.WorkingSet64 / 1024 / 1024; // MB
Log.Information($"SaveFinalResult - SampleID: {sampleResult.SampleId}, Memory: {memoryUsed} MB, sampleStationResults Count: {sampleStationResults.Count}");
```

### 預期日誌

**修正前**：
```
SampleID: 100,  Memory: 1500 MB,  sampleStationResults Count: 100
SampleID: 200,  Memory: 3000 MB,  sampleStationResults Count: 200
SampleID: 3600, Memory: 54000 MB, sampleStationResults Count: 3600  // 崩潰
```

**修正後**：
```
SampleID: 100,  Memory: 500 MB,   sampleStationResults Count: 100
SampleID: 200,  Memory: 600 MB,   sampleStationResults Count: 200
SampleID: 3600, Memory: 1200 MB,  sampleStationResults Count: 3600  // 正常
```

---

## 結論

| 問題 | 嚴重性 | 優先級 | 修正方案 |
|------|--------|--------|----------|
| sampleStationResults 洩漏 | 致命 | P0 | 方案 A + B |
| FinalMap 賦值不一致 | 高 | P0 | 方案 D |
| 過度 Clone | 中 | P1 | 階段 2 優化 |

**預期結果**：
- ✅ 消除 216 GB 記憶體洩漏
- ✅ 避免 ObjectDisposedException
- ✅ 支援 24/7 穩定運行
- ✅ 記憶體峰值 < 2 GB
