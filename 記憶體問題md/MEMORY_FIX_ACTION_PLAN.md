# 🔧 記憶體安全修正執行計劃

**日期**: 2025-10-13  
**目標**: 以最少變動修正所有記憶體安全問題  
**原則**: 誰創建誰釋放，誰使用誰擁有

---

## 📊 問題嚴重程度分級

### 🔴 P0 - 導致 Crash（立即修正）
這些問題導致 `ObjectDisposedException`，機台無法運行

### 🟠 P1 - 記憶體洩漏（高優先）
這些問題導致記憶體持續增長，10分鐘內耗盡記憶體

### 🟡 P2 - 效能問題（中優先）
這些問題導致不必要的記憶體分配，影響效能

### 🟢 P3 - 潛在風險（低優先）
這些問題可能在特定情況下發生

---

## 🎯 修正策略

### **核心原則：最少變動**

1. **不改變函數簽名** - 避免連鎖修改
2. **不改變數據流** - 保持原有邏輯
3. **只修改生命週期管理** - 加 using 或 Dispose
4. **統一 Clone 策略** - 避免雙重 Clone

---

## 📝 修正清單

### 🔴 P0-1: SaveImageAsync 必須 Clone

**位置**: `Form1.cs:14839`

**問題**: SaveImageAsync 直接傳引用，但呼叫方可能釋放

**修正**:
```csharp
// 修正前
public void SaveImageAsync(Mat image, string path)
{
    app.Queue_Save.Enqueue(new ImageSave(image, path)); // ❌ 沒 Clone
    app._sv.Set();
}

// 修正後
public void SaveImageAsync(Mat image, string path)
{
    app.Queue_Save.Enqueue(new ImageSave(image.Clone(), path)); // ✅ Clone
    app._sv.Set();
}
```

**影響範圍**: 所有呼叫 SaveImageAsync 的地方
**測試重點**: 存圖功能正常，無 ObjectDisposedException

---

### 🔴 P0-2: OnImageGrabbed 的 using 移除

**位置**: `Camera0.cs:936-957`

**問題**: imageForSave 被 using 釋放，但已經放入 Queue_Save

**修正**:
```csharp
// 修正前
using (Mat imageForSave = src.Clone())
{
    string imagePath = SaveTouchedImageAndGetPath(imageForSave, cameraIndex, intervalMs, "critical_short");
    LogTriggerStatistics(...);
} // ❌ imageForSave 被釋放

// 修正後
Mat imageForSave = src.Clone();
string imagePath = SaveTouchedImageAndGetPath(imageForSave, cameraIndex, intervalMs, "critical_short");
LogTriggerStatistics(...);
// ✅ 不釋放，交給 sv() 處理（SaveImageAsync 已經 Clone）
imageForSave?.Dispose(); // ✅ 因為 SaveImageAsync 會 Clone，這裡可以釋放
```

**影響範圍**: Camera0.cs OnImageGrabbed 函數
**測試重點**: 誤觸圖片正常儲存

---

### 🔴 P0-3: OnImageGrabbed 移除 src 的 using

**位置**: `Camera0.cs:912`

**問題**: src 被 using 釋放，但 Receiver 直接將其放入佇列

**修正**:
```csharp
// 修正前
using (Mat src = GrabResultToMat(grabResult))
{
    form1.Receiver(cameraIndex, src, time_start);
} // ❌ src 被釋放，但佇列還在用

// 修正後
Mat src = GrabResultToMat(grabResult);
try
{
    // ... 所有處理邏輯 ...
    form1.Receiver(cameraIndex, src, time_start);
}
catch (Exception ex)
{
    Console.WriteLine($"OnImageGrabbed error: {ex.Message}");
    src?.Dispose(); // ✅ 只在異常時釋放
}
// ✅ 正常流程不釋放，交給 getMat1-4 的 finally
```

**影響範圍**: Camera0.cs OnImageGrabbed 函數
**測試重點**: 正常檢測流程無 ObjectDisposedException

---

### 🔴 P0-4: getMat1-4 移除 Queue_Save 的 Clone

**位置**: `Form1.cs:738, 1317, 1870, 2510`

**問題**: 雙重 Clone（getMat 和 SaveImageAsync 都 Clone）

**修正**:
```csharp
// 修正前
app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), path)); // ❌ Clone

// 修正後
app.Queue_Save.Enqueue(new ImageSave(input.image, path)); // ✅ 不 Clone
// SaveImageAsync 內部會 Clone
```

**影響範圍**: getMat1-4 存原圖部分
**測試重點**: 原圖正常儲存，無記憶體洩漏

---

### 🟠 P1-1: DetectAndExtractROI 返回值加 using

**位置**: `Form1.cs:830, 887, 1434, 1485, 1992, 2603`

**問題**: roi 和 chamferRoi 從未釋放（每張 15-20 MB）

**修正 - getMat1 範例**:
```csharp
// 修正前
Mat roi = DetectAndExtractROI(input.image, input.stop, input.count);
DetectionResponse detection = await _yoloDetection.PerformObjectDetection(roi, url);
// ❌ roi 未釋放

// 修正後
using (Mat roi = DetectAndExtractROI(input.image, input.stop, input.count))
{
    (bool nonb, Mat nong, List<Point> detectedGapPositions) = findGap(input.image, input.stop);
    try
    {
        // ... NROI 檢測 ...
        DetectionResponse detection = await _yoloDetection.PerformObjectDetection(roi, url);
        // ... 其他處理 ...
    }
    finally
    {
        nong?.Dispose(); // ✅ 同時處理 P1-2
    }
} // ✅ roi 自動釋放
```

**影響範圍**: getMat1-4 所有使用 DetectAndExtractROI 的地方
**測試重點**: 正常檢測，記憶體不增長

---

### 🟠 P1-2: findGap 返回的 nong 加釋放

**位置**: `Form1.cs:833, 1407, 1961`

**問題**: nong Mat 從未釋放（每張 5-10 MB）

**修正**: 見 P1-1，在 roi 的 using 的 finally 中釋放

---

### 🟠 P1-3: chamferRoi 加 using

**位置**: `Form1.cs:887, 1485`

**修正**:
```csharp
// 修正前
Mat chamferRoi = DetectAndExtractROI(input.image, input.stop, input.count, true);
DetectionResponse chamferDetection = await _yoloDetection.PerformObjectDetection(chamferRoi, url);
// ❌ chamferRoi 未釋放

// 修正後
using (Mat chamferRoi = DetectAndExtractROI(input.image, input.stop, input.count, true))
{
    DetectionResponse chamferDetection = await _yoloDetection.PerformObjectDetection(chamferRoi, url);
    // ... 處理檢測結果 ...
} // ✅ chamferRoi 自動釋放
```

**影響範圍**: getMat1-2 倒角檢測部分
**測試重點**: 倒角檢測正常，記憶體不增長

---

### 🟡 P2-1: DrawDetectionResults 統一處理

**位置**: `Form1.cs` 多處

**問題**: 有些地方有 using 有些沒有，且 FinalMap 引用傳遞

**修正範例**:
```csharp
// 修正前
using (Mat resultImage = _yoloDetection.DrawDetectionResults(...))
{
    StationResult result = new StationResult {
        FinalMap = resultImage, // ❌ 引用傳遞，using 後失效
    };
}

// 修正後
using (Mat resultImage = _yoloDetection.DrawDetectionResults(...))
{
    StationResult result = new StationResult {
        FinalMap = resultImage.Clone(), // ✅ Clone 給 FinalMap
    };
} // ✅ resultImage 釋放
```

**影響範圍**: 所有創建 StationResult 的地方
**測試重點**: 結果圖正常顯示和儲存

---

### 🟢 P3-1: 確認 ResultManager 釋放 FinalMap

**位置**: `ResultManager.SaveFinalResult`

**問題**: markedImage 來自 stationResult.FinalMap，需確認生命週期

**檢查項目**:
1. SaveFinalResult 儲存後是否釋放 FinalMap？
2. CombineResults 使用後是否釋放？
3. sampleResult 清理時是否釋放所有 FinalMap？

**修正**: 如果未釋放，在適當位置加 Dispose

---

## 🔍 修正後檢查清單

### 編譯檢查
- [ ] 無編譯錯誤
- [ ] 無編譯警告

### 功能測試
- [ ] 正常檢測流程無 ObjectDisposedException
- [ ] 存原圖功能正常
- [ ] 誤觸圖片正常儲存
- [ ] NROI 檢測正常
- [ ] 倒角檢測正常
- [ ] OTP 色彩檢測正常
- [ ] 結果圖正常顯示

### 記憶體測試
- [ ] 連續運行 10 分鐘記憶體穩定
- [ ] 連續運行 1 小時記憶體不增長
- [ ] 急停/重啟無記憶體洩漏

### 效能測試
- [ ] 檢測速度不低於修正前
- [ ] CPU 使用率正常
- [ ] 無明顯卡頓

---

## 📌 修正順序

1. **P0-1**: SaveImageAsync 加 Clone（基礎設施）
2. **P0-2**: OnImageGrabbed 移除 imageForSave 的 using
3. **P0-3**: OnImageGrabbed 移除 src 的 using
4. **P0-4**: getMat1-4 移除 Queue_Save 的 Clone
5. **編譯測試** - 確認 P0 修正無誤
6. **P1-1**: DetectAndExtractROI 返回值加 using
7. **P1-2**: findGap 返回值加釋放
8. **P1-3**: chamferRoi 加 using
9. **功能測試** - 確認 P1 修正無誤
10. **P2-1**: DrawDetectionResults 統一處理
11. **P3-1**: 確認 ResultManager 釋放 FinalMap
12. **完整測試** - 編譯、功能、記憶體、效能

---

## ⚠️ 注意事項

1. **每次修正後立即編譯** - 確保語法正確
2. **分階段測試** - P0 → P1 → P2 → P3
3. **記錄修正前後差異** - 便於回溯
4. **保留原始代碼註解** - 便於理解修正原因
5. **測試所有站點** - 確保四個站點都正常

---

**準備好開始修正了嗎？**
