# 🔍 記憶體安全完整稽核報告

**稽核日期**: 2025-10-13  
**核心問題**: 過度使用 Clone() 和 using 導致提早釋放，機台無法運行  
**原則**: 以最少變動為前提，確保機台順暢運作

---

## 📋 檢查清單

### ✅ 第一層：相機觸發層 (Camera0.cs)
- [ ] OnImageGrabbed()
- [ ] GrabResultToMat()
- [ ] SaveTouchedImageAndGetPath()
- [ ] Form1.SaveImageAsync()

### ✅ 第二層：影像分發層 (Form1.cs)
- [ ] Receiver()
- [ ] Queue_Bitmap1-4 管理

### ✅ 第三層：核心處理層 (Form1.cs)
- [ ] getMat1()
- [ ] getMat2()
- [ ] getMat3()
- [ ] getMat4()

### ✅ 第四層：影像分析函數
#### 圓心與 ROI 檢測
- [ ] DetectCircles()
- [ ] DetectAndExtractROI()
- [ ] CheckObjectPosition()
- [ ] findGapWidth()
- [ ] findGap()

#### YOLO 檢測
- [ ] YoloDetection.DetectBatch()
- [ ] YoloDetection.DrawDetectionResults()

#### 色彩分析 (OTP)
- [ ] ApplyOtpColorDetection()
- [ ] AnalyzeOtpColorFeatures()
- [ ] ExtractDefectRegion()
- [ ] CreateNonBlackMask()
- [ ] AnalyzeColorSpace()
- [ ] AnalyzeLabColorSpace()

#### 小黑點檢測
- [ ] DetectBlackDots()
- [ ] IsContourBlackInBinary()

#### 瑕疵過濾
- [ ] ApplyOutscOverkillReduction()
- [ ] IsDefectInNonRoiRegion_in()
- [ ] IsDefectInNonRoiRegion_out()

### ✅ 第五層：結果管理層
- [ ] ResultManager.AddResult()
- [ ] ResultManager.CombineResults()
- [ ] ResultManager.SaveFinalResult()
- [ ] ResultManager.CalculateAndSendPLCSignal()

### ✅ 輔助層：檔案儲存
- [ ] sv()
- [ ] SaveImageAsync()

---

## 🎯 檢查標準

### Mat 對象生命週期規則
1. **創建者負責釋放**: 誰創建 Mat，誰負責 Dispose()
2. **傳遞不需 Clone**: 如果接收方只讀取，不需 Clone
3. **共享需要 Clone**: 如果多個執行緒/佇列使用，必須 Clone
4. **返回值需明確**: 返回 Mat 的函數必須註明誰負責釋放

### Clone 使用判定
| 情境 | 是否需要 Clone | 理由 |
|------|---------------|------|
| 放入單一佇列，後續唯一讀取 | ❌ | 所有權轉移，不需複製 |
| 放入佇列後，原函數繼續使用 | ✅ | 避免共享引用 |
| 跨執行緒傳遞且雙方都使用 | ✅ | 避免 Race Condition |
| 函數內部暫存變數 | ❌ | 局部變數用完即釋放 |
| 返回值給呼叫方 | ❌ | 所有權轉移給呼叫方 |

### Using 語句判定
| 情境 | 是否需要 Using | 理由 |
|------|---------------|------|
| 函數內部建立並使用完畢 | ✅ | 確保釋放 |
| 建立後放入佇列供其他執行緒 | ❌ | 接收方負責釋放 |
| 從佇列取出的物件 | ❌ | 取出後成為所有者，手動釋放 |
| 返回值 | ❌ | 呼叫方負責釋放 |

---

## 🔬 詳細檢查結果

---

## ❌ 發現的嚴重問題

### **問題 1: OnImageGrabbed() 過度使用 Clone() 導致效能問題**

**位置**: Camera0.cs 行 936-941, 952-957

**問題代碼**:
```csharp
// ❌ 錯誤: imageForSave 是從 src Clone 來的，傳給 SaveImageAsync 又沒被用完就被 Dispose
using (Mat imageForSave = src.Clone())
{
    string imagePath = SaveTouchedImageAndGetPath(imageForSave, cameraIndex, intervalMs, "critical_short");
    LogTriggerStatistics(...);
} // imageForSave 被釋放

// SaveTouchedImageAndGetPath 內部:
form1.SaveImageAsync(image, fullPath);  // 把 image 放入 Queue_Save
// ❌ 但 using 立即釋放了 imageForSave！
```

**結果**: `sv()` 函數從 `Queue_Save` 取出的 Mat 已經被釋放 → **ObjectDisposedException**

**修正方案**:
```csharp
// ✅ 方案A: 移除 using，SaveImageAsync 內部 Clone
Mat imageForSave = src.Clone();
string imagePath = SaveTouchedImageAndGetPath(imageForSave, cameraIndex, intervalMs, "critical_short");
LogTriggerStatistics(...);
// 不釋放，交給 sv() 處理

// 修改 SaveImageAsync:
public void SaveImageAsync(Mat image, string path)
{
    app.Queue_Save.Enqueue(new ImageSave(image.Clone(), path)); // Clone 在這裡
    app._sv.Set();
}
```

**理由**: 
- 原本設計 `SaveImageAsync` 只是將引用放入佇列
- 但加了 `using (Mat imageForSave = src.Clone())` 後立即釋放
- 應該讓 `SaveImageAsync` 負責 Clone，或者移除 using

---

### **問題 2: SaveImageAsync() 沒有 Clone，與 Queue_Save 共享引用**

**位置**: Form1.cs 行 14839-14843

**問題代碼**:
```csharp
// ❌ 直接傳引用，沒有 Clone
public void SaveImageAsync(Mat image, string path)
{
    app.Queue_Save.Enqueue(new ImageSave(image, path));
    app._sv.Set();
}
```

**結果**: 
- 如果呼叫方釋放了 `image`，`sv()` 會存取已釋放的記憶體
- Camera0.cs 的 `using (Mat imageForSave = src.Clone())` 立即釋放

**修正方案**:
```csharp
public void SaveImageAsync(Mat image, string path)
{
    // ✅ 在這裡 Clone，確保 Queue_Save 擁有獨立副本
    app.Queue_Save.Enqueue(new ImageSave(image.Clone(), path));
    app._sv.Set();
}
```

---

### **問題 3: DetectAndExtractROI() 返回的 Mat 未被釋放**

**位置**: Form1.cs 行 830, 887, 1434, 1485, 1992, 2603

**問題代碼**:
```csharp
Mat roi = DetectAndExtractROI(input.image, input.stop, input.count);
// ❌ roi 從未被釋放！

Mat chamferRoi = DetectAndExtractROI(input.image, input.stop, input.count, true);
// ❌ chamferRoi 從未被釋放！
```

**DetectAndExtractROI 內部** (行 8169-8469):
```csharp
private Mat DetectAndExtractROI(...)
{
    Mat roi_final = new Mat();
    try {
        // ... 處理 ...
        return roi_final; // ✅ 返回新創建的 Mat
    }
    finally {
        mask?.Dispose();
        roi_full?.Dispose();
        // ❌ roi_final 不能在這裡釋放，因為要返回
    }
}
```

**呼叫方** (getMat1-4):
```csharp
Mat roi = DetectAndExtractROI(...);
DetectionResponse detection = await _yoloDetection.PerformObjectDetection(roi, url);
// ❌ roi 用完後未釋放 → 記憶體洩漏 (每張 15-20 MB)
```

**修正方案**:
```csharp
// ✅ 方案: 在 getMat1-4 中用完 roi 後釋放
using (Mat roi = DetectAndExtractROI(input.image, input.stop, input.count))
{
    DetectionResponse detection = await _yoloDetection.PerformObjectDetection(roi, url);
    // ... 使用檢測結果 ...
} // roi 自動釋放

// 倒角檢測同理
using (Mat chamferRoi = DetectAndExtractROI(input.image, input.stop, input.count, true))
{
    DetectionResponse chamferDetection = await _yoloDetection.PerformObjectDetection(chamferRoi, url);
    // ... 處理 ...
} // chamferRoi 自動釋放
```

---

### **問題 4: findGap() 返回的 nong Mat 未被釋放**

**位置**: Form1.cs 行 833, 1407, 1961

**問題代碼**:
```csharp
(bool nonb, Mat nong, List<Point> detectedGapPositions) = findGap(input.image, input.stop);
// ❌ nong 從未被釋放！
```

**修正方案**:
```csharp
(bool nonb, Mat nong, List<Point> detectedGapPositions) = findGap(input.image, input.stop);
try {
    // ... 使用 nong ...
}
finally {
    nong?.Dispose(); // ✅ 確保釋放
}
```

---

### **問題 5: Receiver() 直接傳遞 Mat 引用，導致共享**

**位置**: Form1.cs 行 578-612, Camera0.cs 行 989

**問題代碼**:
```csharp
// Camera0.cs - OnImageGrabbed():
using (Mat src = GrabResultToMat(grabResult))
{
    form1.Receiver(cameraIndex, src, time_start); // ❌ 直接傳 src 引用
} // src 被釋放

// Form1.cs - Receiver():
app.Queue_Bitmap1.Enqueue(new ImagePosition(Src, 1, ...)); // ❌ Src 已被釋放
```

**結果**: 佇列中的 Mat 已經無效 → **ObjectDisposedException**

**修正方案 A (建議)**: 移除 using，讓 getMat1-4 負責釋放
```csharp
// Camera0.cs - OnImageGrabbed():
Mat src = GrabResultToMat(grabResult);
form1.Receiver(cameraIndex, src, time_start);
// ✅ 不釋放，交給 getMat1-4 的 finally

// getMat1-4 已有:
try {
    app.Queue_Bitmap1.TryDequeue(out input);
    // ... 處理 ...
}
finally {
    input.image?.Dispose(); // ✅ 在這裡釋放
}
```

**修正方案 B (更安全但效能較差)**: Receiver 內部 Clone
```csharp
// Form1.cs - Receiver():
app.Queue_Bitmap1.Enqueue(new ImagePosition(Src.Clone(), 1, ...));
// ✅ Clone 確保獨立副本

// Camera0.cs - OnImageGrabbed():
using (Mat src = GrabResultToMat(grabResult))
{
    form1.Receiver(cameraIndex, src, time_start);
} // ✅ 可以安全釋放
```

---

### **問題 6: getMat1-4 中 input.image 的 Clone 是多餘的**

**位置**: Form1.cs 行 738, 1317, 1870, 2510

**問題代碼**:
```csharp
try {
    if (原圖ToolStripMenuItem.Checked)
    {
        app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), path)); // ❌ 多餘的 Clone
    }
    // ... 處理 ...
}
finally {
    input.image?.Dispose(); // ✅ 釋放原始圖
}
```

**分析**:
- `input.image` 由 finally 負責釋放
- `Queue_Save` 需要獨立副本 → Clone 是**必要的**
- **但** 如果 SaveImageAsync 內部已經 Clone，這裡就多餘了

**修正方案**:
```csharp
// 方案 A: SaveImageAsync 內部 Clone (推薦)
app.Queue_Save.Enqueue(new ImageSave(input.image, path)); // ❌ 移除這裡的 Clone
// SaveImageAsync 內部會 Clone

// 方案 B: 保持現狀
app.Queue_Save.Enqueue(new ImageSave(input.image.Clone(), path)); // ✅ 這裡 Clone
// SaveImageAsync 不 Clone
```

---

### **問題 7: ResultManager.SaveFinalResult() 中 FinalMap Clone 必要性存疑**

**位置**: Form1.cs 行 16210, 16218

**問題代碼**:
```csharp
app.Queue_Save.Enqueue(new ImageSave(markedImage.Clone(), savePath));
app.Queue_Save.Enqueue(new ImageSave(markedImage.Clone(), stationSavePath));
```

**分析**:
- `markedImage` 來自 `stationResult.FinalMap`
- 如果 `stationResult.FinalMap` 用完後會被釋放，Clone 是**必要的**
- 如果 `stationResult.FinalMap` 沒人釋放，這會導致**記憶體洩漏**

**需要檢查**: StationResult.FinalMap 的生命週期

---

### **問題 8: DrawDetectionResults() 返回的 Mat 未被一致釋放**

**位置**: Form1.cs 行 974, 1091, 1211, 等多處

**問題代碼**:
```csharp
// ✅ 有些地方有 using
using (Mat chamfer_resultImage = _yoloDetection.DrawDetectionResults(...))
{
    StationResult chamferResult = new StationResult {
        FinalMap = chamfer_resultImage, // ❌ 引用傳遞
    };
} // chamfer_resultImage 被釋放 → FinalMap 無效

// ❌ 有些地方沒有 using
Mat resultImage = _yoloDetection.DrawDetectionResults(...);
StationResult result = new StationResult {
    FinalMap = resultImage, // ❌ resultImage 未釋放
};
```

**修正方案**:
```csharp
// ✅ FinalMap 必須擁有獨立副本
using (Mat resultImage = _yoloDetection.DrawDetectionResults(...))
{
    StationResult result = new StationResult {
        FinalMap = resultImage.Clone(), // ✅ Clone 給 FinalMap
    };
} // resultImage 被釋放，FinalMap 保留副本

// StationResult 使用完畢後，由 ResultManager 負責釋放 FinalMap
```

---

## 📊 記憶體洩漏總覽

| 位置 | Mat 對象 | 大小估計 | 洩漏頻率 | 影響 |
|------|---------|---------|---------|------|
| DetectAndExtractROI() 返回的 roi | 15-20 MB | 每個樣品 4-8 次 | **嚴重** |
| DetectAndExtractROI() 返回的 chamferRoi | 15-20 MB | 每個樣品 0-2 次 | **中等** |
| findGap() 返回的 nong | 5-10 MB | 每個樣品 4 次 | **嚴重** |
| DrawDetectionResults() 返回的 Mat | 15-20 MB | 每個樣品 1-4 次 | **中等** |
| StationResult.FinalMap | 15-20 MB | 每個樣品 4 次 | **需確認** |

**總計**: 每個樣品洩漏 **100-300 MB**，連續運行 10 分鐘會耗盡記憶體！

---

## ✅ 修正優先級

### 🔴 P0 (立即修正 - 導致 Crash)
1. **修正問題 1 & 2**: SaveImageAsync 內部 Clone
2. **修正問題 5**: 移除 OnImageGrabbed 的 using 或 Receiver Clone

### 🟠 P1 (高優先 - 記憶體洩漏)
3. **修正問題 3**: 為 DetectAndExtractROI 返回值加 using
4. **修正問題 4**: 為 findGap 返回的 nong 加釋放

### 🟡 P2 (中優先 - 優化效能)
5. **修正問題 6**: 統一 Clone 策略，避免雙重 Clone
6. **修正問題 8**: 統一 DrawDetectionResults 返回值處理

### 🟢 P3 (低優先 - 確認正確性)
7. **修正問題 7**: 確認 StationResult.FinalMap 生命週期

