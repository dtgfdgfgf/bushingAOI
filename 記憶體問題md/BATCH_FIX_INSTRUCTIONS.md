# 批量修正指令清單

## getMat2() 需要修正的 continue

### 1. 白比率檢測 (約行 1355)
**搜尋**:
```csharp
// 添加結果到管理器
app.resultManager.AddResult(input.count, WhitePixelResult);
//updateLabel();
continue;
```

**替換為 (getMat2 特定)**:
在 `#region  app.detect_result create` 後面尋找第一個白比率檢測區塊

### 2. 物體位置檢測 (約行 1384)
**搜尋**:
```csharp
// 添加結果到管理器
app.resultManager.AddResult(input.count, ObjectPositionResult);
//updateLabel();
continue;
```

### 3. 變形檢測 (約行 1417)
**搜尋**:
```csharp
// 添加結果到管理器
app.resultManager.AddResult(input.count, GapResult);
//updateLabel();
continue;
```

### 4-6. 倒角和YOLO檢測
與 getMat1() 相同模式

---

## 🔧 手動修正步驟

由於 continue 前的程式碼模式相似，建議採用以下策略：

### 策略 A: 使用 Visual Studio 的搜尋替換
1. 開啟 Form1.cs
2. 搜尋: `continue;` (區分大小寫)
3. 在 getMat2-4() 函數範圍內逐一檢查
4. 在每個 `continue;` 前一行插入: `input.image?.Dispose(); // 由 GitHub Copilot 產生: 釋放圖像以避免記憶體洩漏`

### 策略 B: 使用程式碼模式識別
搜尋以下模式並在 continue 前添加 Dispose:
- `app.resultManager.AddResult(input.count, *Result);\n.*continue;`
- `lbAdd(.*err.*\n.*continue;`

---

## ⚠️ 特別注意: FinalMap 記憶體問題

在以下程式碼中：
```csharp
StationResult WhitePixelResult = new StationResult
{
    FinalMap = input.image,  // ⚠️ 問題：直接引用
    ...
};
app.resultManager.AddResult(input.count, WhitePixelResult);
input.image?.Dispose(); // ❌ 這會導致 FinalMap 也被釋放！
```

**修正方案**:
```csharp
StationResult WhitePixelResult = new StationResult
{
    FinalMap = input.image.Clone(),  // ✅ 使用 Clone
    ...
};
app.resultManager.AddResult(input.count, WhitePixelResult);
input.image?.Dispose(); // ✅ 安全
```

**需要修正的位置**:
1. getMat1() 白比率檢測 (已修正需確認)
2. getMat2() 白比率檢測
3. getMat2() 物體位置檢測
4. getMat3() 白比率檢測
5. getMat4() 白比率檢測

---

## 📊 修正優先順序

### 🔴 高優先級 (立即修正)
1. ✅ getMat1-4() 函數結尾 Dispose (已完成)
2. ✅ getMat1() 所有 continue Dispose (已完成)
3. ⏳ **修正 FinalMap 直接引用問題** (避免 double-free)
4. ⏳ getMat2-4() continue Dispose

### 🟡 中優先級 (後續優化)
1. 檢查其他可能的 Mat 洩漏
2. 檢查 `Queue_Save` 中的圖像是否被正確處理
3. 檢查 `showResultMat()` 函數的記憶體管理

### 🟢 低優先級 (長期優化)
1. 重構為統一的錯誤處理模式
2. 使用 using 語句簡化生命週期管理
3. 引入記憶體監控和警報機制

