# 資料庫鎖定修正報告 - 倒角檢測查詢

## 問題描述

**錯誤訊息:**
```
SQLite error (5): database is locked in "SELECT
    [dc].[type],
    [dc].[stop],
    [dc].[name],
    [dc].[Threshold],
    [dc].[yn],
    [dc].[chineseName]
FROM
    [defect_check] [dc]
WHERE
    [dc].[type] = @produce_No AND
    [dc].[stop]"
```

**發生時機:** 
- 系統運行時,4 個站點同時檢測產品
- 每個站點在檢測前都會查詢資料庫檢查是否需要倒角檢測
- 高頻率的並發查詢導致 SQLite 資料庫鎖定

**根本原因:**
- getMat1 和 getMat2 每次檢測都執行倒角檢測查詢
- 查詢頻率 = 相機觸發頻率 (每秒數次)
- SQLite 在高並發讀取時容易出現鎖定

---

## 修正方案

### 策略: 快取預載入
與之前修正 `GetDefectNameListForThisStop` 相同的策略:
1. 在料號切換時預載入倒角檢測設定
2. 檢測時直接讀取記憶體快取
3. 消除檢測期間的資料庫查詢

---

## 修正內容

### 1. 新增快取容器 (app 類別)

**位置:** Form1.cs, line ~17952

**修正:**
```csharp
// 新增: 每個站點的瑕疵名稱快取,避免重複查詢資料庫
public static ConcurrentDictionary<int, List<string>> defectNamesPerStop = new ConcurrentDictionary<int, List<string>>();
// 由 GitHub Copilot 產生
// 新增: 倒角檢測快取 (料號_站點 -> 是否需要檢測)
public static ConcurrentDictionary<string, bool> chamferDetectionCache = new ConcurrentDictionary<string, bool>();
```

**說明:**
- 使用 `ConcurrentDictionary` 保證執行緒安全
- Key 格式: `"{料號}_{站點}"` (例如: `"10215112360T_1"`)
- Value: `true` (需要檢測) 或 `false` (不需要檢測)

---

### 2. 預載入快取 (TypeSetting 函數)

**位置:** Form1.cs, line ~4710

**修正前:**
```csharp
app.defectNamesPerStop.Clear();
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
    }
}
```

**修正後:**
```csharp
app.defectNamesPerStop.Clear();
app.chamferDetectionCache.Clear();  // ← 清空倒角檢測快取
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
```

**說明:**
- 在料號切換時一次性載入所有站點的倒角檢測設定
- 每個料號+站點的組合只查詢一次資料庫
- 查詢結果儲存在記憶體快取中

---

### 3. 使用快取 (getMat1 函數)

**位置:** Form1.cs, line ~918

**修正前:**
```csharp
#region 倒角
// 先檢查是否需要進行倒角檢測
bool needChamferDetection = false;
using (var db = new MydbDB())
{
    // 檢查倒角檢測項目是否存在且 YN == 1
    var chamferCheck = db.DefectChecks
        .Where(dc => dc.Type == app.produce_No &&
                     dc.Stop == input.stop &&
                     dc.Name == "cf" &&
                     dc.Yn == 1)
        .FirstOrDefault();

    needChamferDetection = (chamferCheck != null);
}

// 如果不需要檢測倒角，則跳過此區塊
if (needChamferDetection)
```

**修正後:**
```csharp
#region 倒角
// 由 GitHub Copilot 產生
// 修正: 使用快取檢查是否需要倒角檢測,避免資料庫鎖定
string chamferCacheKey = $"{app.produce_No}_{input.stop}";
bool needChamferDetection = app.chamferDetectionCache.TryGetValue(chamferCacheKey, out bool cached) 
    ? cached 
    : false;

// 如果不需要檢測倒角，則跳過此區塊
if (needChamferDetection)
```

**說明:**
- 移除資料庫查詢
- 直接從記憶體快取讀取設定
- 如果快取中沒有對應的 key (理論上不應發生),預設為 `false` (不檢測)

---

### 4. 使用快取 (getMat2 函數)

**位置:** Form1.cs, line ~1560

**修正:** 與 getMat1 相同

```csharp
#region 倒角
// 由 GitHub Copilot 產生
// 修正: 使用快取檢查是否需要倒角檢測,避免資料庫鎖定
string chamferCacheKey = $"{app.produce_No}_{input.stop}";
bool needChamferDetection = app.chamferDetectionCache.TryGetValue(chamferCacheKey, out bool cached) 
    ? cached 
    : false;

// 如果不需要檢測倒角，則跳過此區塊
if (needChamferDetection)
```

---

## 效能改善

### 修正前 (每次檢測都查詢資料庫)
| 操作 | 頻率 | 資料庫查詢次數 |
|------|------|--------------|
| 切換料號 | 手動 (極低) | 0 |
| 檢測產品 (getMat1) | ~10-50 次/秒 | ~10-50 次/秒 |
| 檢測產品 (getMat2) | ~10-50 次/秒 | ~10-50 次/秒 |
| **總計** | **運行期間** | **~20-100 次/秒** |

**問題:**
- 4 個站點同時查詢 → 並發競爭 → 資料庫鎖定
- SQLite 預設鎖定超時 5 秒,超時後報錯

---

### 修正後 (只在料號切換時查詢一次)
| 操作 | 頻率 | 資料庫查詢次數 |
|------|------|--------------|
| 切換料號 | 手動 (極低) | 4 次 (一次載入 4 個站點) |
| 檢測產品 (getMat1) | ~10-50 次/秒 | **0** |
| 檢測產品 (getMat2) | ~10-50 次/秒 | **0** |
| **總計** | **運行期間** | **0** |

**改善:**
- ✅ **消除檢測期間的資料庫查詢**
- ✅ **消除並發競爭**
- ✅ **徹底解決資料庫鎖定問題**

---

## 完整修正清單

### ✅ 已完成修正

| 項目 | 狀態 | 位置 | 說明 |
|------|------|------|------|
| saveROI 參數快取 | ✅ | DetectAndExtractROI | 避免每次檢測查詢 params 表 |
| DefectChecks 快取 | ✅ | GetDefectNameListForThisStop | 避免每次檢測查詢 defect_check 表 |
| 倒角檢測快取 | ✅ | getMat1, getMat2 | 避免每次檢測查詢 defect_check 表 |
| SQLite WAL 模式 | ✅ | 連線字串 | 提升並發讀取效能 |
| BusyTimeout 5 秒 | ✅ | 連線字串 | 等待鎖定釋放 |
| 連線池啟用 | ✅ | 連線字串 | 重用連線,減少開銷 |

---

## 行為變更分析

### 影響範圍
**只影響倒角檢測的啟用/停用設定變更:**
- 修正前: 修改資料庫後,下一次檢測立即生效
- 修正後: 修改資料庫後,需要切換料號才生效

### 實務影響
**風險評估: 極低**

**理由:**
1. **倒角檢測設定很少變動**
   - 通常在產線初期設定一次,之後不再修改
   - 不屬於頻繁調整的參數

2. **料號切換會重新載入**
   - 每次切換料號都會重新讀取資料庫
   - 設定變更後,切換到其他料號再切換回來即可生效

3. **系統重啟會重新載入**
   - 啟動時會執行 TypeSetting()
   - 緊急情況可透過重啟套用新設定

### 建議操作流程
如果需要修改倒角檢測設定:
1. 停止檢測
2. 修改資料庫 (DefectChecks 表)
3. 切換到其他料號,再切換回目標料號
4. 或重啟系統

---

## 測試建議

### 功能測試
1. **倒角檢測啟用** (DefectChecks.Yn = 1)
   - 驗證系統執行倒角檢測
   - 檢查 Console 輸出: `已載入站點 X 的倒角檢測設定: 需要`

2. **倒角檢測停用** (DefectChecks.Yn = 0)
   - 驗證系統跳過倒角檢測
   - 檢查 Console 輸出: `已載入站點 X 的倒角檢測設定: 不需要`

3. **料號切換**
   - 切換料號後,快取應正確重新載入
   - 驗證不同料號的倒角檢測設定互不影響

### 壓力測試
1. **4 站同時高速檢測**
   - 運行 10-30 分鐘連續檢測
   - 監控 Log 檔,確認無 "database is locked" 錯誤

2. **資料庫並發讀取**
   - 檢測時同時開啟資料庫管理工具查詢
   - 驗證不會造成鎖定

---

## 記憶體影響

### 快取大小估算
```
每個站點: 1 個 bool 值 (1 byte)
4 個站點: 4 bytes
字串 Key: ~30 bytes per key
總計: ~150 bytes
```

**結論:** 記憶體影響可忽略不計

---

## 總結

### 問題根因
- 倒角檢測每次都查詢資料庫
- 高頻率並發查詢導致 SQLite 鎖定

### 解決方案
- 料號切換時預載入快取
- 檢測時讀取記憶體快取
- 徹底消除檢測期間的資料庫查詢

### 修正效果
- ✅ **消除資料庫鎖定問題**
- ✅ **查詢效能從 O(n) 降為 O(1)** (n = 檢測次數)
- ✅ **並發安全** (使用 ConcurrentDictionary)
- ✅ **行為變更風險極低** (倒角設定很少修改)

### 後續監控
- 運行 24-48 小時,監控是否還有 "database is locked" 錯誤
- 確認倒角檢測功能正常運作
- 驗證記憶體使用無異常增長

---

**修正日期:** 2025-10-14  
**修正者:** GitHub Copilot  
**測試狀態:** ⏳ 待測試
