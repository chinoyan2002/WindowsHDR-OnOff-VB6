# VB6 Windows HDR 原生 API 模組

## 正式版修改計畫 v1.0

### 1. 專案目標

將目前的 `modWindowsHDR.bas` 修改為一個：

* 完全不依賴 HDRTray
* 完全不依賴 HDRCmd
* 不依賴第三方 HDR 控制程式
* 直接呼叫 Windows Display Configuration API
* 支援 Windows 11 24H2 / 25H2 新 HDR API
* 保留舊版 Windows 的 Advanced Color API 相容性
* 能正確區分 SDR / WCG / HDR
* 設定後自動驗證結果
* API 錯誤可追蹤
* VB6 可直接使用
* 不因 Display Configuration 在查詢期間變動而造成錯誤
* 不把 Advanced Color 錯誤宣稱成 HDR capability

的正式底層 API 模組。

---

# 2. 修改原則

## 2.1 不使用 HDRTray

不得：

* 呼叫 HDRTray
* 執行 HDRTray.exe
* 呼叫 HDRCmd.exe
* 讀取 HDRTray 設定
* 依賴 HDRTray DLL
* 依賴 HDRTray 的 undocumented helper

所有功能必須直接由 Windows API 完成。

---

## 2.2 Windows API 為唯一控制來源

核心 API：

```text
GetDisplayConfigBufferSizes
QueryDisplayConfig
DisplayConfigGetDeviceInfo
DisplayConfigSetDeviceInfo
```

記憶體操作：

```text
RtlMoveMemory
```

Windows 版本判斷：

```text
RtlGetVersion
```

---

# 3. API 版本策略

## 3.1 Windows 11 24H2 / Build >= 26100

使用新版 API：

```text
DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2 = 15
DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE             = 16
```

此版本必須使用：

```text
DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2
```

判斷：

* HDR capability
* HDR user enabled
* Advanced Color active
* Active Color Mode

並使用：

```text
DISPLAYCONFIG_SET_HDR_STATE
```

進行 HDR ON/OFF。

---

## 3.2 Windows 11 24H2 以下

使用：

```text
DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO = 9
DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE = 10
```

但是必須明確標示：

> 舊 API 的 Advanced Color 不等同於精確 HDR capability。

因此：

```text
advancedColorSupported
```

不得直接存入：

```text
HDRSupported
```

並不得對舊 API 的結果做出：

> 此螢幕一定支援 HDR

的錯誤判斷。

---

# 4. HDR 狀態模型重新定義

正式版必須把以下概念分開。

## 4.1 Capability

```text
HDRSupported
```

只有新版：

```text
GET_ADVANCED_COLOR_INFO_2
```

能明確取得 HDR capability 時，才填入可靠結果。

舊版 API：

```text
HDRCapabilityKnown = False
```

---

## 4.2 Advanced Color

新增：

```text
AdvancedColorSupported
AdvancedColorActive
AdvancedColorLimitedByPolicy
```

避免把 Advanced Color 與 HDR 混為一談。

---

## 4.3 HDR User State

保留：

```text
HDRUserEnabled
```

代表 Windows 使用者設定層級是否開啟 HDR。

---

## 4.4 HDR Active State

保留：

```text
HDRActive
```

但正式版必須以：

```text
ActiveColorMode = DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR
```

作為實際 HDR active 的主要判斷依據。

不能只看：

```text
HDRUserEnabled = True
```

就判斷 HDR 已經實際生效。

---

# 5. 新增 Active Color Mode

正式版增加：

```text
ActiveColorMode
```

定義：

```text
0 = SDR
1 = WCG
2 = HDR
```

常數：

```text
DISPLAYCONFIG_ADVANCED_COLOR_MODE_SDR = 0
DISPLAYCONFIG_ADVANCED_COLOR_MODE_WCG = 1
DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR = 2
```

狀態判斷：

```text
ActiveColorMode = 0
    → SDR

ActiveColorMode = 1
    → WCG

ActiveColorMode = 2
    → HDR
```

---

# 6. HDR_DISPLAY_INFO 結構修改

正式版資料結構至少包含：

```text
Index

AdapterLow
AdapterHigh
TargetId

Name

HDRSupported
HDRCapabilityKnown

HDRUserEnabled
HDRActive

AdvancedColorSupported
AdvancedColorActive
AdvancedColorLimitedByPolicy

WCGSupported
WCGUserEnabled

ActiveColorMode

FriendlyNameFromEdid
FriendlyNameForced

LastError
```

目的：

避免不同 Windows API 的資訊被混在同一個欄位。

---

# 7. QueryDisplayConfig 重試機制

這是正式版必修項目。

目前：

```text
GetDisplayConfigBufferSizes
↓
Allocate
↓
QueryDisplayConfig
```

可能發生：

```text
GetDisplayConfigBufferSizes
```

取得大小後，到：

```text
QueryDisplayConfig
```

執行前，Windows 顯示組態發生變化。

因此必須處理：

```text
ERROR_INSUFFICIENT_BUFFER
```

---

## 7.1 新增內部函數

建立：

```text
HDR_QueryActiveDisplays()
```

流程：

```text
Retry 1
    GetDisplayConfigBufferSizes
    Allocate buffers
    QueryDisplayConfig

    如果 ERROR_INSUFFICIENT_BUFFER
        Retry

Retry 2
    同上

Retry 3
    同上

三次仍失敗
    回傳 False
    保存最後錯誤碼
```

最大重試次數：

```text
3
```

不得無限迴圈。

---

# 8. Display Configuration Buffer

必須正確處理：

```text
pathCount
modeCount
```

並依照：

```text
DISPLAYCONFIG_PATH_INFO
DISPLAYCONFIG_MODE_INFO
```

所需大小配置 Byte Array。

不得假設固定顯示器數量。

---

# 9. Adapter / Target Identification

每一個 active path 必須取得：

```text
adapterId.LowPart
adapterId.HighPart
targetId
```

並以：

```text
adapterId + targetId
```

作為 Display Configuration API 的唯一識別資訊。

不得使用：

```text
Index
```

作為 Windows API 的實際 Target Identifier。

Index 僅供 VB6 UI 使用。

---

# 10. 顯示器名稱

使用：

```text
DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME = 2
```

取得：

```text
monitorFriendlyDeviceName
```

並解析：

```text
friendlyNameFromEdid
friendlyNameForced
```

正式版應保留這兩個 flag。

若 Friendly Name 不存在：

```text
Name = ""
```

不得因名稱取得失敗而導致 HDR API 整體失敗。

---

# 11. HDR 設定策略

## 11.1 Windows 11 24H2+

開啟：

```text
DISPLAYCONFIG_SET_HDR_STATE
```

關閉：

```text
DISPLAYCONFIG_SET_HDR_STATE
```

不得 fallback：

```text
SET_ADVANCED_COLOR_STATE
```

原因：

舊 API 的 Advanced Color 與 HDR 不完全等價。

如果新版 HDR API 回：

```text
ERROR_NOT_SUPPORTED
```

必須直接回報：

```text
HDR_NOT_SUPPORTED
```

而不是偷偷改用舊 API。

---

# 12. 舊版 Windows 設定策略

Windows 11 24H2 以下：

使用：

```text
DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE
```

但是函數文件與狀態回傳必須明確標示：

```text
Advanced Color API
```

而不是聲稱：

```text
純 HDR API
```

因此正式版可以保留：

```text
HDR_Enable
HDR_Disable
```

但內部必須知道：

```text
Windows >= 26100
    → 真正 HDR API

Windows < 26100
    → Advanced Color API compatibility mode
```

---

# 13. 設定後驗證

所有：

```text
Enable
Disable
Toggle
```

都必須：

```text
Set API
↓
重新 Query
↓
重新取得 Display Status
↓
確認實際狀態
```

不能只依：

```text
DisplayConfigSetDeviceInfo = SUCCESS
```

就判定成功。

---

# 14. 驗證規則

## HDR Enable

新版 API：

```text
Set HDR = ON
↓
Query
↓
ActiveColorMode = HDR
```

才算成功。

---

## HDR Disable

新版 API：

```text
Set HDR = OFF
↓
Query
↓
ActiveColorMode <> HDR
```

才算成功。

---

## Toggle

先讀：

```text
HDRActive
```

然後：

```text
HDRActive = True
    → Disable

HDRActive = False
    → Enable
```

設定完成後再次 Query。

---

# 15. 錯誤處理

所有 Windows API 失敗都必須保存：

```text
LastError
```

並提供：

```text
HDR_GetLastError()
HDR_GetLastErrorText()
```

錯誤文字可透過：

```text
FormatMessageA
```

取得。

---

# 16. Per Display Error

每個：

```text
HDR_DISPLAY_INFO
```

保留：

```text
LastError
```

因此多螢幕環境下可以知道：

```text
Display 1 → 成功
Display 2 → ERROR_NOT_SUPPORTED
Display 3 → 成功
```

不能因單一螢幕失敗就覆蓋全部結果。

---

# 17. 多螢幕處理

所有 API 必須支援：

```text
1 Display
2 Displays
3 Displays
...
N Displays
```

不得寫死：

```text
Display 0
Display 1
```

查詢：

```text
HDR_GetDisplays()
```

時應回傳目前所有 active display。

---

# 18. 公開 API

正式版至少保留：

```text
HDR_GetLastError
HDR_GetLastErrorText

HDR_GetWindowsBuild
HDR_IsWindows11_24H2

HDR_GetDisplayCount
HDR_GetDisplays

HDR_GetStatus
HDR_GetStatusText

HDR_Enable
HDR_Disable
HDR_Toggle

HDR_SetDisplayByIndex
```

原有 API 若沒有必要，不應任意刪除。

---

# 19. 狀態常數

保留：

```text
HDR_UNSUPPORTED = 0
HDR_OFF         = 1
HDR_ON          = 2
HDR_ERROR       = 3
```

但：

```text
HDR_UNSUPPORTED
```

必須只在能確認 HDR 不支援時使用。

如果舊版 Windows API 無法判斷 HDR capability：

不得錯誤回傳：

```text
HDR_UNSUPPORTED
```

應增加：

```text
HDRCapabilityKnown = False
```

以區分：

```text
不支援
```

與：

```text
無法由目前 API 判定
```

---

# 20. Windows 版本判斷

使用：

```text
RtlGetVersion
```

取得：

```text
Major
Minor
Build
```

核心判斷：

```text
Build >= 26100
```

視為 Windows 11 24H2+ API 路徑。

不得單純使用：

```text
Windows version string
```

判斷。

---

# 21. 不加入不必要的 undocumented API

正式 HDR 開關功能不得依賴：

```text
InternalRefreshCalibration
```

也不得把：

```text
SetSDRWhiteLevel
```

混入 HDR ON/OFF 核心流程。

這些屬於額外功能，不應影響核心 HDR 控制。

---

# 22. 監控機制

底層 API 模組先保持純 API 模組。

後續若 VMini6 需要常駐 HDR 自動偵測，再建立 UI/監控層。

監控可以使用：

```text
WM_DISPLAYCHANGE
WM_DEVICECHANGE
```

搭配重新：

```text
HDR_GetDisplays()
```

取得最新狀態。

不要在底層 API module 中建立 Timer 或 UI 相依性。

---

# 23. VB6 相容性

必須維持：

```text
VB6
32-bit
```

相容。

不得使用：

```text
VB.NET
VBA-only syntax
64-bit Declare syntax
```

所有 API Declare 必須符合 VB6。

---

# 24. Byte Array + RtlMoveMemory

目前採用：

```text
Byte Array
+
RtlMoveMemory
```

的設計保留。

原因：

VB6 對 Windows API：

* Union
* bit field
* alignment
* pointer
* variable-size structure

處理容易產生錯位。

因此正式版不要為了「看起來比較漂亮」而改成大量複雜 UDT。

---

# 25. 24H2 GET_ADVANCED_COLOR_INFO_2 欄位解析

必須重新逐一確認目前 Byte Offset。

特別確認：

```text
advancedColorSupported
advancedColorEnabled
wideColorGamut
advancedColorForceDisabled
advancedColorLimitedByPolicy
highDynamicRangeSupported
highDynamicRangeUserEnabled
highDynamicRangeEnabled
activeColorMode
```

所有 bit offset 都必須與 Windows SDK 定義一致。

不得依賴猜測的 bit position。

---

# 26. Code Review 驗證

完成修改後必須進行第二輪 Code Review。

檢查：

### API

* Declare 正確
* Entry Point 正確
* Calling Convention 正確
* Data Type 正確

### Structure

* Header size 正確
* Path offset 正確
* Target ID offset 正確
* Target Name size 正確
* Buffer size 正確

### State

* SDR
* WCG
* HDR

三者不可混淆。

### Version

* Build < 26100
* Build >= 26100

兩條 API 路徑都要檢查。

---

# 27. 實機測試矩陣

至少測試：

| 測試                         | 預期         |
| -------------------------- | ---------- |
| 單螢幕 SDR                    | 正確顯示 SDR   |
| 單螢幕 HDR Off                | 正確顯示 OFF   |
| 單螢幕 HDR On                 | 正確顯示 ON    |
| HDR Enable                 | 設定後驗證 ON   |
| HDR Disable                | 設定後驗證 OFF  |
| Toggle                     | 狀態反轉       |
| 雙螢幕                        | 各自取得正確狀態   |
| HDR + SDR 雙螢幕              | 分別辨識       |
| WCG Display                | 不誤判成 HDR   |
| 不支援 HDR Display            | 正確回報       |
| Display Configuration 動態變更 | Query 重試成功 |
| API 失敗                     | 正確保存錯誤碼    |
| 顯示器重新插拔                    | 可重新 Query  |
| Windows 11 24H2            | 使用新 API    |
| Windows 11 25H2            | 使用新 API    |

---

# 28. 特別驗證：WCG ≠ HDR

這是本次修改最重要的測試之一。

必須確認：

```text
ActiveColorMode = WCG
```

時：

```text
HDRActive = False
```

即使：

```text
AdvancedColorActive = True
```

也不能：

```text
HDRActive = True
```

---

# 29. 特別驗證：User Enabled ≠ Active

必須測試：

```text
HDRUserEnabled = True
ActiveColorMode <> HDR
```

情況。

程式不得因此直接顯示：

```text
HDR ON
```

真正 HDR ON 的判斷：

```text
ActiveColorMode = HDR
```

---

# 30. 特別驗證：API Race Condition

測試：

```text
QueryDisplayConfig
```

期間改變：

* 螢幕
* HDR
* Display arrangement
* DP/HDMI connection

確認：

```text
ERROR_INSUFFICIENT_BUFFER
```

時最多重新查詢 3 次。

不能：

```text
死循環
```

也不能直接將第一次錯誤當成最終結果。

---

# 31. 測試程式

建立或更新 VB6 測試 Form。

至少提供：

```text
[讀取狀態]
[HDR ON]
[HDR OFF]
[HDR Toggle]
[列出所有顯示器]
```

並輸出：

```text
Windows Build
API Mode

Display Index
Display Name

HDR Capability
HDR Capability Known

HDR User Enabled
HDR Active

Advanced Color Supported
Advanced Color Active
Advanced Color Limited By Policy

WCG Supported
WCG User Enabled

Active Color Mode

Last Error
```

---

# 32. 測試輸出格式

例如：

```text
Windows Build : 26200
API Mode      : HDR_STATE

Display #0
Name          : LG ULTRAGEAR

HDR Supported : True
HDR Known     : True
HDR User      : True
HDR Active    : True

Advanced Color Supported : True
Advanced Color Active    : True
Limited By Policy        : False

Active Color Mode        : HDR

Last Error : 0
```

WCG：

```text
Active Color Mode : WCG
HDR Active        : False
```

舊版：

```text
API Mode          : ADVANCED_COLOR
HDR Capability    : Unknown
HDR Capability Known : False
```

---

# 33. 不允許的修改

執行修改時不得：

1. 引入 HDRTray
2. 引入 HDRCmd
3. 引入第三方 DLL
4. 修改 Windows Registry 作為 HDR 開關
5. 使用 SendKeys 操作 Windows 設定
6. 模擬 Win+Alt+B
7. 使用 undocumented API 作為核心 HDR ON/OFF
8. 將 WCG 當 HDR
9. 將 Advanced Color Supported 當 HDR Supported
10. 在新 API 失敗時偷偷 fallback 到舊 API
11. 刪除既有公開 API 而沒有必要
12. 為了修改 HDR API 而大幅重構其他無關程式

---

# 34. 修改後交付內容

完成後必須交付：

```text
modWindowsHDR.bas
```

以及：

```text
測試 Form
```

如有需要再提供：

```text
modWindowsHDR_Test.bas
```

但核心 API 與 UI 必須分離。

---

# 35. 最終驗收標準

只有同時符合以下條件才視為完成：

### A. 原生 API

不依賴 HDRTray / HDRCmd。

### B. 24H2+

使用：

```text
GET_ADVANCED_COLOR_INFO_2
SET_HDR_STATE
```

### C. 舊版

使用 Advanced Color compatibility mode，且不冒充 HDR capability。

### D. 狀態正確

能區分：

```text
SDR
WCG
HDR
```

### E. Race Condition

正確處理：

```text
ERROR_INSUFFICIENT_BUFFER
```

並最多重試 3 次。

### F. 設定驗證

Set API 成功後一定重新 Query。

### G. 多螢幕

每個 active display 都能獨立取得及設定狀態。

### H. 錯誤追蹤

保留 Win32 error code 與文字。

### I. VB6

可直接編譯。

### J. Code Review

修改完成後重新檢查：

```text
API
Structure
Offset
Bit field
Version path
Error handling
Multi-monitor
State validation
```

全部通過後才標記：

```text
正式版完成
```

---

# 36. 建議實作順序

```text
Phase 1
重新整理常數與資料結構
        ↓
Phase 2
建立 HDR_QueryActiveDisplays()
        ↓
Phase 3
修正 24H2+ Advanced Color Info 2 解析
        ↓
Phase 4
修正舊版 Advanced Color 語意
        ↓
Phase 5
修正 HDR Enable / Disable
        ↓
Phase 6
加入 Set 後 Verify
        ↓
Phase 7
完善多螢幕及錯誤處理
        ↓
Phase 8
更新 VB6 Test Form
        ↓
Phase 9
實機測試
        ↓
Phase 10
第二輪 Code Review
        ↓
Phase 11
正式版封存
```

---

# 37. 核心成功條件

本次修改不是單純讓：

```text
HDR ON / OFF
```

可以動作。

真正目標是建立一個可靠的：

```text
Windows Display Configuration
        ↓
Display Detection
        ↓
HDR Capability Detection
        ↓
HDR State Detection
        ↓
HDR State Control
        ↓
State Verification
```

完整鏈路。

最終 VMini6 只需要呼叫這個模組，不需要知道 Windows HDR API 的結構、Buffer、bit field、DisplayConfig Target ID 或版本差異。
