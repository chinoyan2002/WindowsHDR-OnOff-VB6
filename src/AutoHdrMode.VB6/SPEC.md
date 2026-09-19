# AutoHdrMode 規格書（校驗用）

> 版本對應原始碼：VB6 專案 `AutoHdrMode.vbp`，Startup=`frmSettings`  
> 產出目標：單一 `AutoHdrMode.exe`，與 `AutoHdrMode.ini` 同目錄。  
> 用途：交給其他 AI／開發者實作核對、回歸測試、行為驗收。

---

## 1. 產品目標

監視**實體螢幕電源狀態**（通電／斷電）。狀態穩定變更後，依 INI 可組合地：

1. 使用**程式內建 Windows Display Configuration API** 開啟或關閉 HDR（不依賴 `HDRCmd.exe` 等外部工具）。
2. 可選執行 **Shell** 命令（路徑空白或不存在則略過）。
3. 通電時若 HDR 開啟失敗，可選**提權清除虛擬顯示卡**後重試。

常駐托盤；主介面為設定窗（含即時 log 分頁）。

### 1.1 非目標／邊界

- 不依賴專案外自備 exe（唯一系統工具：`pnputil`，僅清虛擬卡時使用）。
- 斷電當下螢幕常已無 active display path，原生關 HDR 可能只能「盡力＋快取目標」；通電會再強制設定。
- 遠端桌面虛擬螢幕不應被 WMI 計為實體螢幕（查詢來源見 §5）。

---

## 2. 技術約束

| 項目 | 要求 |
|------|------|
| 語言／IDE | Visual Basic 6 |
| 啟動物件 | **窗體** `frmSettings`（不是 `Sub Main`） |
| 編碼 | 所有 `.bas` / `.frm` / `.vbp` / `.ini`：**系統 ANSI（繁中 CP950）+ CRLF** |
| Timer | `Timer.Interval` 上限 32767ms → 輪詢最多 30 秒 |
| 單機例 | `App.PrevInstance`；重複啟動 MsgBox「在桌面右下角已執行!!」後結束 |
| 資訊清單 | 建議 `asInvoker`（開機自啟不被 UAC 擋）；清卡才 `runas` 自己 |

詳見專案內 `ENCODING.md`。

---

## 3. 專案結構

### 3.1 必備原始檔（vbp 清單）

```
Form=frmSettings.frm      ' 主啟動窗＝設定 UI＋記錄分頁
Form=frmTray.frm          ' 隱藏：托盤＋輪詢＋狀態機
Module=modMain.bas        ' AppBootstrap、CLI、開機自啟、FireCli
Module=modLang.bas        ' 中英字串
Module=modConfig.bas      ' INI 讀寫、g_FirstRun、ShellTargetExists
Module=modLog.bas         ' autohdrmode.log 附加寫入
Module=modWMI.bas         ' 實體螢幕計數
Module=modShell.bas       ' ShellRunHidden
Module=modWindowsHDR.bas  ' 原生 HDR API
Module=modCleanVirtual.bas' 清虛擬卡 + IsElevated
```

### 3.2 可忽略／非啟動清單

- `frmLog.frm`：舊即時 log 窗，功能已併入 `frmSettings`「記錄」分頁。
- `frmHDRTest.frm`、`HDR1.vbp`、`modWindowsHDR(1).bas`：測試／備份，不應在主 vbp。
- `SetHDR.ps1`：歷史參考（清卡邏輯已移植進 `modCleanVirtual`）。

### 3.3 執行期檔案

| 檔案 | 說明 |
|------|------|
| `AutoHdrMode.exe` | 主程式 |
| `AutoHdrMode.ini` | 設定（可無＝首次執行） |
| `autohdrmode.log` | ANSI 附加 log（預設 exe 同目錄） |

---

## 4. 啟動與生命週期

### 4.1 正常 GUI 啟動（Startup = frmSettings）

```
frmSettings.Form_Load
  → AppBootstrap()
       ConfigLoad / LangInit / LogInit
       若 Command$ 為 CLI 參數 → 處理後 ExitProcess（見 §8）
       若 App.PrevInstance → MsgBox → ExitProcess
  → Load frmTray          ' 掛托盤、開始輪詢
  → 初始化設定 UI
  → 若 Not g_FirstRun → Me.Hide   ' 非首次：只留托盤
  → 若 g_FirstRun → 保持顯示設定窗
```

### 4.2 首次執行定義

- `ConfigLoad` 時：`AutoHdrMode.ini` 在 `App.Path` **不存在** → `g_FirstRun = True`。
- `ConfigSave` 成功寫入後 → `g_FirstRun = False`。
- 首次在設定窗按「儲存」後：寫 INI，並 **Hide** 主窗。

### 4.3 主窗顯示／關閉規則

| 操作 | 行為 |
|------|------|
| 點標題列 X | `Cancel = True`，`Hide`（不結束行程） |
| 按鈕「關閉」 | `Hide` |
| 托盤結束 | `Unload frmSettings` + `Unload frmTray` + `End` |
| 托盤開啟設定 | `frmSettings.Show`，`WindowState = vbNormal` |

### 4.4 frmTray

- `Visible = False`，不佔工作列。
- `Form_Load`：ScaleMode=Pixel；可先移出螢幕外 Show 再 Hide 以確保 hwnd；`TrayAdd`；啟動 `tmrPoll`。
- 負責：托盤圖示、選單、電源狀態輪詢、HDR/Shell 狀態機。

---

## 5. 實體螢幕偵測（modWMI，DDC/CI D6）

```vb
WmiPhysicalCount() As Long    ' D6=1 通電台數（每次呼叫重新列舉）
WmiStandbyCount() As Long     ' D6=2/3 待命台數（需先呼叫 WmiPhysicalCount）
' 訊號：EnumDisplayMonitors → GetPhysicalMonitorsFromHMONITOR(dxva2)
'       → GetVCPFeatureAndVCPFeatureReply(D6 電源模式)
' 虛擬螢幕（遠端軟體虛擬卡）拿不到 physical monitor，不計數
' WMI 舊版已刪除（純 DDC 取代，不留後備）
```

**D6 電源模式：**

| D6 值 | 意義 | 計數 |
|------|------|------|
| 1 | 開機 | 通電 +1 |
| 2 / 3 | 待命 / 暫停 | 待命 +1 |
| 4 / 5 | 關機 / 斷電 | 不計（關） |
| 查詢失敗（不支援 DDC/CI、拿不到代碼） | — | 不計（字面關） |

**判定（PollOnce 三叉）：**

| 條件 | 意義 |
|------|------|
| 通電數 > 0 | 螢幕「開」 |
| 通電數 = 0 且待命數 > 0 | **HOLD**：不排程、不轉態，只寫 log（進出各記一次） |
| 兩者皆 0 | 螢幕「關」 |

注意：DDC 問不到的螢幕（筆電內建面板、OSD 關 DDC/CI 者）永遠讀 0，
部署前以 `--probe` 確認；螢幕 OSD 的 DDC/CI 須開啟。

`--probe` 結束碼：0 = 通電中，1 = 全關，3 = 待命 HOLD；log 記 `on=X standby=Y`。

### 5.1 輪詢

- 每 `PollSeconds`（1~30）查一次。
- 候選狀態連續相同 `StablePolls` 次才視為穩定。
- 第一次穩定只寫「初始狀態」log，**不觸發**通電／斷電動作。
- HOLD 不碰候選值與穩定計數，離開 HOLD 後正常判定。

---

## 6. 自動動作狀態機（frmTray）

全域還受 `g_AutoOn`（INI `Main.AutoEnabled`／托盤「自動切換」）控制；為 False 時只更新狀態與 tip，不排程。

事件認定後：

1. 寫 log「事件」。
2. 若該方向 `Enabled=0` → log「動作已停用」並結束。
3. 否則進入階段機（`tmrWait`，約 500ms 節拍）。

### 6.1 階段常數

| 常數 | 含義 |
|------|------|
| PH_IDLE | 空閒 |
| PH_WAIT_HDR | 等待 `DelayHDR` 後做第一次 HDR |
| PH_WAIT_VERIFY | 通電：HDR ON 後等待 `VerifySeconds` 再查狀態 |
| PH_WAIT_AFTER_CLEAN | 清卡後等待 `DelayAfterClean` 再重試 HDR ON |

排程期間若 `m_lastSeen` 與目標相反 → 取消並 log。

### 6.2 通電／斷電同一流程（結構對稱）

`PowerOn` 與 `PowerOff` **可選項目與階段完全相同**，僅預設值與「成功判定」不同。

```
StartTransition(wantOn)
  等 DelayHDR 秒
  若 NativeHDR = "on" 或 "off"：
      執行 HDR_Enable 或 HDR_Disable（依 NativeHDR）
      記錄 ok + HDR_GetStatus(0)
      若 VerifySeconds > 0：
          等 VerifySeconds
          驗證是否達標（見下表）
          達標 → Shell 段
          未達標且 CleanHelper=1：
              RunCleanElevated()
              等 DelayAfterClean
              再執行同一 NativeHDR 動作一次
          → Shell 段
      否則直接 Shell 段
  若 NativeHDR 空白：跳過 HDR → Shell 段

Shell 段：
  空白或不存在 → 略過；否則隱藏執行
  log「通電/斷電流程結束」
```

| NativeHDR | 驗證成功條件 |
|-----------|----------------|
| on | `HDR_GetStatus = HDR_ON` |
| off | 僅 `HDR_OFF` 算通過；`UNSUPPORTED` 視為未通過，可走清卡重試 |

### 6.2.1 DelayShell（Shell 執行前等待）

驗收通過／失敗／清卡重試三條路最後都會進 `DoShellStep`；進去先看 `DelayShell`：

- `> 0`：記 log，排 `PH_WAIT_SHELL`，到期才真正跑 Shell（碼表從進入此步起算）。
- `= 0`：直跑 Shell。

CLI `FireCli`／`FireTransitionCore` 同語意（blocking 等）。

### 6.3 （已併入 6.2）

### 6.4 HDR 與 Shell 可組合性

| NativeHDR | Shell 有效 | 結果 |
|-----------|------------|------|
| on/off | 是 | 先 HDR 段，再 Shell |
| on/off | 否 | 只 HDR |
| 空白 | 是 | 只 Shell |
| 空白 | 否 | 只記事件後無動作（Enabled=1 時仍走完略過 log） |
| Enabled=0 | * | 不進入狀態機 |

### 6.5 ShellTargetExists 規則（modConfig）

- Trim 後空字串 → False。
- 取第一個 token（支援 `"C:\path\a.exe" args` 引號形式）。
- 無磁碟機與開頭 `\` → 視為相對 `g_WorkDir`。
- `Dir$(path) <> ""` → True。

### 6.6 手動觸發

托盤「立即執行通電／斷電」：`StartTransition True/False`（**不**檢查該方向 Enabled，仍走同一狀態機）。  
CLI `--fire-on/--fire-off`：同步簡化路徑（HDR + Shell，清卡狀態機以 GUI 為準；FireCli 實作應與規格對齊核對）。

---

## 7. 原生 HDR（modWindowsHDR）

公開 API（驗收時可對照）：

- `HDR_Enable() / HDR_Disable() / HDR_GetStatus(index) As HDR_STATUS`
- `HDR_STATUS`：`UNSUPPORTED=0, OFF=1, ON=2, ERROR=3, STATUS_UNKNOWN=4`

行為要點：

- 使用 Display Configuration API；新版優先 `SET_HDR_STATE`，舊版 Advanced Color。
- `HDR_GetDisplays` 成功時更新**快取** adapter LUID + targetId。
- `HDR_Disable`：一般路徑失敗且有快取時，對快取目標直接 `DisplayConfigSetDeviceInfo` 關 HDR。
- 所有狀態／錯誤訊息為繁中（模組內部）。

---

## 8. 清虛擬卡（modCleanVirtual + 提權）

### 8.1 何時觸發

任一方（通電或斷電）在驗證失敗且該方向 `CleanHelper=1` 時：由 `frmTray.RunCleanElevated` 以 `ShellExecuteEx` + `runas` 啟動：

```
AutoHdrMode.exe --clean
```

並等待子行程結束（最多約 3 秒，DoEvents 切片）。

### 8.2 `--clean` 行為（modMain.CleanCli）

1. `Not IsElevated` → `ShellExecute runas` 自己 `--clean` 後結束（父行程若已是 elevate 等待則等子行程）。
2. 已 elevate → `CleanVirtualGpus`：
   - WMI `Win32_PnPEntity`，`PNPClass='Display'`
   - Friendly/Name **不含** AMD / NVIDIA / Intel → `pnputil /remove-device`
   - 優先 `%WINDIR%\sysnative\pnputil.exe` 否則 System32
   - 最多約 10 秒輪詢直到無虛擬卡或逾時

`IsElevated`：`OpenProcessToken` + `TokenElevation`。

---

### 8.3 逐台輪詢（取代傻等）

拔卡後不再 `Sleep 1000` 硬等：`PnpDeviceGone(instId)` 先查一次，存在才每秒一輪、最多 5 輪；5 秒還在記逾時走下一台。收尾確認圈 10 秒（`left = 10`，2 秒一驗）。

## 9. 命令列介面

經 `AppBootstrap`（或兼容 `Sub Main`）處理，處理完 `ExitProcess`：

| 參數 | 行為 | 結束碼（概要） |
|------|------|----------------|
| （無） | GUI | — |
| `--probe` | 查 DDC，log | 0 開 / 1 關 / 3 待命 |
| `--fire-on` | 同步通電向 HDR+Shell | 0/1 |
| `--fire-off` | 同步斷電向 HDR+Shell | 0/1 |
| `--clean` | 清虛擬卡（可自提權） | 0 成功等 |
| `--install-autostart` | HKCU Run 寫入 | 0/1 |
| `--remove-autostart` | HKCU Run 刪除 | 0/1 |

開機自啟鍵：

```
HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AutoHdrMode
= "完整路徑\AutoHdrMode.exe"
```

---

## 10. INI 規格（`AutoHdrMode.ini`）

路徑：`App.Path & "\AutoHdrMode.ini"`。  
註解行以 `;` 開頭。讀取用 `GetPrivateProfileStringA`；儲存用 `WritePrivateProfileStringA`。

### 10.1 [Main]

| 鍵 | 型別 | 預設 | 說明 |
|----|------|------|------|
| Language | string | auto | auto / zh / en |
| PollSeconds | int | 4（範例 ini 可為 1） | 1～30 |
| StablePolls | int | 3（範例可 1） | ≥1；介面顯示「穩定確認次數」 |
| AutoEnabled | 0/1 | 1 | 自動總開關 |
| LogDir | string | 空 | 空＝App.Path |

### 10.2 [PowerOff] / [PowerOn]（**結構完全相同**）

| 鍵 | 預設 Off | 預設 On | 說明 |
|----|----------|---------|------|
| Enabled | 1 | 1 | 0＝只記事件不動作 |
| DelayHDR | 1 | 2 | 事件後～第一次 HDR 前（秒），0～120 |
| NativeHDR | off | on | on / off / 空＝不做 HDR |
| VerifySeconds | 2 | 2 | HDR 後查狀態前等待，0～60；0＝不驗證 |
| CleanHelper | 1 | 1 | 驗證失敗時清虛擬卡後重試（兩邊皆可） |
| DelayAfterClean | 5 | 15 |
| DelayShell | 0 | 0 | Shell 前等待，全路徑皆等，0~120
| ShellTimeout | 120 | 120 | Shell 等待秒數，0~3600；0=啟動即走 清卡後～重試 HDR，0～120 |
| Shell | 空 | 空 | 見 §6.5 |

兼容：若仍存在舊鍵 `DelaySeconds`，`DelayHDR` 可回退讀取該值。

---

## 11. 設定窗（frmSettings）

### 11.1 分頁（按鈕切換 PictureBox，無 OCX 依賴；介面**繁體中文**）

1. **一般**
   - 輪詢間隔（秒）＋說明
   - 穩定確認次數＋說明
   - 啟用自動切換
   - 開機時自動執行本程式
2. **斷電**（與通電欄位一致）
   - 啟用斷電自動動作
   - HDR 動作前等待（秒）
   - 原生 HDR 動作：下拉「（不動作）／關閉 HDR／開啟 HDR」
   - 驗證等待（秒）
   - 驗證失敗時清除虛擬顯示卡後重試
   - 清卡後再試前等待（秒）
   - 額外 Shell 命令＋說明
3. **通電**（同上結構；HDR 下拉順序為「不動作／開啟／關閉」）
4. **記錄**：即時 log，可「暫停自動更新」

### 11.2 儲存

- UI → 全域 → `ConfigSave`
- chkAutostart → `AutostartSet True/False`
- 套用 `frmTray.tmrPoll.Interval`、`mnuAuto.Checked`
- Hide 主窗

### 11.3 開機自啟核對

- 讀：`AutostartInstalled()` 讀上述 HKCU 鍵是否存在且非空。
- 寫：與 CLI install/remove 相同。

---

## 12. 托盤（frmTray）

### 12.1 圖示回呼

- `Shell_NotifyIcon` + `uCallbackMessage = WM_USER+1`
- `Form_MouseMove`：ScaleMode=Pixel 時 `msg = X`；若 `msg>1024` 再按 Twips 換算
- 必須定義：

```vb
WM_LBUTTONDOWN=&H201, WM_LBUTTONUP=&H202, WM_LBUTTONDBLCLK=&H203
WM_RBUTTONDOWN=&H204, WM_RBUTTONUP=&H205
```

- 右鍵／左鍵 up 或右鍵 down → `PopupMenu mnuTray`（不傳座標）
- 事前 `SetForegroundWindow`，事後 `PostMessage WM_NULL`
- 雙擊 → 開啟設定窗

### 12.2 選單

| 項 | 行為 |
|----|------|
| 自動切換 | 切換 g_AutoOn + Checked |
| 立即執行斷電命令 | StartTransition False |
| 立即執行通電命令 | StartTransition True |
| 設定… | OpenSettings |
| 開啟紀錄資料夾 | explorer /select, log 檔 |
| 結束程式 | 卸載並 End |

Tooltip：實體螢幕開／關。

---

## 13. Log（modLog）

- 檔案：`LogDir` 或 `App.Path` + `\autohdrmode.log`
- 格式：`yyyy-mm-dd hh:nn:ss ` + 訊息 + CRLF
- ANSI `Print #`；失敗吞掉不中斷主流程

關鍵事件（校驗時應能在 log 見到對應句，中或英依 Language）：

- 啟動、初始狀態、事件、方向 Disabled、HDR 排程
- 原生 HDR 嘗試／結果、驗證等待／通過／失敗
- 清卡、提權、重試、Shell 略過／完成、流程結束
- WMI 錯誤字串、TrayAdd failed（若有）

---

## 14. 驗收測試清單（建議）

### 14.1 啟動

1. 刪除 INI → 啟動應顯示設定窗。
2. 儲存後 INI 出現，窗隱藏，托盤在。
3. 再啟動 → 不顯示主窗，僅托盤。
4. 第二次啟動同目錄 → MsgBox 已執行，不出現第二托盤。

### 14.2 托盤

1. 右鍵／左鍵出現選單。
2. 雙擊開啟設定。
3. 結束後行程與圖示消失。

### 14.3 通電狀態機

1. `PowerOn.Enabled=1`, `NativeHDR=on`, `VerifySeconds=2`, `CleanHelper=0`, `Shell=`  
   → 通電後 log：Delay → HDR ON → 驗證 → Shell 略過 → 結束。
2. `NativeHDR=` 空，`Shell=存在的cmd` → 只跑 Shell。
3. `Shell=不存在的.exe` → log 略過，不崩潰。
4. `Enabled=0` → 有事件 log，無 HDR/Shell。

### 14.4 斷電

1. `NativeHDR=off` → 嘗試 Disable；允許 UNSUPPORTED + 快取嘗試 log。
2. 斷電不得跳 UAC 清卡。

### 14.5 清卡

1. 人為讓 HDR ON 後驗證仍失敗 + CleanHelper=1 → 應出現 UAC（若非 admin）與清卡 log。
2. 僅 AMD/NVIDIA/Intel 時應「無虛擬卡」。

### 14.6 開機自啟

1. 設定勾選儲存 → Run 鍵存在。
2. 取消勾選儲存 → 鍵移除。

### 14.7 CLI

`--probe` / `--fire-off` / `--fire-on` / `--clean` 結束碼與 log 合理。

---

## 15. 已知限制（校驗時勿當回歸 bug）

1. 螢幕已斷電後，DisplayConfig 可能無 active path；關 HDR 無法保證系統 UI 切換與通電前一致，靠通電強制 ON/OFF 策略補回。
2. WMI 失敗被當作斷電是設計行為。
3. VB6 托盤在部分高 DPI／遠端桌面環境仍可能異常；主窗作 Startup 並 Hide 是為降低此問題。
4. 源碼必須 CP950+CRLF；UTF-8 會導致繁中機亂碼或無法載入 frm。

---

## 16. 給實作／校驗 AI 的注意事項

1. 改碼後維持 **CP950 + CRLF** 寫回 `.bas/.frm/.vbp/.ini`。
2. 勿把 `Startup` 改回純 `Sub Main` 而不保留「主窗 Startup + Hide」行為。
3. 所有在 `Select Case` / 條件用到的 `WM_*` 必須在同一窗體有 `Private Const`。
4. HDR 與 Shell 必須保持可獨立組合；Shell 必須存在性檢查。
5. 清卡僅通電驗證失敗路徑；manifest 保持 asInvoker。
6. 規格與程式不一致時，以**本文件 §6～§12 行為**為驗收準則，並開 issue 對照原始碼修正。

---

## 17. 文件修訂

| 日期 | 說明 |
|------|------|
| 2026-09-19 | 初版：對應 Startup=frmSettings、DelayHDR/Verify/Clean/Shell 組合、設定分頁、托盤與首次執行行為 |
| 2026-09-19 | 通電／斷電流程與 INI／設定窗選項完全對稱；設定窗全面中文化 |
| 2026-09-19 | Code review 修正：off 驗證僅 HDR_OFF；CLI 含 DelayHDR；Icon.Handle；記錄 200/2000 提示 |

