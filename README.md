# AutoHdrMode（VB6 ONLY）

> 終極目標：螢幕斷電自動關閉 HDR，通電自動開啟 HDR。

## 目錄結構

VB6 常駐主線，C# C# 已歸檔，不再維護。

```
AutoHdrMode/
  *.vbp/.bas/.frm/.ini  主程式全部（根目錄，扁平擺放）
  archive/               C# 備份+net472 備份+快照備份
  tools/                 ProbeDisplay.ps1（診斷用）
  usbmmidd_v2/           虛擬顯示卡驅動相關
```

## 運作原理

1. DDC/CI（`GetVCPFeatureAndVCPFeatureReply` D6）輪詢實體螢幕電源；通電後 10 秒寬限期（無 D6=4/5 不認關）；連續穩定才認。
2. 斷電方向：原生關閉（無路徑時走快取目標重試），不動虛擬卡；虛擬卡留給遠端軟體。
3. 通電方向：原生開啟 → 等待驗證秒數 → 失敗就提權清卡 → 重試開啟。
4. 全串行 background 執行（托盤不凍結）。
5. 托盤圖示+右鍵選單+即時記錄窗；無窗體常駐。
6. 中文語系自動（英文原生+Strings 表）。

## 編譯

```bat
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make AutoHdrMode.vbp /out build.log
```

- 先刪 build.log（它是累加的）。
- Declare/Type/Const 必須在所有程序之前。
- 源文件 CP950+CRLF；含中文改動走 Python+encoding='cp950'。
- IDE 開著會互鎖，先關；存檔再關。
- 上游 v1.4 缺宣告本地補（mWindowsVersionSource）。

## 設定（AutoHdrMode.ini，參數說明中文）

| 參數 | 說明 |
|---|---|
| Language | auto=跟隨系統，en=英文，zh=中文 |
| PollSeconds | 偵測間隔秒數（VB6 Timer 上限 30；僅舊 ini 回退用） |
| StablePolls | 連續幾次一致才動作（僅回退用） |
| PowerOff/PowerOn.PollSeconds | 各方向偵測間隔（出貨：通電 1／斷電 3） |
| PowerOff/PowerOn.StablePolls | 各方向連續確認次數（出貨：通電 2／斷電 8） |
| AutoEnabled | 1=自動執行，0=只記錄 |
| DelayHDR / NativeHDR / VerifySeconds | 事件後等幾秒做 HDR、on/off/空、驗證等幾秒 |
| CleanHelper / DelayAfterClean | 驗證失敗清不清卡、清完等幾秒重試 |
| Shell / DelayShell / ShellTimeout | 額外命令、跑前等幾秒、等它幾秒（0=放生） |

## CLI

| 命令 | 結束碼 |
|---|---|
| `--probe` | 只偵測不動作（0 開、1 關、3 待命；log 含明確關／問不到拆分） |
| `--fire-off` / `--fire-on` | 完整流程 |
| `--clean` | 提權清卡（單 exe 自舉） |
| 打錯的參數 | 提示框＋不進 GUI |
| `--install-autostart` / `--remove-autostart` | 開機自啟讀寫 HKCU Run |
| exit code | 成功 0／失敗 1（清卡：提權失敗 5、逾時有殘留 2） |

## 第三方模組

`modWindowsHDR.bas`（作者自有專案；已按正式版計畫重寫結構：Capability/Active 分離、race 重試、Verify，不再是 v1.4 原樣）。

## 叉子炸彈

- 叉子炸彈：`CheckTokenMembership` 回傳假（1309），現改讀 `TokenElevation`（已驗 True/False）。
- 結構防呆：`--clean-elevated` 不回頭旗標（單 exe 自舉）。
- 曾同時存在 1000+ 行程，提權 taskkill 循環排空。

## 授權注意

上游無 LICENSE，自用；散佈前處理。

## 待辦

- 真關開循環實測。
- 常駐 soak（一週看 log）
- 公司機部署（.NET Framework 4.7.2 不相干，VB6 零依賴）
