# AutoHdrMode（VB6 ONLY）

> 終極目標：螢幕斷電自動關閉 HDR，通電自動開啟 HDR。

## 目錄結構

VB6 常駐主線，C# C# 已歸檔，不再維護。

```
AutoHdrMode/
  src/AutoHdrMode.VB6/   主程式全部（.vbp/.bas/.frm/.ini）
  archive/               C# 備份+net472 備份+使用者自理檔備份（HDR1 等自理檔）
  tools/                 ProbeDisplay.ps1（診斷用）
  dist/                  空的（HDR HDR 自理）
```

## 運作原理

1. WMI（`WmiMonitorBasicDisplayParams`）輪詢實體螢幕有無；連續穩定才認。
2. 斷電方向：原生關閉，不動虛擬卡；虛擬卡留給遠端軟體。
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
| PollSeconds | 偵測間隔秒數（VB6 Timer 上限 30） |
| StablePolls | 連續幾次一致才動作 |
| AutoEnabled | 1=自動執行，0=只記錄 |
| PowerOff.DelaySeconds / Shell / HdrAction | 斷電命令 |
| PowerOn.DelaySeconds / Shell / HdrAction | 通電命令＋驗證＋清卡重試 |
| PowerOn.VerifySeconds / CleanOnFail | 驗證秒數與清卡開關 |

## CLI

| 命令 | 結束碼 |
|---|---|
| `--probe` | 只偵測不動作（0 開、1 關、2 未知） |
| `--fire-off` / `--fire-on` | 完整流程 |
| `--clean` | 提權清卡（單 exe 自舉） |
| `--test-admin` | 是否提權診斷 |
| `--install-autostart` / `--remove-autostart` | 開機自啟讀寫 HKCU Run |
| exit code | 成功 0／失敗 1／未設定 3／例外 4（清卡另有 2/6） |

## 第三方模組

`modWindowsHDR.bas` v1.4（作者自有專案），轉碼入庫不改一字+CRLF 以利覆蓋同步，另見 ENCODING.md。

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
