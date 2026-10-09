# AutoHdrMode.VB6 編碼知識：VB6、Big5 與 Unicode 相關

> 一句話：VB6 內部是 Unicode，輸出入全是 ANSI；Big5 只是繁中機上 ANSI 的長相。

## 1. 三層現實

| 層 | 編碼 | 影響範圍 |
|---|---|---|
| 內部 | UTF-16（BSTR/COM） | WMI 走 COM，全程 |
| 介面 | 系統 ANSI | menu、tooltip、Declare 預設全是 A 版 |
| 檔案 | 系統 ANSI | `Print #`、INI、`CreateProcessA` |

- 繁中機 ANSI＝CP950 Big5：中文選單/log/INI 全通。
- 英文機 ANSI＝CP1252：中文變問號，先天限制，非 bug。

## 2. 源文件規則（血淚）

- .bas/.frm/.vbp/.ini 必須 ANSI（繁中機＝CP950）+CRLF。
- UTF-8 存檔：中文字節被當兩個 ANSI 字（亂碼）（亂碼）；LF-only：VB6 連載入都失敗（實測：無法載入 frmTray.frm）。
- 含中文的改動一律 Python+encoding='cp950'，寫完跑 vb6_crlf.py。
- .frm：`Begin` 行尾空格要留；`0 'False` 註解是 IDE 原生格式，可保留。

## 3. 編譯流程

1. 先刪 build.log（它是累加的，不刪會一直看到舊錯誤）。
2. `vb6.exe /make AutoHdrMode.vbp /out build.log`
3. 有錯看 build.log（旁邊的 .frm.log 不一定存在）。
4. 歷史坑：Declare 宣告必須在所有程序之前；Timer.Interval 上限 32767ms（輪詢最多 30 秒）。

## 4. API 都是 A 版

- 不碰 ...A 傳入傳出自動轉（繁體中文就不用管）：繁中機傳中文路徑/參數 OK；真要傳 Unicode 給 API，改 W 版+手動轉（目前不需要）。

## 5. 預案：UTF-8 log 與英文機中文

- log 改 UTF-8：標準解 ADODB.Stream（不是手刻）。
- StrConv 64/128（vbUnicode/vbFromUnicode）：系統 ANSI↔內部 Unicode 互轉。
- Font.Charset：繁體 Big5＝136；要在 Unicode 控制項或 twinBASIC，VB6 原生無解。

## 6. 考據

- StrConv 官方文件：https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strconv-function
- twinBASIC StrConv 東亞相容問題：https://github.com/twinbasic/twinbasic/issues/1871
- vbForums Unicode-to-ANSI 社群討論：https://www.vbforums.com/showthread.php?597901-convert-Unicode-to-ANSI
- frm 格式考據：https://github.com/jasoncychueh/claude-vb6-plugin/blob/master/skills/vb6-project/references/frm-format.md
- Code Page 對照表：https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
- ADODB.Stream 轉檔標準答案：https://stackoverflow.com/questions/5182102/vb6-vbscript-change-file-encoding-to-ansi
