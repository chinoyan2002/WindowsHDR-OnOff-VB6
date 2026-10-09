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

## 7. AI 改碼 SOP（CP950 先讀入→全程 UTF-8→寫回轉 CP950）

> 一句話：亂碼是通道問題不是檔案問題。轉換只發生在讀入和寫回兩步，改碼過程全是可讀中文，不猜。

### 7.1 開工（轉工作複本＋驗無損）

```powershell
python -c "import pathlib; src=pathlib.Path('src/AutoHdrMode.VB6'); dst=pathlib.Path('E:/Temp/opencode/w'); dst.mkdir(parents=True,exist_ok=True); [(dst/(f+'.utf8')).write_bytes((src/f).read_bytes().decode('cp950').encode('utf-8')) for f in ['modDDC.bas','frmTray.frm']]"
```

- 先驗 round-trip：`decode('cp950')` 再 `encode('cp950')` 位元組完全一樣才開工，不一樣就停手。
- 工作複本放 `E:/Temp/opencode/w/`，`read`/`edit` 正常改（中文看得見）。
- 源文件禁 `edit` 工具、禁 `shell` 直接印中文。

### 7.2 寫回（一次轉回＋三驗）

- 工作複本 `splitlines` 重組 `\r\n`，`encode('cp950', errors='strict')`；轉不過的字報行號停手，不硬轉。
- 驗：CRLF 數對上、LF-only 歸零、`git diff` 只看到預期的段。
- 含中文改動走 Python＋`encoding='cp950'`；寫完跑 `vb6_crlf.py` 語義檢查（無 LF-only 行）。

### 7.3 收尾（VB6 重編＋版本對照）

1. 關 IDE（互鎖），先刪 `build.log`（它是累加的）。
2. `VB6.EXE /make AutoHdrMode.vbp /out build.log`，看到「成功地製成」才算。
3. 跑 `--probe` 對版本行：exe 版號必須＝vbp `RevisionVer - 1`（AutoIncrement 機制），對不上就停手查。
4. 更動驗證完寫回本手冊（改了什麼、怎麼驗、回滾點）。

### 7.4 踩過的坑（2026-10-06 實錄）

- `read`/`shell`/`grep` 三個通道都只保 ASCII，中文爛掉先轉 `unicode_escape` 或轉碼後再看。
- PowerShell 會吃雙引號：`Declare Lib "kernel32"` 這種行用腳本檔或 `chr(34)`，禁現場組字串。
- py3.7 的 `pathlib.write_text` 沒 `newline` 參數，用 `open(..., newline='\n')`。
- `edit` 會動工作複本的換行，寫回腳本一律 `splitlines` 重組，不准假設。
- `read` 行號和 python 行號可能差一，引用改動處寫函式名最穩。
- `grep` 遇到 CP950 二進位會爛，用 python 讀檔比對代替。
