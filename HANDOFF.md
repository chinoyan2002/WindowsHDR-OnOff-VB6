# 接手檔 HANDOFF（2026-10-10，exe v1.0.33 / vbp Rev34）

## 現在是什麼狀態
- VB6 單 exe 常駐：螢幕斷電關 HDR、通電開 HDR，托盤＋設定窗＋log。
- 剛做完：LOOP 代碼審查（連續 2 輪乾淨）＋ 啟動補跑（開機已通電且 HDR 不對就跑一次通電鏈）。
- 編譯通過：「成功地製成」，exe v1.0.33 = vbp Rev34-1，`--probe` exit 1 與 log 問不到一致。

## 怎麼編譯
1. 先殺殭屍：`tasklist | findstr AutoHdr` → `taskkill /F /PID xxx`（exe 常駐會鎖檔）。
2. 刪 `build.log`（它是累加的，不刪會看到舊錯誤）。
3. `"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make AutoHdrMode.vbp /out build.log`
4. 看 build.log（CP950，需轉碼）：要看到「成功地製成」。
5. 新鮮度：vbp `RevisionVer - 1` 必須等於 exe 版號（`--probe` 的 log 第一行）。

## 編碼紅線（會炸）
- `.bas/.frm/.vbp/.ini` 必須 CP950＋CRLF，UTF-8 存檔會亂碼、LF-only 會載入失敗。
- 含中文改動走 Python（讀 `decode('cp950')` → 改 → `encode('cp950', strict)` → `\r\n` 重組）。
- PowerShell 只下英文指令，不組中文。
- `Declare` 全部在程序之前；模組級變數放檔案上方；Timer 沒有 TabIndex。

## 關鍵行為（SPEC 為準）
- SPEC.md §5.1：第一次穩定是 ON＋自動開＋通電啟用＋HDR 不對（或有 Shell 可跑）→ 補跑 `StartTransition True`，其餘只記初始 log。
- 輪詢方向分開：ON 候選用通電側 StablePolls；穩定 ON 後用斷電側間隔等關，反之亦然。
- 待命 HOLD／通電 10 秒寬限期：凍結，不碰候選計數。
- off 驗證新 API 要 Supported＋非 Active；UNSUPPORTED 算未通過可走清卡。

## 回滾點
- 各源文件有 `.bak_r1`（改前原檔），蓋回去刪 build.log 重編即可。
- 本 commit 即乾淨基線；exe/log 不進版控（見 .gitignore）。

## TODO
- `S_LogDdcErr` 死字串（零呼叫，無害，留著）。
- log 無鎖併發、多 CLI 並寫會互蓋（單機可接受）。
- DDC 在本機問不到 1 台（筆電面板正常，部署前 `--probe`＋開螢幕 OSD 的 DDC/CI 確認）。
