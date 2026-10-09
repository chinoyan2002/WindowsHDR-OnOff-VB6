# REVIEW LOOP 計畫檔（2026-10-10）

目標：重讀手冊+VB6 SKILL，不斷循環代碼審查，直到連續2輪零錯誤才停。

## R0（三線初審）
- w 複本 round-trip 全 match
- 發現 6 項：HDR英文訊息 / DelayHDR預設On=1應2 / 註解5應15 / 只比Name / 無逐台輪詢+45秒過期 / 存檔不Hide + 微缺口(handle洩漏)

## R1 親驗+修碼（已編過）
- 全確認成立，已修：
  - modWindowsHDR.bas:1173 英文→「此顯示器不支援 HDR。」+註解中文化
  - modConfig.bas:62 加 Dim defDelay；69 按isOn給2/1；63註解5→15
  - modShell.bas Fail 加關handle；modCleanVirtual IsElevated/ShellRunSimple Fail 加關handle
  - modCleanVirtual：Name/FriendlyName/Caption三欄查；PnpDeviceGone轉正（WMI雙斜線修好）；逐台等5秒；45→10秒；Dim挪程序頂
  - frmSettings.frm btnSave_Click 加 Me.Hide
- 編譯：殺殭屍19940→第2次成功「成功地製成」；exe v1.0.31=vbp32-1；二進位UTF16有新字串、無舊英文

## R2 複審（修後再審）
- 發現1 minor：Count/Clean Name讀錯被後續Err蓋掉
- 已修：Name失敗即清空 Frag；刪Count dead Err檢查
- 重編成功；exe v1.0.32=vbp33-1；probe exit 1與log問不到一致

## R2-final（本輪，乾淨）
- w3 複本重讀：Config預設/夾值/回退對；HDR分流/重試3次/常數/Verify off要求Supported+非Active合SPEC；Tray三叉/HOLD/寬限期/方向間隔/取消/DelayShell全對；Clean三欄+Pnp 5輪+10秒圈；Settings Hide；Shell隱藏/timeout語意；Log 2000/20次；Lang全覆蓋
- 結論：零實問題 PASS

## R3-final（連續第2輪乾淨）
- 同w3再跑一次全對照：無新問題，編譯已過版號對上，二進位新鮮
- 結論：零實問題 PASS → 達停機條件（連續2輪乾淨），停

## 驗證
- `VB6.EXE /make` 兩次皆「成功地製成」；build.log先刪後看CP950轉碼
- `(Start-Process .\AutoHdrMode.exe '--probe' -Wait -PassThru).ExitCode`=1；log末兩行 v1.0.32 + 問不到1台
- 二進位：UTF16「此顯示器不支援 HDR」有、ASCII舊英文無

## 回滾點
- 各檔 `.bak_r1` 為改前原檔；要回滾 copy回來刪build.log重編即可
- vbp Revision已33，exe 1.0.32；回滾後版號續增不影響

## TODO/盲區
- S_LogDdcErr 死字串留著（零呼叫，無害）
- Log無鎖併發（SPEC允許吞錯，單機可接受）
- Shell逾時回-1與建失敗同值（呼叫方只記log，不分流可接受）
- DDC本機問不到1台（筆電面板/OSD關DDC常見，部署前--probe確認，非bug）

## 啟動補跑（2026-10-10，用戶選A）
- frmTray.PollOnce 初始分支：已通電+g_AutoOn+PowerOn.Enabled，且 HDR_Verify不過（或有Shell可跑）才 StartTransition True；無事不跑
- 編譯成功；exe v1.0.33=vbp34-1；probe exit 1正常
