' 本模組：語系表（對照 C# 版 Strings）
' g_Chinese=True 走中文，False 走英文
Option Explicit

' 全域旗標：目前是否使用中文
Public g_Chinese As Boolean

' 讀系統介面語系（&H404 台灣、&H804 大陸、&HC04 香港）
Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer

' 依設定決定語系：zh 開頭=中文，en=英文，auto=跟系統走
Public Sub LangInit(ByVal mode As String)
    Dim m As String
    m = LCase$(Trim$(mode))
    Select Case m
        Case "zh", "zh-tw", "zh-cn", "zh-hk", "chs", "cht", "chinese"
            g_Chinese = True
        Case "en", "english"
            g_Chinese = False
        Case Else
            g_Chinese = IsChineseWindows()
    End Select
End Sub


' 主語系是否中文（&H4=LANG_CHINESE）
Private Function IsChineseWindows() As Boolean
    IsChineseWindows = ((GetUserDefaultUILanguage() And &H3FF) = &H4)
End Function

' ---- Tray menu ----
' ---- 以下：托盤選單文字 ----
Public Function S_MenuAuto() As String
    If g_Chinese Then S_MenuAuto = "自動切換" Else S_MenuAuto = "Auto switch"
End Function
Public Function S_MenuFireOff() As String
    If g_Chinese Then S_MenuFireOff = "立即執行斷電命令" Else S_MenuFireOff = "Run power-off command now"
End Function
Public Function S_MenuFireOn() As String
    If g_Chinese Then S_MenuFireOn = "立即執行通電命令" Else S_MenuFireOn = "Run power-on command now"
End Function
Public Function S_MenuOpenLog() As String
    If g_Chinese Then S_MenuOpenLog = "開啟記錄資料夾" Else S_MenuOpenLog = "Open log folder"
End Function
Public Function S_MenuExit() As String
    If g_Chinese Then S_MenuExit = "結束程式" Else S_MenuExit = "Exit"
End Function

' ---- Tooltip ----
' ---- 以下：圖示提示文字 ----
Public Function S_TipStarting() As String
    If g_Chinese Then S_TipStarting = "AutoHdrMode 啟動中" Else S_TipStarting = "AutoHdrMode starting"
End Function
Public Function S_TipUnknown() As String
    If g_Chinese Then S_TipUnknown = "AutoHdrMode：螢幕狀態未知" Else S_TipUnknown = "AutoHdrMode: monitor state unknown"
End Function
Public Function S_TipState(ByVal isOn As Boolean) As String
    If g_Chinese Then
        If isOn Then S_TipState = "AutoHdrMode：實體螢幕開" Else S_TipState = "AutoHdrMode：實體螢幕關"
    Else
        If isOn Then S_TipState = "AutoHdrMode: physical ON" Else S_TipState = "AutoHdrMode: physical OFF"
    End If
End Function

' ---- Log ----
' ---- 以下：紀錄檔文字 ----
Public Function S_LogStart(ByVal poll As Long, ByVal stable As Long, ByVal autoOn As Boolean) As String
    If g_Chinese Then
        S_LogStart = "啟動 輪詢" & poll & "秒 去拖" & stable & "次 自動切換=" & IIf(autoOn, "開", "關")
    Else
        S_LogStart = "START poll=" & poll & "s stable=" & stable & " auto=" & autoOn
    End If
End Function
Public Function S_LogAutoToggle(ByVal autoOn As Boolean) As String
    If g_Chinese Then S_LogAutoToggle = "自動切換 " & IIf(autoOn, "開", "關") Else S_LogAutoToggle = "AUTO " & IIf(autoOn, "on", "off")
End Function
Public Function S_LogManualFireOff(ByVal rc As Long) As String
    If g_Chinese Then S_LogManualFireOff = "手動執行斷電命令 rc=" & rc Else S_LogManualFireOff = "MANUAL fire power-off rc=" & rc
End Function
Public Function S_LogManualFireOn(ByVal rc As Long) As String
    If g_Chinese Then S_LogManualFireOn = "手動執行通電命令 rc=" & rc Else S_LogManualFireOn = "MANUAL fire power-on rc=" & rc
End Function
Public Function S_LogInit(ByVal isOn As Boolean, ByVal wmi As Long) As String
    If g_Chinese Then
        S_LogInit = "初始狀態 實體螢幕=" & IIf(isOn, "開", "關") & " WMI筆數=" & wmi
    Else
        S_LogInit = "INIT physical=" & IIf(isOn, "ON", "OFF") & " wmi=" & wmi
    End If
End Function
Public Function S_LogEvent(ByVal isOn As Boolean, ByVal wmi As Long) As String
    If g_Chinese Then
        S_LogEvent = "事件 實體螢幕=" & IIf(isOn, "開", "關") & " WMI筆數=" & wmi
    Else
        S_LogEvent = "EVENT physical=" & IIf(isOn, "ON", "OFF") & " wmi=" & wmi
    End If
End Function
Public Function S_LogPlanOff(ByVal d As Long) As String
    If g_Chinese Then S_LogPlanOff = "斷電命令排程" & d & "秒後執行" Else S_LogPlanOff = "AUTO power-off shell in " & d & "s"
End Function
Public Function S_LogPlanOn(ByVal d As Long) As String
    If g_Chinese Then S_LogPlanOn = "通電命令排程" & d & "秒後執行" Else S_LogPlanOn = "AUTO power-on shell in " & d & "s"
End Function
Public Function S_LogFiredOff(ByVal rc As Long) As String
    If g_Chinese Then S_LogFiredOff = "斷電命令完成 rc=" & rc Else S_LogFiredOff = "AUTO power-off done rc=" & rc
End Function
Public Function S_LogFiredOn(ByVal rc As Long) As String
    If g_Chinese Then S_LogFiredOn = "通電命令完成 rc=" & rc Else S_LogFiredOn = "AUTO power-on done rc=" & rc
End Function
Public Function S_LogSkipOff() As String
    If g_Chinese Then S_LogSkipOff = "排程期間狀態已變回，跳過斷電命令" Else S_LogSkipOff = "AUTO state changed, skip power-off"
End Function
Public Function S_LogSkipOn() As String
    If g_Chinese Then S_LogSkipOn = "排程期間狀態已變回，跳過通電命令" Else S_LogSkipOn = "AUTO state changed, skip power-on"
End Function
Public Function S_LogEmptyOff() As String
    If g_Chinese Then S_LogEmptyOff = "斷電命令未設定，跳過" Else S_LogEmptyOff = "AUTO power-off shell not set, skip"
End Function
Public Function S_LogEmptyOn() As String
    If g_Chinese Then S_LogEmptyOn = "通電命令未設定，跳過" Else S_LogEmptyOn = "AUTO power-on shell not set, skip"
End Function
Public Function S_LogAutoErr(ByVal m As String) As String
    If g_Chinese Then S_LogAutoErr = "自動：例外 " & m Else S_LogAutoErr = "AUTO-ERR " & m
End Function
Public Function S_LogWmiErr(ByVal m As String) As String
    If g_Chinese Then S_LogWmiErr = "WMI查詢失敗 " & m Else S_LogWmiErr = "WMI-ERR " & m
End Function
Public Function S_LogPollErr(ByVal m As String) As String
    If g_Chinese Then S_LogPollErr = "輪詢例外 " & m Else S_LogPollErr = "POLL-ERR " & m
End Function
Public Function S_LogOpenLogErr(ByVal m As String) As String
    If g_Chinese Then S_LogOpenLogErr = "開啟記錄資料夾失敗 " & m Else S_LogOpenLogErr = "OPENLOG-ERR " & m
End Function
Public Function S_LogProbe(ByVal wmi As Long) As String
    If g_Chinese Then S_LogProbe = "探測 WMI筆數=" & wmi Else S_LogProbe = "PROBE wmi=" & wmi
End Function
Public Function S_LogAsIn() As String
    If g_Chinese Then S_LogAsIn = "開機自動啟動 已安裝" Else S_LogAsIn = "AUTOSTART installed"
End Function
Public Function S_LogAsRm() As String
    If g_Chinese Then S_LogAsRm = "開機自動啟動 已移除" Else S_LogAsRm = "AUTOSTART removed"
End Function
Public Function S_LogAsErr(ByVal m As String) As String
    If g_Chinese Then S_LogAsErr = "開機自動啟動失敗 " & m Else S_LogAsErr = "AUTOSTART-ERR " & m
End Function
Public Function S_LogExit() As String
    If g_Chinese Then S_LogExit = "結束" Else S_LogExit = "EXIT"
End Function
Public Function S_LogHdrOn(ByVal rc As Long) As String
    If g_Chinese Then S_LogHdrOn = "原生開啟HDR完成 rc=" & rc Else S_LogHdrOn = "NATIVE HDR on done rc=" & rc
End Function
Public Function S_LogHdrOff(ByVal rc As Long) As String
    If g_Chinese Then S_LogHdrOff = "原生關閉HDR完成 rc=" & rc Else S_LogHdrOff = "NATIVE HDR off done rc=" & rc
End Function
Public Function S_LogVerifyWait(ByVal s As Long) As String
    If g_Chinese Then S_LogVerifyWait = "等待驗證" & s & "秒" Else S_LogVerifyWait = "waiting verify " & s & "s"
End Function
Public Function S_LogVerified() As String
    If g_Chinese Then S_LogVerified = "驗證通過，HDR 已開啟" Else S_LogVerified = "verified, HDR is on"
End Function
Public Function S_LogVerifyFail() As String
    If g_Chinese Then S_LogVerifyFail = "驗證未通過，準備清卡" Else S_LogVerifyFail = "verify failed, cleaning"
End Function
Public Function S_LogRetry() As String
    If g_Chinese Then S_LogRetry = "重試開啟HDR" Else S_LogRetry = "retry HDR enable"
End Function
Public Function S_LogCleanStart() As String
    If g_Chinese Then S_LogCleanStart = "開始清除虛擬顯示卡（提權執行）" Else S_LogCleanStart = "cleaning (elevated)"
End Function
Public Function S_LogCleanDone(ByVal rc As Long) As String
    If g_Chinese Then S_LogCleanDone = "清除虛擬顯示卡完成 rc=" & rc Else S_LogCleanDone = "clean done rc=" & rc
End Function
Public Function S_LogCleanNone() As String
    If g_Chinese Then S_LogCleanNone = "沒有虛擬顯示卡" Else S_LogCleanNone = "no virtual display"
End Function
Public Function S_LogCleanRemoving(ByVal name As String) As String
    If g_Chinese Then S_LogCleanRemoving = "正在移除：" & name Else S_LogCleanRemoving = "removing: " & name
End Function
Public Function S_LogCleanOk() As String
    If g_Chinese Then S_LogCleanOk = "虛擬顯示卡已清空" Else S_LogCleanOk = "virtual displays cleared"
End Function
Public Function S_LogCleanLeft() As String
    If g_Chinese Then S_LogCleanLeft = "逾時仍有殘留" Else S_LogCleanLeft = "still present after timeout"
End Function
Public Function S_LogCleanErr(ByVal m As String) As String
    If g_Chinese Then S_LogCleanErr = "清除虛擬顯示卡執行錯誤 " & m Else S_LogCleanErr = "clean error " & m
End Function
Public Function S_MenuLogWin() As String
    If g_Chinese Then S_MenuLogWin = "查看即時記錄" Else S_MenuLogWin = "View live log"
End Function
Public Function S_LogWinCap() As String
    If g_Chinese Then S_LogWinCap = "AutoHdrMode 即時記錄" Else S_LogWinCap = "AutoHdrMode live log"
End Function
Public Function S_PauseCap() As String
    If g_Chinese Then S_PauseCap = "暫停更新" Else S_PauseCap = "Pause"
End Function
Public Function S_CloseCap() As String
    If g_Chinese Then S_CloseCap = "關閉" Else S_CloseCap = "Close"
End Function
Public Function S_MenuSettings() As String
    If g_Chinese Then S_MenuSettings = "設定…" Else S_MenuSettings = "Settings..."
End Function
Public Function S_MenuHdrOn() As String
    If g_Chinese Then S_MenuHdrOn = "開啟HDR" Else S_MenuHdrOn = "Turn HDR on"
End Function
Public Function S_MenuHdrOff() As String
    If g_Chinese Then S_MenuHdrOff = "關閉HDR" Else S_MenuHdrOff = "Turn HDR off"
End Function
Public Function S_SetCap() As String
    If g_Chinese Then S_SetCap = "AutoHdrMode 設定" Else S_SetCap = "AutoHdrMode Settings"
End Function
Public Function S_TabGeneral() As String
    If g_Chinese Then S_TabGeneral = "一般" Else S_TabGeneral = "General"
End Function
Public Function S_TabPowerOff() As String
    If g_Chinese Then S_TabPowerOff = "斷電" Else S_TabPowerOff = "On power off"
End Function
Public Function S_TabPowerOn() As String
    If g_Chinese Then S_TabPowerOn = "通電" Else S_TabPowerOn = "On power on"
End Function
Public Function S_TabLog() As String
    If g_Chinese Then S_TabLog = "記錄" Else S_TabLog = "Log"
End Function
Public Function S_LblPoll() As String
    If g_Chinese Then S_LblPoll = "輪詢間隔（秒）" Else S_LblPoll = "Poll interval (s)"
End Function
Public Function S_LblStable() As String
    If g_Chinese Then S_LblStable = "去拖次數" Else S_LblStable = "Stable polls"
End Function
Public Function S_ChkAuto() As String
    If g_Chinese Then S_ChkAuto = "啟用自動切換" Else S_ChkAuto = "Enable auto switch"
End Function
Public Function S_ChkAutostart() As String
    If g_Chinese Then S_ChkAutostart = "開機時自動執行本程式" Else S_ChkAutostart = "Run this program at startup"
End Function
Public Function S_ChkBalloon() As String
    If g_Chinese Then S_ChkBalloon = "狀態改變時氣球提示" Else S_ChkBalloon = "Balloon on state change"
End Function
Public Function S_LblEnabled() As String
    If g_Chinese Then S_LblEnabled = "啟用此方向自動動作" Else S_LblEnabled = "Enable this direction"
End Function
Public Function S_LblDelayHdr() As String
    If g_Chinese Then S_LblDelayHdr = "HDR 動作前等待（秒）" Else S_LblDelayHdr = "Wait before HDR (s)"
End Function
Public Function S_LblNative() As String
    If g_Chinese Then S_LblNative = "原生 HDR 動作" Else S_LblNative = "Native HDR action"
End Function
Public Function S_OptHdrNone() As String
    If g_Chinese Then S_OptHdrNone = "（不動作）" Else S_OptHdrNone = "(none)"
End Function
Public Function S_OptHdrOn() As String
    If g_Chinese Then S_OptHdrOn = "開啟 HDR" Else S_OptHdrOn = "Turn HDR on"
End Function
Public Function S_OptHdrOff() As String
    If g_Chinese Then S_OptHdrOff = "關閉 HDR" Else S_OptHdrOff = "Turn HDR off"
End Function
Public Function S_LblVerify() As String
    If g_Chinese Then S_LblVerify = "驗證等待（秒）" Else S_LblVerify = "Verify wait (s)"
End Function
Public Function S_ChkClean() As String
    If g_Chinese Then S_ChkClean = "驗證失敗時清除虛擬顯示卡後重試" Else S_ChkClean = "Clean virtual displays and retry on verify fail"
End Function
Public Function S_LblAfterClean() As String
    If g_Chinese Then S_LblAfterClean = "清卡後再試前等待（秒）" Else S_LblAfterClean = "Wait after clean (s)"
End Function
Public Function S_LblShell() As String
    If g_Chinese Then S_LblShell = "額外 Shell 命令" Else S_LblShell = "Extra shell command"
End Function
Public Function S_BtnSave() As String
    If g_Chinese Then S_BtnSave = "儲存" Else S_BtnSave = "Save"
End Function
Public Function S_BtnClose() As String
    If g_Chinese Then S_BtnClose = "關閉" Else S_BtnClose = "Close"
End Function
Public Function S_LblStatus() As String
    If g_Chinese Then S_LblStatus = "狀態" Else S_LblStatus = "Status"
End Function
Public Function S_LblPhy() As String
    If g_Chinese Then S_LblPhy = "實體螢幕" Else S_LblPhy = "Physical monitor"
End Function
Public Function S_LblHdr2() As String
    If g_Chinese Then S_LblHdr2 = "HDR" Else S_LblHdr2 = "HDR"
End Function
Public Function S_LblLastEvt() As String
    If g_Chinese Then S_LblLastEvt = "最後事件" Else S_LblLastEvt = "Last event"
End Function
Public Function S_WordOn() As String
    If g_Chinese Then S_WordOn = "開" Else S_WordOn = "On"
End Function
Public Function S_WordOff() As String
    If g_Chinese Then S_WordOff = "關" Else S_WordOff = "Off"
End Function
Public Function S_LogSaved() As String
    If g_Chinese Then S_LogSaved = "設定已儲存並套用" Else S_LogSaved = "Settings saved and applied"
End Function
Public Function S_LogDisabledOff() As String
    If g_Chinese Then S_LogDisabledOff = "斷電動作已停用" Else S_LogDisabledOff = "Power-off action disabled"
End Function
Public Function S_LogDisabledOn() As String
    If g_Chinese Then S_LogDisabledOn = "通電動作已停用" Else S_LogDisabledOn = "Power-on action disabled"
End Function
Public Function S_LogShellSkip() As String
    If g_Chinese Then S_LogShellSkip = "Shell 略過（空白或不存在）" Else S_LogShellSkip = "Shell skipped (empty or missing)"
End Function
Public Function S_LogShellDone(ByVal rc As Long) As String
    If g_Chinese Then S_LogShellDone = "Shell 完成 rc=" & rc Else S_LogShellDone = "Shell done rc=" & rc
End Function
Public Function S_LogFlowDoneOff() As String
    If g_Chinese Then S_LogFlowDoneOff = "斷電流程結束" Else S_LogFlowDoneOff = "Power-off flow done"
End Function
Public Function S_LogFlowDoneOn() As String
    If g_Chinese Then S_LogFlowDoneOn = "通電流程結束" Else S_LogFlowDoneOn = "Power-on flow done"
End Function
Public Function S_LogTrayAddFail() As String
    If g_Chinese Then S_LogTrayAddFail = "托盤圖示建立失敗" Else S_LogTrayAddFail = "Tray icon failed"
End Function
Public Function S_LogHdrTryOn() As String
    If g_Chinese Then S_LogHdrTryOn = "原生 HDR 嘗試：開啟" Else S_LogHdrTryOn = "Native HDR try: on"
End Function
Public Function S_LogHdrTryOff() As String
    If g_Chinese Then S_LogHdrTryOff = "原生 HDR 嘗試：關閉" Else S_LogHdrTryOff = "Native HDR try: off"
End Function
Public Function S_BalOffTitle() As String
    If g_Chinese Then S_BalOffTitle = "HDR 已關閉" Else S_BalOffTitle = "HDR off"
End Function
Public Function S_BalOffInfo() As String
    If g_Chinese Then S_BalOffInfo = "實體螢幕斷電，已關閉 HDR" Else S_BalOffInfo = "Monitor off, HDR disabled"
End Function
Public Function S_BalOnTitle() As String
    If g_Chinese Then S_BalOnTitle = "HDR 已開啟" Else S_BalOnTitle = "HDR on"
End Function
Public Function S_BalOnInfo() As String
    If g_Chinese Then S_BalOnInfo = "實體螢幕通電，已開啟 HDR" Else S_BalOnInfo = "Monitor on, HDR enabled"
End Function
Public Function S_BalFailTitle() As String
    If g_Chinese Then S_BalFailTitle = "動作失敗" Else S_BalFailTitle = "Action failed"
End Function
Public Function S_BalFailInfo() As String
    If g_Chinese Then S_BalFailInfo = "詳見記錄" Else S_BalFailInfo = "See log"
End Function

Public Function S_WordUnknown() As String
    If g_Chinese Then S_WordUnknown = "未知" Else S_WordUnknown = "Unknown"
End Function
