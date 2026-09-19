Attribute VB_Name = "modMain"
' 本模組：程式進入點（對照 C# 版 Program.Main）
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uCode As Long)
Private Declare Function ShellExecuteA Lib "shell32" (ByVal hwnd As Long, ByVal lpOp As String, ByVal lpFile As String, ByVal lpParam As String, ByVal lpDir As String, ByVal nShow As Long) As Long

Private Const SW_HIDE = 0

' 命令列模式：若有 --xxx 則處理完結束並回 True（呼叫端勿再跑 GUI）
' 正常 GUI 啟動請以 frmSettings 為 Startup，在 Form_Load 呼叫 AppBootstrap
Public Function AppBootstrap() As Boolean
    Dim args As String
    args = LCase$(Trim$(Command$))
    ConfigLoad
    LangInit g_Language
    LogInit g_LogDir

    If args = "--install-autostart" Then
        If AutostartSet(True) Then ExitProcess 0 Else ExitProcess 1
    ElseIf args = "--remove-autostart" Then
        If AutostartSet(False) Then ExitProcess 0 Else ExitProcess 1
    ElseIf args = "--probe" Then
        ExitProcess ProbeRun()
    ElseIf args = "--fire-off" Then
        ExitProcess FireCli(False)
    ElseIf args = "--fire-on" Then
        ExitProcess FireCli(True)
    ElseIf args = "--clean" Then
        ExitProcess CleanCli()
    End If

    If App.PrevInstance Then
        MsgBox "在桌面右下角已執行!!", vbSystemModal + vbMsgBoxSetForeground
        ExitProcess 0
    End If

    ' False = 繼續 GUI（主窗體已是 Startup）
    AppBootstrap = False
End Function

' 保留 Sub Main 僅供手動改回 Startup=Sub Main 時相容；正式以 frmSettings 啟動
Public Sub Main()
    ' 正式 Startup 為 frmSettings；此處僅供手動改回 Sub Main 或純 CLI。
    ' 勿與 frmSettings.Form_Load 的 AppBootstrap / Load frmTray 重複並用。
    If AppBootstrap() Then Exit Sub
    Load frmSettings
    If Not g_FirstRun Then frmSettings.Hide
End Sub

' --clean：提權清虛擬卡。非 admin 就 runas 自己，做完退出。
Private Function CleanCli() As Long
    On Error GoTo Fail
    If Not IsElevated() Then
        LogMsg S_LogCleanElevate()
        Dim exe As String, rc As Long
        exe = App.Path
        If Right$(exe, 1) <> "\" Then exe = exe & "\"
        exe = exe & App.EXEName & ".exe"
        rc = ShellExecuteA(0, "runas", exe, "--clean", App.Path, SW_HIDE)
        If rc > 32 Then
            CleanCli = 0
        Else
            LogMsg S_LogCleanElevateFail(rc)
            CleanCli = 5
        End If
        Exit Function
    End If
    CleanCli = CleanVirtualGpus()
    Exit Function
Fail:
    LogMsg S_LogCleanErr(Err.Description)
    CleanCli = 1
End Function

' 用途：--probe：輸出通電與待命台數；回傳：0開 1關 3待命
Private Function ProbeRun() As Long
    Dim n As Long, sb As Long
    n = WmiPhysicalCount()
    sb = WmiStandbyCount()
    LogMsg S_LogProbe(n, sb)
    If n > 0 Then
        ProbeRun = 0
    ElseIf sb > 0 Then
        ProbeRun = 3
    Else
        ProbeRun = 1
    End If
End Function

' --fire-off/on：完整流程（通電含驗證＋可選清卡）
Private Function FireCli(ByVal wantOn As Boolean) As Long
    ' CLI：與 GUI 相同（DelayHDR → HDR → 驗證 → 可選清卡重試 → Shell）
    On Error GoTo Fail
    Dim cfg As TransCfg
    Dim ok As Boolean
    Dim st As HDR_STATUS
    Dim crc As Long
    Dim i As Long
    Dim t0 As Single
    Dim rc As Long

    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff

    ' DelayHDR
    If cfg.DelayHDR > 0 Then
        LogMsg S_LogPlanHdr(wantOn, cfg.DelayHDR)
        For i = 1 To cfg.DelayHDR
            t0 = Timer
            Do While Timer < t0 + 1
                DoEvents
            Loop
        Next
    End If

    If cfg.NativeHDR = "on" Or cfg.NativeHDR = "off" Then
        If cfg.NativeHDR = "on" Then
            LogMsg S_LogNativeTryOn()
            ok = HDR_Enable()
        Else
            LogMsg S_LogNativeTryOff()
            ok = HDR_Disable()
        End If
        st = HDR_GetStatus(0)
        LogMsg S_LogNativeResult(ok, st)
        If cfg.NativeHDR = "off" And (st = HDR_UNSUPPORTED Or st = HDR_ERROR Or st = HDR_STATUS_UNKNOWN) Then
            LogMsg S_LogNativeOffNoPath()
        End If

        If cfg.VerifySeconds > 0 Then
            LogMsg S_LogVerifyWait(cfg.VerifySeconds)
            For i = 1 To cfg.VerifySeconds
                t0 = Timer
                Do While Timer < t0 + 1
                    DoEvents
                Loop
            Next
            st = HDR_GetStatus(0)
            ok = False
            If cfg.NativeHDR = "on" Then
                ok = (st = HDR_ON)
            ElseIf cfg.NativeHDR = "off" Then
                ok = (st = HDR_OFF)
            End If
            If ok Then
                LogMsg S_LogVerifyOk()
            Else
                LogMsg S_LogVerifyFail()
                If cfg.CleanHelper Then
                    crc = CleanCli()
                    LogMsg S_LogCleanDone(crc)
                    If cfg.DelayAfterClean > 0 Then
                        LogMsg S_LogWaitAfterClean(cfg.DelayAfterClean)
                        For i = 1 To cfg.DelayAfterClean
                            t0 = Timer
                            Do While Timer < t0 + 1
                                DoEvents
                            Loop
                        Next
                    End If
                    LogMsg S_LogRetryEnable()
                    If cfg.NativeHDR = "on" Then
                        ok = HDR_Enable()
                    Else
                        ok = HDR_Disable()
                    End If
                    LogMsg S_LogNativeResult(ok, HDR_GetStatus(0))
                End If
            End If
        End If
    Else
        LogMsg S_LogNativeSkip()
    End If

    If Len(Trim$(cfg.Shell)) = 0 Then
        LogMsg S_LogShellSkipEmpty()
    ElseIf Not ShellTargetExists(cfg.Shell) Then
        LogMsg S_LogShellSkipMissing(cfg.Shell)
    Else
        rc = ShellRunHidden(cfg.Shell, g_WorkDir)
        LogMsg S_LogShellRun(rc)
    End If

    LogMsg S_LogPhaseDone(wantOn)
    FireCli = 0
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FireCli = 1
End Function

' 共用：執行一次通電或斷電（NativeHDR 優先，否則／另外跑 Shell）
Public Sub FireTransitionCore(ByVal wantOn As Boolean, Optional ByVal fromCli As Boolean = False)
    ' CLI / 相容入口：同步跑 HDR + Shell（不含清卡狀態機）
    On Error GoTo Fail
    Dim cfg As TransCfg, ok As Boolean, st As HDR_STATUS, rc As Long
    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff
    If cfg.NativeHDR = "on" Or cfg.NativeHDR = "off" Then
        If cfg.NativeHDR = "on" Then
            LogMsg S_LogNativeTryOn()
            ok = HDR_Enable()
        Else
            LogMsg S_LogNativeTryOff()
            ok = HDR_Disable()
        End If
        st = HDR_GetStatus(0)
        LogMsg S_LogNativeResult(ok, st)
    End If
    If Len(Trim$(cfg.Shell)) = 0 Then
        LogMsg S_LogShellSkipEmpty()
    ElseIf Not ShellTargetExists(cfg.Shell) Then
        LogMsg S_LogShellSkipMissing(cfg.Shell)
    Else
        rc = ShellRunHidden(cfg.Shell, g_WorkDir)
        LogMsg S_LogShellRun(rc)
    End If
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
End Sub

' 用途：查 HKCU Run 是否已註冊開機啟動；回傳：有=True
Public Function AutostartInstalled() As Boolean
    On Error Resume Next
    Dim ws As Object, key As String, cur As String
    Set ws = CreateObject("WScript.Shell")
    key = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AutoHdrMode"
    cur = ws.RegRead(key)
    AutostartInstalled = (Err.Number = 0 And Len(cur) > 0)
    Err.Clear
End Function

' 用途：寫入或刪除 HKCU Run 開機啟動；參數 install=True 註冊、False 刪除
Public Function AutostartSet(ByVal install As Boolean) As Boolean
    On Error GoTo Fail
    Dim ws As Object, key As String, exe As String
    Set ws = CreateObject("WScript.Shell")
    key = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AutoHdrMode"
    exe = App.Path
    If Right$(exe, 1) <> "\" Then exe = exe & "\"
    exe = """" & exe & App.EXEName & ".exe" & """"
    If install Then
        ws.RegWrite key, exe, "REG_SZ"
        LogMsg S_LogAsIn()
    Else
        On Error Resume Next
        Dim cur As String
        cur = ws.RegRead(key)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            On Error GoTo Fail
            ws.RegDelete key
        End If
        LogMsg S_LogAsRm()
    End If
    AutostartSet = True
    Exit Function
Fail:
    LogMsg S_LogAsErr(Err.Description)
    AutostartSet = False
End Function

' 版本字串：vbp 版號 Major.Minor.Revision，Revision 每次編譯自動加一
Public Function AppVersionText() As String
    On Error Resume Next
    AppVersionText = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Function

' 編譯日期：取 exe 檔寫入時間；IDE 內執行無 exe 時回傳 IDE
Public Function AppBuildText() As String
    On Error GoTo Fail
    Dim exe As String
    exe = App.Path
    If Right$(exe, 1) <> "\" Then exe = exe & "\"
    exe = exe & App.EXEName & ".exe"
    If Dir$(exe) = "" Then AppBuildText = "IDE": Exit Function
    AppBuildText = Format$(FileDateTime(exe), "yyyy-mm-dd hh:nn")
    Exit Function
Fail:
    AppBuildText = "IDE"
End Function

' 標題列用：AutoHdrMode v1.0.3 (2026-09-19 23:06)，數字無需翻譯
Public Function AppVersionLine() As String
    AppVersionLine = "AutoHdrMode " & AppVersionText() & " (" & AppBuildText() & ")"
End Function
