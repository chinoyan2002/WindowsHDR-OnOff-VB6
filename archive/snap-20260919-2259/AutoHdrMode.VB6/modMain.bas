' (orig header lost on 2026-09-19, re-add in IDE)
Option Explicit

' (orig header lost on 2026-09-19, re-add in IDE)
Private Declare Function CreateMutexA Lib "kernel32" Alias "CreateMutexA" (ByVal lpAttr As Long, ByVal bInit As Long, ByVal lpName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uCode As Long)

Private Const ERROR_ALREADY_EXISTS = 183&

Public g_Quitting As Boolean

' (orig header lost on 2026-09-19, re-add in IDE)
' 進入點總管（由設定窗 Form_Load 呼叫；CLI 與重複執行在這裡直接結束）
Public Sub AppBootstrap()
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
        ExitProcess CleanMode(False)
    ElseIf args = "--clean-elevated" Then
        ExitProcess CleanMode(True)
    ElseIf args = "--test-admin" Then
        ExitProcess TestAdmin()
    End If
    If App.PrevInstance Then
        MsgBox "在桌面右下角已執行!!", vbExclamation, "AutoHdrMode"
        ExitProcess 0
    End If
End Sub

' (orig header lost on 2026-09-19, re-add in IDE)
Private Function ProbeRun() As Long
    Dim n As Long
    n = WmiPhysicalCount()
    LogMsg S_LogProbe(n)
    If n > 0 Then
        ProbeRun = 0
    ElseIf n = 0 Then
        ProbeRun = 1
    Else
        ProbeRun = 2
    End If
End Function

' (orig header lost on 2026-09-19, re-add in IDE)
Private Function FireCli(ByVal wantOn As Boolean) As Long
    On Error GoTo Fail
    Dim r1 As Long, r2 As Long, finalRc As Long
    If wantOn Then
        r2 = StepNative(True)
        r1 = RunShellChecked(g_PowerOn.Shell, True)
    Else
        r2 = StepNative(False)
        r1 = RunShellChecked(g_PowerOff.Shell, False)
    End If
    finalRc = 0
    If r1 <> 0 And r1 <> 3 Then finalRc = r1
    If r2 <> 0 And r2 <> 3 Then finalRc = r2
    If r1 = 3 And r2 = 3 Then finalRc = 3
    FireCli = finalRc
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FireCli = 4
End Function

Private Function CleanMode(ByVal noRelaunch As Boolean) As Long
    On Error GoTo Fail
    If IsElevated() Then
        CleanMode = CleanVirtualGpus()
    ElseIf noRelaunch Then
        LogMsg S_LogCleanErr("already-elevated-child refused relaunch (fork guard)")
        CleanMode = 6
    Else
        Dim exe As String
        exe = App.Path
        If Right$(exe, 1) <> "\" Then exe = exe & "\"
        exe = exe & App.EXEName & ".exe"
        CleanMode = ShellRunElevated(exe, "--clean-elevated", g_WorkDir)
    End If
    Exit Function
Fail:
    LogMsg S_LogCleanErr(Err.Description)
    CleanMode = 2
End Function

Private Function TestAdmin() As Long
    On Error GoTo Fail
    If IsElevated() Then
        LogMsg "ADMIN-TEST: elevated=True " & IsAdminDiag()
        TestAdmin = 0
    Else
        LogMsg "ADMIN-TEST: elevated=False " & IsAdminDiag()
        TestAdmin = 1
    End If
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    TestAdmin = 2
End Function

' (orig header lost on 2026-09-19, re-add in IDE)
Private Function AutostartSet(ByVal install As Boolean) As Boolean
    On Error GoTo Fail
    Dim ws As Object, key As String, exe As String
    Set ws = CreateObject("WScript.Shell")
    key = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AutoHdrMode\"
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

' 讀開機自啟（供設定窗勾選初值）
Public Function AutostartInstalled() As Boolean
    On Error GoTo Fail
    Dim ws As Object, v As String
    Set ws = CreateObject("WScript.Shell")
    v = ws.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AutoHdrMode\")
    AutostartInstalled = (Trim$(v) <> "")
    Exit Function
Fail:
    AutostartInstalled = False
End Function
