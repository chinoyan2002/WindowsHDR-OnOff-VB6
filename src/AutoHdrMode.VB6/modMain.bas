' ���ҲաG�{���i�J�I�]��� C# �� Program.Main�^
Option Explicit

' �H�U�G�����P�ۭq�����X�Ϊ��t�Ψ禡
Private Declare Function CreateMutexA Lib "kernel32" Alias "CreateMutexA" (ByVal lpAttr As Long, ByVal bInit As Long, ByVal lpName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uCode As Long)

Private Const ERROR_ALREADY_EXISTS = 183&

' �i�J�I�GŪ�]�w���w�y�t�����R�O�C�ѼƴN��ưh�X�A�_�h���J���L�`�n
Public Sub Main()
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
    Dim hM As Long
    hM = CreateMutexA(0, 0, "Global\AutoHdrMode")
    If Err.LastDllError = ERROR_ALREADY_EXISTS Then ExitProcess 0
    Load frmTray
End Sub

' --probe�G�u�������ʧ@�]0 �}�B1 ���B2 �����^
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

' --fire-off/on�G�����]�R�O�]���եΡA��������^
Private Function FireCli(ByVal wantOn As Boolean) As Long
    On Error GoTo Fail
    Dim r1 As Long, r2 As Long, finalRc As Long
    If Not wantOn Then
        FireCli = DoPowerAction(False)
        Exit Function
    End If
    r1 = StepShell(True)
    r2 = StepNative(True)
    If r1 = 3 And r2 = 3 Then
        LogMsg S_LogEmptyOn()
        FireCli = 3
        Exit Function
    End If
    If r2 = 3 Then
        If r1 = 3 Then FireCli = 3 Else FireCli = 0
        Exit Function
    End If
    SleepMs g_VerifySec * 1000
    If CheckOn() Then
        LogMsg S_LogVerified()
        FireCli = 0
        Exit Function
    End If
    LogMsg S_LogVerifyFail()
    If g_CleanOnFail Then
        StepClean
        LogMsg S_LogRetry()
        r2 = StepNative(True)
    End If
    finalRc = 0
    If r1 <> 0 And r1 <> 3 Then finalRc = r1
    If r2 <> 0 And r2 <> 3 Then finalRc = r2
    FireCli = finalRc
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FireCli = 4
End Function

Private Function CleanMode(ByVal noRelaunch As Boolean) As Long
    On Error GoTo Fail
    If IsAdmin() Then
        CleanMode = CleanVirtualDisplays()
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
    If IsAdmin() Then
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

' --install/remove-autostart�G�g HKCU �}���Ұ�
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
