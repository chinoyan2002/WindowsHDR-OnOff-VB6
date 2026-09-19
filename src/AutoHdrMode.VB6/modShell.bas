Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApp As String, ByVal lpCmd As String, ByVal lpProcAttr As Long, ByVal lpThreadAttr As Long, ByVal bInherit As Long, ByVal dwFlags As Long, ByVal lpEnv As Long, ByVal lpDir As String, lpStart As STARTUPINFO, lpProc As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMs As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const CREATE_NO_WINDOW = &H8000000
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const WAIT_OBJECT_0 = 0

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal hProc As Long, ByVal access As Long, hTok As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal hTok As Long, ByVal InfoClass As Long, pInfo As Any, ByVal InfoLen As Long, retLen As Long) As Long

Private Type TOKEN_ELEVATION
    TokenIsElevated As Long
End Type

Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const TokenElevation As Long = 20

'=====================================================================
' 追加：提權執行、原生 HDR、組合動作（融合 modWindowsHDR）
'=====================================================================

' 原生 HDR 上次失敗原因（中文，由 modWindowsHDR 提供）
Public g_HdrErr As String

Public Function ShellRunHidden(ByVal shellCmd As String, ByVal workDir As String) As Long
    On Error GoTo Fail
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION
    Dim cmd As String, rc As Long, wr As Long, waited As Long
    cmd = "cmd.exe /s /c " & Chr$(34) & shellCmd & Chr$(34)
    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    If CreateProcessA(vbNullString, cmd, 0, 0, 0, CREATE_NO_WINDOW, 0, workDir, si, pi) = 0 Then
        ShellRunHidden = -1
        Exit Function
    End If
    waited = 0
    Do
        wr = WaitForSingleObject(pi.hProcess, 500)
        If wr = WAIT_OBJECT_0 Then Exit Do
        DoEvents
        waited = waited + 500
    Loop While waited < 120000
    If wr <> WAIT_OBJECT_0 Then
        TerminateProcess pi.hProcess, 99
        rc = -1
    Else
        GetExitCodeProcess pi.hProcess, rc
    End If
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    ShellRunHidden = rc
    Exit Function
Fail:
    ShellRunHidden = -1
End Function

' 毫秒等待（CLI 流程用）
Public Sub SleepMs(ByVal ms As Long)
    If ms > 0 Then Sleep ms
End Sub

Public Function IsAdmin() As Boolean
    On Error GoTo Fail
    Dim tok As Long, elev As TOKEN_ELEVATION, retLen As Long
    IsAdmin = False
    If OpenProcessToken(GetCurrentProcess(), &H8, tok) = 0 Then Exit Function
    If GetTokenInformation(tok, TokenElevation, elev, 4, retLen) = 0 Then
        CloseHandle tok
        Exit Function
    End If
    CloseHandle tok
    IsAdmin = (elev.TokenIsElevated <> 0)
    Exit Function
Fail:
    On Error Resume Next
    IsAdmin = False
End Function

Public Function IsAdminDiag() As String
    On Error GoTo Fail
    Dim tok As Long, elev As TOKEN_ELEVATION, retLen As Long, r1 As Long, r2 As Long
    r1 = OpenProcessToken(GetCurrentProcess(), &H8, tok)
    If r1 = 0 Then
        IsAdminDiag = "openT=0 le=" & Err.LastDllError
        Exit Function
    End If
    r2 = GetTokenInformation(tok, TokenElevation, elev, 4, retLen)
    CloseHandle tok
    IsAdminDiag = "openT=1 getInfo=" & r2 & " elevated=" & elev.TokenIsElevated & " retLen=" & retLen & " le=" & Err.LastDllError
    Exit Function
Fail:
    IsAdminDiag = "ex=" & Err.Description
End Function

' 以 runas 提權執行並等待，逾時 150 秒
Public Function ShellRunElevated(ByVal exePath As String, ByVal args As String, ByVal workDir As String) As Long
    On Error GoTo Fail
    Dim sei As SHELLEXECUTEINFO, wr As Long, waited As Long, rc As Long
    sei.cbSize = Len(sei)
    sei.fMask = SEE_MASK_NOCLOSEPROCESS
    sei.lpVerb = "runas"
    sei.lpFile = exePath
    sei.lpParameters = args
    sei.lpDirectory = workDir
    sei.nShow = SW_HIDE
    If ShellExecuteEx(sei) = 0 Then ShellRunElevated = -1: Exit Function
    If sei.hProcess = 0 Then ShellRunElevated = -1: Exit Function
    waited = 0
    Do
        wr = WaitForSingleObject(sei.hProcess, 500)
        If wr = WAIT_OBJECT_0 Then Exit Do
        DoEvents
        waited = waited + 500
    Loop While waited < 150000
    If wr <> WAIT_OBJECT_0 Then
        TerminateProcess sei.hProcess, 99
        rc = -1
    Else
        GetExitCodeProcess sei.hProcess, rc
    End If
    CloseHandle sei.hProcess
    ShellRunElevated = rc
    Exit Function
Fail:
    ShellRunElevated = -1
End Function

' 執行原生 HDR 開關（底層 modWindowsHDR，不經過任何外部程式）
Public Function NativeHDR(ByVal action As String) As Long
    On Error GoTo Fail
    Dim ok As Boolean
    Dim a As String
    a = UCase$(Trim$(action))
    g_HdrErr = ""
    If a = "ON" Then
        ok = HDR_Enable()
    ElseIf a = "OFF" Then
        ok = HDR_Disable()
    Else
        NativeHDR = 3
        Exit Function
    End If
    If ok Then
        NativeHDR = 0
    Else
        g_HdrErr = HDR_GetLastErrorText()
        NativeHDR = 1
    End If
    Exit Function
Fail:
    g_HdrErr = Err.Description
    NativeHDR = 1
End Function

Public Function StepShell(ByVal wantOn As Boolean) As Long
    On Error GoTo Fail
    Dim s As String
    If wantOn Then s = g_PowerOn.Shell Else s = g_PowerOff.Shell
    If Trim$(s) = "" Then StepShell = 3: Exit Function
    Dim rc As Long
    rc = ShellRunHidden(s, g_WorkDir)
    If wantOn Then LogMsg S_LogFiredOn(rc) Else LogMsg S_LogFiredOff(rc)
    StepShell = rc
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    StepShell = 4
End Function

Public Function StepNative(ByVal wantOn As Boolean) As Long
    On Error GoTo Fail
    Dim a As String
    If wantOn Then a = g_PowerOnHdr Else a = g_PowerOffHdr
    If Trim$(a) = "" Then StepNative = 3: Exit Function
    Dim rc As Long
    rc = NativeHDR(a)
    If wantOn Then LogMsg S_LogHdrOn(rc) Else LogMsg S_LogHdrOff(rc)
    If rc <> 0 Then
        If Len(g_HdrErr) > 0 Then LogMsg g_HdrErr
    End If
    StepNative = rc
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    StepNative = 4
End Function

Public Function CheckOn() As Boolean
    On Error GoTo Fail
    Dim displays() As HDR_DISPLAY_INFO
    Dim i As Long, lo As Long, hi As Long
    CheckOn = False
    If Not HDR_GetDisplays(displays) Then Exit Function
    On Error Resume Next
    lo = LBound(displays)
    hi = UBound(displays)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo Fail
    For i = lo To hi
        If displays(i).HDRActive Then CheckOn = True: Exit Function
    Next
    Exit Function
Fail:
End Function

Public Function StepClean() As Long
    On Error GoTo Fail
    Dim exe As String
    LogMsg S_LogCleanStart()
    exe = App.Path
    If Right$(exe, 1) <> "\" Then exe = exe & "\"
    exe = exe & App.EXEName & ".exe"
    StepClean = ShellRunElevated(exe, "--clean", g_WorkDir)
    LogMsg S_LogCleanDone(StepClean)
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    StepClean = 4
End Function

Public Function DoPowerAction(ByVal wantOn As Boolean) As Long
    On Error GoTo Fail
    Dim r1 As Long, r2 As Long
    r1 = StepShell(wantOn)
    r2 = StepNative(wantOn)
    If r1 = 3 And r2 = 3 Then
        If wantOn Then LogMsg S_LogEmptyOn() Else LogMsg S_LogEmptyOff()
        DoPowerAction = 3
        Exit Function
    End If
    If r1 <> 0 And r1 <> 3 Then DoPowerAction = r1: Exit Function
    If r2 <> 0 And r2 <> 3 Then DoPowerAction = r2: Exit Function
    DoPowerAction = 0
    Exit Function
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    DoPowerAction = 4
End Function
