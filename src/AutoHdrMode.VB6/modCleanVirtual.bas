Attribute VB_Name = "modCleanVirtual"
' 本模組：清除虛擬顯示卡（pnputil），僅在通電且需要時由提權 --clean 模式呼叫
' 邏輯：AMD/NVIDIA/Intel 以外全砍 + 45 秒驗證迴圈
Option Explicit
' 清卡模組：辨識並移除虛擬顯示卡，必要時提權

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
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Const CREATE_NO_WINDOW = &H8000000
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const WAIT_OBJECT_0 = 0
Private Const TOKEN_QUERY As Long = &H8
Private Const TokenElevation As Long = 20

' 是否以系統管理員執行
Public Function IsElevated() As Boolean
    On Error GoTo Fail
    Dim hToken As Long, elev As Long, retLen As Long
    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hToken) = 0 Then
        IsElevated = False
        Exit Function
    End If
    elev = 0
    If GetTokenInformation(hToken, TokenElevation, elev, 4, retLen) <> 0 Then
        IsElevated = (elev <> 0)
    Else
        IsElevated = False
    End If
    CloseHandle hToken
    Exit Function
Fail:
    IsElevated = False
End Function

' 簡單同步跑命令回結束碼（隱藏）
Private Function ShellRunSimple(ByVal cmdLine As String) As Long
    On Error GoTo Fail
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION
    Dim wr As Long, rc As Long
    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    If CreateProcessA(vbNullString, cmdLine, 0, 0, 0, CREATE_NO_WINDOW, 0, vbNullString, si, pi) = 0 Then ' 移除 cmd.exe /s /c 避免 & 符號截斷指令
        ShellRunSimple = -1
        Exit Function
    End If
    wr = WaitForSingleObject(pi.hProcess, 10000)
    If wr = WAIT_OBJECT_0 Then
        GetExitCodeProcess pi.hProcess, rc
    Else
        rc = -1
    End If
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    ShellRunSimple = rc
    Exit Function
Fail:
    ShellRunSimple = -1
End Function

' 取得 pnputil 路徑（優先 sysnative 避開 WOW64）
Private Function PnpUtilPath() As String
    Dim p As String
    p = Environ$("WINDIR")
    If Right$(p, 1) <> "\" Then p = p & "\"
    If Len(dir$(p & "sysnative\pnputil.exe")) > 0 Then
        PnpUtilPath = p & "sysnative\pnputil.exe"
    Else
        PnpUtilPath = p & "System32\pnputil.exe"
    End If
End Function

' 核心：清除虛擬顯卡，回傳 0=成功/無需清理，非0=失敗
' 只應在 elevated 行程中呼叫
Public Function CleanVirtualGpus() As Long
    On Error GoTo Fail
    Dim col As Object, mo As Object, n As Long, instId As String, Name As String
    Dim pnp As String, rc As Long, left As Long
    
    LogMsg S_LogCleanStart()
    pnp = PnpUtilPath()
    
    Set col = GetObject("winmgmts:\\.\root\cimv2").ExecQuery( _
    "SELECT * FROM Win32_PnPEntity WHERE PNPClass='Display'")
    
    n = 0
    For Each mo In col
        On Error Resume Next
        Name = CStr(mo.Name)
        instId = CStr(mo.DeviceID)
        If Err.Number <> 0 Then Err.Clear: GoTo NextDev
        On Error GoTo Fail
        ' 保留 AMD / NVIDIA / Intel
        If InStr(1, Name, "AMD", vbTextCompare) > 0 Then GoTo NextDev
        If InStr(1, Name, "NVIDIA", vbTextCompare) > 0 Then GoTo NextDev
        If InStr(1, Name, "Intel", vbTextCompare) > 0 Then GoTo NextDev
        ' 其餘視為虛擬卡，砍
        LogMsg S_LogCleanRemove(Name)
        rc = ShellRunSimple("""" & pnp & """ /remove-device """ & instId & """")
        LogMsg S_LogCleanRemoveRc(rc)
        n = n + 1
        Sleep 1000
NextDev:
    Next
    
    If n = 0 Then
        LogMsg S_LogCleanNone()
        CleanVirtualGpus = 0
        Exit Function
    End If
    
    ' 45 秒驗證迴圈
    left = 45
    Do While left > 0
        Sleep 2000
        left = left - 2
        If CountVirtualLeft() = 0 Then
            LogMsg S_LogCleanOk()
            CleanVirtualGpus = 0
            Exit Function
        End If
    Loop
    LogMsg S_LogCleanTimeout()
    CleanVirtualGpus = 2
    Exit Function
Fail:
    LogMsg S_LogCleanErr(Err.Description)
    CleanVirtualGpus = 1
End Function

' 用途：數殘留虛擬卡（驗證清卡成效）
Private Function CountVirtualLeft() As Long
    On Error GoTo Fail
    Dim col As Object, mo As Object, Name As String, n As Long
    Set col = GetObject("winmgmts:\\.\root\cimv2").ExecQuery( _
        "SELECT * FROM Win32_PnPEntity WHERE PNPClass='Display'")
    n = 0
    For Each mo In col
        On Error Resume Next
        Name = CStr(mo.Name)
        If Err.Number <> 0 Then Err.Clear: GoTo Nx
        On Error GoTo Fail
        If InStr(1, Name, "AMD", vbTextCompare) = 0 And _
           InStr(1, Name, "NVIDIA", vbTextCompare) = 0 And _
           InStr(1, Name, "Intel", vbTextCompare) = 0 Then
            n = n + 1
        End If
Nx:
    Next
    CountVirtualLeft = n
    Exit Function
Fail:
    CountVirtualLeft = -1
End Function
