' 本模組：隱藏執行 Shell 命令（對照 C# 版 ShellRunner）
' 經 cmd.exe /s /c 執行，工作目錄=程式目錄；只回傳結束碼，不擷取輸出
Option Explicit

' 建立處理程序用的啟動參數結構
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

' 建立處理程序後取回的 handle 結構
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

' 以下：Windows 原生函式宣告
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApp As String, ByVal lpCmd As String, ByVal lpProcAttr As Long, ByVal lpThreadAttr As Long, ByVal bInherit As Long, ByVal dwFlags As Long, ByVal lpEnv As Long, ByVal lpDir As String, lpStart As STARTUPINFO, lpProc As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMs As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const CREATE_NO_WINDOW = &H8000000
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const WAIT_OBJECT_0 = 0

' Run hidden via cmd.exe with workDir as cwd. Waits in 500ms slices with
' DoEvents so the tray stays responsive. Returns exit code, -1 on fail/timeout.
' NOTE: no output capture (unlike the .NET build) - log records rc only.
' 隱藏執行並等最多 120 秒；每 500 毫秒讓出控制權以免托盤凍結
Public Function ShellRunHidden(ByVal shellCmd As String, ByVal workDir As String) As Long
    On Error GoTo Fail
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION
    Dim cmd As String, rc As Long, wr As Long, waited As Long
    cmd = "cmd.exe /s /c """ & shellCmd & """"
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
    If wr <> WAIT_OBJECT_0 Then TerminateProcess pi.hProcess, 99
    If wr <> WAIT_OBJECT_0 Then
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
