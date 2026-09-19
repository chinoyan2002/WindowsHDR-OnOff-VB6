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

Private Const CREATE_NO_WINDOW = &H8000000 ' 建行程不配視窗
Private Const STARTF_USESHOWWINDOW = &H1 ' 啟用顯示設定欄位
Private Const SW_HIDE = 0 ' 隱藏視窗代碼
Private Const WAIT_OBJECT_0 = 0 ' 等到結束代碼

' Run hidden via cmd.exe with workDir as cwd. Waits in 500ms slices with
' DoEvents so the tray stays responsive. Returns exit code, -1 on fail/timeout.
' NOTE: no output capture (unlike the .NET build) - log records rc only.
' 等待上限吃參數秒數（0=啟動即走）；每 500 切一片讓托盤有反應
Public Function ShellRunHidden(ByVal shellCmd As String, ByVal workDir As String, ByVal timeoutSec As Long) As Long
    On Error GoTo Fail
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION
    Dim cmd As String, rc As Long, wr As Long, waited As Long
    cmd = "cmd.exe /s /c """ & shellCmd & """" ' 包 cmd 跑整串：& 符號不斷行
    si.cb = Len(si) ' 結構大小必填，API 對版用
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    If CreateProcessA(vbNullString, cmd, 0, 0, 0, CREATE_NO_WINDOW, 0, workDir, si, pi) = 0 Then
        ShellRunHidden = -1
        Exit Function
    End If
    If timeoutSec <= 0 Then
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
        ShellRunHidden = 0 ' 0 秒：啟動即走，不等結束不卡排程
        Exit Function
    End If
    waited = 0
    Do
        wr = WaitForSingleObject(pi.hProcess, 500) ' 500 毫秒一切片
        If wr = WAIT_OBJECT_0 Then Exit Do
        DoEvents ' 讓托盤有反應，不凍結
        waited = waited + 500 ' 累計等待毫秒
    Loop While waited < CLng(timeoutSec) * 1000 ' 上限=設定秒數
    If wr <> WAIT_OBJECT_0 Then TerminateProcess pi.hProcess, 99 ' 逾時殺掉，碼 99 識別
    If wr <> WAIT_OBJECT_0 Then
        rc = -1 ' 逾時或建失敗一律 -1
    Else
        GetExitCodeProcess pi.hProcess, rc ' 取真正結束碼
    End If
    CloseHandle pi.hThread ' 收執行緒代碼防漏
    CloseHandle pi.hProcess ' 收行程代碼防漏
    ShellRunHidden = rc
    Exit Function
Fail:
    ShellRunHidden = -1
End Function
