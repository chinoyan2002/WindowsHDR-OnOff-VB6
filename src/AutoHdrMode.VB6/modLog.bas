Attribute VB_Name = "modLog"
' 本模組：寫紀錄檔（ANSI）；最多保留 LOG_MAX_LINES 行，超出刪舊留新
Option Explicit
' 記錄模組：autohdrmode.log 寫入與截斷

Public g_LogFile As String

Private Const LOG_MAX_LINES As Long = 2000
Private m_writeCount As Long

' 用途：決定記錄檔路徑（LogDir 空白則放 exe 目錄）
Public Sub LogInit(ByVal dir As String)
    On Error Resume Next
    If dir = "" Then dir = App.Path
    If Right$(dir, 1) = "\" Then dir = Left$(dir, Len(dir) - 1)
    MkDir dir
    g_LogFile = dir & "\autohdrmode.log"
    m_writeCount = 0
    ' 啟動時若已超過上限先裁一次
    Call LogTrimToMax
End Sub

' 用途：寫一行記錄（含時間戳）；定期檢查截斷
Public Sub LogMsg(ByVal msg As String)
    On Error Resume Next
    Dim fn As Integer
    If Len(g_LogFile) = 0 Then Exit Sub
    fn = FreeFile
    Open g_LogFile For Append As #fn
    Print #fn, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #fn
    m_writeCount = m_writeCount + 1
    ' 每寫入約 20 行檢查一次，避免每行都重寫整檔
    If m_writeCount >= 20 Then
        m_writeCount = 0
        Call LogTrimToMax
    End If
End Sub

' 超過 LOG_MAX_LINES 時只保留最後 N 行
Public Sub LogTrimToMax()
    On Error GoTo Fail
    Dim fn As Integer
    Dim line As String
    Dim buf As String
    Dim lines() As String
    Dim i As Long
    Dim n As Long
    Dim start As Long
    Dim out As String

    If Len(g_LogFile) = 0 Then Exit Sub
    If Dir$(g_LogFile) = "" Then Exit Sub

    fn = FreeFile
    Open g_LogFile For Input As #fn
    buf = ""
    Do While Not EOF(fn)
        Line Input #fn, line
        If Len(buf) > 0 Then buf = buf & vbCrLf
        buf = buf & line
    Loop
    Close #fn

    If Len(buf) = 0 Then Exit Sub
    lines = Split(buf, vbCrLf)
    n = UBound(lines) - LBound(lines) + 1
    ' 去掉結尾可能的空元素
    If n > 0 Then
        If Len(lines(UBound(lines))) = 0 Then n = n - 1
    End If
    If n <= LOG_MAX_LINES Then Exit Sub

    start = UBound(lines) - LOG_MAX_LINES + 1
    If start < LBound(lines) Then start = LBound(lines)
    out = ""
    For i = start To UBound(lines)
        If Len(lines(i)) > 0 Or i < UBound(lines) Then
            If Len(out) > 0 Then out = out & vbCrLf
            out = out & lines(i)
        End If
    Next

    fn = FreeFile
    Open g_LogFile For Output As #fn
    Print #fn, out;
    Close #fn
    Exit Sub
Fail:
    On Error Resume Next
    Close #fn
End Sub
