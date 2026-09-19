' 本模組：寫紀錄檔（ANSI 存檔，中文機用記事本開可讀）
Option Explicit

' 紀錄檔完整路徑
Public g_LogFile As String

' 決定紀錄檔位置（空字串=程式目錄）
Public Sub LogInit(ByVal dir As String)
    On Error Resume Next
    If dir = "" Then dir = App.Path
    If Right$(dir, 1) = "\" Then dir = Left$(dir, Len(dir) - 1)
    MkDir dir
    g_LogFile = dir & "\autohdrmode.log"
End Sub

' 附加一行（含時間）；失敗就吞掉，不中斷主程式
Public Sub LogMsg(ByVal msg As String)
    On Error Resume Next
    Dim fn As Integer
    fn = FreeFile
    Open g_LogFile For Append As #fn
    Print #fn, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #fn
End Sub
