' 本模組：查實體螢幕數量（對照 C# 版 WmiProbe）
' 遠端軟體的虛擬螢幕不會出現在這裡；螢幕斷電時查詢會丟錯，是正常現象
Option Explicit

' 上次查詢的錯誤文字（空字串=正常）
Public g_WmiErr As String

' Returns active physical monitor count, or -1 with g_WmiErr set.
' Virtual displays (remote software virtual screens) do not appear here.
' NOTE: with the physical monitor powered off, this query throws
' ("not supported") instead of returning zero rows.
' 回傳作用中實體螢幕數；-1=查詢失敗
Public Function WmiPhysicalCount() As Long
    On Error GoTo Fail
    Dim svc As Object, col As Object, mo As Object
    Dim n As Long
    Set svc = GetObject("winmgmts:\\.\root\WMI")
    Set col = svc.ExecQuery("SELECT Active FROM WmiMonitorBasicDisplayParams")
    n = 0
    For Each mo In col
        On Error Resume Next
        If CBool(mo.Active) Then n = n + 1
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo Fail
    Next
    g_WmiErr = ""
    WmiPhysicalCount = n
    Exit Function
Fail:
    g_WmiErr = "Err" & Err.Number & " " & Err.Description
    WmiPhysicalCount = -1
End Function
