Option Explicit

' 本模組：清除遠端軟體虛擬顯示卡（邏輯對照 SetHDR.ps1，VB6 實作）
' 只在提權後的 --clean 行程內執行；回傳 0=已清空或本來就沒有、1=逾時仍有殘留、2=執行錯誤

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 清除虛擬顯示卡主函數（含 45 秒驗證迴圈）
Public Function CleanVirtualDisplays() As Long
    On Error GoTo Fail
    Dim ids() As String, names() As String, n As Long
    Dim i As Long, t0 As Long, ok As Boolean, targets As Long
    If Not EnumAdapters(ids, names, n) Then
        LogMsg S_LogCleanErr("enum")
        CleanVirtualDisplays = 2
        Exit Function
    End If
    targets = 0
    For i = 0 To n - 1
        If IsVirtualAdapter(names(i)) Then
            targets = targets + 1
            LogMsg S_LogCleanRemoving(names(i))
            ShellRunHidden "pnputil.exe /remove-device """ & ids(i) & """", g_WorkDir
            Sleep 1000
        End If
    Next
    If targets = 0 Then
        LogMsg S_LogCleanNone()
        CleanVirtualDisplays = 0
        Exit Function
    End If
    t0 = GetTickCount()
    Do
        If Not EnumAdapters(ids, names, n) Then
            LogMsg S_LogCleanErr("re-enum")
            CleanVirtualDisplays = 2
            Exit Function
        End If
        ok = True
        For i = 0 To n - 1
            If IsVirtualAdapter(names(i)) Then ok = False: Exit For
        Next
        If ok Then
            LogMsg S_LogCleanOk()
            CleanVirtualDisplays = 0
            Exit Function
        End If
        Sleep 2000
    Loop While (GetTickCount() - t0 >= 0) And (GetTickCount() - t0 < 45000)
    LogMsg S_LogCleanLeft()
    CleanVirtualDisplays = 1
    Exit Function
Fail:
    LogMsg S_LogCleanErr(Err.Description)
    CleanVirtualDisplays = 2
End Function

' 列舉 Display 類裝置（經 pnputil，輸出轉暫存檔再解析）
Private Function EnumAdapters(ByRef ids() As String, ByRef names() As String, ByRef count As Long) As Boolean
    On Error GoTo Fail
    Dim tmp As String, fn As Integer, line As String, t As String
    Dim curId As String, curDesc As String
    Dim capN As Long
    tmp = Environ$("TEMP") & "\ahv_enum.txt"
    On Error Resume Next
    Kill tmp
    On Error GoTo Fail
    ShellRunHidden "pnputil.exe /enum-devices /class Display > """ & tmp & """", g_WorkDir
    If Dir$(tmp) = "" Then Exit Function
    ReDim ids(0 To 15)
    ReDim names(0 To 15)
    capN = 16
    count = 0
    curId = ""
    curDesc = ""
    fn = FreeFile
    Open tmp For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, line
        t = Trim$(line)
        If Left$(t, 12) = "Instance ID:" Then
            curId = Trim$(Mid$(t, 13))
        ElseIf Left$(t, 19) = "Device Description:" Then
            curDesc = Trim$(Mid$(t, 20))
        ElseIf t = "" Then
            If curId <> "" Then
                If count >= capN Then
                    capN = capN * 2
                    ReDim Preserve ids(0 To capN - 1)
                    ReDim Preserve names(0 To capN - 1)
                End If
                ids(count) = curId
                names(count) = curDesc
                count = count + 1
                curId = ""
                curDesc = ""
            End If
        End If
    Loop
    Close #fn
    If curId <> "" Then
        If count >= capN Then
            ReDim Preserve ids(0 To count)
            ReDim Preserve names(0 To count)
        End If
        ids(count) = curId
        names(count) = curDesc
        count = count + 1
    End If
    EnumAdapters = True
    Exit Function
Fail:
    On Error Resume Next
    Close #fn
    EnumAdapters = False
End Function

' 顯示卡名稱不含 AMD/NVIDIA/INTEL 即視為虛擬（空名稱亦然）
Private Function IsVirtualAdapter(ByVal desc As String) As Boolean
    Dim u As String
    u = UCase$(desc)
    IsVirtualAdapter = (InStr(u, "AMD") = 0 And InStr(u, "NVIDIA") = 0 And InStr(u, "INTEL") = 0)
End Function
