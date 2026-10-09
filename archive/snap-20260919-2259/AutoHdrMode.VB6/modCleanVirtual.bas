Option Explicit

' 本模組：清除遠端軟體虛擬顯示卡（WMI 列舉＋pnputil 移除）
' 只在提權後的 --clean 行程內執行；回傳 0=已清空或本來就沒有、1=逾時仍有殘留、2=執行錯誤

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 清除虛擬顯示卡主函數（含 45 秒驗證迴圈）
Public Function CleanVirtualGpus() As Long
    On Error GoTo Fail
    Dim ids() As String, names() As String, n As Long
    Dim i As Long, t0 As Long, ok As Boolean, targets As Long
    If Not EnumDisplayAdapters(ids, names, n) Then
        LogMsg S_LogCleanErr("enum")
        CleanVirtualGpus = 2
        Exit Function
    End If
    targets = 0
    For i = 0 To n - 1
        If IsVirtualAdapter(names(i)) Then
            targets = targets + 1
            LogMsg S_LogCleanRemoving(names(i))
            RemoveAdapter ids(i)
            Sleep 1000
        End If
    Next
    If targets = 0 Then
        LogMsg S_LogCleanNone()
        CleanVirtualGpus = 0
        Exit Function
    End If
    t0 = GetTickCount()
    Do
        If Not EnumDisplayAdapters(ids, names, n) Then
            LogMsg S_LogCleanErr("re-enum")
            CleanVirtualGpus = 2
            Exit Function
        End If
        ok = True
        For i = 0 To n - 1
            If IsVirtualAdapter(names(i)) Then ok = False: Exit For
        Next
        If ok Then
            LogMsg S_LogCleanOk()
            CleanVirtualGpus = 0
            Exit Function
        End If
        Sleep 2000
    Loop While (GetTickCount() - t0 >= 0) And (GetTickCount() - t0 < 45000)
    LogMsg S_LogCleanLeft()
    CleanVirtualGpus = 1
    Exit Function
Fail:
    LogMsg S_LogCleanErr(Err.Description)
    CleanVirtualGpus = 2
End Function

' 經 WMI Win32_PnPEntity 列舉 Display 類（取 Name 與 PNPDeviceID）
Private Function EnumDisplayAdapters(ByRef ids() As String, ByRef names() As String, ByRef count As Long) As Boolean
    On Error GoTo Fail
    Dim svc As Object, col As Object, mo As Object
    Dim capN As Long
    Set svc = GetObject("winmgmts:\\.\root\CIMV2")
    Set col = svc.ExecQuery("SELECT Name, PNPDeviceID FROM Win32_PnPEntity WHERE PNPClass = 'Display'")
    ReDim ids(0 To 15)
    ReDim names(0 To 15)
    capN = 16
    count = 0
    For Each mo In col
        On Error Resume Next
        Dim did As String, dname As String
        did = CStr(mo.PNPDeviceID)
        dname = CStr(mo.Name)
        If Err.Number <> 0 Then
            Err.Clear
        ElseIf did <> "" Then
            On Error GoTo Fail
            If count >= capN Then
                capN = capN * 2
                ReDim Preserve ids(0 To capN - 1)
                ReDim Preserve names(0 To capN - 1)
            End If
            ids(count) = did
            names(count) = dname
            count = count + 1
        End If
        On Error GoTo Fail
    Next
    EnumDisplayAdapters = True
    Exit Function
Fail:
    EnumDisplayAdapters = False
End Function

' pnputil 路徑解析（sysnative 優先）與單一裝置移除
Private Sub RemoveAdapter(ByVal instanceId As String)
    On Error Resume Next
    Dim pnp As String, windir As String
    windir = Environ$("windir")
    If Right$(windir, 1) <> "\" Then windir = windir & "\"
    pnp = windir & "sysnative\pnputil.exe"
    If Dir$(pnp) = "" Then pnp = windir & "System32\pnputil.exe"
    ShellRunHidden """" & pnp & """ /remove-device """ & instanceId & """", g_WorkDir
End Sub

' 顯示卡名稱不含 AMD/NVIDIA/INTEL 即視為虛擬（空名稱亦然）
Private Function IsVirtualAdapter(ByVal desc As String) As Boolean
    Dim u As String
    u = UCase$(desc)
    IsVirtualAdapter = (InStr(u, "AMD") = 0 And InStr(u, "NVIDIA") = 0 And InStr(u, "INTEL") = 0)
End Function
