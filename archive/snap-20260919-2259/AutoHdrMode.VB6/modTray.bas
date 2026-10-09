Option Explicit

' 本模組：托盤圖示＋氣球（改寫自使用者提供版本）
' 結構補 guidItem 到標準 508；氣球等級參數化

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
    guidItem As String * 16
    hBalloonIcon As Long
End Type

Private Const NIF_ICON As Long = &H2
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_TIP As Long = &H4
Private Const NIF_INFO As Long = &H10
Private Const NIF_STATE As Long = &H8

Private Const NIM_ADD As Long = &H0
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_DELETE As Long = &H2
Private Const NIM_SETVERSION As Long = &H4

Private Const NIIF_NONE As Long = &H0
Private Const NIIF_INFO As Long = &H1
Private Const NIIF_WARNING As Long = &H2
Private Const NIIF_ERROR As Long = &H3
Private Const NIIF_NOSOUND As Long = &H10

Private Const NOTIFYICON_VERSION_4 As Long = 4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMsg As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInst As Long, ByVal lpName As Long) As Long

Private Const IDI_APPLICATION = 32512

Private m_nid As NOTIFYICONDATA
Private m_added As Boolean

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = -4
Private m_oldProc As Long
Private m_hookHwnd As Long
Private m_taskbarMsg As Long

' 掛上托盤圖示（自訂圖缺檔退回系統圖）
Public Function TrayInit(ByVal hwnd As Long, ByVal tip As String, ByVal iconFile As String, ByVal callbackMsg As Long) As Boolean
    On Error GoTo Fail
    Dim pic As StdPicture
    Dim hIco As Long
    hIco = 0
    If Len(Trim$(iconFile)) > 0 Then
        On Error Resume Next
        Set pic = LoadPicture(iconFile)
        If Err.Number = 0 Then
            If Not pic Is Nothing Then hIco = pic.Handle
        End If
        Err.Clear
        On Error GoTo Fail
    End If
    If hIco = 0 Then hIco = LoadIcon(0, IDI_APPLICATION)
    With m_nid
        .cbSize = Len(m_nid)
        .hwnd = hwnd
        .uID = 1
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = callbackMsg
        .hIcon = hIco
        .szTip = tip & vbNullChar
    End With
    If Shell_NotifyIcon(NIM_ADD, m_nid) = 0 Then GoTo Fail
    m_nid.uTimeoutOrVersion = NOTIFYICON_VERSION_4
    Shell_NotifyIcon NIM_SETVERSION, m_nid
    m_added = True
    TrayInit = True
    Exit Function
Fail:
    m_added = False
    TrayInit = False
End Function

' 更新提示文字
Public Sub TrayTip(ByVal tip As String)
    If Not m_added Then Exit Sub
    m_nid.szTip = tip & vbNullChar
    m_nid.uFlags = NIF_TIP
    Shell_NotifyIcon NIM_MODIFY, m_nid
    m_nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
End Sub

' 發氣球，等級 1 資訊／2 警告／3 錯誤
Public Sub TrayBalloon(ByVal title As String, ByVal info As String, ByVal level As Long)
    If Not m_added Then Exit Sub
    Dim flags As Long
    Select Case level
        Case 2
            flags = NIIF_WARNING
        Case 3
            flags = NIIF_ERROR Or NIIF_NOSOUND
        Case Else
            flags = NIIF_INFO Or NIIF_NOSOUND
    End Select
    With m_nid
        .uFlags = NIF_INFO
        .szInfoTitle = title & vbNullChar
        .szInfo = info & vbNullChar
        .dwInfoFlags = flags
        .uTimeoutOrVersion = 15000
    End With
    Shell_NotifyIcon NIM_MODIFY, m_nid
    m_nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
End Sub

' 拔掉圖示
Public Sub TrayRemove()
    If Not m_added Then Exit Sub
    Shell_NotifyIcon NIM_DELETE, m_nid
    m_added = False
End Sub

Public Function TrayIsAdded() As Boolean
    TrayIsAdded = m_added
End Function

' 本模組追加：訊息鉤子（TaskbarCreated 重掛圖示；IDE 內嚴禁按 Stop，先卸鉤）



' 掛鉤（成功回 True；失敗不影響輪詢）
Public Function TrayHook(ByVal hwnd As Long) As Boolean
    On Error GoTo Fail
    m_taskbarMsg = RegisterWindowMessage("TaskbarCreated")
    m_oldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf TrayWndProc)
    m_hookHwnd = hwnd
    TrayHook = (m_oldProc <> 0)
    Exit Function
Fail:
    TrayHook = False
End Function

Public Sub TrayUnhook()
    On Error Resume Next
    If m_oldProc <> 0 And m_hookHwnd <> 0 Then SetWindowLong m_hookHwnd, GWL_WNDPROC, m_oldProc
    m_oldProc = 0
    m_hookHwnd = 0
End Sub

Public Function TrayWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If m_taskbarMsg <> 0 And uMsg = m_taskbarMsg Then frmTray.OnTaskbarCreated
    TrayWndProc = CallWindowProc(m_oldProc, hwnd, uMsg, wParam, lParam)
End Function
