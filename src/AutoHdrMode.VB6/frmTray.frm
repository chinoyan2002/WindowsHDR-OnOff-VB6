VERSION 5.00
Begin VB.Form frmTray
   Caption         =   "AutoHdrMode"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1200
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   1200
   ShowInTaskbar   =   0
   StartUpPosition =   3
   Visible         =   0
   Begin VB.Timer tmrPoll
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer tmrWait
      Enabled         =   0
      Interval        =   500
      Left            =   600
      Top             =   120
   End
   Begin VB.Menu mnuTray
      Caption         =   "Tray"
      Visible         =   0
      Begin VB.Menu mnuAuto
         Caption         =   "Auto"
      End
      Begin VB.Menu mnuFireOff
         Caption         =   "FireOff"
      End
      Begin VB.Menu mnuFireOn
         Caption         =   "FireOn"
      End
      Begin VB.Menu mnuSep1
         Caption         =   "-"
      End
      Begin VB.Menu mnuLog
         Caption         =   "Log"
      End
      Begin VB.Menu mnuLogWin
         Caption         =   "LogWin"
      End
      Begin VB.Menu mnuSep2
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
' 本窗體：常駐本體（隱藏，不佔工作列）
' tmrPoll=輪詢，tmrWait=延遲執行；狀態機：輪詢→去拖→穩定轉換→排程→執行
Attribute VB_Exposed = False
Option Explicit

' 托盤圖示用的系統結構
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' 以下：托盤／圖示／計時／開資料夾用的系統函式
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMsg As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInst As Long, ByVal lpName As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOp As String, ByVal lpFile As String, ByVal lpParam As String, ByVal lpDir As String, ByVal nShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_USER = &H400
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const IDI_APPLICATION = 32512
Private Const SW_SHOWNORMAL = 1

' 狀態機變數（-1 未知、0 關、1 開）
Private m_lastSeen As Integer   ' -1 unknown, 0 off, 1 on
Private m_candidate As Integer  ' -1 unknown, 0 off, 1 on
Private m_stable As Long
Private m_lastErr As String
Private m_pending As Integer    ' -1 none, 0 off, 1 on
Private m_pendingDue As Long    ' GetTickCount target
Private m_phase As Integer      ' 0 initial delay, 1 verify wait
Private m_busy As Boolean
Private m_nid As NOTIFYICONDATA
Private m_trayAdded As Boolean

' 啟動：套語系選單文字、掛托盤圖示、開輪詢
Private Sub Form_Load()
    mnuTray.Visible = False
    mnuAuto.Caption = S_MenuAuto()
    mnuFireOff.Caption = S_MenuFireOff()
    mnuFireOn.Caption = S_MenuFireOn()
    mnuLog.Caption = S_MenuOpenLog()
    mnuLogWin.Caption = S_MenuLogWin()
    mnuExit.Caption = S_MenuExit()
    mnuAuto.Checked = g_AutoOn
    tmrPoll.Interval = g_PollSec * 1000
    tmrWait.Enabled = False
    m_lastSeen = -1
    m_candidate = -1
    m_stable = 0
    m_pending = -1
    m_lastErr = ""
    m_phase = 0
    LogMsg S_LogStart(g_PollSec, g_StableN, g_AutoOn)
    TrayAdd
    PollOnce
End Sub

' 布林轉 1/0（VB6 沒有三元運算子）
Private Function PresentInt(ByVal present As Boolean) As Integer
    If present Then PresentInt = 1 Else PresentInt = 0
End Function

' 依通電或斷電取延遲秒數
Private Function DelayOf(ByVal isOn As Boolean) As Long
    If isOn Then DelayOf = g_PowerOn.DelaySec Else DelayOf = g_PowerOff.DelaySec
End Function

' 核心：查 WMI→去拖→穩定才認→轉換就排程
Private Sub PollOnce()
    On Error GoTo Fail
    Dim n As Long
    n = WmiPhysicalCount()
    If g_WmiErr <> "" Then
        If g_WmiErr <> m_lastErr Then
            LogMsg S_LogWmiErr(g_WmiErr)
            m_lastErr = g_WmiErr
        End If
    Else
        m_lastErr = ""
    End If
    Dim present As Boolean
    present = (n > 0)
    If m_candidate <> PresentInt(present) Then
        m_candidate = PresentInt(present)
        m_stable = 1
    Else
        m_stable = m_stable + 1
    End If
    TrayTip S_TipState(present)
    If m_stable < g_StableN Then Exit Sub
    If m_lastSeen = -1 Then
        m_lastSeen = PresentInt(present)
        LogMsg S_LogInit(present, n)
        Exit Sub
    End If
    If m_lastSeen = PresentInt(present) Then Exit Sub
    m_lastSeen = PresentInt(present)
    LogMsg S_LogEvent(present, n)
    If Not g_AutoOn Then Exit Sub
    m_pending = PresentInt(present)
    m_pendingDue = GetTickCount() + DelayOf(present) * 1000
    m_phase = 0
    tmrWait.Enabled = True
    If present Then LogMsg S_LogPlanOn(DelayOf(True)) Else LogMsg S_LogPlanOff(DelayOf(False))
    Exit Sub
Fail:
    LogMsg S_LogPollErr(Err.Description)
End Sub

' 輪詢節拍
Private Sub tmrPoll_Timer()
    PollOnce
End Sub

' 判斷延遲時間到（可承受計時器 49 天迴轉）
Private Function TickPassed(ByVal due As Long) As Boolean
    TickPassed = (GetTickCount() - due >= 0)
End Function

' 延遲節拍：時間到先重驗狀態，變了就跳過
Private Sub tmrWait_Timer()
    If m_pending = -1 Then tmrWait.Enabled = False: Exit Sub
    If m_busy Then Exit Sub
    If Not TickPassed(m_pendingDue) Then Exit Sub
    Dim wantOn As Boolean
    wantOn = (m_pending = 1)
    If m_lastSeen <> PresentInt(wantOn) Then
        m_pending = -1
        tmrWait.Enabled = False
        Exit Sub
    End If
    m_busy = True
    If m_phase = 0 Then
        If wantOn Then
            StepShell True
            StepNative True
            m_phase = 1
            m_pendingDue = GetTickCount() + g_VerifySec * 1000
            LogMsg S_LogVerifyWait(g_VerifySec)
        Else
            DoPowerAction False
            m_pending = -1
            tmrWait.Enabled = False
        End If
    Else
        If m_pending <> 1 Then
            m_pending = -1
            tmrWait.Enabled = False
        ElseIf CheckOn() Then
            LogMsg S_LogVerified()
            m_pending = -1
            tmrWait.Enabled = False
        Else
            LogMsg S_LogVerifyFail()
            If g_CleanOnFail Then
                StepClean
                LogMsg S_LogRetry()
                StepNative True
            End If
            m_pending = -1
            tmrWait.Enabled = False
        End If
    End If
    m_busy = False
End Sub

' 執行斷電或通電命令並記結果

' 選單：自動切換開關
Private Sub mnuAuto_Click()
    mnuAuto.Checked = Not mnuAuto.Checked
    g_AutoOn = mnuAuto.Checked
    LogMsg S_LogAutoToggle(g_AutoOn)
End Sub

' 選單：立即執行斷電命令（不等延遲）
Private Sub mnuFireOff_Click()
    FireNow False
End Sub

' 選單：立即執行通電命令（不等延遲）
Private Sub mnuFireOn_Click()
    FireNow True
End Sub

' 立即執行本體（500 毫秒讓步一次）
Private Sub FireNow(ByVal wantOn As Boolean)
    Call DoPowerAction(wantOn)
End Sub

' 選單：開紀錄檔
Private Sub mnuLog_Click()
    OpenLogFolder
End Sub
Private Sub mnuLogWin_Click()
    On Error GoTo Fail
    Load frmLog
    frmLog.Show vbModeless
    frmLog.ZOrder 0
    On Error Resume Next
    frmLog.SetFocus
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
End Sub

' 用檔案總管選取紀錄檔
Private Sub OpenLogFolder()
    On Error GoTo Fail
    ShellExecute Me.hwnd, "open", "explorer.exe", "/select,""" & g_LogFile & """", "", SW_SHOWNORMAL
    Exit Sub
Fail:
    LogMsg S_LogOpenLogErr(Err.Description)
End Sub

' 選單：結束（先拔托盤圖示再卸載窗體）
Private Sub mnuExit_Click()
    TrayRemove
    LogMsg S_LogExit()
    Unload Me
End Sub

' 把圖示掛上通知區域
Private Sub TrayAdd()
    With m_nid
        .cbSize = Len(m_nid)
        .hwnd = Me.hwnd
        .uID = 1
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = WM_USER + 1
        .hIcon = LoadIcon(0, IDI_APPLICATION)
        .szTip = S_TipStarting() & vbNullChar
    End With
    m_trayAdded = (Shell_NotifyIcon(NIM_ADD, m_nid) <> 0)
End Sub

' 更新圖示提示文字
Private Sub TrayTip(ByVal txt As String)
    If Not m_trayAdded Then Exit Sub
    m_nid.szTip = txt & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, m_nid
End Sub

' 把圖示拔掉（結束前必做，否則圖示殘留）
Private Sub TrayRemove()
    If m_trayAdded Then
        Shell_NotifyIcon NIM_DELETE, m_nid
        m_trayAdded = False
    End If
End Sub

' 托盤回呼：右鍵彈選單，雙擊開紀錄檔
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    If Not m_trayAdded Then Exit Sub
    msg = X / Screen.TwipsPerPixelX
    If msg = WM_RBUTTONUP Then
        ' 先搶前景，否則選單點了沒反應（VB6 托盤老坑）
        SetForegroundWindow Me.hwnd
        PopupMenu mnuTray
    ElseIf msg = WM_LBUTTONDBLCLK Then
        OpenLogFolder
    End If
End Sub
