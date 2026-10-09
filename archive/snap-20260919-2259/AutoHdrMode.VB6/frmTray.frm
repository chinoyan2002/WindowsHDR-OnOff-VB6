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
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3   'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrPoll
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer tmrWait
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   120
   End
   Begin VB.Menu mnuTray
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuAuto
         Caption         =   "Auto"
      End
      Begin VB.Menu mnuFireOff
         Caption         =   "FireOff"
      End
      Begin VB.Menu mnuFireOn
         Caption         =   "FireOn"
      End
      Begin VB.Menu mnuHdrOn
         Caption         =   "HdrOn"
      End
      Begin VB.Menu mnuHdrOff
         Caption         =   "HdrOff"
      End
      Begin VB.Menu mnuSep1
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuLog
         Caption         =   "Log"
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
Attribute VB_Exposed = False
Option Explicit
' @@TR_H1@@
' @@TR_H2@@
Private Const WM_USER = &H400
Private Const WM_NULL = &H0
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOp As String, ByVal lpFile As String, ByVal lpParam As String, ByVal lpDir As String, ByVal nShow As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SW_SHOWNORMAL = 1

Private Const PH_IDLE = 0
Private Const PH_WAIT_HDR = 1
Private Const PH_WAIT_VERIFY = 2
Private Const PH_WAIT_AFTER_CLEAN = 3

Private m_lastSeen As Integer
Private m_candidate As Integer
Private m_stable As Long
Private m_lastErr As String
Private m_phase As Integer
Private m_targetOn As Boolean
Private m_due As Long
Private m_busy As Boolean

Private Sub Form_Load()
    Me.ScaleMode = vbPixels
    mnuTray.Visible = False
    mnuAuto.Caption = S_MenuAuto()
    mnuFireOff.Caption = S_MenuFireOff()
    mnuFireOn.Caption = S_MenuFireOn()
    mnuHdrOn.Caption = S_MenuHdrOn()
    mnuHdrOff.Caption = S_MenuHdrOff()
    mnuSettings.Caption = S_MenuSettings()
    mnuLog.Caption = S_MenuOpenLog()
    mnuExit.Caption = S_MenuExit()
    mnuAuto.Checked = g_AutoOn
    tmrPoll.Interval = g_PollSec * 1000
    tmrWait.Enabled = False
    m_lastSeen = -1
    m_candidate = -1
    m_stable = 0
    m_lastErr = ""
    m_phase = PH_IDLE
    m_busy = False
    Logger_Write_Start
    If TrayInit(Me.hwnd, S_TipUnknown(), g_WorkDir & "tray.ico", WM_USER + 1) Then
        TrayHook Me.hwnd
    Else
        LogMsg S_LogTrayAddFail()
    End If
    PollOnce
End Sub

Private Sub Logger_Write_Start()
    LogMsg S_LogStart(g_PollSec, g_StableN, g_AutoOn)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not g_Quitting Then
        Cancel = True
        Me.Hide
    End If
End Sub

' @@TR_TASKBAR_DOC@@
Public Sub OnTaskbarCreated()
    On Error Resume Next
    If TrayInit(Me.hwnd, S_TipUnknown(), g_WorkDir & "tray.ico", WM_USER + 1) Then
        TrayHook Me.hwnd
        PollOnce
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    msg = CLng(X)
    If msg > 1024 Then msg = CLng(X / Screen.TwipsPerPixelX)
    If msg = WM_RBUTTONUP Or msg = WM_LBUTTONUP Or msg = WM_RBUTTONDOWN Then
        SetForegroundWindow Me.hwnd
        PopupMenu mnuTray
        PostMessage Me.hwnd, WM_NULL, 0, 0
    ElseIf msg = WM_LBUTTONDBLCLK Then
        OpenSettings
    End If
End Sub

Private Sub mnuAuto_Click()
    mnuAuto.Checked = Not mnuAuto.Checked
    g_AutoOn = mnuAuto.Checked
    LogMsg S_LogAutoToggle(g_AutoOn)
End Sub

Private Sub mnuFireOff_Click()
    StartTransition False
End Sub

Private Sub mnuFireOn_Click()
    StartTransition True
End Sub

Private Sub mnuHdrOn_Click()
    DirectHDR True
End Sub

Private Sub mnuHdrOff_Click()
    DirectHDR False
End Sub

Private Sub mnuSettings_Click()
    OpenSettings
End Sub

Private Sub mnuLog_Click()
    OpenLogFolder
End Sub

Private Sub mnuExit_Click()
    QuitApp
End Sub

' Live-apply entry for settings window (menus are form-private).
Public Sub ApplyLiveSettings()
    On Error Resume Next
    tmrPoll.Interval = g_PollSec * 1000
    mnuAuto.Checked = g_AutoOn
End Sub

Private Sub OpenSettings()
    On Error GoTo Fail
    Load frmSettings
    frmSettings.Show vbModeless
    frmSettings.WindowState = vbNormal
    SetForegroundWindow frmSettings.hwnd
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
End Sub

Private Sub OpenLogFolder()
    On Error GoTo Fail
    ShellExecute Me.hwnd, "open", "explorer.exe", "/select,""" & g_LogFile & """", "", SW_SHOWNORMAL
    Exit Sub
Fail:
    LogMsg S_LogOpenLogErr(Err.Description)
End Sub

Private Sub QuitApp()
    g_Quitting = True
    tmrPoll.Enabled = False
    tmrWait.Enabled = False
    TrayUnhook
    TrayRemove
    LogMsg S_LogExit()
    Unload frmSettings
    Unload Me
    End
End Sub

Private Function PresentInt(ByVal present As Boolean) As Integer
    If present Then PresentInt = 1 Else PresentInt = 0
End Function

Private Function DelayHDRof(ByVal isOn As Boolean) As Long
    If isOn Then DelayHDRof = g_PowerOn.DelayHDR Else DelayHDRof = g_PowerOff.DelayHDR
End Function

Private Function TickPassed(ByVal due As Long) As Boolean
    TickPassed = (GetTickCount() - due >= 0)
End Function

Private Sub SetTip(ByVal present As Boolean)
    modTray.TrayTip S_TipState(present)
End Sub

Private Sub Blast(ByVal ok As Boolean, ByVal isOn As Boolean)
    If Not g_Balloon Then Exit Sub
    If ok Then
        If isOn Then modTray.TrayBalloon S_BalOnTitle(), S_BalOnInfo(), 1 _
        Else modTray.TrayBalloon S_BalOffTitle(), S_BalOffInfo(), 1
    Else
        modTray.TrayBalloon S_BalFailTitle(), S_BalFailInfo(), 2
    End If
End Sub

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
    SetTip present
    If m_stable < g_StableN Then Exit Sub
    If m_lastSeen = -1 Then
        m_lastSeen = PresentInt(present)
        LogMsg S_LogInit(present, n)
        Exit Sub
    End If
    If m_lastSeen = PresentInt(present) Then Exit Sub
    m_lastSeen = PresentInt(present)
    LogMsg S_LogEvent(present, n)
    If Not DirEnabled(present) Then
        If present Then LogMsg S_LogDisabledOn() Else LogMsg S_LogDisabledOff()
        Exit Sub
    End If
    If Not g_AutoOn Then Exit Sub
    StartTransition present
    Exit Sub
Fail:
    LogMsg S_LogPollErr(Err.Description)
End Sub

Private Function DirEnabled(ByVal isOn As Boolean) As Boolean
    If isOn Then DirEnabled = g_PowerOn.Enabled Else DirEnabled = g_PowerOff.Enabled
End Function

' @@TR_START_DOC@@
Public Sub StartTransition(ByVal wantOn As Boolean)
    m_targetOn = wantOn
    m_phase = PH_WAIT_HDR
    m_due = GetTickCount() + DelayHDRof(wantOn) * 1000
    If wantOn Then LogMsg S_LogPlanOn(DelayHDRof(True)) Else LogMsg S_LogPlanOff(DelayHDRof(False))
    tmrWait.Enabled = True
End Sub

Private Sub tmrPoll_Timer()
    PollOnce
End Sub

Private Sub tmrWait_Timer()
    If m_phase = PH_IDLE Then tmrWait.Enabled = False: Exit Sub
    If m_busy Then Exit Sub
    If Not TickPassed(m_due) Then Exit Sub
    Dim wantOn As Boolean
    wantOn = m_targetOn
    If m_lastSeen <> PresentInt(wantOn) Then
        If wantOn Then LogMsg S_LogSkipOn() Else LogMsg S_LogSkipOff()
        m_phase = PH_IDLE
        tmrWait.Enabled = False
        Exit Sub
    End If
    m_busy = True
    If m_phase = PH_WAIT_HDR Then
        DoHdrPhase wantOn
    ElseIf m_phase = PH_WAIT_VERIFY Then
        DoVerifyPhase wantOn
    ElseIf m_phase = PH_WAIT_AFTER_CLEAN Then
        DoRetryPhase wantOn
    End If
    m_busy = False
End Sub

Private Sub DoHdrPhase(ByVal wantOn As Boolean)
    Dim a As String
    If wantOn Then a = g_PowerOn.NativeHDR Else a = g_PowerOff.NativeHDR
    If Trim$(a) = "" Then
        DoShellPhase wantOn
        Exit Sub
    End If
    If wantOn Then LogMsg S_LogHdrTryOn() Else LogMsg S_LogHdrTryOff()
    Dim rc As Long
    rc = NativeHDR(a)
    If wantOn Then LogMsg S_LogHdrOn(rc) Else LogMsg S_LogHdrOff(rc)
    If rc <> 0 Then
        If Len(g_HdrErr) > 0 Then LogMsg g_HdrErr
        Blast False, wantOn
    End If
    Dim vsec As Long
    If wantOn Then vsec = g_PowerOn.VerifySec Else vsec = g_PowerOff.VerifySec
    If vsec > 0 Then
        m_phase = PH_WAIT_VERIFY
        m_due = GetTickCount() + vsec * 1000
        LogMsg S_LogVerifyWait(vsec)
    Else
        DoShellPhase wantOn
    End If
End Sub

Private Sub DoVerifyPhase(ByVal wantOn As Boolean)
    Dim ok As Boolean
    If wantOn Then ok = CheckOn() Else ok = CheckOff()
    If ok Then
        LogMsg S_LogVerified()
        Blast True, wantOn
        DoShellPhase wantOn
        Exit Sub
    End If
    LogMsg S_LogVerifyFail()
    Dim doClean As Boolean
    If wantOn Then doClean = g_PowerOn.CleanHelper Else doClean = g_PowerOff.CleanHelper
    If doClean Then
        StepClean
        Dim dac As Long
        If wantOn Then dac = g_PowerOn.DelayAfterClean Else dac = g_PowerOff.DelayAfterClean
        m_phase = PH_WAIT_AFTER_CLEAN
        m_due = GetTickCount() + dac * 1000
        LogMsg S_LogPlanOn(dac)
    Else
        DoShellPhase wantOn
    End If
End Sub

Private Sub DoRetryPhase(ByVal wantOn As Boolean)
    Dim a As String
    If wantOn Then a = g_PowerOn.NativeHDR Else a = g_PowerOff.NativeHDR
    If Trim$(a) <> "" Then
        LogMsg S_LogRetry()
        Dim rc As Long
        rc = NativeHDR(a)
        If wantOn Then LogMsg S_LogHdrOn(rc) Else LogMsg S_LogHdrOff(rc)
        If rc <> 0 Then
            If Len(g_HdrErr) > 0 Then LogMsg g_HdrErr
            Blast False, wantOn
        Else
            Blast True, wantOn
        End If
    End If
    DoShellPhase wantOn
End Sub

Private Sub DoShellPhase(ByVal wantOn As Boolean)
    Dim s As String
    If wantOn Then s = g_PowerOn.Shell Else s = g_PowerOff.Shell
    If Not ShellTargetExists(s) Then
        LogMsg S_LogShellSkip()
    Else
        Dim rc As Long
        rc = ShellRunHidden(s, g_WorkDir)
        LogMsg S_LogShellDone(rc)
    End If
    If wantOn Then LogMsg S_LogFlowDoneOn() Else LogMsg S_LogFlowDoneOff()
    m_phase = PH_IDLE
    tmrWait.Enabled = False
End Sub

' @@TR_DIRECT_DOC@@
Private Sub DirectHDR(ByVal wantOn As Boolean)
    On Error GoTo Fail
    Dim rc As Long
    If wantOn Then
        LogMsg S_LogHdrTryOn()
        rc = NativeHDR("ON")
        LogMsg S_LogHdrOn(rc)
    Else
        LogMsg S_LogHdrTryOff()
        rc = NativeHDR("OFF")
        LogMsg S_LogHdrOff(rc)
    End If
    If rc <> 0 Then
        If Len(g_HdrErr) > 0 Then LogMsg g_HdrErr
        Blast False, wantOn
    Else
        Blast True, wantOn
    End If
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
End Sub
