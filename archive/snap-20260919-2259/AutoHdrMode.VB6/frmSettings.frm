VERSION 5.00
Begin VB.Form frmSettings
   Caption         =   "Settings"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "frmSettings"
   ScaleHeight     =   6240
   ScaleWidth      =   7560
   StartUpPosition =   2   'CenterScreen
   Begin VB.CommandButton cmdTabG
      Caption         =   "G"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1480
   End
   Begin VB.CommandButton cmdTabOff
      Caption         =   "Off"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1480
   End
   Begin VB.CommandButton cmdTabOn
      Caption         =   "On"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1480
   End
   Begin VB.CommandButton cmdTabLog
      Caption         =   "Log"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1480
   End
   Begin VB.PictureBox picGeneral
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7320
      Begin VB.Label lblStatusCap
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1000
      End
      Begin VB.Label lblPhyCap
         Caption         =   "Phy"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblPhyVal
         Caption         =   "-"
         Height          =   255
         Left            =   2500
         TabIndex        =   32
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblHdrCap
         Caption         =   "HDR"
         Height          =   255
         Left            =   4100
         TabIndex        =   33
         Top             =   120
         Width           =   800
      End
      Begin VB.Label lblHdrVal
         Caption         =   "-"
         Height          =   255
         Left            =   5000
         TabIndex        =   34
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblEvtCap
         Caption         =   "Evt"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblEvtVal
         Caption         =   "-"
         Height          =   255
         Left            =   1400
         TabIndex        =   36
         Top             =   480
         Width           =   5500
      End
      Begin VB.Label lblPoll
         Caption         =   "Poll"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   900
         Width           =   1500
      End
      Begin VB.TextBox txtPoll
         Height          =   315
         Left            =   1700
         TabIndex        =   5
         Text            =   "1"
         Top             =   870
         Width           =   800
      End
      Begin VB.Label lblStable
         Caption         =   "Stable"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1300
         Width           =   1500
      End
      Begin VB.TextBox txtStable
         Height          =   315
         Left            =   1700
         TabIndex        =   6
         Text            =   "3"
         Top             =   1270
         Width           =   800
      End
      Begin VB.CheckBox chkAuto
         Caption         =   "Auto"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1700
         Width           =   3000
      End
      Begin VB.CheckBox chkAutostart
         Caption         =   "Autostart"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2100
         Width           =   3500
      End
      Begin VB.CheckBox chkBalloon
         Caption         =   "Balloon"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2500
         Width           =   3500
      End
   End
   Begin VB.PictureBox picOff
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CheckBox chkEnabledOff
         Caption         =   "Enabled"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   3000
      End
      Begin VB.Label lblDelayHdrOff
         Caption         =   "Delay"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   560
         Width           =   2200
      End
      Begin VB.TextBox txtDelayHdrOff
         Height          =   315
         Left            =   2500
         TabIndex        =   12
         Text            =   "1"
         Top             =   530
         Width           =   800
      End
      Begin VB.Label lblNativeOff
         Caption         =   "Native"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   2200
      End
      Begin VB.ComboBox cboNativeOff
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   930
         Width           =   2200
      End
      Begin VB.Label lblVerifyOff
         Caption         =   "Verify"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1360
         Width           =   2200
      End
      Begin VB.TextBox txtVerifyOff
         Height          =   315
         Left            =   2500
         TabIndex        =   14
         Text            =   "2"
         Top             =   1330
         Width           =   800
      End
      Begin VB.CheckBox chkCleanOff
         Caption         =   "Clean"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1760
         Width           =   4500
      End
      Begin VB.Label lblAfterCleanOff
         Caption         =   "AfterClean"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   2200
      End
      Begin VB.TextBox txtAfterCleanOff
         Height          =   315
         Left            =   2500
         TabIndex        =   16
         Text            =   "5"
         Top             =   2130
         Width           =   800
      End
      Begin VB.Label lblShellOff
         Caption         =   "Shell"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2560
         Width           =   2200
      End
      Begin VB.TextBox txtShellOff
         Height          =   315
         Left            =   2500
         TabIndex        =   17
         Top             =   2530
         Width           =   4400
      End
   End
   Begin VB.PictureBox picOn
      Height          =   4575
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CheckBox chkEnabledOn
         Caption         =   "Enabled"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   3000
      End
      Begin VB.Label lblDelayHdrOn
         Caption         =   "Delay"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   560
         Width           =   2200
      End
      Begin VB.TextBox txtDelayHdrOn
         Height          =   315
         Left            =   2500
         TabIndex        =   20
         Text            =   "2"
         Top             =   530
         Width           =   800
      End
      Begin VB.Label lblNativeOn
         Caption         =   "Native"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   2200
      End
      Begin VB.ComboBox cboNativeOn
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   930
         Width           =   2200
      End
      Begin VB.Label lblVerifyOn
         Caption         =   "Verify"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1360
         Width           =   2200
      End
      Begin VB.TextBox txtVerifyOn
         Height          =   315
         Left            =   2500
         TabIndex        =   22
         Text            =   "2"
         Top             =   1330
         Width           =   800
      End
      Begin VB.CheckBox chkCleanOn
         Caption         =   "Clean"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1760
         Width           =   4500
      End
      Begin VB.Label lblAfterCleanOn
         Caption         =   "AfterClean"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2160
         Width           =   2200
      End
      Begin VB.TextBox txtAfterCleanOn
         Height          =   315
         Left            =   2500
         TabIndex        =   24
         Text            =   "5"
         Top             =   2130
         Width           =   800
      End
      Begin VB.Label lblShellOn
         Caption         =   "Shell"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2560
         Width           =   2200
      End
      Begin VB.TextBox txtShellOn
         Height          =   315
         Left            =   2500
         TabIndex        =   25
         Top             =   2530
         Width           =   4400
      End
   End
   Begin VB.PictureBox picLog
      Height          =   4575
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   7320
      Begin VB.TextBox txtLogView
         Height          =   3800
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   27
         Top             =   120
         Width           =   7080
      End
      Begin VB.CheckBox chkPauseLog
         Caption         =   "Pause"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   4050
         Width           =   2500
      End
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   49
      Top             =   5700
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   6150
      TabIndex        =   50
      Top             =   5700
      Width           =   1200
   End
   Begin VB.Timer tmrStatus
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   5700
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 本窗體：主設定（Startup 物件；CLI 與重複執行直接結束）
' X 與關閉鈕只隱藏；結束唯一走托盤選單
Private Sub Form_Load()
    AppBootstrap
    Load frmTray
    ApplyLang
    LoadValues
    ShowPage 0
    tmrStatus.Interval = 2000
    tmrStatus.Enabled = True
    If Not g_FirstRun Then Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not g_Quitting Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub ApplyLang()
    Me.Caption = S_SetCap()
    cmdTabG.Caption = S_TabGeneral()
    cmdTabOff.Caption = S_TabPowerOff()
    cmdTabOn.Caption = S_TabPowerOn()
    cmdTabLog.Caption = S_TabLog()
    lblStatusCap.Caption = S_LblStatus()
    lblPhyCap.Caption = S_LblPhy()
    lblHdrCap.Caption = S_LblHdr2()
    lblEvtCap.Caption = S_LblLastEvt()
    lblPoll.Caption = S_LblPoll()
    lblStable.Caption = S_LblStable()
    chkAuto.Caption = S_ChkAuto()
    chkAutostart.Caption = S_ChkAutostart()
    chkBalloon.Caption = S_ChkBalloon()
    chkEnabledOff.Caption = S_LblEnabled()
    lblDelayHdrOff.Caption = S_LblDelayHdr()
    lblNativeOff.Caption = S_LblNative()
    cboNativeOff.Clear
    cboNativeOff.AddItem S_OptHdrNone()
    cboNativeOff.AddItem S_OptHdrOff()
    cboNativeOff.AddItem S_OptHdrOn()
    lblVerifyOff.Caption = S_LblVerify()
    chkCleanOff.Caption = S_ChkClean()
    lblAfterCleanOff.Caption = S_LblAfterClean()
    lblShellOff.Caption = S_LblShell()
    chkEnabledOn.Caption = S_LblEnabled()
    lblDelayHdrOn.Caption = S_LblDelayHdr()
    lblNativeOn.Caption = S_LblNative()
    cboNativeOn.Clear
    cboNativeOn.AddItem S_OptHdrNone()
    cboNativeOn.AddItem S_OptHdrOn()
    cboNativeOn.AddItem S_OptHdrOff()
    lblVerifyOn.Caption = S_LblVerify()
    chkCleanOn.Caption = S_ChkClean()
    lblAfterCleanOn.Caption = S_LblAfterClean()
    lblShellOn.Caption = S_LblShell()
    chkPauseLog.Caption = S_PauseCap()
    cmdSave.Caption = S_BtnSave()
    cmdClose.Caption = S_BtnClose()
End Sub

Private Sub LoadValues()
    txtPoll.Text = CStr(g_PollSec)
    txtStable.Text = CStr(g_StableN)
    If g_AutoOn Then chkAuto.Value = vbChecked Else chkAuto.Value = vbUnchecked
    If AutostartInstalled() Then chkAutostart.Value = vbChecked Else chkAutostart.Value = vbUnchecked
    If g_Balloon Then chkBalloon.Value = vbChecked Else chkBalloon.Value = vbUnchecked
    LoadPage False
    LoadPage True
    RefreshStatus
End Sub

Private Sub LoadPage(ByVal isOn As Boolean)
    If isOn Then
        If g_PowerOn.Enabled Then chkEnabledOn.Value = vbChecked Else chkEnabledOn.Value = vbUnchecked
        txtDelayHdrOn.Text = CStr(g_PowerOn.DelayHDR)
        SetComboOn g_PowerOn.NativeHDR
        txtVerifyOn.Text = CStr(g_PowerOn.VerifySec)
        If g_PowerOn.CleanHelper Then chkCleanOn.Value = vbChecked Else chkCleanOn.Value = vbUnchecked
        txtAfterCleanOn.Text = CStr(g_PowerOn.DelayAfterClean)
        txtShellOn.Text = g_PowerOn.Shell
    Else
        If g_PowerOff.Enabled Then chkEnabledOff.Value = vbChecked Else chkEnabledOff.Value = vbUnchecked
        txtDelayHdrOff.Text = CStr(g_PowerOff.DelayHDR)
        SetComboOff g_PowerOff.NativeHDR
        txtVerifyOff.Text = CStr(g_PowerOff.VerifySec)
        If g_PowerOff.CleanHelper Then chkCleanOff.Value = vbChecked Else chkCleanOff.Value = vbUnchecked
        txtAfterCleanOff.Text = CStr(g_PowerOff.DelayAfterClean)
        txtShellOff.Text = g_PowerOff.Shell
    End If
End Sub

Private Sub SetComboOff(ByVal v As String)
    v = UCase$(Trim$(v))
    If v = "OFF" Then
        cboNativeOff.ListIndex = 1
    ElseIf v = "ON" Then
        cboNativeOff.ListIndex = 2
    Else
        cboNativeOff.ListIndex = 0
    End If
End Sub

Private Sub SetComboOn(ByVal v As String)
    v = UCase$(Trim$(v))
    If v = "ON" Then
        cboNativeOn.ListIndex = 1
    ElseIf v = "OFF" Then
        cboNativeOn.ListIndex = 2
    Else
        cboNativeOn.ListIndex = 0
    End If
End Sub

Private Function GetComboOff() As String
    Select Case cboNativeOff.ListIndex
        Case 1: GetComboOff = "OFF"
        Case 2: GetComboOff = "ON"
        Case Else: GetComboOff = ""
    End Select
End Function

Private Function GetComboOn() As String
    Select Case cboNativeOn.ListIndex
        Case 1: GetComboOn = "ON"
        Case 2: GetComboOn = "OFF"
        Case Else: GetComboOn = ""
    End Select
End Function

Private Sub ShowPage(ByVal idx As Long)
    picGeneral.Visible = (idx = 0)
    picOff.Visible = (idx = 1)
    picOn.Visible = (idx = 2)
    picLog.Visible = (idx = 3)
End Sub

Private Sub cmdTabG_Click()
    ShowPage 0
End Sub

Private Sub cmdTabOff_Click()
    ShowPage 1
End Sub

Private Sub cmdTabOn_Click()
    ShowPage 2
End Sub

Private Sub cmdTabLog_Click()
    ShowPage 3
    RefreshLogView
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Fail
    g_PollSec = ClampLng(Val(txtPoll.Text), 1, 30)
    g_StableN = ClampLng(Val(txtStable.Text), 1, 99)
    g_AutoOn = (chkAuto.Value = vbChecked)
    g_Balloon = (chkBalloon.Value = vbChecked)
    If chkAutostart.Value = vbChecked Then AutostartSet True Else AutostartSet False
    SavePage False
    SavePage True
    If ConfigSave() Then
        frmTray.ApplyLiveSettings
        LogMsg S_LogSaved()
    Else
        LogMsg S_LogAutoErr("ConfigSave")
    End If
    LoadValues
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
End Sub

Private Sub SavePage(ByVal isOn As Boolean)
    If isOn Then
        g_PowerOn.Enabled = (chkEnabledOn.Value = vbChecked)
        g_PowerOn.DelayHDR = ClampLng(Val(txtDelayHdrOn.Text), 0, 120)
        g_PowerOn.NativeHDR = GetComboOn()
        g_PowerOn.VerifySec = ClampLng(Val(txtVerifyOn.Text), 0, 60)
        g_PowerOn.CleanHelper = (chkCleanOn.Value = vbChecked)
        g_PowerOn.DelayAfterClean = ClampLng(Val(txtAfterCleanOn.Text), 0, 120)
        g_PowerOn.Shell = txtShellOn.Text
    Else
        g_PowerOff.Enabled = (chkEnabledOff.Value = vbChecked)
        g_PowerOff.DelayHDR = ClampLng(Val(txtDelayHdrOff.Text), 0, 120)
        g_PowerOff.NativeHDR = GetComboOff()
        g_PowerOff.VerifySec = ClampLng(Val(txtVerifyOff.Text), 0, 60)
        g_PowerOff.CleanHelper = (chkCleanOff.Value = vbChecked)
        g_PowerOff.DelayAfterClean = ClampLng(Val(txtAfterCleanOff.Text), 0, 120)
        g_PowerOff.Shell = txtShellOff.Text
    End If
End Sub

Private Function ClampLng(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then v = lo
    If v > hi Then v = hi
    ClampLng = v
End Function

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub tmrStatus_Timer()
    If Not Me.Visible Then Exit Sub
    RefreshStatus
    If picLog.Visible Then
        If chkPauseLog.Value <> vbChecked Then RefreshLogView
    End If
End Sub

Private Sub RefreshStatus()
    On Error Resume Next
    Dim n As Long
    n = WmiPhysicalCount()
    If n > 0 Then lblPhyVal.Caption = S_WordOn() Else lblPhyVal.Caption = S_WordOff()
    Dim st As String
    st = S_WordUnknown()
    Dim ds() As HDR_DISPLAY_INFO
    If HDR_GetDisplays(ds) Then
        Dim lo As Long
        lo = LBound(ds)
        If Err.Number = 0 Then
            If ds(lo).HDRActive Then st = S_WordOn() Else st = S_WordOff()
        Else
            Err.Clear
        End If
    End If
    lblHdrVal.Caption = st
    lblEvtVal.Caption = g_LastEvt
End Sub

Private Sub RefreshLogView()
    On Error Resume Next
    Dim buf(0 To 199) As String
    Dim n As Long, fn As Integer, s As String
    Dim start As Long, count As Long, k As Long, out As String
    If Dir$(g_LogFile) = "" Then Exit Sub
    fn = FreeFile
    Open g_LogFile For Input As #fn
    n = 0
    Do While Not EOF(fn)
        Line Input #fn, s
        buf(n Mod 200) = s
        n = n + 1
    Loop
    Close #fn
    If n = 0 Then Exit Sub
    If n > 200 Then
        count = 200
        start = n - 200
    Else
        count = n
        start = 0
    End If
    out = ""
    For k = 0 To count - 1
        out = out & buf((start + k) Mod 200) & vbCrLf
    Next
    If out <> txtLogView.Text Then
        txtLogView.Text = out
        txtLogView.SelStart = Len(txtLogView.Text)
    End If
End Sub
