VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "AutoHdrMode 設定"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7800
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton btnTab0 
      Caption         =   "一般"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton btnTab1 
      Caption         =   "斷電"
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton btnTab2 
      Caption         =   "通電"
      Height          =   360
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton btnTab3 
      Caption         =   "記錄"
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1200
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  '沒有框線
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4800
      ScaleWidth      =   7560
      TabIndex        =   4
      Top             =   540
      Width           =   7560
      Begin VB.TextBox txtPoll 
         Height          =   300
         Left            =   2700
         TabIndex        =   7
         Top             =   560
         Width           =   900
      End
      Begin VB.TextBox txtStable 
         Height          =   300
         Left            =   2700
         TabIndex        =   10
         Top             =   1040
         Width           =   900
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "啟用自動切換（依螢幕通電／斷電執行動作）"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   6000
      End
      Begin VB.CheckBox chkAutostart 
         Caption         =   "開機時自動執行本程式"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   6000
      End
      Begin VB.CheckBox chkBalloon 
         Caption         =   "顯示氣球提示"
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   6000
      End
      Begin VB.Label lblMainTitle 
         Caption         =   "一般設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4000
      End
      Begin VB.Label lblPoll 
         Caption         =   "輪詢間隔（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label lblPollHint 
         Caption         =   "每隔幾秒查一次實體螢幕狀態（1～30）"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   3800
         TabIndex        =   8
         Top             =   600
         Width           =   3600
      End
      Begin VB.Label lblStable 
         Caption         =   "穩定確認次數"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label lblStableHint 
         Caption         =   "需連續幾次偵測結果相同，才認定狀態已變更"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   3800
         TabIndex        =   11
         Top             =   1080
         Width           =   3600
      End
   End
   Begin VB.PictureBox picOff 
      BorderStyle     =   0  '沒有框線
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4800
      ScaleWidth      =   7560
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   7560
      Begin VB.CheckBox chkOffEn 
         Caption         =   "啟用斷電自動動作"
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   4000
      End
      Begin VB.TextBox txtOffDelayHDR 
         Height          =   300
         Left            =   2700
         TabIndex        =   18
         Top             =   860
         Width           =   900
      End
      Begin VB.ComboBox cboOffHDR 
         Height          =   315
         Left            =   2700
         Style           =   2  '單純下拉式
         TabIndex        =   20
         Top             =   1340
         Width           =   2200
      End
      Begin VB.TextBox txtOffVerify 
         Height          =   300
         Left            =   2700
         TabIndex        =   22
         Top             =   1820
         Width           =   900
      End
      Begin VB.CheckBox chkOffClean 
         Caption         =   "驗證失敗時清除虛擬顯示卡後重試"
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   2340
         Width           =   6000
      End
      Begin VB.TextBox txtOffDelayClean 
         Height          =   300
         Left            =   2700
         TabIndex        =   26
         Top             =   2780
         Width           =   900
      End
      Begin VB.TextBox txtOffShell 
         Height          =   300
         Left            =   2700
         TabIndex        =   28
         Top             =   3260
         Width           =   4600
      End
      Begin VB.Label lblOffTitle 
         Caption         =   "斷電時動作（與通電相同結構）"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   60
         Width           =   6000
      End
      Begin VB.Label lblOffDelayHDR 
         Caption         =   "HDR 動作前等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   900
         Width           =   2400
      End
      Begin VB.Label lblOffHDR 
         Caption         =   "原生 HDR 動作"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1380
         Width           =   2400
      End
      Begin VB.Label lblOffVerify 
         Caption         =   "驗證等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1860
         Width           =   2400
      End
      Begin VB.Label lblOffVerifyHint 
         Caption         =   "HDR 後等多久再檢查；0＝不驗證"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   3800
         TabIndex        =   23
         Top             =   1860
         Width           =   3600
      End
      Begin VB.Label lblOffDelayClean 
         Caption         =   "清卡後再試前等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2820
         Width           =   2400
      End
      Begin VB.Label lblOffShell 
         Caption         =   "額外 Shell 命令"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3300
         Width           =   2400
      End
      Begin VB.Label lblOffShellHint 
         Caption         =   "空白或不存在則不執行；在 HDR／驗證／清卡流程全部結束後執行"
         ForeColor       =   &H00808080&
         Height          =   480
         Left            =   2700
         TabIndex        =   29
         Top             =   3660
         Width           =   4600
      End
   End
   Begin VB.PictureBox picOn 
      BorderStyle     =   0  '沒有框線
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4800
      ScaleWidth      =   7560
      TabIndex        =   30
      Top             =   540
      Visible         =   0   'False
      Width           =   7560
      Begin VB.CheckBox chkOnEn 
         Caption         =   "啟用通電自動動作"
         Height          =   300
         Left            =   120
         TabIndex        =   32
         Top             =   420
         Width           =   4000
      End
      Begin VB.TextBox txtOnDelayHDR 
         Height          =   300
         Left            =   2700
         TabIndex        =   34
         Top             =   860
         Width           =   900
      End
      Begin VB.ComboBox cboOnHDR 
         Height          =   315
         Left            =   2700
         Style           =   2  '單純下拉式
         TabIndex        =   36
         Top             =   1340
         Width           =   2200
      End
      Begin VB.TextBox txtOnVerify 
         Height          =   300
         Left            =   2700
         TabIndex        =   38
         Top             =   1820
         Width           =   900
      End
      Begin VB.CheckBox chkOnClean 
         Caption         =   "驗證失敗時清除虛擬顯示卡後重試"
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   2340
         Width           =   6000
      End
      Begin VB.TextBox txtOnDelayClean 
         Height          =   300
         Left            =   2700
         TabIndex        =   42
         Top             =   2780
         Width           =   900
      End
      Begin VB.TextBox txtOnShell 
         Height          =   300
         Left            =   2700
         TabIndex        =   44
         Top             =   3260
         Width           =   4600
      End
      Begin VB.Label lblOnTitle 
         Caption         =   "通電時動作（與斷電相同結構）"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   31
         Top             =   60
         Width           =   6000
      End
      Begin VB.Label lblOnDelayHDR 
         Caption         =   "HDR 動作前等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   900
         Width           =   2400
      End
      Begin VB.Label lblOnHDR 
         Caption         =   "原生 HDR 動作"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1380
         Width           =   2400
      End
      Begin VB.Label lblOnVerify 
         Caption         =   "驗證等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1860
         Width           =   2400
      End
      Begin VB.Label lblOnVerifyHint 
         Caption         =   "HDR 後等多久再檢查；0＝不驗證"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   3800
         TabIndex        =   39
         Top             =   1860
         Width           =   3600
      End
      Begin VB.Label lblOnDelayClean 
         Caption         =   "清卡後再試前等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2820
         Width           =   2400
      End
      Begin VB.Label lblOnShell 
         Caption         =   "額外 Shell 命令"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   3300
         Width           =   2400
      End
      Begin VB.Label lblOnShellHint 
         Caption         =   "空白或不存在則不執行；在 HDR／驗證／清卡流程全部結束後執行"
         ForeColor       =   &H00808080&
         Height          =   480
         Left            =   2700
         TabIndex        =   45
         Top             =   3660
         Width           =   4600
      End
   End
   Begin VB.PictureBox picLog 
      BorderStyle     =   0  '沒有框線
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4800
      ScaleWidth      =   7560
      TabIndex        =   46
      Top             =   540
      Visible         =   0   'False
      Width           =   7560
      Begin VB.TextBox txtLog 
         Height          =   4200
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   47
         Top             =   0
         Width           =   7560
      End
      Begin VB.CheckBox chkPause 
         Caption         =   "暫停自動更新"
         Height          =   300
         Left            =   0
         TabIndex        =   48
         Top             =   4320
         Width           =   1800
      End
      Begin VB.Timer tmrLog 
         Interval        =   500
         Left            =   2800
         Top             =   4320
      End
      Begin VB.Label lblLogHint 
         Caption         =   "僅顯示最新 200 行；完整檔最多保留 2000 行"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2000
         TabIndex        =   51
         Top             =   4350
         Width           =   5400
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "儲存"
      Height          =   400
      Left            =   5040
      TabIndex        =   49
      Top             =   5460
      Width           =   1200
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "隱藏"
      Height          =   400
      Left            =   6360
      TabIndex        =   50
      Top             =   5460
      Width           =   1200
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_tab As Integer
Private m_lastLog As String
Private m_pause As Boolean

Public Sub ShowTab(ByVal idx As Integer)
    m_tab = idx
    If Me.Visible Then ApplyTab
End Sub

Private Sub Command1_Click()
    Call CleanVirtualGpus
End Sub

Private Sub Form_Load()
    Call AppBootstrap
    Load frmTray
    
    Me.Caption = "AutoHdrMode 設定"
    btnTab0.Caption = "一般"
    btnTab1.Caption = "斷電"
    btnTab2.Caption = "通電"
    btnTab3.Caption = "記錄"
    btnSave.Caption = "儲存"
    btnClose.Caption = "隱藏"
    
    cboOffHDR.Clear
    cboOffHDR.AddItem "（不動作）"
    cboOffHDR.ItemData(0) = 0
    cboOffHDR.AddItem "關閉 HDR"
    cboOffHDR.ItemData(1) = 1
    cboOffHDR.AddItem "開啟 HDR"
    cboOffHDR.ItemData(2) = 2
    
    cboOnHDR.Clear
    cboOnHDR.AddItem "（不動作）"
    cboOnHDR.ItemData(0) = 0
    cboOnHDR.AddItem "開啟 HDR"
    cboOnHDR.ItemData(1) = 1
    cboOnHDR.AddItem "關閉 HDR"
    cboOnHDR.ItemData(2) = 2
    
    LoadToUI
    If m_tab < 0 Or m_tab > 3 Then m_tab = 0
    ApplyTab
    m_pause = False
    m_lastLog = ""
    tmrLog.Enabled = True
    
    If Not g_FirstRun Then
        Me.Hide
    End If
End Sub

Private Function HdrToIndex(ByVal cbo As ComboBox, ByVal v As String, ByVal isOn As Boolean) As Long
    v = LCase$(Trim$(v))
    If isOn Then
        If v = "on" Then
            HdrToIndex = 1
        ElseIf v = "off" Then
            HdrToIndex = 2
        Else
            HdrToIndex = 0
        End If
    Else
        If v = "off" Then
            HdrToIndex = 1
        ElseIf v = "on" Then
            HdrToIndex = 2
        Else
            HdrToIndex = 0
        End If
    End If
End Function

Private Function IndexToHdr(ByVal cbo As ComboBox, ByVal isOn As Boolean) As String
    Dim i As Long
    i = cbo.ListIndex
    If i < 0 Then i = 0
    If isOn Then
        Select Case i
            Case 1: IndexToHdr = "on"
            Case 2: IndexToHdr = "off"
            Case Else: IndexToHdr = ""
        End Select
    Else
        Select Case i
            Case 1: IndexToHdr = "off"
            Case 2: IndexToHdr = "on"
            Case Else: IndexToHdr = ""
        End Select
    End If
End Function

Private Sub LoadToUI()
    txtPoll.Text = CStr(g_PollSec)
    txtStable.Text = CStr(g_StableN)
    chkAuto.Value = IIf(g_AutoOn, vbChecked, vbUnchecked)
    chkAutostart.Value = IIf(AutostartInstalled(), vbChecked, vbUnchecked)
    chkBalloon.Value = IIf(g_Balloon, vbChecked, vbUnchecked)

    chkOffEn.Value = IIf(g_PowerOff.Enabled, vbChecked, vbUnchecked)
    txtOffDelayHDR.Text = CStr(g_PowerOff.DelayHDR)
    cboOffHDR.ListIndex = HdrToIndex(cboOffHDR, g_PowerOff.NativeHDR, False)
    txtOffVerify.Text = CStr(g_PowerOff.VerifySeconds)
    chkOffClean.Value = IIf(g_PowerOff.CleanHelper, vbChecked, vbUnchecked)
    txtOffDelayClean.Text = CStr(g_PowerOff.DelayAfterClean)
    txtOffShell.Text = g_PowerOff.Shell

    chkOnEn.Value = IIf(g_PowerOn.Enabled, vbChecked, vbUnchecked)
    txtOnDelayHDR.Text = CStr(g_PowerOn.DelayHDR)
    cboOnHDR.ListIndex = HdrToIndex(cboOnHDR, g_PowerOn.NativeHDR, True)
    txtOnVerify.Text = CStr(g_PowerOn.VerifySeconds)
    chkOnClean.Value = IIf(g_PowerOn.CleanHelper, vbChecked, vbUnchecked)
    txtOnDelayClean.Text = CStr(g_PowerOn.DelayAfterClean)
    txtOnShell.Text = g_PowerOn.Shell
End Sub

Private Sub UIToGlobals()
    g_PollSec = val(txtPoll.Text)
    If g_PollSec < 1 Then g_PollSec = 1
    If g_PollSec > 30 Then g_PollSec = 30
    g_StableN = val(txtStable.Text)
    If g_StableN < 1 Then g_StableN = 1
    g_AutoOn = (chkAuto.Value = vbChecked)
    g_Balloon = (chkBalloon.Value = vbChecked)

    g_PowerOff.Enabled = (chkOffEn.Value = vbChecked)
    g_PowerOff.DelayHDR = val(txtOffDelayHDR.Text)
    g_PowerOff.NativeHDR = IndexToHdr(cboOffHDR, False)
    g_PowerOff.VerifySeconds = val(txtOffVerify.Text)
    g_PowerOff.CleanHelper = (chkOffClean.Value = vbChecked)
    g_PowerOff.DelayAfterClean = val(txtOffDelayClean.Text)
    g_PowerOff.Shell = Trim$(txtOffShell.Text)

    g_PowerOn.Enabled = (chkOnEn.Value = vbChecked)
    g_PowerOn.DelayHDR = val(txtOnDelayHDR.Text)
    g_PowerOn.NativeHDR = IndexToHdr(cboOnHDR, True)
    g_PowerOn.VerifySeconds = val(txtOnVerify.Text)
    g_PowerOn.CleanHelper = (chkOnClean.Value = vbChecked)
    g_PowerOn.DelayAfterClean = val(txtOnDelayClean.Text)
    g_PowerOn.Shell = Trim$(txtOnShell.Text)
End Sub

Private Sub ApplyTab()
    picMain.Visible = (m_tab = 0)
    picOff.Visible = (m_tab = 1)
    picOn.Visible = (m_tab = 2)
    picLog.Visible = (m_tab = 3)
    ' 分頁外觀（不透過 ByRef CommandButton，避免 VB6 控制項參數問題）
    btnTab0.FontBold = (m_tab = 0)
    btnTab1.FontBold = (m_tab = 1)
    btnTab2.FontBold = (m_tab = 2)
    btnTab3.FontBold = (m_tab = 3)
    btnTab0.BackColor = IIf(m_tab = 0, &H800000, &H0)
    btnTab1.BackColor = IIf(m_tab = 1, &H800000, &H0)
    btnTab2.BackColor = IIf(m_tab = 2, &H800000, &H0)
    btnTab3.BackColor = IIf(m_tab = 3, &H800000, &H0)
    tmrLog.Enabled = (m_tab = 3 And Me.Visible)
    If m_tab = 3 Then RefreshLog
End Sub

Private Sub btnTab0_Click()
    m_tab = 0: ApplyTab
End Sub
Private Sub btnTab1_Click()
    m_tab = 1: ApplyTab
End Sub
Private Sub btnTab2_Click()
    m_tab = 2: ApplyTab
End Sub
Private Sub btnTab3_Click()
    m_tab = 3: ApplyTab
End Sub

Private Sub btnSave_Click()
    On Error Resume Next
    UIToGlobals
    ConfigSave
    If chkAutostart.Value = vbChecked Then
        Call AutostartSet(True)
    Else
        Call AutostartSet(False)
    End If
    frmTray.tmrPoll.Interval = g_PollSec * 1000
    frmTray.mnuAuto.Checked = g_AutoOn
    frmTray.mnuBalloon.Checked = g_Balloon
    LogMsg S_LogSettingsSaved()
    ' 儲存後不關閉視窗，方便繼續調整
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        tmrLog.Enabled = False
        Exit Sub
    End If
    tmrLog.Enabled = False
End Sub

Private Sub chkPause_Click()
    m_pause = (chkPause.Value = vbChecked)
End Sub

Private Sub tmrLog_Timer()
    If Not picLog.Visible Or m_pause Then Exit Sub
    RefreshLog
End Sub

Private Sub RefreshLog()
    On Error GoTo Fail
    Dim fn As Integer, line As String, buf As String
    Dim lines() As String, i As Long, start As Long, n As Long
    If g_LogFile = "" Then Exit Sub
    If dir$(g_LogFile) = "" Then Exit Sub
    fn = FreeFile
    Open g_LogFile For Input As #fn
    buf = ""
    Do While Not EOF(fn)
        Line Input #fn, line
        buf = buf & line & vbCrLf
    Loop
    Close #fn
    lines = Split(buf, vbCrLf)
    n = UBound(lines) - LBound(lines) + 1
    If n > 200 Then
        start = UBound(lines) - 199
        buf = ""
        For i = start To UBound(lines)
            If i > start Then buf = buf & vbCrLf
            buf = buf & lines(i)
        Next
    End If
    If buf <> m_lastLog Then
        m_lastLog = buf
        txtLog.Text = buf
        txtLog.SelStart = Len(txtLog.Text)
    End If
    Exit Sub
Fail:
    On Error Resume Next
    Close #fn
End Sub
