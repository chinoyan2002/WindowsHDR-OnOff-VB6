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
' ---- 一般頁：輪詢與自動 ----
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
      Begin VB.Label lblFlow 
         Caption         =   "流程（兩方向同結構）：偵測（輪詢秒×穩定次）→ 等DelayHDR → 開/關HDR → 等VerifySeconds驗收 → 清卡＋沉澱（失敗且開啟時）→ 等DelayShell → 跑Shell（等ShellTimeout）｜每段等待都從上段做完起算"
         Height          =   1600
         Left            =   120
         TabIndex        =   64
         Top             =   3050
         Width           =   7200
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
' ---- 斷電頁：螢幕關的整套動作 ----
   Begin VB.PictureBox picOff 
      BorderStyle     =   0  '沒有框線
      Height          =   4900
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
      Begin VB.Label lblOffDelayShell 
         Caption         =   "Shell 執行前等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   4200
         Width           =   2400
      End
      Begin VB.TextBox txtOffDelayShell 
         Height          =   300
         Left            =   2700
         TabIndex        =   53
         Top             =   4180
         Width           =   900
      End
      Begin VB.Label lblOffDelayShellHint 
         Caption         =   "HDR/驗證/清卡跑完、Shell前再等；0=不等"
         Height          =   255
         Left            =   3800
         TabIndex        =   54
         Top             =   4200
         Width           =   3400
      End
      Begin VB.Label lblOffShellWait 
         Caption         =   "Shell 等待秒數"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   4580
         Width           =   2400
      End
      Begin VB.TextBox txtOffShellWait 
         Height          =   300
         Left            =   2700
         TabIndex        =   56
         Top             =   4560
         Width           =   900
      End
      Begin VB.Label lblOffShellWaitHint 
         Caption         =   "0=啟動即走，預設120秒"
         Height          =   255
         Left            =   3800
         TabIndex        =   57
         Top             =   4580
         Width           =   3400
      End
   End
' ---- 通電頁：螢幕開的整套動作（與斷電頁鏡像） ----
   Begin VB.PictureBox picOn 
      BorderStyle     =   0  '沒有框線
      Height          =   4900
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
      Begin VB.Label lblOnDelayShell 
         Caption         =   "Shell 執行前等待（秒）"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   4200
         Width           =   2400
      End
      Begin VB.TextBox txtOnDelayShell 
         Height          =   300
         Left            =   2700
         TabIndex        =   59
         Top             =   4180
         Width           =   900
      End
      Begin VB.Label lblOnDelayShellHint 
         Caption         =   "HDR/驗證/清卡跑完、Shell前再等；0=不等"
         Height          =   255
         Left            =   3800
         TabIndex        =   60
         Top             =   4200
         Width           =   3400
      End
      Begin VB.Label lblOnShellWait 
         Caption         =   "Shell 等待秒數"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   4580
         Width           =   2400
      End
      Begin VB.TextBox txtOnShellWait 
         Height          =   300
         Left            =   2700
         TabIndex        =   62
         Top             =   4560
         Width           =   900
      End
      Begin VB.Label lblOnShellWaitHint 
         Caption         =   "0=啟動即走，預設120秒"
         Height          =   255
         Left            =   3800
         TabIndex        =   63
         Top             =   4580
         Width           =   3400
      End
   End
' ---- 記錄頁：讀記錄尾 ----
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
' 本窗：設定 UI（Startup 進入點）。四分頁＋記錄尾讀；X 只隱藏

Private m_tab As Integer ' 0一般 1斷電 2通電 3記錄
Private m_lastLog As String ' 上次顯示內容：相同不重繪
Private m_pause As Boolean ' 記錄頁暫停捲動

' 用途：切分頁；參數 idx：0一般 1斷電 2通電 3記錄
Public Sub ShowTab(ByVal idx As Integer)
    m_tab = idx
    If Me.Visible Then ApplyTab
End Sub

' 用途：隱藏測試鈕：直接清虛擬卡（除錯用）
Private Sub Command1_Click()
    Call CleanVirtualGpus
End Sub

' 用途：程式進入點 Startup：跑 AppBootstrap、載托盤窗、套 UI；首次執行才顯示
Private Sub Form_Load()
    Call AppBootstrap ' CLI 在此分流：帶參數做完就走，不進 GUI
    Load frmTray ' 載常駐本體：只取值不用顯示
    
        Me.Caption = AppVersionLine()
    btnTab0.Caption = "一般"
    btnTab1.Caption = "斷電"
    btnTab2.Caption = "通電"
    btnTab3.Caption = "記錄"
    btnSave.Caption = "儲存"
    btnClose.Caption = "隱藏"
    
    cboOffHDR.Clear
    cboOffHDR.AddItem "（不動作）"
    cboOffHDR.ItemData(0) = 0 ' 斷電頁：0略過 1關HDR 2開HDR
    cboOffHDR.AddItem "關閉 HDR"
    cboOffHDR.ItemData(1) = 1
    cboOffHDR.AddItem "開啟 HDR"
    cboOffHDR.ItemData(2) = 2
    
    cboOnHDR.Clear
    cboOnHDR.AddItem "（不動作）"
    cboOnHDR.ItemData(0) = 0 ' 通電頁：0略過 1開HDR 2關HDR（與斷電鏡像）
    cboOnHDR.AddItem "開啟 HDR"
    cboOnHDR.ItemData(1) = 1
    cboOnHDR.AddItem "關閉 HDR"
    cboOnHDR.ItemData(2) = 2
    
    LoadToUI
    If m_tab < 0 Or m_tab > 3 Then m_tab = 0 ' 範圍保護：越界回一般頁
    ApplyTab
    m_pause = False ' 預設跟著捲
    m_lastLog = ""
    tmrLog.Enabled = True ' 記錄節拍先開，ApplyTab 會再管一次
    
    If Not g_FirstRun Then ' 非首次：直接藏，只留托盤
        Me.Hide
    End If
End Sub

' 用途：NativeHDR 字串轉下拉索引；回傳：0略過 1開 2關
Private Function HdrToIndex(ByVal cbo As ComboBox, ByVal v As String, ByVal isOn As Boolean) As Long
    v = LCase$(Trim$(v)) ' 正規化：去空白轉小寫再比
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

' 用途：下拉索引轉 NativeHDR 字串
Private Function IndexToHdr(ByVal cbo As ComboBox, ByVal isOn As Boolean) As String
    Dim i As Long
    i = cbo.ListIndex
    If i < 0 Then i = 0 ' 未選視為略過
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

' 用途：全域設定灌入控制項
Private Sub LoadToUI()
    txtPoll.Text = CStr(g_PollSec) ' 讀出：輪詢秒數
    txtStable.Text = CStr(g_StableN) ' 讀出：穩定次數
    chkAuto.Value = IIf(g_AutoOn, vbChecked, vbUnchecked) ' 讀出：自動總開關
    chkAutostart.Value = IIf(AutostartInstalled(), vbChecked, vbUnchecked) ' 讀出：機碼有無（非記憶體值）
    chkBalloon.Value = IIf(g_Balloon, vbChecked, vbUnchecked) ' 讀出：氣球開關

    ' ---- 斷電整組讀出 ----
    chkOffEn.Value = IIf(g_PowerOff.Enabled, vbChecked, vbUnchecked) ' 讀出：斷電啟用
    txtOffDelayHDR.Text = CStr(g_PowerOff.DelayHDR) ' 讀出：斷電延遲秒數
    cboOffHDR.ListIndex = HdrToIndex(cboOffHDR, g_PowerOff.NativeHDR, False) ' 讀出：斷電 HDR 選項轉索引
    txtOffVerify.Text = CStr(g_PowerOff.VerifySeconds) ' 讀出：斷電驗證秒數
    chkOffClean.Value = IIf(g_PowerOff.CleanHelper, vbChecked, vbUnchecked) ' 讀出：斷電清卡
    txtOffDelayClean.Text = CStr(g_PowerOff.DelayAfterClean) ' 讀出：斷電清後等待
    txtOffShell.Text = g_PowerOff.Shell ' 讀出：斷電 Shell
    txtOffDelayShell.Text = CStr(g_PowerOff.DelayShell) ' 讀出：驗收後等待
    txtOffShellWait.Text = CStr(g_PowerOff.ShellTimeout) ' 讀出：Shell 等待

    ' ---- 通電整組讀出（鏡像） ----
    chkOnEn.Value = IIf(g_PowerOn.Enabled, vbChecked, vbUnchecked) ' 讀出：通電啟用
    txtOnDelayHDR.Text = CStr(g_PowerOn.DelayHDR) ' 讀出：通電延遲秒數
    cboOnHDR.ListIndex = HdrToIndex(cboOnHDR, g_PowerOn.NativeHDR, True) ' 讀出：通電 HDR 選項轉索引
    txtOnVerify.Text = CStr(g_PowerOn.VerifySeconds) ' 讀出：通電驗證秒數
    chkOnClean.Value = IIf(g_PowerOn.CleanHelper, vbChecked, vbUnchecked) ' 讀出：通電清卡
    txtOnDelayClean.Text = CStr(g_PowerOn.DelayAfterClean) ' 讀出：通電清後等待
    txtOnShell.Text = g_PowerOn.Shell ' 讀出：通電 Shell
    txtOnDelayShell.Text = CStr(g_PowerOn.DelayShell) ' 讀出：驗收後等待
    txtOnShellWait.Text = CStr(g_PowerOn.ShellTimeout) ' 讀出：Shell 等待
End Sub

' 用途：控制項寫回全域（含範圍夾限）
Private Sub UIToGlobals()
    g_PollSec = val(txtPoll.Text) ' 非數字 val 給 0，下面夾回 1
    If g_PollSec < 1 Then g_PollSec = 1
    If g_PollSec > 30 Then g_PollSec = 30 ' Timer 上限 32767 毫秒，30 秒封頂
    g_StableN = val(txtStable.Text) ' 寫入：穩定次數（下行夾限）
    If g_StableN < 1 Then g_StableN = 1 ' 至少一次才有意義
    g_AutoOn = (chkAuto.Value = vbChecked) ' 寫入：自動總開關
    g_Balloon = (chkBalloon.Value = vbChecked) ' 寫入：氣球開關

    ' ---- 斷電整組 ----
    g_PowerOff.Enabled = (chkOffEn.Value = vbChecked) ' 寫入：斷電啟用
    g_PowerOff.DelayHDR = val(txtOffDelayHDR.Text) ' 寫入：斷電延遲秒數
    g_PowerOff.NativeHDR = IndexToHdr(cboOffHDR, False) ' 寫入：斷電 HDR 選項轉字串
    g_PowerOff.VerifySeconds = val(txtOffVerify.Text) ' 寫入：斷電驗證秒數
    g_PowerOff.CleanHelper = (chkOffClean.Value = vbChecked) ' 寫入：斷電清卡
    g_PowerOff.DelayAfterClean = val(txtOffDelayClean.Text) ' 寫入：斷電清後等待
    g_PowerOff.Shell = Trim$(txtOffShell.Text) ' 寫入：斷電 Shell（去空白）
    g_PowerOff.DelayShell = val(txtOffDelayShell.Text) ' 寫入：驗收後等待（夾0~120）
    If g_PowerOff.DelayShell < 0 Then g_PowerOff.DelayShell = 0
    If g_PowerOff.DelayShell > 120 Then g_PowerOff.DelayShell = 120
    g_PowerOff.ShellTimeout = val(txtOffShellWait.Text) ' 寫入：Shell 等待（夾0~3600）
    If g_PowerOff.ShellTimeout < 0 Then g_PowerOff.ShellTimeout = 0
    If g_PowerOff.ShellTimeout > 3600 Then g_PowerOff.ShellTimeout = 3600

    ' ---- 通電整組（鏡像） ----
    g_PowerOn.Enabled = (chkOnEn.Value = vbChecked) ' 寫入：通電啟用
    g_PowerOn.DelayHDR = val(txtOnDelayHDR.Text) ' 寫入：通電延遲秒數
    g_PowerOn.NativeHDR = IndexToHdr(cboOnHDR, True) ' 寫入：通電 HDR 選項轉字串
    g_PowerOn.VerifySeconds = val(txtOnVerify.Text) ' 寫入：通電驗證秒數
    g_PowerOn.CleanHelper = (chkOnClean.Value = vbChecked) ' 寫入：通電清卡
    g_PowerOn.DelayAfterClean = val(txtOnDelayClean.Text) ' 寫入：通電清後等待
    g_PowerOn.Shell = Trim$(txtOnShell.Text) ' 寫入：通電 Shell（去空白）
    g_PowerOn.DelayShell = val(txtOnDelayShell.Text) ' 寫入：驗收後等待（夾0~120）
    If g_PowerOn.DelayShell < 0 Then g_PowerOn.DelayShell = 0
    If g_PowerOn.DelayShell > 120 Then g_PowerOn.DelayShell = 120
    g_PowerOn.ShellTimeout = val(txtOnShellWait.Text) ' 寫入：Shell 等待（夾0~3600）
    If g_PowerOn.ShellTimeout < 0 Then g_PowerOn.ShellTimeout = 0
    If g_PowerOn.ShellTimeout > 3600 Then g_PowerOn.ShellTimeout = 3600
End Sub

' 用途：依 m_tab 顯示對應頁；記錄頁順手重讀
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
    tmrLog.Enabled = (m_tab = 3 And Me.Visible) ' 只有記錄頁現身才讀檔
    If m_tab = 3 Then RefreshLog ' 切過去立刻刷一次
End Sub

' 用途：分頁鈕：切一般頁
Private Sub btnTab0_Click()
    m_tab = 0: ApplyTab ' 切一般頁再套版
End Sub
' 用途：分頁鈕：切斷電頁
Private Sub btnTab1_Click()
    m_tab = 1: ApplyTab ' 切斷電頁再套版
End Sub
' 用途：分頁鈕：切通電頁
Private Sub btnTab2_Click()
    m_tab = 2: ApplyTab ' 切通電頁再套版
End Sub
' 用途：分頁鈕：切記錄頁
Private Sub btnTab3_Click()
    m_tab = 3: ApplyTab ' 切記錄頁再套版
End Sub

' 用途：存檔：UI→全域→寫 INI→套自啟→同步托盤；存完不關窗
Private Sub btnSave_Click()
    On Error Resume Next
    UIToGlobals
    ConfigSave
    If chkAutostart.Value = vbChecked Then ' 自啟勾選與機碼同步寫入或刪除
        Call AutostartSet(True)
    Else
        Call AutostartSet(False)
    End If
    frmTray.tmrPoll.Interval = g_PollSec * 1000 ' 間隔立即生效，不用重開
    frmTray.mnuAuto.Checked = g_AutoOn
    frmTray.mnuBalloon.Checked = g_Balloon
    LogMsg S_LogSettingsSaved()
    ' 儲存後不關閉視窗，方便繼續調整
End Sub

' 用途：關窗鈕只隱藏（常駐繼續）
Private Sub btnClose_Click()
    Me.Hide
End Sub

' 用途：點 X 只隱藏，擋掉真正關閉
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then ' 只有點 X 才擋；程式結束的 Unload 放行
        Cancel = True
        Me.Hide
        tmrLog.Enabled = False ' 藏窗就停讀，省磁碟
        Exit Sub
    End If
    tmrLog.Enabled = False ' 結束卸載：停讀記錄
End Sub

' 用途：記錄頁暫停或繼續自動捲動
Private Sub chkPause_Click()
    m_pause = (chkPause.Value = vbChecked)
End Sub

' 用途：記錄頁節拍：可見且未暫停時重讀記錄尾
Private Sub tmrLog_Timer()
    If Not picLog.Visible Or m_pause Then Exit Sub ' 頁面藏起或暫停：跳過
    RefreshLog
End Sub

' 用途：讀記錄尾 200 行；內容相同不重繪（防閃爍）
Private Sub RefreshLog()
    On Error GoTo Fail
    Dim fn As Integer, line As String, buf As String
    Dim lines() As String, i As Long, start As Long, n As Long
    If g_LogFile = "" Then Exit Sub ' 路徑空：無檔可讀
    If dir$(g_LogFile) = "" Then Exit Sub ' 檔不存在：等下一輪
    fn = FreeFile
    Open g_LogFile For Input As #fn
    buf = ""
    Do While Not EOF(fn)
        Line Input #fn, line
        buf = buf & line & vbCrLf ' 逐行串起，保留換行
    Loop
    Close #fn
    lines = Split(buf, vbCrLf)
    n = UBound(lines) - LBound(lines) + 1
    If n > 200 Then
        start = UBound(lines) - 199 ' 超 200 行：只留尾段
        buf = ""
        For i = start To UBound(lines)
            If i > start Then buf = buf & vbCrLf
            buf = buf & lines(i)
        Next
    End If
    If buf <> m_lastLog Then
        m_lastLog = buf ' 記住本次：下次相同跳過
        txtLog.Text = buf
        txtLog.SelStart = Len(txtLog.Text) ' 游標跳尾：自動跟最新
    End If
    Exit Sub
Fail:
    On Error Resume Next ' 讀到一半出錯也要關檔，防鎖死
    Close #fn
End Sub
