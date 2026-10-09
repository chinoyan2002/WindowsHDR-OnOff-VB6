VERSION 5.00
Begin VB.Form frmLog
   Caption         =   "Log"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "frmLog"
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   StartUpPosition =   3   'Windows Default
   Begin VB.TextBox txtLog
      Height          =   5415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8760
   End
   Begin VB.CheckBox chkPause
      Caption         =   "Pause"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5625
      Width           =   2000
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   5625
      Width           =   1320
   End
   Begin VB.Timer tmrRefresh
      Interval        =   1000
      Left            =   2400
      Top             =   5625
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 本窗體：即時記錄檢視（唯讀，不影響輪詢）
' 單例由托盤選單開啟；關閉只卸載本窗體
Private Sub Form_Load()
    Me.Caption = S_LogWinCap()
    chkPause.Caption = S_PauseCap()
    cmdClose.Caption = S_CloseCap()
    tmrRefresh.Interval = 1000
    tmrRefresh.Enabled = True
    RefreshLog
End Sub
Private Sub tmrRefresh_Timer()
    If chkPause.Value = vbChecked Then Exit Sub
    RefreshLog
End Sub
' 讀尾 200 行，有變才更新並捲到底
Private Sub RefreshLog()
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
    If out <> txtLog.Text Then
        txtLog.Text = out
        txtLog.SelStart = Len(txtLog.Text)
    End If
End Sub
Private Sub cmdClose_Click()
    tmrRefresh.Enabled = False
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmrRefresh.Enabled = False
End Sub
