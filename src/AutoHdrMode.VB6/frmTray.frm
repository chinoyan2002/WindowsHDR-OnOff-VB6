VERSION 5.00
Begin VB.Form frmTray 
   Caption         =   "AutoHdrMode"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1200
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   60
   ScaleMode       =   3  '像素
   ScaleWidth      =   80
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
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
      Begin VB.Menu mnuBalloon 
         Caption         =   "Balloon"
      End
      Begin VB.Menu mnuFireOff 
         Caption         =   "FireOff"
      End
      Begin VB.Menu mnuFireOn 
         Caption         =   "FireOn"
      End
      Begin VB.Menu mnuSepHdr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHdrOn 
         Caption         =   "HdrOn"
      End
      Begin VB.Menu mnuHdrOff 
         Caption         =   "HdrOff"
      End
      Begin VB.Menu mnuClean 
         Caption         =   "Clean"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "LogFolder"
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
' 本窗：常駐本體。托盤圖示＋輪詢＋通電斷電狀態機；唯一出口 mnuExit

Private Type NOTIFYICONDATA
    cbSize As Long          ' 結構大小，API 對版用
    hwnd As Long            ' 回呼視窗（本窗代碼）
    uId As Long             ' 圖示 ID，固定值防衝突
    uFlags As Long          ' 哪些欄位有效
    uCallBackMessage As Long ' 滑鼠訊息轉哪個訊息號
    hIcon As Long           ' 托盤圖示代碼
    szTip As String * 128   ' 懸停文字，最長 127 加結尾
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256  ' 氣球內文，最長 255 加結尾
    uTimeoutOrVersion As Long ' 氣球停留秒數
    szInfoTitle As String * 64 ' 氣球標題，最長 63 加結尾
    dwInfoFlags As Long     ' 氣球圖示種類
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String        ' 動詞，runas 即提權執行
    lpFile As String
    lpParameters As String  ' 參數，此處固定 --clean
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long        ' 子行程代碼（等結束用）
End Type

Private Declare Function Shell_NotifyIconA Lib "SHELL32.DLL" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInst As Long, ByVal lpName As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOp As String, ByVal lpFile As String, ByVal lpParam As String, ByVal lpDir As String, ByVal nShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (sei As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMs As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpCode As Long) As Long

Private Const NIF_ICON As Long = &H2
Private Const NIF_INFO As Long = &H10 ' 氣球欄位有效旗標
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_TIP As Long = &H4
Private Const NIIF_INFO = &H1
Private Const NIM_ADD As Long = &H0 ' 新增圖示
Private Const NIM_MODIFY As Long = &H1 ' 修改圖示或提示或氣球
Private Const NIM_DELETE As Long = &H2 ' 刪除圖示
Private Const WM_MOUSEMOVE = &H200 ' 托盤回呼一律走此訊息號
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_NULL = &H0
Private Const IDI_APPLICATION = 32512&
Private Const SW_SHOWNORMAL = 1
Private Const SEE_MASK_NOCLOSEPROCESS = &H40 ' 要回子行程代碼才等得到
Private Const WAIT_OBJECT_0 = 0
Private Const SW_HIDE = 0

Private Const PH_IDLE As Integer = 0 ' 閒置：無排程
Private Const PH_WAIT_HDR As Integer = 1 ' 等 DelayHDR 做 HDR
Private Const PH_WAIT_VERIFY As Integer = 2 ' 等 VerifySeconds 查狀態
Private Const PH_WAIT_AFTER_CLEAN As Integer = 3 ' 等 DelayAfterClean 重做 HDR
Private Const PH_WAIT_SHELL As Integer = 4 ' Shell 執行前等待中

Private m_lastSeen As Integer ' 上次穩定狀態（-1=未知）
Private m_candidate As Integer ' 候選狀態（去抖中）
Private m_stable As Long ' 候選連續相同次數
Private m_pending As Integer ' 排程目標（-1=無排程）
Private m_pendingDue As Long ' 排程到期 tick
Private m_phase As Integer ' 排程階段（PH_*）
Private m_lastErr As String
Private m_nid As NOTIFYICONDATA ' 托盤結構本體
Private m_trayAdded As Boolean ' 圖示已註冊旗標
Private m_holdLogged As Boolean ' HOLD 已記過，防每輪洗版

' 用途：托盤窗啟動：建圖示、套選單文字、啟動輪詢並做第一次偵測
Private Sub Form_Load()
    Me.ScaleMode = vbPixels                                                     ' 像素座標：托盤滑鼠訊息直接讀 X
    Me.Move -32000, -32000                                                      ' 藏到螢幕外：要 hwnd 不要畫面
    Me.Show
    Me.Refresh
    
    mnuTray.Visible = False                                                     ' 根選單只當容器，不顯示
    mnuAuto.Caption = S_MenuAuto()
    mnuBalloon.Caption = S_MenuBalloon()
    mnuFireOff.Caption = S_MenuFireOff()
    mnuFireOn.Caption = S_MenuFireOn()
    mnuHdrOn.Caption = S_MenuHdrOn()
    mnuHdrOff.Caption = S_MenuHdrOff()
    mnuClean.Caption = S_MenuClean()
    mnuSettings.Caption = S_MenuSettings()
    mnuLog.Caption = S_MenuOpenLog()
    mnuExit.Caption = S_MenuExit()
    mnuAuto.Checked = g_AutoOn
    mnuBalloon.Checked = g_Balloon
    tmrPoll.Interval = g_PollSec * 1000                                         ' 秒轉毫秒
    tmrWait.Enabled = False                                                     ' 排程節拍關閉
    m_lastSeen = -1                                                             ' 未知初值：首輪只記錄不動作
    m_candidate = -1
    m_stable = 0
    m_pending = -1                                                              ' 無排程
    m_phase = PH_IDLE
    m_lastErr = ""
    LogMsg S_LogStart(g_PollSec, g_StableN, g_AutoOn)
    LogMsg AppVersionLine()
    
    Call InitTray(S_TipStarting())
    Me.Hide
    
    PollOnce
End Sub

' 用途：布林轉 0/1；回傳：開=1、關=0
Private Function PresentInt(ByVal present As Boolean) As Integer
    If present Then PresentInt = 1 Else PresentInt = 0
End Function

' 用途：每輪偵測一次；三叉：開/待命HOLD/關
' 注意：HOLD 不碰候選值與穩定計數，離開後正常判定
Private Sub PollOnce()
    On Error GoTo Fail
    Dim n As Long, sb As Long
    Dim d6s As String
    n = DdcOnCount() ' 通電台數（內含重新列舉）
    sb = DdcStandbyCount() ' 待命台數（讀上次列舉快取）
    d6s = DdcD6Summary() ' 本輪各台 D6 原始值，事件記錄用
    If n = 0 And sb > 0 Then
        TrayTip S_TipStandby()
        If Not m_holdLogged Then
            m_holdLogged = True
            LogMsg S_LogHoldStandby(n, sb, d6s)
            Call TrayBalloon(S_TipStandby(), S_LogHoldStandby(n, sb, d6s))
        End If
        Exit Sub ' 凍結：不碰候選與計數
    End If
    m_holdLogged = False ' 離開待命，重置進出旗標
    Dim present As Boolean
    present = (n > 0) ' 有一台通電即視為開
    If m_candidate <> PresentInt(present) Then
        m_candidate = PresentInt(present)
        m_stable = 1 ' 換候選：重數
    Else
        m_stable = m_stable + 1 ' 相同：累加
    End If
    TrayTip S_TipState(present)
    If m_stable < g_StableN Then Exit Sub ' 未達穩定次數：繼續等
    If m_lastSeen = -1 Then
        m_lastSeen = PresentInt(present)
        LogMsg S_LogInit(present, d6s, DisplayName0())
        Exit Sub
    End If
    If m_lastSeen = PresentInt(present) Then Exit Sub ' 無翻轉：收工
    m_lastSeen = PresentInt(present)
    LogMsg S_LogEvent(present, d6s, DisplayName0())
    Call TrayBalloon(S_TipState(present), S_LogEvent(present, d6s, DisplayName0()))
    If Not g_AutoOn Then Exit Sub ' 自動關閉：只更新狀態提示
    If present Then
        If Not g_PowerOn.Enabled Then LogMsg S_LogDirDisabled(True): Exit Sub
    Else
        If Not g_PowerOff.Enabled Then LogMsg S_LogDirDisabled(False): Exit Sub
    End If
    StartTransition present
    Exit Sub
Fail:
    LogMsg S_LogPollErr(Err.Description)
End Sub

' 開始一輪：先等 DelayHDR
Private Sub StartTransition(ByVal wantOn As Boolean)
    Dim cfg As TransCfg
    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff
    If m_phase <> PH_IDLE And m_pending <> -1 Then LogMsg S_LogCancel(m_pending = 1) ' 舊排程被新狀態蓋掉：記取消再重排
    m_pending = PresentInt(wantOn)
    m_pendingDue = GetTickCount() + cfg.DelayHDR * 1000 ' 到期 tick：現在加延遲毫秒
    m_phase = PH_WAIT_HDR
    tmrWait.Enabled = True ' 開排程節拍
    LogMsg S_LogPlanHdr(wantOn, cfg.DelayHDR)
End Sub

' 用途：輪詢節拍，回呼 PollOnce
Private Sub tmrPoll_Timer()
    PollOnce
End Sub

' 用途：排程時間是否已到（相減比大小，天然防 GetTickCount 溢位）；回傳：到期=True
Private Function TickPassed(ByVal due As Long) As Boolean
    TickPassed = (GetTickCount() - due >= 0)
End Function

' 用途：排程節拍；中途狀態翻回來則取消排程
Private Sub tmrWait_Timer()
    If m_pending = -1 Or m_phase = PH_IDLE Then ' 無排程：關節拍收工
        tmrWait.Enabled = False ' 排程節拍關閉
        Exit Sub
    End If
    If Not TickPassed(m_pendingDue) Then Exit Sub ' 未到期：等下一拍

    Dim wantOn As Boolean
    wantOn = (m_pending = 1) ' 還原排程目標方向
    If m_lastSeen <> PresentInt(wantOn) Then
        LogMsg S_LogSkip(wantOn) ' 中途翻回來：取消排程
        FinishPhase
        Exit Sub
    End If

    Select Case m_phase
        Case PH_WAIT_HDR
            Call DoHdrStep(wantOn)
        Case PH_WAIT_VERIFY
            Call DoVerifyStep(wantOn)
        Case PH_WAIT_AFTER_CLEAN
            Call DoRetryHdrAfterClean(wantOn)
        Case PH_WAIT_SHELL
            Call DoShellStep(wantOn) ' Shell 前等待到期：跑 Shell
    End Select
End Sub

' HDR 段：有設定才做；通電可進入驗證
' 通電／斷電同一套：HDR 動作 → 可選驗證 → 失敗可清卡再試 → Shell
Private Sub DoHdrStep(ByVal wantOn As Boolean)
    On Error GoTo Fail
    Dim cfg As TransCfg
    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff
    
    If cfg.NativeHDR = "on" Or cfg.NativeHDR = "off" Then
        Dim ok As Boolean, st As HDR_STATUS
        If cfg.NativeHDR = "on" Then
            LogMsg S_LogNativeTryOn()
            ok = HDR_Enable() ' 鏈內：設定要開就開
        Else
            LogMsg S_LogNativeTryOff()
            ok = HDR_Disable() ' 鏈內：設定要關就關
        End If
        st = HDR_GetStatus(0) ' 讀 0 號螢幕當代表
        LogMsg S_LogNativeResult(ok, st)
        If cfg.NativeHDR = "off" And (st = HDR_UNSUPPORTED Or st = HDR_ERROR Or st = HDR_STATUS_UNKNOWN) Then ' 斷電無路徑：記一筆免誤會
            LogMsg S_LogNativeOffNoPath()
        End If
        ' 兩邊都可驗證（VerifySeconds>0）
        If cfg.VerifySeconds > 0 Then
            LogMsg S_LogVerifyWait(cfg.VerifySeconds)
            m_pendingDue = GetTickCount() + cfg.VerifySeconds * 1000 ' 轉驗證等待
            m_phase = PH_WAIT_VERIFY
            Exit Sub
        End If
    Else
        LogMsg S_LogNativeSkip()
    End If
    
    Call DoShellStep(wantOn)
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FinishPhase
End Sub

' 用途：驗證是否達標；on 鏈要 HDR_ON、off 鏈要 HDR_OFF；回傳：達標=True
Private Function HdrVerifyOk(ByRef cfg As TransCfg, ByVal st As HDR_STATUS) As Boolean
    ' on → 必須 ON；off → 必須 OFF（UNSUPPORTED 不算通過，可走清卡重試）
    If cfg.NativeHDR = "on" Then
        HdrVerifyOk = (st = HDR_ON)
    ElseIf cfg.NativeHDR = "off" Then
        HdrVerifyOk = (st = HDR_OFF)
    Else
        HdrVerifyOk = True
    End If
End Function

' 用途：驗證等候到期後讀狀態；失敗且 CleanHelper=1 則提權清卡
Private Sub DoVerifyStep(ByVal wantOn As Boolean)
    On Error GoTo Fail
    Dim cfg As TransCfg, st As HDR_STATUS
    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff
    st = HDR_GetStatus(0)
    If HdrVerifyOk(cfg, st) Then
        LogMsg S_LogVerifyOk()
        Call DoShellStep(wantOn)
        Exit Sub
    End If
    LogMsg S_LogVerifyFail()
    If cfg.CleanHelper Then
        Dim crc As Long
        crc = RunCleanElevated() ' 提權起自己清卡並等結果
        LogMsg S_LogCleanDone(crc)
        If cfg.DelayAfterClean > 0 Then
            LogMsg S_LogWaitAfterClean(cfg.DelayAfterClean)
            m_pendingDue = GetTickCount() + cfg.DelayAfterClean * 1000 ' 清後等裝置沉澱再重做
            m_phase = PH_WAIT_AFTER_CLEAN
            Exit Sub
        Else
            Call DoRetryHdrAfterClean(wantOn)
            Exit Sub
        End If
    End If
    Call DoShellStep(wantOn)
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FinishPhase
End Sub

' 用途：清卡等候到期後重做一次 HDR 開關，再走 Shell
Private Sub DoRetryHdrAfterClean(ByVal wantOn As Boolean)
    On Error GoTo Fail
    Dim cfg As TransCfg, ok As Boolean
    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff
    LogMsg S_LogRetryEnable()
    If cfg.NativeHDR = "on" Then
        ok = HDR_Enable() ' 清卡後重試：再開一次
    ElseIf cfg.NativeHDR = "off" Then
        ok = HDR_Disable() ' 清卡後重試：再關一次
    Else
        ok = False
    End If
    LogMsg S_LogNativeResult(ok, HDR_GetStatus(0))
    Call DoShellStep(wantOn)
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FinishPhase
End Sub

' 用途：排程最後一步：Shell 空白或目標不存在就跳過，否則隱藏執行；收尾 FinishPhase
Private Sub DoShellStep(ByVal wantOn As Boolean)
    On Error GoTo Fail
    Dim cfg As TransCfg, rc As Long
    If wantOn Then cfg = g_PowerOn Else cfg = g_PowerOff
    If cfg.DelayShell > 0 And m_phase <> PH_WAIT_SHELL Then
        LogMsg S_LogWaitShell(cfg.DelayShell)
        m_pendingDue = GetTickCount() + cfg.DelayShell * 1000 ' 碼表重按：Shell 執行前等待
        m_phase = PH_WAIT_SHELL
        Exit Sub
    End If
    If Len(Trim$(cfg.Shell)) = 0 Then ' Shell 空白：跳過不算錯
        LogMsg S_LogShellSkipEmpty()
    ElseIf Not ShellTargetExists(cfg.Shell) Then
        LogMsg S_LogShellSkipMissing(cfg.Shell)
    Else
        rc = ShellRunHidden(cfg.Shell, g_WorkDir, cfg.ShellTimeout) ' 工作目錄跑，結束碼記檔
        If cfg.ShellTimeout <= 0 Then LogMsg S_LogShellLaunch() Else LogMsg S_LogShellRun(rc) ' 0 秒放生記啟動，否則記結束碼
    End If
    Call TrayBalloon(S_TipState(wantOn), S_LogPhaseDone(wantOn, DisplayName0()))
    LogMsg S_LogPhaseDone(wantOn, DisplayName0())
    FinishPhase
    Exit Sub
Fail:
    LogMsg S_LogAutoErr(Err.Description)
    FinishPhase
End Sub

' 用途：清排程狀態回閒置
Private Sub FinishPhase()
    m_pending = -1 ' 無排程
    m_phase = PH_IDLE
    tmrWait.Enabled = False ' 排程節拍關閉
End Sub

' 用途：runas 提權起自己 --clean，只等 3 秒；逾時放生記 -1，提權使用者自理
Private Function RunCleanElevated() As Long
    On Error GoTo Fail
    Dim sei As SHELLEXECUTEINFO
    Dim exe As String, wr As Long, rc As Long, waited As Long
    exe = App.Path
    If Right$(exe, 1) <> "\" Then exe = exe & "\"
    exe = exe & App.EXEName & ".exe"
    With sei
        .cbSize = Len(sei)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .hwnd = Me.hwnd
        .lpVerb = "runas" ' 提權動詞：跳 UAC
        .lpFile = exe
        .lpParameters = "--clean" ' 子行程只做清卡就走
        .lpDirectory = App.Path
        .nShow = SW_HIDE
    End With
    LogMsg S_LogCleanElevate()
    If ShellExecuteEx(sei) = 0 Then
        LogMsg S_LogCleanElevateFail(0)
        RunCleanElevated = 5 ' 叫不起來（多半 UAC 按取消）
        Exit Function
    End If
    If sei.hProcess = 0 Then RunCleanElevated = 0: Exit Function ' 拿不到代碼：當無事發生
    waited = 0
    Do
        wr = WaitForSingleObject(sei.hProcess, 500) ' 500 毫秒一切片等
        If wr = WAIT_OBJECT_0 Then Exit Do
        DoEvents
        waited = waited + 500
    Loop While waited < 3000 ' 上限 3 秒：不等做完，UAC 失敗記一筆就好
    If wr = WAIT_OBJECT_0 Then GetExitCodeProcess sei.hProcess, rc Else rc = -1 ' 正常取碼，逾時記 -1
    CloseHandle sei.hProcess
    RunCleanElevated = rc
    Exit Function
Fail:
    LogMsg S_LogCleanErr(Err.Description)
    RunCleanElevated = 1
End Function

' 用途：托盤選單切換自動（只改記憶體，存檔走設定窗）
Private Sub mnuAuto_Click()
    mnuAuto.Checked = Not mnuAuto.Checked
    g_AutoOn = mnuAuto.Checked
    LogMsg S_LogAutoToggle(g_AutoOn)
End Sub

' 用途：托盤選單切換氣球提示，立即寫 INI
Private Sub mnuBalloon_Click()
    mnuBalloon.Checked = Not mnuBalloon.Checked
    g_Balloon = mnuBalloon.Checked
    Call ConfigSave
    LogMsg S_LogBalloonToggle(g_Balloon)
End Sub

' 用途：手動立即跑斷電鏈
Private Sub mnuFireOff_Click()
    LogMsg S_LogManual(False) ' 手動不斷自動開關，直接排斷電鏈
    StartTransition False
End Sub

' 用途：手動立即跑通電鏈
Private Sub mnuFireOn_Click()
    LogMsg S_LogManual(True) ' 手動不斷自動開關，直接排通電鏈
    StartTransition True
End Sub

' 用途：手動直接開 HDR（不跑鏈）
Private Sub mnuHdrOn_Click()
    On Error Resume Next
    Dim ok As Boolean
    LogMsg S_LogHdrManual(True)
    ok = HDR_Enable() ' 手動直打原生 API，不跑鏈
    LogMsg S_LogNativeResult(ok, HDR_GetStatus(0))
End Sub

' 用途：手動直接關 HDR（不跑鏈）
Private Sub mnuHdrOff_Click()
    On Error Resume Next
    Dim ok As Boolean
    LogMsg S_LogHdrManual(False)
    ok = HDR_Disable() ' 手動直打原生 API，不跑鏈
    LogMsg S_LogNativeResult(ok, HDR_GetStatus(0))
End Sub

' 用途：手動提權清虛擬卡
Private Sub mnuClean_Click()
    On Error Resume Next
    Dim crc As Long
    LogMsg S_LogCleanStart()
    crc = RunCleanElevated() ' 手動清卡也會跳 UAC
    LogMsg S_LogCleanDone(crc)
End Sub

' 用途：開設定窗一般頁
Private Sub mnuSettings_Click()
    OpenSettings 0
End Sub

' 用途：開設定窗指定分頁；參數 tabIdx：0一般 1斷電 2通電 3記錄
Public Sub OpenSettings(ByVal tabIdx As Integer)
    On Error Resume Next
    frmSettings.ShowTab tabIdx
    frmSettings.Show
    frmSettings.WindowState = vbNormal
    frmSettings.SetFocus
End Sub

' 用途：用檔案總管選取開啟記錄檔
Private Sub mnuLog_Click()
    On Error GoTo Fail
    ShellExecute Me.hwnd, "open", "explorer.exe", "/select,""" & g_LogFile & """", "", SW_SHOWNORMAL
    Exit Sub
Fail:
    LogMsg S_LogOpenLogErr(Err.Description)
End Sub

' 用途：唯一出口：拆托盤圖示、寫結束記錄、卸載結束
Private Sub mnuExit_Click()
    TrayRemove
    LogMsg S_LogExit()
    On Error Resume Next
    Unload frmSettings ' 先卸設定窗再卸自己
    Unload Me
    End ' 全劇終：VB6 建議收尾
End Sub

' 用途：註冊托盤圖示；圖示取 Me.Icon（.frx），取不到退回系統預設
Private Sub InitTray(ByVal MouseMoveTip As String)
    Dim hIco As Long
    On Error Resume Next
    hIco = 0
    hIco = Me.Icon.Handle ' 首選：窗體圖示（.frx 內嵌）
    If hIco = 0 Then hIco = CLng(Me.Icon) ' 次選：圖示數值
    If hIco = 0 Then hIco = LoadIcon(0, IDI_APPLICATION) ' 保底：系統預設圖
    On Error GoTo 0

    With m_nid
        .cbSize = Len(m_nid)
        .hwnd = Me.hwnd
        .uId = 19791229 ' 固定 ID：重建工作列也認得
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = hIco
        .szTip = left$(MouseMoveTip & String$(127, vbNullChar), 127) & vbNullChar ' 補結尾防截斷亂碼
        .dwState = 0
        .dwStateMask = 0
        .szInfo = vbNullChar
        .uTimeoutOrVersion = 0
        .szInfoTitle = vbNullChar
        .dwInfoFlags = 0
    End With
    m_trayAdded = (Shell_NotifyIconA(NIM_ADD, m_nid) <> 0) ' 非零即註冊成功
    If Not m_trayAdded Then
        LogMsg "TrayAdd failed hwnd=" & Me.hwnd
    End If
End Sub

' 用途：更新托盤懸停文字
Private Sub TrayTip(ByVal txt As String)
    Dim tip As String
    If Not m_trayAdded Then Exit Sub ' 未註冊：後面全免談
    tip = txt
    If Len(tip) > 127 Then tip = left$(tip, 127) ' 超長截斷，防 API 拒收
    m_nid.szTip = tip & vbNullChar
    m_nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Shell_NotifyIconA NIM_MODIFY, m_nid ' 送修改：只動提示
End Sub

' 用途：移除托盤圖示（結束時）
Private Sub TrayRemove()
    If m_trayAdded Then
        Shell_NotifyIconA NIM_DELETE, m_nid ' 拆圖示：防結束後殘留孤兒
        m_trayAdded = False
    End If
End Sub

' 用途：托盤回呼：右鍵放開彈選單、左鍵或雙擊開設定
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    If Not m_trayAdded Then Exit Sub ' 未註冊：後面全免談

    If Me.ScaleMode = vbPixels Then ' 像素模式：X 即訊息號
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX ' 緹轉像素還原訊息號
    End If

    Select Case msg
        Case WM_RBUTTONUP ' 右鍵放開：彈選單
            Call ShowTrayMenu
        Case WM_LBUTTONUP ' 左鍵放開：開設定
            Call OpenSettings(0)
        Case WM_LBUTTONDBLCLK ' 左鍵雙擊：開設定
            Call OpenSettings(0)
    End Select
End Sub

' 用途：彈右鍵選單（先搶前景，防選單卡住不消失）
Private Sub ShowTrayMenu()
    On Error Resume Next
    Call SetForegroundWindow(Me.hwnd) ' 先搶前景：選單失焦才會自動收
    Call PopupMenu(mnuTray)
    Call PostMessage(Me.hwnd, WM_NULL, 0, 0) ' 補空訊息：老招，防選單卡住
End Sub


' 用途：卸載時確保托盤圖示移除
Private Sub Form_Unload(Cancel As Integer)
    TrayRemove
End Sub

' 用途：氣球通知；g_Balloon 關閉時直接返回
Private Sub TrayBalloon(ByVal tipTitle As String, ByVal tipText As String)
    If Not g_Balloon Then Exit Sub
    If Not m_trayAdded Then Exit Sub ' 未註冊：後面全免談
    On Error Resume Next
    With m_nid
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
        .szInfoTitle = left$(tipTitle & String$(63, vbNullChar), 63) & vbNullChar ' 標題限 63 字
        .szInfo = left$(tipText & String$(255, vbNullChar), 255) & vbNullChar ' 內文限 255 字
        .dwInfoFlags = NIIF_INFO
        .uTimeoutOrVersion = 10 ' 停留 10 秒
    End With
    Call Shell_NotifyIconA(NIM_MODIFY, m_nid) ' 送修改：發氣球
End Sub
