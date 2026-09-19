Attribute VB_Name = "modConfig"
' 讀 AutoHdrMode.ini
Option Explicit
' 設定模組：讀寫 AutoHdrMode.ini；TransCfg 為單方向參數組

Public Type TransCfg
    Enabled As Boolean
    DelayHDR As Long          ' 事件後、第一次 HDR 前
    NativeHDR As String       ' "on" / "off" / "" = 不做 HDR
    VerifySeconds As Long     ' HDR 動作後再查狀態前等待（通電用；斷電可 0）
    CleanHelper As Boolean    ' 僅通電：驗證仍 OFF 時清虛擬卡
    DelayAfterClean As Long   ' 清卡後、重試 HDR 前
    Shell As String           ' 空白或不存在則不執行
End Type

' ---- 全域設定（記憶體駐留，存檔才寫碟） ----
Public g_Language As String ' auto 跟系統，否則指定
Public g_PollSec As Long ' 幾秒偵測一次
Public g_StableN As Long ' 連續幾次相同才算穩
Public g_AutoOn As Boolean ' 自動切換總開關
Public g_LogDir As String ' 空白=exe 目錄
Public g_Balloon As Boolean ' 氣球通知開關
Public g_PowerOff As TransCfg
Public g_PowerOn As TransCfg
Public g_WorkDir As String
Public g_FirstRun As Boolean          ' INI 不存在=第一次執行

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpDef As String, ByVal lpRet As String, ByVal nSize As Long, ByVal lpFile As String) As Long ' kernel32 讀 INI
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpVal As String, ByVal lpFile As String) As Long ' kernel32 寫 INI

' 用途：讀 INI；檔不存在視為首次執行；數值夾上下限
Public Sub ConfigLoad()
    Dim ini As String
    g_WorkDir = App.Path ' 工作目錄即 exe 目錄
    If Right$(g_WorkDir, 1) <> "\" Then g_WorkDir = g_WorkDir & "\" ' 尾斜線正規化，後面直接串檔名
    ini = g_WorkDir & "AutoHdrMode.ini" ' 設定與 exe 同目錄
    g_FirstRun = False ' 預設非首次：缺檔才翻首次
    On Error Resume Next
    If Dir$(ini) = "" Then g_FirstRun = True ' 檔不存在=首次執行，顯示設定窗
    On Error GoTo 0
    g_Language = IniGet(ini, "Main", "Language", "auto")
    g_PollSec = Val(IniGet(ini, "Main", "PollSeconds", "4")) ' 缺省 4 秒
    If g_PollSec < 1 Then g_PollSec = 1
    If g_PollSec > 30 Then g_PollSec = 30 ' 上限 30 秒：Timer 吃不下更大
    g_StableN = Val(IniGet(ini, "Main", "StablePolls", "3"))
    If g_StableN < 1 Then g_StableN = 1 ' 至少一次才有意義
    g_AutoOn = (Val(IniGet(ini, "Main", "AutoEnabled", "1")) <> 0) ' 非零即開
    g_LogDir = IniGet(ini, "Main", "LogDir", "")
    g_Balloon = (Val(IniGet(ini, "Main", "Balloon", "1")) <> 0)

    Call LoadTrans(ini, "PowerOff", g_PowerOff, False) ' 讀斷電組
    Call LoadTrans(ini, "PowerOn", g_PowerOn, True) ' 讀通電組
End Sub

' 用途：讀單方向參數；參數 isOn 決定通電或斷電預設值
Private Sub LoadTrans(ByVal ini As String, ByVal sec As String, ByRef c As TransCfg, ByVal isOn As Boolean)
    Dim defHDR As String, defVer As String, defClean As String, defAfter As String
    If isOn Then
        defHDR = "on": defVer = "2": defClean = "1": defAfter = "5" ' 通電預設：開 HDR、驗 2 秒、可清卡、清後等 5 秒
    Else
        defHDR = "off": defVer = "2": defClean = "1": defAfter = "5" ' 斷電預設：關 HDR（其餘同通電）
    End If
    c.Enabled = (Val(IniGet(ini, sec, "Enabled", "1")) <> 0)
    c.DelayHDR = Val(IniGet(ini, sec, "DelayHDR", IniGet(ini, sec, "DelaySeconds", "1"))) ' 舊鍵 DelaySeconds 相容
    If c.DelayHDR < 0 Then c.DelayHDR = 0
    If c.DelayHDR > 120 Then c.DelayHDR = 120 ' 夾限 0~120 秒
    c.NativeHDR = LCase$(Trim$(IniGet(ini, sec, "NativeHDR", defHDR))) ' 正規化：去空白轉小寫再比
    c.VerifySeconds = Val(IniGet(ini, sec, "VerifySeconds", defVer))
    c.CleanHelper = (Val(IniGet(ini, sec, "CleanHelper", defClean)) <> 0)
    c.DelayAfterClean = Val(IniGet(ini, sec, "DelayAfterClean", defAfter))
    If c.VerifySeconds < 0 Then c.VerifySeconds = 0
    If c.VerifySeconds > 60 Then c.VerifySeconds = 60 ' 驗證 0~60 秒
    If c.DelayAfterClean < 0 Then c.DelayAfterClean = 0
    If c.DelayAfterClean > 120 Then c.DelayAfterClean = 120 ' 清後等 0~120 秒
    c.Shell = Trim$(IniGet(ini, sec, "Shell", ""))
End Sub

' 存回 INI（設定窗用）
Public Sub ConfigSave()
    Dim ini As String
    ini = g_WorkDir & "AutoHdrMode.ini" ' 設定與 exe 同目錄
    g_FirstRun = False ' 存過檔=不再是首次
    Call IniSet(ini, "Main", "Language", g_Language)
    Call IniSet(ini, "Main", "PollSeconds", CStr(g_PollSec))
    Call IniSet(ini, "Main", "StablePolls", CStr(g_StableN))
    Call IniSet(ini, "Main", "AutoEnabled", IIf(g_AutoOn, "1", "0"))
    Call IniSet(ini, "Main", "LogDir", g_LogDir)
    Call IniSet(ini, "Main", "Balloon", IIf(g_Balloon, "1", "0"))
    Call SaveTrans(ini, "PowerOff", g_PowerOff)
    Call SaveTrans(ini, "PowerOn", g_PowerOn)
End Sub

' 用途：寫單方向參數
Private Sub SaveTrans(ByVal ini As String, ByVal sec As String, ByRef c As TransCfg)
    Call IniSet(ini, sec, "Enabled", IIf(c.Enabled, "1", "0"))
    Call IniSet(ini, sec, "DelayHDR", CStr(c.DelayHDR))
    Call IniSet(ini, sec, "NativeHDR", c.NativeHDR)
    Call IniSet(ini, sec, "VerifySeconds", CStr(c.VerifySeconds))
    Call IniSet(ini, sec, "CleanHelper", IIf(c.CleanHelper, "1", "0"))
    Call IniSet(ini, sec, "DelayAfterClean", CStr(c.DelayAfterClean))
    Call IniSet(ini, sec, "Shell", c.Shell)
End Sub

' 用途：讀 INI 鍵值；回傳：取不到回傳預設
Public Function IniGet(ByVal iniFile As String, ByVal sec As String, ByVal key As String, ByVal deft As String) As String
    Dim buf As String, n As Long
    buf = Space$(1024) ' 千字緩衝接 API 回傳
    n = GetPrivateProfileString(sec, key, deft, buf, 1024, iniFile)
    IniGet = Left$(buf, n) ' 截實際長度：去尾端空白
End Function

' 用途：寫 INI 鍵值
Private Sub IniSet(ByVal iniFile As String, ByVal sec As String, ByVal key As String, ByVal val As String)
    Call WritePrivateProfileString(sec, key, val, iniFile)
End Sub

' Shell 第一個 token 是否存在（相對路徑以程式目錄為準）
Public Function ShellTargetExists(ByVal cmd As String) As Boolean
    Dim p As String, i As Long
    cmd = Trim$(cmd)
    If Len(cmd) = 0 Then Exit Function ' 空白直接假
    If Left$(cmd, 1) = """" Then
        i = InStr(2, cmd, """") ' 找結尾引號：取引號內路徑
        If i > 1 Then p = Mid$(cmd, 2, i - 2) Else p = Mid$(cmd, 2)
    Else
        i = InStr(1, cmd, " ") ' 無引號：空格前是路徑、後是參數
        If i > 0 Then p = Left$(cmd, i - 1) Else p = cmd
    End If
    p = Trim$(p)
    If Len(p) = 0 Then Exit Function
    If InStr(p, ":") = 0 And Left$(p, 1) <> "\" Then
        p = g_WorkDir & p ' 相對路徑：配工作目錄
    End If
    On Error Resume Next
    ShellTargetExists = (Dir$(p) <> "") ' 存在即真
End Function
