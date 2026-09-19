' 本模組：讀 INI 設定（對照 C# 版 AppConfig）
' 設定檔與程式同目錄，檔名 AutoHdrMode.ini
Option Explicit

' 斷電／通電命令組：等幾秒＋執行什麼
Public Type TransCfg
    DelaySec As Long
    Shell As String
End Type

' 以下全域變數由 ConfigLoad 讀入
Public g_Language As String
Public g_PollSec As Long
Public g_StableN As Long
Public g_AutoOn As Boolean
Public g_LogDir As String
Public g_PowerOff As TransCfg
Public g_PowerOn As TransCfg
Public g_WorkDir As String
Public g_PowerOffHdr As String
Public g_PowerOnHdr As String
Public g_VerifySec As Long
Public g_CleanOnFail As Boolean

' Windows 內建讀 INI 函式
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpDef As String, ByVal lpRet As String, ByVal nSize As Long, ByVal lpFile As String) As Long

' 讀入全部設定並夾限範圍（VB6 Timer 上限 32767 毫秒，故輪詢最多 30 秒）
Public Sub ConfigLoad()
    Dim ini As String
    g_WorkDir = App.Path
    If Right$(g_WorkDir, 1) <> "\" Then g_WorkDir = g_WorkDir & "\"
    ini = g_WorkDir & "AutoHdrMode.ini"
    g_Language = IniGet(ini, "Main", "Language", "auto")
    ' NOTE: VB6 Timer.Interval is 16-bit (max 32767ms), so cap at 30s.
    g_PollSec = Val(IniGet(ini, "Main", "PollSeconds", "4"))
    If g_PollSec < 1 Then g_PollSec = 1
    If g_PollSec > 30 Then g_PollSec = 30
    g_StableN = Val(IniGet(ini, "Main", "StablePolls", "3"))
    If g_StableN < 1 Then g_StableN = 1
    g_AutoOn = (Val(IniGet(ini, "Main", "AutoEnabled", "1")) <> 0)
    g_LogDir = IniGet(ini, "Main", "LogDir", "")
    g_PowerOff.DelaySec = Val(IniGet(ini, "PowerOff", "DelaySeconds", "3"))
    g_PowerOff.Shell = IniGet(ini, "PowerOff", "Shell", "")
    g_PowerOn.DelaySec = Val(IniGet(ini, "PowerOn", "DelaySeconds", "8"))
    g_PowerOn.Shell = IniGet(ini, "PowerOn", "Shell", "")
    g_PowerOffHdr = UCase$(Trim$(IniGet(ini, "PowerOff", "HdrAction", "")))
    g_PowerOnHdr = UCase$(Trim$(IniGet(ini, "PowerOn", "HdrAction", "")))
    g_VerifySec = Val(IniGet(ini, "PowerOn", "VerifySeconds", "10"))
    If g_VerifySec < 0 Then g_VerifySec = 0
    If g_VerifySec > 120 Then g_VerifySec = 120
    g_CleanOnFail = (Val(IniGet(ini, "PowerOn", "CleanOnFail", "1")) <> 0)
End Sub

' 讀單一鍵值，讀不到回傳預設值
Private Function IniGet(ByVal iniFile As String, ByVal sec As String, ByVal key As String, ByVal deft As String) As String
    Dim buf As String
    buf = Space$(1024)
    Dim n As Long
    n = GetPrivateProfileString(sec, key, deft, buf, 1024, iniFile)
    IniGet = Left$(buf, n)
End Function
