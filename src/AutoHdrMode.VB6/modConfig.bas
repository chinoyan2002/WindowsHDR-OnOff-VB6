Attribute VB_Name = "modConfig"
' 讀 AutoHdrMode.ini
Option Explicit

Public Type TransCfg
    Enabled As Boolean
    DelayHDR As Long          ' 事件後、第一次 HDR 前
    NativeHDR As String       ' "on" / "off" / "" = 不做 HDR
    VerifySeconds As Long     ' HDR 動作後再查狀態前等待（通電用；斷電可 0）
    CleanHelper As Boolean    ' 僅通電：驗證仍 OFF 時清虛擬卡
    DelayAfterClean As Long   ' 清卡後、重試 HDR 前
    Shell As String           ' 空白或不存在則不執行
End Type

Public g_Language As String
Public g_PollSec As Long
Public g_StableN As Long
Public g_AutoOn As Boolean
Public g_LogDir As String
Public g_Balloon As Boolean
Public g_PowerOff As TransCfg
Public g_PowerOn As TransCfg
Public g_WorkDir As String
Public g_FirstRun As Boolean          ' INI 不存在=第一次執行

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpDef As String, ByVal lpRet As String, ByVal nSize As Long, ByVal lpFile As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpVal As String, ByVal lpFile As String) As Long

Public Sub ConfigLoad()
    Dim ini As String
    g_WorkDir = App.Path
    If Right$(g_WorkDir, 1) <> "\" Then g_WorkDir = g_WorkDir & "\"
    ini = g_WorkDir & "AutoHdrMode.ini"
    g_FirstRun = False
    On Error Resume Next
    If Dir$(ini) = "" Then g_FirstRun = True
    On Error GoTo 0
    g_Language = IniGet(ini, "Main", "Language", "auto")
    g_PollSec = Val(IniGet(ini, "Main", "PollSeconds", "4"))
    If g_PollSec < 1 Then g_PollSec = 1
    If g_PollSec > 30 Then g_PollSec = 30
    g_StableN = Val(IniGet(ini, "Main", "StablePolls", "3"))
    If g_StableN < 1 Then g_StableN = 1
    g_AutoOn = (Val(IniGet(ini, "Main", "AutoEnabled", "1")) <> 0)
    g_LogDir = IniGet(ini, "Main", "LogDir", "")
    g_Balloon = (Val(IniGet(ini, "Main", "Balloon", "1")) <> 0)

    Call LoadTrans(ini, "PowerOff", g_PowerOff, False)
    Call LoadTrans(ini, "PowerOn", g_PowerOn, True)
End Sub

Private Sub LoadTrans(ByVal ini As String, ByVal sec As String, ByRef c As TransCfg, ByVal isOn As Boolean)
    Dim defHDR As String, defVer As String, defClean As String, defAfter As String
    If isOn Then
        defHDR = "on": defVer = "2": defClean = "1": defAfter = "5"
    Else
        defHDR = "off": defVer = "2": defClean = "1": defAfter = "5"
    End If
    c.Enabled = (Val(IniGet(ini, sec, "Enabled", "1")) <> 0)
    c.DelayHDR = Val(IniGet(ini, sec, "DelayHDR", IniGet(ini, sec, "DelaySeconds", "1")))
    If c.DelayHDR < 0 Then c.DelayHDR = 0
    If c.DelayHDR > 120 Then c.DelayHDR = 120
    c.NativeHDR = LCase$(Trim$(IniGet(ini, sec, "NativeHDR", defHDR)))
    c.VerifySeconds = Val(IniGet(ini, sec, "VerifySeconds", defVer))
    c.CleanHelper = (Val(IniGet(ini, sec, "CleanHelper", defClean)) <> 0)
    c.DelayAfterClean = Val(IniGet(ini, sec, "DelayAfterClean", defAfter))
    If c.VerifySeconds < 0 Then c.VerifySeconds = 0
    If c.VerifySeconds > 60 Then c.VerifySeconds = 60
    If c.DelayAfterClean < 0 Then c.DelayAfterClean = 0
    If c.DelayAfterClean > 120 Then c.DelayAfterClean = 120
    c.Shell = Trim$(IniGet(ini, sec, "Shell", ""))
End Sub

' 存回 INI（設定窗用）
Public Sub ConfigSave()
    Dim ini As String
    ini = g_WorkDir & "AutoHdrMode.ini"
    g_FirstRun = False
    Call IniSet(ini, "Main", "Language", g_Language)
    Call IniSet(ini, "Main", "PollSeconds", CStr(g_PollSec))
    Call IniSet(ini, "Main", "StablePolls", CStr(g_StableN))
    Call IniSet(ini, "Main", "AutoEnabled", IIf(g_AutoOn, "1", "0"))
    Call IniSet(ini, "Main", "LogDir", g_LogDir)
    Call IniSet(ini, "Main", "Balloon", IIf(g_Balloon, "1", "0"))
    Call SaveTrans(ini, "PowerOff", g_PowerOff)
    Call SaveTrans(ini, "PowerOn", g_PowerOn)
End Sub

Private Sub SaveTrans(ByVal ini As String, ByVal sec As String, ByRef c As TransCfg)
    Call IniSet(ini, sec, "Enabled", IIf(c.Enabled, "1", "0"))
    Call IniSet(ini, sec, "DelayHDR", CStr(c.DelayHDR))
    Call IniSet(ini, sec, "NativeHDR", c.NativeHDR)
    Call IniSet(ini, sec, "VerifySeconds", CStr(c.VerifySeconds))
    Call IniSet(ini, sec, "CleanHelper", IIf(c.CleanHelper, "1", "0"))
    Call IniSet(ini, sec, "DelayAfterClean", CStr(c.DelayAfterClean))
    Call IniSet(ini, sec, "Shell", c.Shell)
End Sub

Public Function IniGet(ByVal iniFile As String, ByVal sec As String, ByVal key As String, ByVal deft As String) As String
    Dim buf As String, n As Long
    buf = Space$(1024)
    n = GetPrivateProfileString(sec, key, deft, buf, 1024, iniFile)
    IniGet = Left$(buf, n)
End Function

Private Sub IniSet(ByVal iniFile As String, ByVal sec As String, ByVal key As String, ByVal val As String)
    Call WritePrivateProfileString(sec, key, val, iniFile)
End Sub

' Shell 第一個 token 是否存在（相對路徑以程式目錄為準）
Public Function ShellTargetExists(ByVal cmd As String) As Boolean
    Dim p As String, i As Long
    cmd = Trim$(cmd)
    If Len(cmd) = 0 Then Exit Function
    If Left$(cmd, 1) = """" Then
        i = InStr(2, cmd, """")
        If i > 1 Then p = Mid$(cmd, 2, i - 2) Else p = Mid$(cmd, 2)
    Else
        i = InStr(1, cmd, " ")
        If i > 0 Then p = Left$(cmd, i - 1) Else p = cmd
    End If
    p = Trim$(p)
    If Len(p) = 0 Then Exit Function
    If InStr(p, ":") = 0 And Left$(p, 1) <> "\" Then
        p = g_WorkDir & p
    End If
    On Error Resume Next
    ShellTargetExists = (Dir$(p) <> "")
End Function
