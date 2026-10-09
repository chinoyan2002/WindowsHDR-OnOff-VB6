Option Explicit

' 本模組：讀寫 INI 設定（參數英文，說明中文）
' 首次執行（無 INI）→ 記憶體預設＋ g_FirstRun

Public Type TransCfg
    Enabled As Boolean
    DelayHDR As Long
    NativeHDR As String
    VerifySec As Long
    CleanHelper As Boolean
    DelayAfterClean As Long
    Shell As String
End Type

' 全域設定值（ConfigLoad 讀入／設定窗寫回）
Public g_Language As String
Public g_PollSec As Long
Public g_StableN As Long
Public g_AutoOn As Boolean
Public g_LogDir As String
Public g_Balloon As Boolean
Public g_PowerOff As TransCfg
Public g_PowerOn As TransCfg
Public g_WorkDir As String
Public g_FirstRun As Boolean

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpDef As String, ByVal lpRet As String, ByVal nSize As Long, ByVal lpFile As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApp As String, ByVal lpKey As String, ByVal lpVal As String, ByVal lpFile As String) As Long

' 讀入全部設定；無 INI 視為首次執行
Public Sub ConfigLoad()
    Dim ini As String
    g_WorkDir = App.Path
    If Right$(g_WorkDir, 1) <> "\" Then g_WorkDir = g_WorkDir & "\"
    ini = g_WorkDir & "AutoHdrMode.ini"
    If Dir$(ini) = "" Then
        g_FirstRun = True
        SetDefaults
        Exit Sub
    End If
    g_FirstRun = False
    g_Language = IniGet(ini, "Main", "Language", "auto")
    g_PollSec = Val(IniGet(ini, "Main", "PollSeconds", "4"))
    If g_PollSec < 1 Then g_PollSec = 1
    If g_PollSec > 30 Then g_PollSec = 30
    g_StableN = Val(IniGet(ini, "Main", "StablePolls", "3"))
    If g_StableN < 1 Then g_StableN = 1
    g_AutoOn = (Val(IniGet(ini, "Main", "AutoEnabled", "1")) <> 0)
    g_Balloon = (Val(IniGet(ini, "Main", "Balloon", "1")) <> 0)
    g_LogDir = IniGet(ini, "Main", "LogDir", "")
    ReadTrans ini, "PowerOff", g_PowerOff, 1, "off"
    ReadTrans ini, "PowerOn", g_PowerOn, 2, "on"
End Sub

Private Sub SetDefaults()
    g_Language = "auto"
    g_PollSec = 4
    g_StableN = 3
    g_AutoOn = True
    g_Balloon = True
    g_LogDir = ""
    g_PowerOff.Enabled = True
    g_PowerOff.DelayHDR = 1
    g_PowerOff.NativeHDR = "OFF"
    g_PowerOff.VerifySec = 2
    g_PowerOff.CleanHelper = True
    g_PowerOff.DelayAfterClean = 5
    g_PowerOff.Shell = ""
    g_PowerOn.Enabled = True
    g_PowerOn.DelayHDR = 2
    g_PowerOn.NativeHDR = "ON"
    g_PowerOn.VerifySec = 2
    g_PowerOn.CleanHelper = True
    g_PowerOn.DelayAfterClean = 5
    g_PowerOn.Shell = ""
End Sub

Private Sub ReadTrans(ByVal ini As String, ByVal sec As String, ByRef cfg As TransCfg, ByVal defDelay As Long, ByVal defNative As String)
    Dim s As String
    cfg.Enabled = (Val(IniGet(ini, sec, "Enabled", "1")) <> 0)
    s = IniGet(ini, sec, "DelayHDR", "")
    If s = "" Then s = IniGet(ini, sec, "DelaySeconds", CStr(defDelay))
    cfg.DelayHDR = Val(s)
    If cfg.DelayHDR < 0 Then cfg.DelayHDR = 0
    If cfg.DelayHDR > 120 Then cfg.DelayHDR = 120
    cfg.NativeHDR = UCase$(Trim$(IniGet(ini, sec, "NativeHDR", defNative)))
    cfg.VerifySec = Val(IniGet(ini, sec, "VerifySeconds", "2"))
    If cfg.VerifySec < 0 Then cfg.VerifySec = 0
    If cfg.VerifySec > 60 Then cfg.VerifySec = 60
    cfg.CleanHelper = (Val(IniGet(ini, sec, "CleanHelper", "1")) <> 0)
    cfg.DelayAfterClean = Val(IniGet(ini, sec, "DelayAfterClean", "5"))
    If cfg.DelayAfterClean < 0 Then cfg.DelayAfterClean = 0
    If cfg.DelayAfterClean > 120 Then cfg.DelayAfterClean = 120
    cfg.Shell = IniGet(ini, sec, "Shell", "")
End Sub

' 寫回全部設定並清首次旗標
Public Function ConfigSave() As Boolean
    On Error GoTo Fail
    Dim ini As String
    ini = g_WorkDir & "AutoHdrMode.ini"
    SaveStr ini, "Main", "Language", g_Language
    SaveStr ini, "Main", "PollSeconds", CStr(g_PollSec)
    SaveStr ini, "Main", "StablePolls", CStr(g_StableN)
    SaveStr ini, "Main", "AutoEnabled", IIf(g_AutoOn, "1", "0")
    SaveStr ini, "Main", "Balloon", IIf(g_Balloon, "1", "0")
    SaveStr ini, "Main", "LogDir", g_LogDir
    WriteTrans ini, "PowerOff", g_PowerOff
    WriteTrans ini, "PowerOn", g_PowerOn
    g_FirstRun = False
    ConfigSave = True
    Exit Function
Fail:
    ConfigSave = False
End Function

Private Sub SaveStr(ByVal ini As String, ByVal sec As String, ByVal key As String, ByVal val As String)
    WritePrivateProfileString sec, key, val, ini
End Sub

Private Sub WriteTrans(ByVal ini As String, ByVal sec As String, ByRef cfg As TransCfg)
    SaveStr ini, sec, "Enabled", IIf(cfg.Enabled, "1", "0")
    SaveStr ini, sec, "DelayHDR", CStr(cfg.DelayHDR)
    SaveStr ini, sec, "NativeHDR", cfg.NativeHDR
    SaveStr ini, sec, "VerifySeconds", CStr(cfg.VerifySec)
    SaveStr ini, sec, "CleanHelper", IIf(cfg.CleanHelper, "1", "0")
    SaveStr ini, sec, "DelayAfterClean", CStr(cfg.DelayAfterClean)
    SaveStr ini, sec, "Shell", cfg.Shell
End Sub

Private Function IniGet(ByVal iniFile As String, ByVal sec As String, ByVal key As String, ByVal deft As String) As String
    Dim buf As String
    buf = Space$(1024)
    Dim n As Long
    n = GetPrivateProfileString(sec, key, deft, buf, 1024, iniFile)
    IniGet = Left$(buf, n)
End Function

' Shell 第一 token 是否存在（引號／相對路徑皆處理）
Public Function ShellTargetExists(ByVal shellCmd As String) As Boolean
    Dim s As String, tok As String, p As Long
    s = Trim$(shellCmd)
    If s = "" Then Exit Function
    If Left$(s, 1) = """" Then
        p = InStr(2, s, """")
        If p > 2 Then tok = Mid$(s, 2, p - 2) Else tok = s
    Else
        p = InStr(s, " ")
        If p > 0 Then tok = Left$(s, p - 1) Else tok = s
    End If
    tok = Trim$(tok)
    If tok = "" Then Exit Function
    If Mid$(tok, 2, 1) <> ":" And Left$(tok, 1) <> "\" Then tok = g_WorkDir & tok
    ShellTargetExists = (Dir$(tok) <> "")
End Function
