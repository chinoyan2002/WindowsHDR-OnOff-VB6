Attribute VB_Name = "modLang"
Option Explicit
' 語言模組：全部 UI 與記錄字串；g_Chinese=True 繁中否則英文，auto 跟系統
Public g_Chinese As Boolean
Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer

Public Sub LangInit(ByVal mode As String) ' 語言初始化：auto 跟系統語系
    Dim m As String
    m = LCase$(Trim$(mode))
    Select Case m
        Case "zh", "zh-tw", "zh-cn", "zh-hk", "chs", "cht", "chinese"
            g_Chinese = True
        Case "en", "english"
            g_Chinese = False
        Case Else
            g_Chinese = ((GetUserDefaultUILanguage() And &H3FF) = &H4)
    End Select
End Sub

' ---- 托盤選單文字 ----
Public Function S_MenuAuto() As String ' 選單：自動切換勾選
    If g_Chinese Then S_MenuAuto = "自動切換" Else S_MenuAuto = "Auto switch"
End Function
Public Function S_MenuFireOff() As String ' 選單：手動跑斷電鏈
    If g_Chinese Then S_MenuFireOff = "立即執行斷電流程" Else S_MenuFireOff = "Run power-off now"
End Function
Public Function S_MenuFireOn() As String ' 選單：手動跑通電鏈
    If g_Chinese Then S_MenuFireOn = "立即執行通電流程" Else S_MenuFireOn = "Run power-on now"
End Function
Public Function S_MenuHdrOn() As String ' 選單：直接開 HDR
    If g_Chinese Then S_MenuHdrOn = "開啟 HDR" Else S_MenuHdrOn = "Turn HDR On"
End Function
Public Function S_MenuHdrOff() As String ' 選單：直接關 HDR
    If g_Chinese Then S_MenuHdrOff = "關閉 HDR" Else S_MenuHdrOff = "Turn HDR Off"
End Function
Public Function S_MenuClean() As String ' 選單：手動清虛擬卡
    If g_Chinese Then S_MenuClean = "清除虛擬顯示卡…" Else S_MenuClean = "Clean virtual GPUs…"
End Function
Public Function S_MenuSettings() As String ' 選單：開設定窗
    If g_Chinese Then S_MenuSettings = "設定…" Else S_MenuSettings = "Settings…"
End Function
Public Function S_MenuOpenLog() As String ' 選單：開記錄檔位置
    If g_Chinese Then S_MenuOpenLog = "開啟紀錄資料夾" Else S_MenuOpenLog = "Open log folder"
End Function
Public Function S_MenuExit() As String ' 選單：結束程式
    If g_Chinese Then S_MenuExit = "結束程式" Else S_MenuExit = "Exit"
End Function

' ---- 托盤懸停提示 ----
Public Function S_TipStarting() As String ' 懸停：啟動中
    If g_Chinese Then S_TipStarting = "AutoHdrMode 啟動中" Else S_TipStarting = "AutoHdrMode starting"
End Function
Public Function S_TipState(ByVal isOn As Boolean) As String ' 懸停：開關狀態
    If g_Chinese Then
        If isOn Then S_TipState = "AutoHdrMode：螢幕電源開啟" Else S_TipState = "AutoHdrMode：螢幕電源關閉"
    Else
        If isOn Then S_TipState = "AutoHdrMode: display ON" Else S_TipState = "AutoHdrMode: display OFF"
    End If
End Function

' ---- 記錄檔字串 ----
Public Function S_LogStart(ByVal poll As Long, ByVal stable As Long, ByVal autoOn As Boolean) As String ' 記錄：常駐啟動參數
    If g_Chinese Then
        S_LogStart = "程式啟動。輪詢間隔 " & poll & " 秒，穩定確認 " & stable & " 次，自動切換=" & IIf(autoOn, "開", "關")
    Else
        S_LogStart = "START poll=" & poll & "s stable=" & stable & " auto=" & autoOn
    End If
End Function
Public Function S_LogAutoToggle(ByVal autoOn As Boolean) As String ' 記錄：自動開關異動
    If g_Chinese Then S_LogAutoToggle = "自動切換已改為：" & IIf(autoOn, "開啟", "關閉") Else S_LogAutoToggle = "AUTO " & IIf(autoOn, "on", "off")
End Function
Public Function S_LogInit(ByVal isOn As Boolean, ByVal wmi As Long) As String ' 記錄：初始狀態
    If g_Chinese Then
        S_LogInit = "初始狀態：螢幕電源=" & IIf(isOn, "開", "關") & "（WMI 筆數=" & wmi & "）"
    Else
        S_LogInit = "INIT physical=" & IIf(isOn, "ON", "OFF") & " wmi=" & wmi
    End If
End Function
Public Function S_LogEvent(ByVal isOn As Boolean, ByVal wmi As Long) As String ' 記錄：翻轉事件
    If g_Chinese Then
        S_LogEvent = "狀態變更：螢幕電源=" & IIf(isOn, "開", "關") & "（WMI 筆數=" & wmi & "）"
    Else
        S_LogEvent = "EVENT physical=" & IIf(isOn, "ON", "OFF") & " wmi=" & wmi
    End If
End Function
Public Function S_LogDirDisabled(ByVal isOn As Boolean) As String ' 記錄：方向被停用
    If g_Chinese Then
        S_LogDirDisabled = IIf(isOn, "通電", "斷電") & "自動動作已關閉（設定 Enabled=0），本次不執行"
    Else
        S_LogDirDisabled = IIf(isOn, "PowerOn", "PowerOff") & " disabled, skip"
    End If
End Function
Public Function S_LogPlanHdr(ByVal isOn As Boolean, ByVal d As Long) As String ' 記錄：排程延遲秒數
    If g_Chinese Then
        S_LogPlanHdr = IIf(isOn, "通電", "斷電") & "流程：先等待 " & d & " 秒再處理 HDR"
    Else
        S_LogPlanHdr = IIf(isOn, "ON", "OFF") & " HDR in " & d & "s"
    End If
End Function
Public Function S_LogSkip(ByVal isOn As Boolean) As String ' 記錄：排程中途取消
    If g_Chinese Then
        S_LogSkip = "等待期間螢幕狀態又變了，取消本次" & IIf(isOn, "通電", "斷電") & "流程"
    Else
        S_LogSkip = "state changed, cancel"
    End If
End Function
Public Function S_LogNativeTryOn() As String ' 記錄：嘗試開 HDR
    If g_Chinese Then S_LogNativeTryOn = "正在開啟原生 HDR…" Else S_LogNativeTryOn = "NATIVE HDR enable…"
End Function
Public Function S_LogNativeTryOff() As String ' 記錄：嘗試關 HDR
    If g_Chinese Then S_LogNativeTryOff = "正在關閉原生 HDR…" Else S_LogNativeTryOff = "NATIVE HDR disable…"
End Function
Public Function S_LogNativeResult(ByVal ok As Boolean, ByVal st As Long) As String ' 記錄：原生 HDR 結果
    Dim s As String
    Select Case st
        Case 0: s = IIf(g_Chinese, "不支援", "UNSUPPORTED")
        Case 1: s = IIf(g_Chinese, "已關閉", "OFF")
        Case 2: s = IIf(g_Chinese, "已開啟", "ON")
        Case 3: s = IIf(g_Chinese, "錯誤", "ERROR")
        Case Else: s = IIf(g_Chinese, "未知", "UNKNOWN")
    End Select
    If g_Chinese Then
        S_LogNativeResult = "原生 HDR 結果：成功=" & IIf(ok, "是", "否") & "，目前狀態=" & s
    Else
        S_LogNativeResult = "NATIVE ok=" & IIf(ok, "Y", "N") & " " & s
    End If
End Function
Public Function S_LogNativeSkip() As String ' 記錄：略過 HDR
    If g_Chinese Then S_LogNativeSkip = "未設定原生 HDR 動作，略過 HDR 步驟" Else S_LogNativeSkip = "NativeHDR empty, skip HDR"
End Function
Public Function S_LogVerifyWait(ByVal sec As Long) As String ' 記錄：驗證等待秒數
    If g_Chinese Then S_LogVerifyWait = "等待 " & sec & " 秒後檢查 HDR 是否生效…" Else S_LogVerifyWait = "VERIFY wait " & sec & "s"
End Function
Public Function S_LogVerifyOk() As String ' 記錄：驗證通過
    If g_Chinese Then S_LogVerifyOk = "檢查通過：HDR 狀態符合預期" Else S_LogVerifyOk = "VERIFY ok"
End Function
Public Function S_LogVerifyFail() As String ' 記錄：驗證失敗
    If g_Chinese Then S_LogVerifyFail = "檢查未通過：HDR 狀態與預期不符" Else S_LogVerifyFail = "VERIFY fail"
End Function
Public Function S_LogCleanStart() As String ' 記錄：開始清卡
    If g_Chinese Then S_LogCleanStart = "開始清除虛擬顯示卡…" Else S_LogCleanStart = "CLEAN start"
End Function
Public Function S_LogCleanNone() As String ' 記錄：無卡可清
    If g_Chinese Then S_LogCleanNone = "未發現需清除的虛擬顯示卡" Else S_LogCleanNone = "CLEAN none"
End Function
Public Function S_LogCleanRemove(ByVal name As String) As String ' 記錄：移除某卡
    If g_Chinese Then S_LogCleanRemove = "移除顯示裝置：" & name Else S_LogCleanRemove = "CLEAN remove " & name
End Function
Public Function S_LogCleanRemoveRc(ByVal rc As Long) As String ' 記錄：移除結束碼
    If g_Chinese Then S_LogCleanRemoveRc = "pnputil 結束代碼=" & rc Else S_LogCleanRemoveRc = "pnputil rc=" & rc
End Function
Public Function S_LogCleanOk() As String ' 記錄：清卡完成
    If g_Chinese Then S_LogCleanOk = "虛擬顯示卡已清除完成" Else S_LogCleanOk = "CLEAN ok"
End Function
Public Function S_LogCleanTimeout() As String ' 記錄：清卡逾時
    If g_Chinese Then S_LogCleanTimeout = "清除虛擬顯示卡逾時" Else S_LogCleanTimeout = "CLEAN timeout"
End Function
Public Function S_LogCleanErr(ByVal m As String) As String ' 記錄：清卡異常
    If g_Chinese Then S_LogCleanErr = "清除虛擬顯示卡發生錯誤：" & m Else S_LogCleanErr = "CLEAN-ERR " & m
End Function
Public Function S_LogCleanElevate() As String ' 記錄：準備提權
    If g_Chinese Then S_LogCleanElevate = "清除虛擬顯示卡需要系統管理員權限，正在要求提權…" Else S_LogCleanElevate = "CLEAN elevate…"
End Function
Public Function S_LogCleanElevateFail(ByVal rc As Long) As String ' 記錄：提權失敗
    If g_Chinese Then S_LogCleanElevateFail = "提權失敗（代碼=" & rc & "）" Else S_LogCleanElevateFail = "ELEVATE fail rc=" & rc
End Function
Public Function S_LogCleanDone(ByVal rc As Long) As String ' 記錄：清卡結束碼
    If g_Chinese Then S_LogCleanDone = "清除虛擬顯示卡結束（代碼=" & rc & "）" Else S_LogCleanDone = "CLEAN done rc=" & rc
End Function
Public Function S_LogRetryEnable() As String ' 記錄：清後重試 HDR
    If g_Chinese Then S_LogRetryEnable = "清卡完成，重新執行 HDR 動作…" Else S_LogRetryEnable = "RETRY HDR after clean"
End Function
Public Function S_LogWaitAfterClean(ByVal sec As Long) As String ' 記錄：清後等待秒數
    If g_Chinese Then S_LogWaitAfterClean = "清卡後再等待 " & sec & " 秒，然後重試 HDR" Else S_LogWaitAfterClean = "wait " & sec & "s after clean"
End Function
Public Function S_LogShellSkipEmpty() As String ' 記錄：Shell 空白跳過
    If g_Chinese Then S_LogShellSkipEmpty = "未設定 Shell 命令，略過" Else S_LogShellSkipEmpty = "Shell empty, skip"
End Function
Public Function S_LogShellSkipMissing(ByVal p As String) As String ' 記錄：Shell 目標遺失
    If g_Chinese Then S_LogShellSkipMissing = "Shell 檔案不存在，略過：" & p Else S_LogShellSkipMissing = "Shell missing, skip: " & p
End Function
Public Function S_LogShellRun(ByVal rc As Long) As String ' 記錄：Shell 結束碼
    If g_Chinese Then S_LogShellRun = "Shell 命令已執行（結束代碼=" & rc & "）" Else S_LogShellRun = "Shell done rc=" & rc
End Function
Public Function S_LogPhaseDone(ByVal isOn As Boolean) As String ' 記錄：整鏈完成
    If g_Chinese Then S_LogPhaseDone = IIf(isOn, "通電", "斷電") & "流程結束" Else S_LogPhaseDone = IIf(isOn, "ON", "OFF") & " phase done"
End Function
Public Function S_LogWmiErr(ByVal m As String) As String ' 記錄：偵測錯誤（舊 WMI 遺留）
    If g_Chinese Then S_LogWmiErr = "螢幕狀態查詢（WMI）：" & m Else S_LogWmiErr = "WMI " & m
End Function
Public Function S_LogPollErr(ByVal m As String) As String ' 記錄：輪詢異常
    If g_Chinese Then S_LogPollErr = "輪詢發生錯誤：" & m Else S_LogPollErr = "POLL-ERR " & m
End Function
Public Function S_LogAutoErr(ByVal m As String) As String ' 記錄：自動鏈異常
    If g_Chinese Then S_LogAutoErr = "自動流程發生錯誤：" & m Else S_LogAutoErr = "AUTO-ERR " & m
End Function
Public Function S_LogOpenLogErr(ByVal m As String) As String ' 記錄：開記錄檔失敗
    If g_Chinese Then S_LogOpenLogErr = "無法開啟紀錄資料夾：" & m Else S_LogOpenLogErr = "OPENLOG-ERR " & m
End Function
Public Function S_LogProbe(ByVal onN As Long, ByVal sbN As Long) As String ' 記錄：探測結果行
    If g_Chinese Then S_LogProbe = "探測 DDC on=" & onN & " standby=" & sbN Else S_LogProbe = "PROBE on=" & onN & " standby=" & sbN
End Function
Public Function S_LogAsIn() As String ' 記錄：自啟寫入
    If g_Chinese Then S_LogAsIn = "已加入開機自動執行" Else S_LogAsIn = "AUTOSTART installed"
End Function
Public Function S_LogAsRm() As String ' 記錄：自啟刪除
    If g_Chinese Then S_LogAsRm = "已取消開機自動執行" Else S_LogAsRm = "AUTOSTART removed"
End Function
Public Function S_LogAsErr(ByVal m As String) As String ' 記錄：自啟異常
    If g_Chinese Then S_LogAsErr = "開機自動執行設定失敗：" & m Else S_LogAsErr = "AUTOSTART-ERR " & m
End Function
Public Function S_LogExit() As String ' 記錄：程式結束
    If g_Chinese Then S_LogExit = "程式結束" Else S_LogExit = "EXIT"
End Function
Public Function S_LogManual(ByVal isOn As Boolean) As String ' 記錄：手動觸發
    If g_Chinese Then S_LogManual = "手動啟動" & IIf(isOn, "通電", "斷電") & "流程" Else S_LogManual = "MANUAL " & IIf(isOn, "ON", "OFF")
End Function
Public Function S_LogNativeOffNoPath() As String ' 記錄：斷電無路徑提示
    If g_Chinese Then S_LogNativeOffNoPath = "關閉 HDR：目前沒有作用中的顯示路徑，已改用快取目標嘗試設定" Else S_LogNativeOffNoPath = "NATIVE off: tried cached targets"
End Function
Public Function S_LogHdrManual(ByVal turningOn As Boolean) As String ' 記錄：手動 HDR
    If g_Chinese Then
        S_LogHdrManual = "手動" & IIf(turningOn, "開啟", "關閉") & " HDR"
    Else
        S_LogHdrManual = "MANUAL HDR " & IIf(turningOn, "ON", "OFF")
    End If
End Function
Public Function S_LogSettingsSaved() As String ' 記錄：設定已存檔
    If g_Chinese Then S_LogSettingsSaved = "設定已儲存" Else S_LogSettingsSaved = "Settings saved"
End Function

' ---- 氣球與待命（新增） ----
Public Function S_MenuBalloon() As String ' 選單：氣球提示勾選
    If g_Chinese Then S_MenuBalloon = "氣球提示" Else S_MenuBalloon = "Balloon tip"
End Function

Public Function S_LogBalloonToggle(ByVal isOn As Boolean) As String ' 記錄：氣球開關異動
    If g_Chinese Then
        If isOn Then S_LogBalloonToggle = "氣球提示：開" Else S_LogBalloonToggle = "氣球提示：關"
    Else
        If isOn Then S_LogBalloonToggle = "Balloon tip: ON" Else S_LogBalloonToggle = "Balloon tip: OFF"
    End If
End Function

Public Function S_TipStandby() As String ' 懸停：待命中
    If g_Chinese Then S_TipStandby = "AutoHdrMode：螢幕待命中" Else S_TipStandby = "AutoHdrMode: display STANDBY"
End Function

Public Function S_LogHoldStandby(ByVal onN As Long, ByVal sbN As Long) As String ' 記錄：待命凍結
    If g_Chinese Then S_LogHoldStandby = "HOLD 待命不動作 on=" & onN & " standby=" & sbN Else S_LogHoldStandby = "HOLD standby no action on=" & onN & " standby=" & sbN
End Function

Public Function S_LogWaitShell(ByVal secs As Long) As String ' 記錄：Shell 前等待秒數
    If g_Chinese Then S_LogWaitShell = "跑 Shell 前再等 " & secs & " 秒" Else S_LogWaitShell = "Wait " & secs & " s before shell"
End Function

Public Function S_LogShellLaunch() As String ' 記錄：Shell 已啟動不等候
    If g_Chinese Then S_LogShellLaunch = "Shell 已啟動，不等候" Else S_LogShellLaunch = "Shell launched, not waiting"
End Function
