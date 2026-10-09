Attribute VB_Name = "modDDC"
' DDC 模組：讀實體螢幕 DDC/CI D6 電源模式
' 斷電不丟錯：讀不到一律記問不到（字面關）
Option Explicit

Private Type PHYSICAL_MONITOR
    hPhysicalMonitor As Long ' 實體螢幕代碼
    szPhysicalMonitorDescription(255) As Byte ' 描述 256 位元組（即 128 寬字）
End Type

Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, ByVal lprcClip As Long, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long ' 列舉顯示器：一台呼叫一次回呼
Private Declare Function GetPhysicalMonitorsFromHMONITOR Lib "dxva2.dll" (ByVal hMonitor As Long, ByVal dwPhysicalMonitorArraySize As Long, pPhysicalMonitorArray As PHYSICAL_MONITOR) As Long ' 取實體代碼：虛擬卡無此物會失敗
Private Declare Function DestroyPhysicalMonitors Lib "dxva2.dll" (ByVal dwPhysicalMonitorArraySize As Long, pPhysicalMonitorArray As PHYSICAL_MONITOR) As Long ' 釋放代碼：必呼叫防洩漏
Private Declare Function GetVCPFeatureAndVCPFeatureReply Lib "dxva2.dll" (ByVal hMonitor As Long, ByVal bVCPCode As Byte, pdwVCPCodeType As Long, pdwCurrentValue As Long, pdwMaximumValue As Long) As Long ' 讀 VCP 功能值：D6 即電源

Private m_MonitorCount As Long                                                  ' 記錄實體螢幕總數
Private m_PowerOnCount As Long                                                  ' 記錄通電中的螢幕數
Private m_StandbyCount As Long
Private m_D6N As Long
Private m_D6Vals(7) As Long                                                  ' 待命中台數
Private m_CachedName As String



' 用途：單次列舉全部螢幕，同步填三組計數（通電、待命、總數）
Private Sub EnumAll()
    m_MonitorCount = 0
    m_PowerOnCount = 0
    m_StandbyCount = 0                                                          ' 單次列舉，多組計數
    m_D6N = 0
    EnumDisplayMonitors 0, 0, AddressOf MonitorEnumProc, 0                      ' 列舉所有顯示器
End Sub

' 用途：D6=1 通電台數；每次呼叫重新列舉
Public Function DdcOnCount() As Long
    EnumAll
    DdcOnCount = m_PowerOnCount                                           ' DDC/CI D6 通電台數
End Function

' 用途：D6=2/3 待命台數；須先呼叫 DdcOnCount 才有新值
Public Function DdcStandbyCount() As Long
    DdcStandbyCount = m_StandbyCount                                            ' 需先呼叫 DdcOnCount
End Function

' 用途：EnumDisplayMonitors 回呼；D6=1 計通電、2/3 計待命，餘不計；回傳 1 繼續下一個
' 用途：EnumDisplayMonitors 回呼；D6 記原始值供除錯；回傳 1 繼續下一個
' VESA MCCS D6 電源模式：1=開機、2=待命、3=休眠、4=關機、5=硬體全關；問不到記 -1
Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByVal lprcMonitor As Long, ByVal dwData As Long) As Long
    Dim physMon As PHYSICAL_MONITOR, vct As Long, currentVal As Long, maxVal As Long
    Dim v As Long
    v = -1                                                              ' 預設問不到
    If GetPhysicalMonitorsFromHMONITOR(hMonitor, 1, physMon) Then
        m_MonitorCount = m_MonitorCount + 1                             ' 發現實體螢幕
        If GetVCPFeatureAndVCPFeatureReply(physMon.hPhysicalMonitor, &HD6, vct, currentVal, maxVal) Then
            v = currentVal                                              ' 記原始回傳值
            If currentVal = 1 Then
                m_PowerOnCount = m_PowerOnCount + 1                     ' D6=1 通電中
            ElseIf currentVal = 2 Or currentVal = 3 Then
                m_StandbyCount = m_StandbyCount + 1                     ' D6=2/3 待命
            End If
        End If
        DestroyPhysicalMonitors 1, physMon                              ' 釋放代碼
    End If
    If m_D6N <= 7 Then m_D6Vals(m_D6N) = v: m_D6N = m_D6N + 1            ' 原始值入列（超 8 台略過）
    MonitorEnumProc = 1                                                 ' 回傳 1 繼續列舉
End Function

' 用途：回傳本次列舉各台 D6 原始值，如 D6=[1]、D6=[問不到]；須先呼叫 DdcOnCount
' 用途：本次列舉 D6 原始值；單台純量（D6=1／D6=問不到），多台列表；須先呼叫 DdcOnCount
Public Function DdcD6Summary() As String
    Dim i As Long, s As String, v As Long
    If m_D6N <= 0 Then DdcD6Summary = "D6=[未列舉]": Exit Function
    If m_D6N = 1 Then
        If m_D6Vals(0) < 0 Then DdcD6Summary = "D6=問不到" Else DdcD6Summary = "D6=" & CStr(m_D6Vals(0))
        Exit Function
    End If
    s = ""
    For i = 0 To m_D6N - 1
        v = m_D6Vals(i)
        If Len(s) > 0 Then s = s & ","
        If v < 0 Then s = s & "問不到" Else s = s & CStr(v)
    Next
    DdcD6Summary = "D6=[" & s & "]"
End Function



' 用途：取 0 號螢幕名稱；成功寫入快取，失敗回快取，再沒有回「螢幕」
' 注意：只在事件／初始／收尾呼叫，不進每輪 poll
Public Function DisplayName0() As String
    Dim ds() As HDR_DISPLAY_INFO
    Dim nm As String
    On Error Resume Next
    Err.Clear
    If HDR_GetDisplays(ds) Then
        nm = Trim$(ds(0).Name)
        If Err.Number = 0 And Len(nm) > 0 And Left$(nm, 3) <> "???" Then m_CachedName = nm
    End If
    Err.Clear
    If Len(m_CachedName) > 0 Then DisplayName0 = m_CachedName Else DisplayName0 = "螢幕"
End Function
