Attribute VB_Name = "modWMI"
' 本模組：查實體螢幕數量（對照 C# 版 WmiProbe）
' 遠端軟體的虛擬螢幕不會出現在這裡；螢幕斷電時查詢會丟錯，是正常現象
Option Explicit

Private Type PHYSICAL_MONITOR
    hPhysicalMonitor As Long
    szPhysicalMonitorDescription(255) As Byte
End Type

Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, ByVal lprcClip As Long, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Private Declare Function GetPhysicalMonitorsFromHMONITOR Lib "dxva2.dll" (ByVal hMonitor As Long, ByVal dwPhysicalMonitorArraySize As Long, pPhysicalMonitorArray As PHYSICAL_MONITOR) As Long
Private Declare Function DestroyPhysicalMonitors Lib "dxva2.dll" (ByVal dwPhysicalMonitorArraySize As Long, pPhysicalMonitorArray As PHYSICAL_MONITOR) As Long
Private Declare Function GetVCPFeatureAndVCPFeatureReply Lib "dxva2.dll" (ByVal hMonitor As Long, ByVal bVCPCode As Byte, pdwVCPCodeType As Long, pdwCurrentValue As Long, pdwMaximumValue As Long) As Long

Private m_MonitorCount As Long                                                  ' 記錄實體螢幕總數
Private m_PowerOnCount As Long                                                  ' 記錄通電中的螢幕數
Private m_StandbyCount As Long                                                  ' 待命中台數



Private Sub EnumAll()
    m_MonitorCount = 0
    m_PowerOnCount = 0
    m_StandbyCount = 0                                                          ' 單次列舉，多組計數
    EnumDisplayMonitors 0, 0, AddressOf MonitorEnumProc, 0                      ' 列舉所有顯示器
End Sub

Public Function WmiPhysicalCount() As Long
    EnumAll
    WmiPhysicalCount = m_PowerOnCount                                           ' DDC/CI D6 通電台數
End Function

Public Function WmiStandbyCount() As Long
    WmiStandbyCount = m_StandbyCount                                            ' 需先呼叫 WmiPhysicalCount
End Function

Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByVal lprcMonitor As Long, ByVal dwData As Long) As Long
    Dim physMon As PHYSICAL_MONITOR, vct As Long, currentVal As Long, maxVal As Long
    If GetPhysicalMonitorsFromHMONITOR(hMonitor, 1, physMon) Then
        m_MonitorCount = m_MonitorCount + 1                                     ' 發現實體螢幕
        If GetVCPFeatureAndVCPFeatureReply(physMon.hPhysicalMonitor, &HD6, vct, currentVal, maxVal) Then
            If currentVal = 1 Then
                m_PowerOnCount = m_PowerOnCount + 1                                          ' D6=1 通電中
            ElseIf currentVal = 2 Or currentVal = 3 Then
                m_StandbyCount = m_StandbyCount + 1                                          ' D6=2/3 待命
            End If
        End If
        DestroyPhysicalMonitors 1, physMon                                      ' 釋放硬體控制代碼資源
    End If
    MonitorEnumProc = 1                                                         ' 回傳 1 繼續列舉下一個螢幕
End Function

