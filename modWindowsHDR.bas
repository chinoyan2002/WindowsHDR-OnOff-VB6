Attribute VB_Name = "modWindowsHDR"
Option Explicit

'=====================================================================
' Windows HDR / Advanced Color 原生 API 模組
'
' 設計原則：
' 1. 完全直接呼叫 Windows Display Configuration API。
' 2. Windows 11 24H2（Build 26100）以上使用新版 HDR API。
' 3. 舊版 Windows 使用 Advanced Color 相容 API，但不誤稱為精確 HDR API。
' 4. 所有 Set 操作完成後都會重新查詢並驗證實際狀態。
' 5. 所有錯誤與狀態文字均使用繁體中文。
' 6. VB6 使用 Byte Array + RtlMoveMemory，避免 UDT / Union / Bit Field 對齊問題。
'=====================================================================

'---------------------------------------------------------------------
' Windows API 宣告
'---------------------------------------------------------------------

' 取得目前 Display Configuration 所需的 Path / Mode Buffer 大小。
Private Declare Function GetDisplayConfigBufferSizes Lib "user32.dll" ( _
    ByVal flags As Long, _
    ByRef numPathArrayElements As Long, _
    ByRef numModeInfoArrayElements As Long) As Long

' 查詢目前 Display Configuration。
' currentTopologyId 在 QDC_ONLY_ACTIVE_PATHS 模式下必須傳入 NULL。
Private Declare Function QueryDisplayConfig Lib "user32.dll" ( _
    ByVal flags As Long, _
    ByRef numPathArrayElements As Long, _
    ByRef pathArray As Any, _
    ByRef numModeInfoArrayElements As Long, _
    ByRef modeInfoArray As Any, _
    ByVal currentTopologyId As Long) As Long

' 取得指定 Display Target 的裝置資訊。
Private Declare Function DisplayConfigGetDeviceInfo Lib "user32.dll" ( _
    ByRef requestPacket As Any) As Long

' 設定指定 Display Target 的裝置資訊。
Private Declare Function DisplayConfigSetDeviceInfo Lib "user32.dll" ( _
    ByRef requestPacket As Any) As Long

' 直接複製記憶體，用於 Byte Array 與 32-bit 整數之間的安全轉換。
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

' 設定完成後短暫等待 Windows 更新 Display Configuration。
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 取得實際 Windows OS Build，不依賴相容性層可能被虛擬化的版本字串。
Private Declare Function RtlGetVersion Lib "ntdll.dll" ( _
    ByRef lpVersionInformation As Any) As Long

'---------------------------------------------------------------------
' Win32 錯誤碼
'---------------------------------------------------------------------

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_INVALID_FUNCTION As Long = 1
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const ERROR_NOT_SUPPORTED As Long = 50
Private Const ERROR_NOT_FOUND As Long = 1168
Private Const ERROR_GEN_FAILURE As Long = 31

'---------------------------------------------------------------------
' Display Configuration 常數
'---------------------------------------------------------------------

' 只查詢目前 active 的 Display Path。
Private Const QDC_ONLY_ACTIVE_PATHS As Long = 2

' DISPLAYCONFIG_DEVICE_INFO_TYPE。
Private Const DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME As Long = 2
Private Const DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO As Long = 9
Private Const DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE As Long = 10
Private Const DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2 As Long = 15
Private Const DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE As Long = 16

' Display Configuration Path flags。
Private Const DISPLAYCONFIG_PATH_ACTIVE As Long = 1

' Windows 11 HDR / Advanced Color Mode。
Private Const DISPLAYCONFIG_ADVANCED_COLOR_MODE_SDR As Long = 0
Private Const DISPLAYCONFIG_ADVANCED_COLOR_MODE_WCG As Long = 1
Private Const DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR As Long = 2

'---------------------------------------------------------------------
' 結構大小
'---------------------------------------------------------------------

' DISPLAYCONFIG_PATH_INFO：sourceInfo 20 + targetInfo 48 + flags 4。
Private Const DISPLAYCONFIG_PATH_INFO_SIZE As Long = 72

' DISPLAYCONFIG_MODE_INFO 的 32-bit Windows API 結構固定配置大小。
Private Const DISPLAYCONFIG_MODE_INFO_SIZE As Long = 64

' DISPLAYCONFIG_TARGET_DEVICE_NAME：
' Header 20 + Flags 4 + OutputTechnology 4 + EDID IDs 4 + Connector 4
' + Friendly Name 128 + Device Path 256 = 420 bytes。
Private Const DISPLAYCONFIG_TARGET_DEVICE_NAME_SIZE As Long = 420

' 舊版 DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO。
Private Const DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_SIZE As Long = 32

' 舊版 DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE。
Private Const DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE_SIZE As Long = 24

' Windows 11 24H2+ DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2。
Private Const DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2_SIZE As Long = 36

' Windows 11 24H2+ DISPLAYCONFIG_SET_HDR_STATE。
Private Const DISPLAYCONFIG_SET_HDR_STATE_SIZE As Long = 24

' RTL_OSVERSIONINFOEXW 在 32-bit VB6 下的完整結構大小。
Private Const RTL_OSVERSIONINFOW_SIZE As Long = 20

' Windows 11 24H2 對應的最低 Build。
Private Const WINDOWS_11_24H2_BUILD As Long = 26100

'---------------------------------------------------------------------
' DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2 的 bitfield
'
' 官方 SDK 定義：
' bit 0 advancedColorSupported
' bit 1 advancedColorActive
' bit 2 reserved
' bit 3 advancedColorLimitedByPolicy
' bit 4 highDynamicRangeSupported
' bit 5 highDynamicRangeUserEnabled
' bit 6 wideColorSupported
' bit 7 wideColorUserEnabled
'---------------------------------------------------------------------

Private Const AC2_ADVANCED_COLOR_SUPPORTED As Long = 0
Private Const AC2_ADVANCED_COLOR_ACTIVE As Long = 1
Private Const AC2_ADVANCED_COLOR_LIMITED_BY_POLICY As Long = 3
Private Const AC2_HDR_SUPPORTED As Long = 4
Private Const AC2_HDR_USER_ENABLED As Long = 5
Private Const AC2_WIDE_COLOR_SUPPORTED As Long = 6
Private Const AC2_WIDE_COLOR_USER_ENABLED As Long = 7

'---------------------------------------------------------------------
' DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO 的舊版 bitfield
'---------------------------------------------------------------------

Private Const AC1_ADVANCED_COLOR_SUPPORTED As Long = 0
Private Const AC1_ADVANCED_COLOR_ENABLED As Long = 1
Private Const AC1_WIDE_COLOR_ENFORCED As Long = 2
Private Const AC1_ADVANCED_COLOR_FORCE_DISABLED As Long = 3

'---------------------------------------------------------------------
' HDR 狀態列舉
'---------------------------------------------------------------------

Public Enum HDR_STATUS
    HDR_UNSUPPORTED = 0
    HDR_OFF = 1
    HDR_ON = 2
    HDR_ERROR = 3
    HDR_STATUS_UNKNOWN = 4
End Enum

'---------------------------------------------------------------------
' 顯示器資訊結構
'
' Index：供 VB6 UI 使用的陣列索引，不是 Windows Target ID。
' AdapterLow / AdapterHigh：Windows Display Configuration 的 Adapter LUID。
' TargetId：Windows Display Configuration Target Identifier。
' Name：顯示器 Friendly Name。
'
' HDRSupported：新版 API 精確取得的 HDR capability。
' HDRCapabilityKnown：HDRSupported 是否可信。
' HDRUserEnabled：Windows HDR 使用者設定狀態。
' HDRActive：實際 Active Color Mode 是否為 HDR。
'
' AdvancedColorSupported / Active：Advanced Color 能力與實際狀態。
' AdvancedColorLimitedByPolicy：Advanced Color 是否受到政策限制。
' AdvancedColorForceDisabled：舊版 API 的強制停用狀態。
'
' WCGSupported / WCGUserEnabled：Wide Color / WCG 資訊。
' WideColorEnforced：舊版 API 的 Wide Color 強制狀態。
'
' ActiveColorMode：0 SDR、1 WCG、2 HDR。
' FriendlyNameFromEdid / Forced：Friendly Name 來源旗標。
' LastError：此顯示器最後一次 API 操作錯誤碼。
'---------------------------------------------------------------------

Public Type HDR_DISPLAY_INFO
    ' UI 顯示索引。
    Index As Long

    ' Windows Adapter LUID。
    AdapterLow As Long
    AdapterHigh As Long

    ' Windows Display Target ID。
    TargetId As Long

    ' 顯示器 Friendly Name。
    Name As String

    ' 新版 API 的 HDR capability。
    HDRSupported As Boolean
    ' False 表示目前 API 無法精確判定 HDR capability。
    HDRCapabilityKnown As Boolean
    ' Windows 使用者設定是否啟用 HDR。
    HDRUserEnabled As Boolean
    ' 實際 Active Color Mode 是否為 HDR。
    HDRActive As Boolean

    ' Advanced Color 相關能力與狀態。
    AdvancedColorSupported As Boolean
    AdvancedColorActive As Boolean
    AdvancedColorLimitedByPolicy As Boolean
    AdvancedColorForceDisabled As Boolean

    ' Wide Color / WCG 相關資訊。
    WCGSupported As Boolean
    WCGUserEnabled As Boolean
    WideColorEnforced As Boolean

    ' 0=SDR、1=WCG、2=HDR；舊版 API 為 -1。
    ActiveColorMode As Long

    ' Friendly Name 來源旗標。
    FriendlyNameFromEdid As Boolean
    FriendlyNameForced As Boolean

    ' 每一個 Display 的最後錯誤。
    LastError As Long
End Type

' 模組層級最後錯誤碼。
Private mLastError As Long

' 模組層級繁體中文錯誤訊息。
Private mLastErrorText As String

'=====================================================================
' 公開函數：取得最後錯誤碼
'=====================================================================

' 回傳模組最近一次 Windows API 操作的 Win32 錯誤碼。
Public Function HDR_GetLastError() As Long
    HDR_GetLastError = mLastError
End Function

'=====================================================================
' 公開函數：取得最後錯誤文字
'=====================================================================

' 回傳只包含繁體中文描述與錯誤碼的錯誤訊息。
Public Function HDR_GetLastErrorText() As String
    If Len(mLastErrorText) > 0 Then
        HDR_GetLastErrorText = mLastErrorText
    ElseIf mLastError = ERROR_SUCCESS Then
        HDR_GetLastErrorText = "操作成功。"
    Else
        HDR_GetLastErrorText = HDR_ErrorTextChinese(mLastError)
    End If
End Function

'=====================================================================
' 公開函數：取得 Windows Build
'=====================================================================

' 使用 RtlGetVersion 取得實際 Windows Build。
' osInfo：RTL_OSVERSIONINFOEXW 的 284-byte Byte Array。
' result：RtlGetVersion 回傳值。
Public Function HDR_GetWindowsVersionInfo( _
    ByRef MajorVersion As Long, _
    ByRef MinorVersion As Long, _
    ByRef BuildNumber As Long) As Boolean

    Dim osInfo() As Byte           ' RTL_OSVERSIONINFOW 的 20-byte 原始結構。
    Dim result As Long             ' RtlGetVersion 回傳值。

    MajorVersion = 0
    MinorVersion = 0
    BuildNumber = 0
    HDR_GetWindowsVersionInfo = False

    ReDim osInfo(0 To RTL_OSVERSIONINFOW_SIZE - 1)

    ' dwOSVersionInfoSize 位於 offset 0，必須先指定結構大小。
    HDR_WriteLong osInfo, 0, RTL_OSVERSIONINFOW_SIZE

    result = RtlGetVersion(osInfo(0))

    If result <> ERROR_SUCCESS Then
        HDR_SetLastError result, "取得 Windows 版本失敗。"
        Exit Function
    End If

    ' RTL_OSVERSIONINFOW 欄位：
    ' offset 0  = dwOSVersionInfoSize
    ' offset 4  = dwMajorVersion
    ' offset 8  = dwMinorVersion
    ' offset 12 = dwBuildNumber
    ' offset 16 = dwPlatformId
    MajorVersion = HDR_ReadLong(osInfo, 4)
    MinorVersion = HDR_ReadLong(osInfo, 8)
    BuildNumber = HDR_ReadLong(osInfo, 12)

    HDR_ClearLastError
    HDR_GetWindowsVersionInfo = True
End Function

'=====================================================================
' 公開函數：取得 Windows Build
'=====================================================================

' 只回傳目前 Windows 的 Build Number。
Public Function HDR_GetWindowsBuild() As Long
    Dim majorVersion As Long         ' Windows Major Version。
    Dim minorVersion As Long         ' Windows Minor Version。
    Dim buildNumber As Long          ' Windows Build Number。

    If HDR_GetWindowsVersionInfo(majorVersion, minorVersion, buildNumber) Then
        HDR_GetWindowsBuild = buildNumber
    Else
        HDR_GetWindowsBuild = 0
    End If
End Function

'=====================================================================
' 公開函數：判斷是否使用 Windows 11 24H2+ API
'=====================================================================

' 依 Windows Build 判斷是否可使用新版 HDR API。
Public Function HDR_IsWindows11_24H2() As Boolean
    Dim majorVersion As Long         ' Windows Major Version。
    Dim minorVersion As Long         ' Windows Minor Version。
    Dim buildNumber As Long          ' Windows Build Number。

    If Not HDR_GetWindowsVersionInfo(majorVersion, minorVersion, buildNumber) Then
        HDR_IsWindows11_24H2 = False
        Exit Function
    End If

    ' Windows 10 / 11 都使用 Major=10；再以 Build 判斷 24H2+ HDR API。
    HDR_IsWindows11_24H2 = (majorVersion >= 10 And buildNumber >= WINDOWS_11_24H2_BUILD)
End Function

'=====================================================================
' 公開函數：取得目前 API 模式文字
'=====================================================================

' 回傳目前程式使用的 HDR API 路徑。
Public Function HDR_GetApiModeText() As String
    If HDR_IsWindows11_24H2() Then
        HDR_GetApiModeText = "Windows 11 24H2 以上：新版 HDR API"
    Else
        HDR_GetApiModeText = "舊版 Windows：Advanced Color 相容 API"
    End If
End Function

'=====================================================================
' 公開函數：取得顯示器數量
'=====================================================================

' 取得目前所有 active Display Path 的數量；失敗回傳 -1。
Public Function HDR_GetDisplayCount() As Long
    Dim displays() As HDR_DISPLAY_INFO       ' 所有 active 顯示器資訊。

    If Not HDR_GetDisplays(displays) Then
        HDR_GetDisplayCount = -1
        Exit Function
    End If

    If HDR_ArrayAllocated(displays) Then
        HDR_GetDisplayCount = UBound(displays) - LBound(displays) + 1
    Else
        HDR_GetDisplayCount = 0
    End If
End Function

'=====================================================================
' 公開函數：取得所有 active 顯示器資訊
'=====================================================================

' 查詢所有目前 active 的顯示器，並填入 Displays 陣列。
' Displays：輸出的顯示器資訊陣列。
Public Function HDR_GetDisplays(ByRef Displays() As HDR_DISPLAY_INFO) As Boolean
    Dim pathBuffer() As Byte         ' DISPLAYCONFIG_PATH_INFO 原始 Buffer。
    Dim modeBuffer() As Byte         ' DISPLAYCONFIG_MODE_INFO 原始 Buffer。
    Dim pathCount As Long            ' Active Path 數量。
    Dim modeCount As Long            ' Mode Info 數量。
    Dim i As Long                    ' 顯示器迴圈索引。

    Erase Displays
    HDR_ClearLastError

    If Not HDR_QueryActiveDisplays(pathBuffer, pathCount, modeBuffer, modeCount) Then Exit Function

    If pathCount <= 0 Then
        HDR_GetDisplays = True
        Exit Function
    End If

    ReDim Displays(0 To pathCount - 1)

    For i = 0 To pathCount - 1
        Displays(i) = HDR_CreateDisplayInfo(i, pathBuffer)
    Next i

    HDR_GetDisplays = True
End Function

'=====================================================================
' 公開函數：取得指定顯示器 HDR 狀態
'=====================================================================

' 依 UI Index 取得 HDR 狀態。
' DisplayIndex：HDR_GetDisplays 回傳陣列中的索引。
Public Function HDR_GetStatus(ByVal DisplayIndex As Long) As HDR_STATUS
    Dim displays() As HDR_DISPLAY_INFO       ' 目前所有顯示器。

    HDR_GetStatus = HDR_ERROR

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_NOT_FOUND, "目前沒有可用的作用中顯示器。"
        Exit Function
    End If

    If DisplayIndex < LBound(displays) Or DisplayIndex > UBound(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "指定的顯示器索引無效。"
        Exit Function
    End If

    If Not displays(DisplayIndex).HDRCapabilityKnown Then
        HDR_GetStatus = HDR_STATUS_UNKNOWN
        Exit Function
    End If

    If Not displays(DisplayIndex).HDRSupported Then
        HDR_GetStatus = HDR_UNSUPPORTED
    ElseIf displays(DisplayIndex).HDRActive Then
        HDR_GetStatus = HDR_ON
    Else
        HDR_GetStatus = HDR_OFF
    End If
End Function

'=====================================================================
' 公開函數：HDR 狀態文字
'=====================================================================

' 將 HDR_STATUS 轉成繁體中文狀態文字。
Public Function HDR_GetStatusText(ByVal Status As HDR_STATUS) As String
    Select Case Status
        Case HDR_UNSUPPORTED
            HDR_GetStatusText = "HDR 不支援"
        Case HDR_OFF
            HDR_GetStatusText = "HDR 已關閉"
        Case HDR_ON
            HDR_GetStatusText = "HDR 已開啟"
        Case HDR_ERROR
            HDR_GetStatusText = "HDR 操作錯誤"
        Case HDR_STATUS_UNKNOWN
            HDR_GetStatusText = "HDR 能力未知（目前 API 無法精確判定）"
        Case Else
            HDR_GetStatusText = "未知狀態"
    End Select
End Function

'=====================================================================
' 公開函數：開啟所有顯示器 HDR
'=====================================================================

' 對所有 active 顯示器執行 HDR ON，完成後重新查詢驗證。
Public Function HDR_Enable() As Boolean
    HDR_Enable = HDR_SetAllDisplays(True)
End Function

'=====================================================================
' 公開函數：關閉所有顯示器 HDR
'=====================================================================

' 對所有 active 顯示器執行 HDR OFF，完成後重新查詢驗證。
Public Function HDR_Disable() As Boolean
    HDR_Disable = HDR_SetAllDisplays(False)
End Function

'=====================================================================
' 公開函數：切換所有顯示器 HDR
'=====================================================================

' 依每一個顯示器目前實際狀態逐一切換 HDR。
Public Function HDR_Toggle() As Boolean
    Dim displays() As HDR_DISPLAY_INFO       ' 所有 active 顯示器。
    Dim i As Long                            ' 顯示器索引。
    Dim okAll As Boolean                     ' 是否全部成功。
    Dim firstError As Long                   ' 第一個錯誤碼。
    Dim targetState As Boolean               ' 此顯示器的目標狀態。

    HDR_Toggle = False

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_NOT_FOUND, "目前沒有可用的作用中顯示器。"
        Exit Function
    End If

    okAll = True
    firstError = ERROR_SUCCESS

    For i = LBound(displays) To UBound(displays)
        If HDR_IsWindows11_24H2() Then
            ' 新版 API 必須有可確認的 HDR capability。
            If Not displays(i).HDRCapabilityKnown Or Not displays(i).HDRSupported Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = ERROR_NOT_SUPPORTED
            Else
                targetState = Not displays(i).HDRActive

                If Not HDR_SetDisplay(displays(i), targetState) Then
                    okAll = False
                    If firstError = ERROR_SUCCESS Then firstError = mLastError
                End If
            End If
        Else
            ' 舊版只能切換 Advanced Color。
            targetState = Not displays(i).AdvancedColorActive

            If Not HDR_SetDisplay(displays(i), targetState) Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = mLastError
            End If
        End If
    Next i

    If okAll Then
        HDR_ClearLastError
    Else
        If firstError = ERROR_SUCCESS Then firstError = ERROR_GEN_FAILURE
        HDR_SetLastError firstError, "部分或全部顯示器切換失敗。"
    End If

    HDR_Toggle = okAll
End Function

'=====================================================================
' 公開函數：指定顯示器 HDR ON / OFF
'=====================================================================

' 依 UI Index 設定單一顯示器 HDR。
' DisplayIndex：目標顯示器索引。
' EnableHDR：True=開啟；False=關閉。
Public Function HDR_SetDisplayByIndex(ByVal DisplayIndex As Long, ByVal EnableHDR As Boolean) As Boolean
    Dim displays() As HDR_DISPLAY_INFO       ' 目前所有顯示器。

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_NOT_FOUND, "目前沒有可用的作用中顯示器。"
        Exit Function
    End If

    If DisplayIndex < LBound(displays) Or DisplayIndex > UBound(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "指定的顯示器索引無效。"
        Exit Function
    End If

    HDR_SetDisplayByIndex = HDR_SetDisplay(displays(DisplayIndex), EnableHDR)
End Function

'=====================================================================
' 私有函數：查詢 Active Display Configuration
'=====================================================================

' 取得 Path / Mode Buffer 並查詢 active Display Path。
' Windows 顯示組態可能在兩次 API 呼叫間改變，因此遇到
' ERROR_INSUFFICIENT_BUFFER 時最多重試三次。
Private Function HDR_QueryActiveDisplays( _
    ByRef pathBuffer() As Byte, _
    ByRef pathCount As Long, _
    ByRef modeBuffer() As Byte, _
    ByRef modeCount As Long) As Boolean

    Dim attempt As Long               ' 目前第幾次嘗試。
    Dim result As Long                ' API 回傳值。
    Dim requestedPaths As Long        ' API 要求的 Path 數量。
    Dim requestedModes As Long        ' API 要求的 Mode 數量。
    Dim pathBytes As Long             ' Path Buffer bytes。
    Dim modeBytes As Long             ' Mode Buffer bytes。

    Erase pathBuffer
    Erase modeBuffer
    pathCount = 0
    modeCount = 0

    For attempt = 1 To 3
        requestedPaths = 0
        requestedModes = 0

        result = GetDisplayConfigBufferSizes( _
                    QDC_ONLY_ACTIVE_PATHS, _
                    requestedPaths, _
                    requestedModes)

        If result <> ERROR_SUCCESS Then
            HDR_SetLastError result, "取得顯示組態 Buffer 大小失敗。"
            Exit Function
        End If

        pathCount = requestedPaths
        modeCount = requestedModes

        ' VB6 不允許零長度 Byte Array，因此最少配置一個元素。
        If pathCount < 1 Then pathCount = 1
        If modeCount < 1 Then modeCount = 1

        pathBytes = pathCount * DISPLAYCONFIG_PATH_INFO_SIZE
        modeBytes = modeCount * DISPLAYCONFIG_MODE_INFO_SIZE

        ReDim pathBuffer(0 To pathBytes - 1)
        ReDim modeBuffer(0 To modeBytes - 1)

        ' QDC_ONLY_ACTIVE_PATHS 時 currentTopologyId 必須是 NULL。
        result = QueryDisplayConfig( _
                    QDC_ONLY_ACTIVE_PATHS, _
                    pathCount, _
                    pathBuffer(0), _
                    modeCount, _
                    modeBuffer(0), _
                    0&)

        If result = ERROR_SUCCESS Then
            HDR_ClearLastError
            HDR_QueryActiveDisplays = True
            Exit Function
        End If

        If result <> ERROR_INSUFFICIENT_BUFFER Then
            HDR_SetLastError result, "查詢目前顯示組態失敗。"
            Exit Function
        End If
    Next attempt

    HDR_SetLastError ERROR_INSUFFICIENT_BUFFER, _
                     "顯示組態在查詢期間持續變更，連續三次取得 Buffer 都不足。"
End Function

'=====================================================================
' 私有函數：建立單一顯示器資訊
'=====================================================================

' 從 DISPLAYCONFIG_PATH_INFO Byte Array 建立 HDR_DISPLAY_INFO。
' Index：目前 active path 的陣列索引。
' pathBuffer：QueryDisplayConfig 回傳的 Path Buffer。
Private Function HDR_CreateDisplayInfo(ByVal Index As Long, ByRef pathBuffer() As Byte) As HDR_DISPLAY_INFO
    Dim info As HDR_DISPLAY_INFO       ' 輸出的顯示器資訊。
    Dim offset As Long                 ' 此 Path 在 Buffer 中的起始 offset。
    Dim pathFlags As Long              ' Display Path flags。
    Dim nameError As Long              ' 顯示器名稱查詢錯誤碼。

    info.Index = Index
    info.ActiveColorMode = -1
    info.LastError = ERROR_SUCCESS

    offset = Index * DISPLAYCONFIG_PATH_INFO_SIZE

    ' DISPLAYCONFIG_PATH_INFO.targetInfo：
    ' adapterId low  offset 20
    ' adapterId high offset 24
    ' targetId       offset 28
    ' flags          offset 68
    info.AdapterLow = HDR_ReadLong(pathBuffer, offset + 20)
    info.AdapterHigh = HDR_ReadLong(pathBuffer, offset + 24)
    info.TargetId = HDR_ReadLong(pathBuffer, offset + 28)
    pathFlags = HDR_ReadLong(pathBuffer, offset + 68)

    If (pathFlags And DISPLAYCONFIG_PATH_ACTIVE) = 0 Then
        info.LastError = ERROR_INVALID_PARAMETER
        HDR_CreateDisplayInfo = info
        Exit Function
    End If

    ' Friendly Name 是輔助資訊，名稱失敗不應阻止 HDR 狀態取得。
    info.Name = HDR_GetDisplayName( _
                    info.AdapterLow, _
                    info.AdapterHigh, _
                    info.TargetId, _
                    info.FriendlyNameFromEdid, _
                    info.FriendlyNameForced, _
                    nameError)

    If HDR_IsWindows11_24H2() Then
        Call HDR_ReadDisplayStatus2(info)
    Else
        Call HDR_ReadDisplayStatusLegacy(info)
    End If

    If info.LastError = ERROR_SUCCESS And nameError <> ERROR_SUCCESS Then
        info.LastError = nameError
    End If

    HDR_CreateDisplayInfo = info
End Function

'=====================================================================
' 私有函數：讀取 Windows 11 24H2+ HDR 狀態
'=====================================================================

' 使用 DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2 讀取精確 HDR / WCG 狀態。
' info：要填入的顯示器資訊結構。
Private Function HDR_ReadDisplayStatus2(ByRef info As HDR_DISPLAY_INFO) As Boolean
    Dim packet(0 To DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2_SIZE - 1) As Byte ' API Packet。
    Dim result As Long                   ' API 回傳碼。
    Dim value As Long                    ' 4-byte bitfield。
    Dim activeMode As Long               ' 0=SDR、1=WCG、2=HDR。

    HDR_InitHeader packet, _
                   DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2, _
                   DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2_SIZE, _
                   info.AdapterLow, _
                   info.AdapterHigh, _
                   info.TargetId

    result = DisplayConfigGetDeviceInfo(packet(0))
    If result <> ERROR_SUCCESS Then
        info.LastError = result
        HDR_SetLastError result, "讀取新版 HDR 狀態失敗。"
        Exit Function
    End If

    ' bitfield 位於 offset 20。
    value = HDR_ReadLong(packet, 20)

    ' Active Color Mode 位於 offset 32。
    activeMode = HDR_ReadLong(packet, 32)

    info.AdvancedColorSupported = HDR_TestBit(value, AC2_ADVANCED_COLOR_SUPPORTED)
    info.AdvancedColorActive = HDR_TestBit(value, AC2_ADVANCED_COLOR_ACTIVE)
    info.AdvancedColorLimitedByPolicy = HDR_TestBit(value, AC2_ADVANCED_COLOR_LIMITED_BY_POLICY)

    info.HDRSupported = HDR_TestBit(value, AC2_HDR_SUPPORTED)
    info.HDRCapabilityKnown = True
    info.HDRUserEnabled = HDR_TestBit(value, AC2_HDR_USER_ENABLED)

    ' 實際 HDR 必須看 Active Color Mode，不只看 User Enabled。
    info.HDRActive = (activeMode = DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR)

    info.WCGSupported = HDR_TestBit(value, AC2_WIDE_COLOR_SUPPORTED)
    info.WCGUserEnabled = HDR_TestBit(value, AC2_WIDE_COLOR_USER_ENABLED)

    info.AdvancedColorForceDisabled = False
    info.WideColorEnforced = False

    If activeMode >= DISPLAYCONFIG_ADVANCED_COLOR_MODE_SDR And _
       activeMode <= DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR Then
        info.ActiveColorMode = activeMode
    Else
        info.ActiveColorMode = -1
    End If

    info.LastError = ERROR_SUCCESS
    HDR_ClearLastError
    HDR_ReadDisplayStatus2 = True
End Function

'=====================================================================
' 私有函數：讀取舊版 Advanced Color 狀態
'=====================================================================

' 使用舊版 API 讀取 Advanced Color。
' 舊版 API 無法可靠區分 WCG 與 HDR，因此 HDR capability 必須標記 Unknown。
Private Function HDR_ReadDisplayStatusLegacy(ByRef info As HDR_DISPLAY_INFO) As Boolean
    Dim packet(0 To DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_SIZE - 1) As Byte ' API Packet。
    Dim result As Long                   ' API 回傳碼。
    Dim value As Long                    ' 4-byte bitfield。

    HDR_InitHeader packet, _
                   DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO, _
                   DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_SIZE, _
                   info.AdapterLow, _
                   info.AdapterHigh, _
                   info.TargetId

    result = DisplayConfigGetDeviceInfo(packet(0))
    If result <> ERROR_SUCCESS Then
        info.LastError = result
        HDR_SetLastError result, "讀取舊版 Advanced Color 狀態失敗。"
        Exit Function
    End If

    value = HDR_ReadLong(packet, 20)

    info.AdvancedColorSupported = HDR_TestBit(value, AC1_ADVANCED_COLOR_SUPPORTED)
    info.AdvancedColorActive = HDR_TestBit(value, AC1_ADVANCED_COLOR_ENABLED)
    info.WideColorEnforced = HDR_TestBit(value, AC1_WIDE_COLOR_ENFORCED)
    info.AdvancedColorForceDisabled = HDR_TestBit(value, AC1_ADVANCED_COLOR_FORCE_DISABLED)

    ' 舊 API 沒有 highDynamicRangeSupported 欄位。
    info.HDRCapabilityKnown = False
    info.HDRSupported = False
    info.HDRUserEnabled = False
    info.HDRActive = False
    info.WCGSupported = False
    info.WCGUserEnabled = False
    info.AdvancedColorLimitedByPolicy = False
    info.ActiveColorMode = -1

    info.LastError = ERROR_SUCCESS
    HDR_ClearLastError
    HDR_ReadDisplayStatusLegacy = True
End Function

'=====================================================================
' 私有函數：設定單一顯示器
'=====================================================================

' 依 Windows 版本使用正確的 Set API，並在 Set 後重新 Query 驗證。
' info：要設定的顯示器。
' EnableHDR：True=開啟；False=關閉。
Private Function HDR_SetDisplay(ByRef info As HDR_DISPLAY_INFO, ByVal EnableHDR As Boolean) As Boolean
    Dim packet(0 To DISPLAYCONFIG_SET_HDR_STATE_SIZE - 1) As Byte ' 新版 HDR Packet。
    Dim packetLegacy(0 To DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE_SIZE - 1) As Byte ' 舊版 Packet。
    Dim result As Long                       ' API 回傳值。

    If HDR_IsWindows11_24H2() Then
        ' 新版 API 不得 fallback 到舊版 Advanced Color API。
        If Not info.HDRCapabilityKnown Then
            HDR_SetLastError ERROR_NOT_SUPPORTED, "無法取得此顯示器的 HDR 能力。"
            Exit Function
        End If

        If EnableHDR And Not info.HDRSupported Then
            HDR_SetLastError ERROR_NOT_SUPPORTED, "此顯示器不支援 HDR。"
            Exit Function
        End If

        If Not EnableHDR And Not info.HDRSupported Then
            ' 不支援 HDR 的顯示器本身就不會處於 HDR。
            HDR_ClearLastError
            HDR_SetDisplay = True
            Exit Function
        End If

        HDR_InitHeader packet, _
                       DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE, _
                       DISPLAYCONFIG_SET_HDR_STATE_SIZE, _
                       info.AdapterLow, _
                       info.AdapterHigh, _
                       info.TargetId

        If EnableHDR Then
            HDR_WriteLong packet, 20, 1
        Else
            HDR_WriteLong packet, 20, 0
        End If

        result = DisplayConfigSetDeviceInfo(packet(0))
        If result <> ERROR_SUCCESS Then
            info.LastError = result
            HDR_SetLastError result, "設定新版 HDR 狀態失敗。"
            Exit Function
        End If
    Else
        ' 舊版 API 實際控制的是 Advanced Color，相容名稱保留 HDR_Enable/HDR_Disable。
        HDR_InitHeader packetLegacy, _
                       DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE, _
                       DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE_SIZE, _
                       info.AdapterLow, _
                       info.AdapterHigh, _
                       info.TargetId

        If EnableHDR Then
            HDR_WriteLong packetLegacy, 20, 1
        Else
            HDR_WriteLong packetLegacy, 20, 0
        End If

        result = DisplayConfigSetDeviceInfo(packetLegacy(0))
        If result <> ERROR_SUCCESS Then
            info.LastError = result
            HDR_SetLastError result, "設定舊版 Advanced Color 狀態失敗。"
            Exit Function
        End If
    End If

    ' Set API 成功不代表實際狀態已完成，必須重新查詢驗證。
    If HDR_VerifyDisplay(info, EnableHDR) Then
        HDR_ClearLastError
        HDR_SetDisplay = True
    End If
End Function

'=====================================================================
' 私有函數：設定所有顯示器
'=====================================================================

' 對所有 active 顯示器設定相同目標狀態。
' EnableHDR：True=開啟；False=關閉。
Private Function HDR_SetAllDisplays(ByVal EnableHDR As Boolean) As Boolean
    Dim displays() As HDR_DISPLAY_INFO       ' 所有 active 顯示器。
    Dim i As Long                            ' 顯示器索引。
    Dim okAll As Boolean                     ' 是否全部成功。
    Dim firstError As Long                   ' 第一個錯誤碼。

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_NOT_FOUND, "目前沒有可用的作用中顯示器。"
        Exit Function
    End If

    okAll = True
    firstError = ERROR_SUCCESS

    For i = LBound(displays) To UBound(displays)
        If HDR_IsWindows11_24H2() Then
            If EnableHDR And Not displays(i).HDRSupported Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = ERROR_NOT_SUPPORTED
            ElseIf Not HDR_SetDisplay(displays(i), EnableHDR) Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = mLastError
            End If
        Else
            If Not HDR_SetDisplay(displays(i), EnableHDR) Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = mLastError
            End If
        End If
    Next i

    If okAll Then
        HDR_ClearLastError
    Else
        If firstError = ERROR_SUCCESS Then firstError = ERROR_GEN_FAILURE
        HDR_SetLastError firstError, "部分或全部顯示器設定失敗。"
    End If

    HDR_SetAllDisplays = okAll
End Function

'=====================================================================
' 私有函數：設定後驗證
'=====================================================================

' 重新列舉 Display Configuration，找到原 Adapter LUID + Target ID，
' 並確認設定後的實際狀態。
' OriginalInfo：原始顯示器資訊，成功後更新為最新資訊。
' ExpectedHDR：新版 API 預期的 HDR Active；舊版則代表 Advanced Color Active。
Private Function HDR_VerifyDisplay(ByRef OriginalInfo As HDR_DISPLAY_INFO, ByVal ExpectedHDR As Boolean) As Boolean
    Dim displays() As HDR_DISPLAY_INFO       ' 重新查詢後的顯示器清單。
    Dim retry As Long                        ' 驗證重試次數。
    Dim i As Long                            ' 清單索引。
    Dim found As Boolean                     ' 是否找到相同 Target。
    Dim current As HDR_DISPLAY_INFO          ' 最新顯示器資訊。

    For retry = 1 To 3
        If HDR_GetDisplays(displays) Then
            found = False

            If HDR_ArrayAllocated(displays) Then
                For i = LBound(displays) To UBound(displays)
                    If displays(i).AdapterLow = OriginalInfo.AdapterLow And _
                       displays(i).AdapterHigh = OriginalInfo.AdapterHigh And _
                       displays(i).TargetId = OriginalInfo.TargetId Then

                        current = displays(i)
                        found = True
                        Exit For
                    End If
                Next i
            End If

            If found Then
                OriginalInfo = current

                If HDR_IsWindows11_24H2() Then
                    If ExpectedHDR Then
                        If current.HDRActive Then
                            HDR_VerifyDisplay = True
                            Exit Function
                        End If
                    Else
                        If Not current.HDRActive Then
                            HDR_VerifyDisplay = True
                            Exit Function
                        End If
                    End If
                Else
                    ' 舊版驗證的是 Advanced Color Active。
                    If current.AdvancedColorActive = ExpectedHDR Then
                        HDR_VerifyDisplay = True
                        Exit Function
                    End If
                End If
            Else
                HDR_SetLastError ERROR_NOT_FOUND, "設定後找不到原本的顯示器 Target。"
            End If
        End If

        Sleep 100
    Next retry

    If HDR_IsWindows11_24H2() Then
        HDR_SetLastError ERROR_GEN_FAILURE, "設定完成，但重新查詢後 HDR 實際狀態未符合預期。"
    Else
        HDR_SetLastError ERROR_GEN_FAILURE, "設定完成，但重新查詢後 Advanced Color 實際狀態未符合預期。"
    End If
End Function

'=====================================================================
' 私有函數：取得顯示器名稱
'=====================================================================

' 使用 DISPLAYCONFIG_TARGET_DEVICE_NAME 取得 Friendly Name 與來源旗標。
' ErrorCode：回傳取得名稱時的 API 錯誤碼。
Private Function HDR_GetDisplayName( _
    ByVal adapterLow As Long, _
    ByVal adapterHigh As Long, _
    ByVal targetId As Long, _
    ByRef FriendlyNameFromEdid As Boolean, _
    ByRef FriendlyNameForced As Boolean, _
    ByRef ErrorCode As Long) As String

    Dim packet(0 To DISPLAYCONFIG_TARGET_DEVICE_NAME_SIZE - 1) As Byte ' 420-byte Packet。
    Dim result As Long                       ' API 回傳值。
    Dim flags As Long                        ' Friendly Name flags。
    Dim nameText As String                   ' 顯示器名稱。

    FriendlyNameFromEdid = False
    FriendlyNameForced = False
    ErrorCode = ERROR_SUCCESS

    HDR_InitHeader packet, _
                   DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME, _
                   DISPLAYCONFIG_TARGET_DEVICE_NAME_SIZE, _
                   adapterLow, _
                   adapterHigh, _
                   targetId

    result = DisplayConfigGetDeviceInfo(packet(0))
    If result <> ERROR_SUCCESS Then
        ErrorCode = result
        HDR_GetDisplayName = "顯示器 " & CStr(targetId)
        Exit Function
    End If

    ' flags 位於 offset 20。
    flags = HDR_ReadLong(packet, 20)

    ' bit 0：Friendly Name 由 EDID 組成。
    FriendlyNameFromEdid = HDR_TestBit(flags, 0)

    ' bit 1：Forced Target。
    FriendlyNameForced = HDR_TestBit(flags, 1)

    ' Header 20 + flags 4 + outputTechnology 4 + manufacture 2
    ' + product 2 + connectorInstance 4 = offset 36。
    nameText = HDR_ReadUnicodeString(packet, 36, 64)

    If Len(nameText) = 0 Then
        nameText = "顯示器 " & CStr(targetId)
    End If

    HDR_GetDisplayName = nameText
End Function

'=====================================================================
' 私有函數：初始化 DISPLAYCONFIG_DEVICE_INFO_HEADER
'=====================================================================

' 將 Windows Device Info Header 寫入 Byte Array。
' packet：輸出的 API Packet。
' infoType：DISPLAYCONFIG_DEVICE_INFO_TYPE。
' packetSize：完整 Packet 大小。
' adapterLow / adapterHigh：Adapter LUID。
' targetId：Display Target ID。
Private Sub HDR_InitHeader( _
    ByRef packet() As Byte, _
    ByVal infoType As Long, _
    ByVal packetSize As Long, _
    ByVal adapterLow As Long, _
    ByVal adapterHigh As Long, _
    ByVal targetId As Long)

    HDR_WriteLong packet, 0, infoType
    HDR_WriteLong packet, 4, packetSize
    HDR_WriteLong packet, 8, adapterLow
    HDR_WriteLong packet, 12, adapterHigh
    HDR_WriteLong packet, 16, targetId
End Sub

'=====================================================================
' 私有函數：從 Byte Array 讀取 32-bit Long
'=====================================================================

' 以原始記憶體方式讀取 Windows API 的 32-bit UINT32。
Private Function HDR_ReadLong(ByRef buffer() As Byte, ByVal Offset As Long) As Long
    Dim value As Long                 ' 暫存 32-bit 整數。

    CopyMemory value, buffer(Offset), 4
    HDR_ReadLong = value
End Function

'=====================================================================
' 私有函數：寫入 32-bit Long
'=====================================================================

' 以原始記憶體方式將 32-bit Long 寫入 Windows API Packet。
Private Sub HDR_WriteLong(ByRef buffer() As Byte, ByVal Offset As Long, ByVal value As Long)
    CopyMemory buffer(Offset), value, 4
End Sub

'=====================================================================
' 私有函數：讀取 16-bit Unicode Code Unit
'=====================================================================

' 讀取 Little-Endian WCHAR 的 16-bit code unit。
Private Function HDR_ReadUInt16(ByRef buffer() As Byte, ByVal Offset As Long) As Long
    HDR_ReadUInt16 = CLng(buffer(Offset)) + (CLng(buffer(Offset + 1)) * 256)
End Function

'=====================================================================
' 私有函數：讀取 Unicode 字串
'=====================================================================

' 從 Byte Array 讀取固定長度的 UTF-16LE WCHAR 陣列。
' CharacterCount：最大 WCHAR 數量。
Private Function HDR_ReadUnicodeString( _
    ByRef buffer() As Byte, _
    ByVal Offset As Long, _
    ByVal CharacterCount As Long) As String

    Dim i As Long                    ' WCHAR 索引。
    Dim codeUnit As Long             ' 目前 WCHAR。
    Dim resultText As String         ' 組成的中文字串。

    For i = 0 To CharacterCount - 1
        codeUnit = HDR_ReadUInt16(buffer, Offset + (i * 2))

        If codeUnit = 0 Then Exit For

        resultText = resultText & ChrW$(codeUnit)
    Next i

    HDR_ReadUnicodeString = resultText
End Function

'=====================================================================
' 私有函數：測試 bitfield
'=====================================================================

' 測試 32-bit Long 指定 bit 是否為 1。
Private Function HDR_TestBit(ByVal value As Long, ByVal bitIndex As Long) As Boolean
    Select Case bitIndex
        Case 0
            HDR_TestBit = ((value And 1) <> 0)
        Case 1
            HDR_TestBit = ((value And 2) <> 0)
        Case 2
            HDR_TestBit = ((value And 4) <> 0)
        Case 3
            HDR_TestBit = ((value And 8) <> 0)
        Case 4
            HDR_TestBit = ((value And 16) <> 0)
        Case 5
            HDR_TestBit = ((value And 32) <> 0)
        Case 6
            HDR_TestBit = ((value And 64) <> 0)
        Case 7
            HDR_TestBit = ((value And 128) <> 0)
        Case Else
            HDR_TestBit = False
    End Select
End Function

'=====================================================================
' 私有函數：清除錯誤
'=====================================================================

' 清除模組層級最後錯誤狀態。
Private Sub HDR_ClearLastError()
    mLastError = ERROR_SUCCESS
    mLastErrorText = vbNullString
End Sub

'=====================================================================
' 私有函數：設定錯誤
'=====================================================================

' 設定模組層級錯誤碼與繁體中文錯誤內容。
' ContextText：目前操作的中文上下文。
Private Sub HDR_SetLastError(ByVal ErrorCode As Long, ByVal ContextText As String)
    mLastError = ErrorCode

    If ErrorCode = ERROR_SUCCESS Then
        mLastErrorText = vbNullString
    ElseIf Len(ContextText) > 0 Then
        mLastErrorText = ContextText & "錯誤碼：" & CStr(ErrorCode) & "；" & HDR_ErrorTextChinese(ErrorCode)
    Else
        mLastErrorText = HDR_ErrorTextChinese(ErrorCode)
    End If
End Sub

'=====================================================================
' 私有函數：Win32 錯誤碼繁體中文化
'=====================================================================

' 將常用 Windows 錯誤碼轉成固定繁體中文，避免輸出英文系統訊息。
Private Function HDR_ErrorTextChinese(ByVal ErrorCode As Long) As String
    Select Case ErrorCode
        Case ERROR_SUCCESS
            HDR_ErrorTextChinese = "成功。"
        Case ERROR_INVALID_FUNCTION
            HDR_ErrorTextChinese = "要求的功能無法使用。"
        Case ERROR_ACCESS_DENIED
            HDR_ErrorTextChinese = "存取被拒絕。"
        Case ERROR_INVALID_PARAMETER
            HDR_ErrorTextChinese = "參數無效。"
        Case ERROR_INSUFFICIENT_BUFFER
            HDR_ErrorTextChinese = "提供的 Buffer 大小不足。"
        Case ERROR_NOT_SUPPORTED
            HDR_ErrorTextChinese = "目前系統或顯示器不支援此操作。"
        Case ERROR_NOT_FOUND
            HDR_ErrorTextChinese = "找不到指定的顯示器或 Target。"
        Case ERROR_GEN_FAILURE
            HDR_ErrorTextChinese = "一般性 Windows API 操作失敗。"
        Case Else
            HDR_ErrorTextChinese = "Windows API 發生未分類錯誤。"
    End Select
End Function

'=====================================================================
' 私有函數：判斷 UDT 動態陣列是否已配置
'=====================================================================

' 安全判斷 HDR_DISPLAY_INFO 動態陣列是否已配置。
Private Function HDR_ArrayAllocated(ByRef Value() As HDR_DISPLAY_INFO) As Boolean
    On Error GoTo NotAllocated

    HDR_ArrayAllocated = (UBound(Value) >= LBound(Value))
    Exit Function

NotAllocated:
    HDR_ArrayAllocated = False
End Function
