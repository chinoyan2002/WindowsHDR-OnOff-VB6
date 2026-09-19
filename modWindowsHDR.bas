Attribute VB_Name = "modWindowsHDR"
Option Explicit

'=====================================================================
' Windows HDR / Advanced Color ­ěĄÍ API ĽŇ˛Ő
'
' ĄD­nł]­pˇG
' 1. §ąĄţŞ˝±µ©IĄs Windows Display Configuration APIˇC
' 2. Windows 11 24H2ˇ]Build 26100ˇ^ĄH¤W¨ĎĄÎ·sŞ© HDR APIˇC
' 3. ÂÂŞ© Windows ¨ĎĄÎ Advanced Color ¬Ű®eĽŇ¦ˇˇA¦ý¤Ł«_ĄRşë˝T HDRˇC
' 4. ©Ň¦ł Set ľŢ§@§ą¦¨«áłŁĄ˛¶·­«·s¬d¸ß¨ĂĹçĂŇąę»ÚŞ¬şAˇC
' 5. ©Ň¦łąďĄ~Ş¬şA¤ĺ¦r»Pżů»~°T®§©T©w¨ĎĄÎÁcĹé¤¤¤ĺˇC
' 6. ¨ĎĄÎ Byte Array + RtlMoveMemoryˇAÁ×§K VB6 UDT / Union / Bit Field ąď»ô°ÝĂDˇC
'=====================================================================

'-----------------------------
' Win32 API «Ĺ§i
'-----------------------------

' ¨ú±oĄŘ«eĹăĄÜ˛ŐşA©Ň»Ý­nŞş Path »P Mode Buffer ¤j¤pˇC
Private Declare Function GetDisplayConfigBufferSizes Lib "user32.dll" ( _
    ByVal flags As Long, _
    ByRef numPathArrayElements As Long, _
    ByRef numModeInfoArrayElements As Long) As Long

' ¨ú±oĄŘ«eŞşĹăĄÜ¸ô®|»PĽŇ¦ˇ¸ę°TˇFĄ»ĽŇ˛ŐĄu¬d¸ßĄŘ«e±ŇĄÎŞş PathˇC
Private Declare Function QueryDisplayConfig Lib "user32.dll" ( _
    ByVal flags As Long, _
    ByRef numPathArrayElements As Long, _
    ByRef pathArray As Any, _
    ByRef numModeInfoArrayElements As Long, _
    ByRef modeInfoArray As Any, _
    ByVal currentTopologyId As Long) As Long

' ¨Ě Adapter LUID + Target ID ¨ú±o«ü©w Display ŞşĂBĄ~¸ę°TˇC
Private Declare Function DisplayConfigGetDeviceInfo Lib "user32.dll" ( _
    ByRef requestPacket As Any) As Long

' ¨Ě Adapter LUID + Target ID ­×§ď«ü©w Display ŞşĹăĄÜł]©wˇC
Private Declare Function DisplayConfigSetDeviceInfo Lib "user32.dll" ( _
    ByRef requestPacket As Any) As Long

' ĄH¦ě¤¸˛Ő¤č¦ˇ¦wĄţĹŞĽg Windows API µ˛şcˇC
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

' ł]©w«áµĄ«Ý Windows §ą¦¨ĹăĄÜ˛ŐşA§ó·sˇA¦A¶i¦ć­«·s¬d¸ßĹçĂŇˇC
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ¨ú±oąę»Ú Windows Ş©Ą»ˇAÁ×§K¨ü¬Ű®e©Ę«Ĺ§iĽvĹTˇC
Private Declare Function RtlGetVersion Lib "ntdll.dll" ( _
    ByRef lpVersionInformation As Any) As Long

'-----------------------------
' Win32 żů»~˝X
'-----------------------------

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_INVALID_FUNCTION As Long = 1
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const ERROR_GEN_FAILURE As Long = 31
Private Const ERROR_NOT_SUPPORTED As Long = 50
Private Const ERROR_NOT_FOUND As Long = 1168

'-----------------------------
' QueryDisplayConfig ±`ĽĆ
'-----------------------------

' Ąu¬d¸ßĄŘ«e±ŇĄÎŞşĹăĄÜ¸ô®|ˇF¦b¦ąĽŇ¦ˇ¤U currentTopologyId Ą˛¶·¬° NULLˇC
Private Const QDC_ONLY_ACTIVE_PATHS As Long = 2

'-----------------------------
' DISPLAYCONFIG µ˛şc¤j¤p
'-----------------------------

' 32-bit VB6 ¨ĎĄÎŞş Windows Display Configuration µ˛şc¤j¤pˇC
Private Const DISPLAYCONFIG_PATH_INFO_SIZE As Long = 72
Private Const DISPLAYCONFIG_MODE_INFO_SIZE As Long = 64
Private Const DISPLAYCONFIG_TARGET_DEVICE_NAME_SIZE As Long = 420
Private Const DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_SIZE As Long = 32
Private Const DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE_SIZE As Long = 24
Private Const DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2_SIZE As Long = 36
Private Const DISPLAYCONFIG_SET_HDR_STATE_SIZE As Long = 24

' RTL_OSVERSIONINFOEXW ¦b 32-bit VB6 ¤¤°t¸m 284 bytesˇC
Private Const RTL_OSVERSIONINFOEXW_SIZE As Long = 284

'-----------------------------
' DISPLAYCONFIG_DEVICE_INFO_TYPE
'-----------------------------

Private Const DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME As Long = 2
Private Const DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO As Long = 9
Private Const DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE As Long = 10
Private Const DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2 As Long = 15
Private Const DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE As Long = 16

'-----------------------------
' Active Color Mode
'-----------------------------

Private Const DISPLAYCONFIG_ADVANCED_COLOR_MODE_SDR As Long = 0
Private Const DISPLAYCONFIG_ADVANCED_COLOR_MODE_WCG As Long = 1
Private Const DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR As Long = 2

' DISPLAYCONFIG_PATH_INFO.targetInfo.status Şş ACTIVE bitˇC
Private Const DISPLAYCONFIG_PATH_ACTIVE As Long = 1

'-----------------------------
' GET_ADVANCED_COLOR_INFO_2 ¦ě¤¸Äć¦ě
'-----------------------------

Private Const AC2_ADVANCED_COLOR_SUPPORTED As Long = 0
Private Const AC2_ADVANCED_COLOR_ACTIVE As Long = 1
Private Const AC2_ADVANCED_COLOR_LIMITED_BY_POLICY As Long = 3
Private Const AC2_HDR_SUPPORTED As Long = 4
Private Const AC2_HDR_USER_ENABLED As Long = 5
Private Const AC2_WIDE_COLOR_SUPPORTED As Long = 6
Private Const AC2_WIDE_COLOR_USER_ENABLED As Long = 7

'-----------------------------
' ÂÂŞ© GET_ADVANCED_COLOR_INFO ¦ě¤¸Äć¦ě
'-----------------------------

Private Const AC1_ADVANCED_COLOR_SUPPORTED As Long = 0
Private Const AC1_ADVANCED_COLOR_ENABLED As Long = 1
Private Const AC1_WIDE_COLOR_ENFORCED As Long = 2
Private Const AC1_ADVANCED_COLOR_FORCE_DISABLED As Long = 3

'-----------------------------
' ¤˝¶}Ş¬şA¦CÁ|
'-----------------------------

' HDR Ş¬şAˇFUNKNOWN ĄÎ©óÂÂŞ© API µLŞkşë˝T°Ď¤Ŕ HDR / WCG Şş±ˇŞpˇC
Public Enum HDR_STATUS
    HDR_UNSUPPORTED = 0
    HDR_OFF = 1
    HDR_ON = 2
    HDR_ERROR = 3
    HDR_STATUS_UNKNOWN = 4
End Enum

'-----------------------------
' ĹăĄÜľą¸ę°Tµ˛şc
'-----------------------------

' ¨C­ÓĄŘ«e±ŇĄÎŞş Display Path ąďŔł¤@µ§¸ę®ĆˇC
Public Type HDR_DISPLAY_INFO
    ' UI ¨ĎĄÎŞşĹăĄÜľąŻÁ¤ŢˇA¤Ł¬O Windows DisplayConfig Target ĂŃ§O˝XˇC
    Index As Long

    ' Windows Display Configuration Şş Adapter LUIDˇA¤Ŕ¦¨§C¦ě»P°Ş¦ěŔx¦sˇC
    AdapterLow As Long
    AdapterHigh As Long

    ' Windows Display Configuration Şş Target IDˇC
    TargetId As Long

    ' ±q DISPLAYCONFIG_TARGET_DEVICE_NAME ¨ú±oŞş¤Íµ˝ĹăĄÜľą¦WşŮˇC
    Name As String

    ' 24H2+ Ąişë˝T¨ú±oŞş HDR Capability »PŞ¬şAˇC
    HDRSupported As Boolean
    HDRCapabilityKnown As Boolean
    HDRUserEnabled As Boolean
    HDRActive As Boolean

    ' Advanced Color Ş¬şAˇF¦ąÄć¦ě»P HDR ¤Ŕ¶}łB˛zˇC
    AdvancedColorSupported As Boolean
    AdvancedColorActive As Boolean
    AdvancedColorLimitedByPolicy As Boolean
    AdvancedColorForceDisabled As Boolean

    ' Wide Color / WCG ¬ŰĂöŞ¬şAˇC
    WCGSupported As Boolean
    WCGUserEnabled As Boolean
    WideColorEnforced As Boolean

    ' ĄŘ«eąę»Ú¦â±mĽŇ¦ˇˇG0=SDRˇB1=WCGˇB2=HDRˇFÂÂŞ© API µLŞkşë˝T¨ú±o®É¬° -1ˇC
    ActiveColorMode As Long

    ' ĹăĄÜľą¦WşŮ¨Ó·˝şXĽĐˇG¨Ó¦Ű EDID ©ÎłQ±j¨î«ü©wˇC
    FriendlyNameFromEdid As Boolean
    FriendlyNameForced As Boolean

    ' ¦ąĹăĄÜľąłĚ«á¤@¦¸ API ¬d¸ß©ÎĹçĂŇżů»~˝XˇF0 ŞíĄÜ¦¨Ą\ˇC
    LastError As Long
End Type

'-----------------------------
' ĽŇ˛Ő¤şżů»~Ş¬şA
'-----------------------------

' Ŕx¦słĚ«á¤@¦¸ĽŇ˛ŐľŢ§@Şş Win32 żů»~˝XˇC
Private mLastError As Long

' Ŕx¦s©T©w¬°ÁcĹé¤¤¤ĺŞşżů»~´y­zˇAÁ×§KŞ˝±µżéĄX Windows ­^¤ĺ¤ĺ¦rˇC
Private mLastErrorText As String

'=====================================================================
' ¤˝¶}¨çĽĆˇG¨ú±ołĚ«áżů»~˝X
' Ą\ŻŕˇG¨Ń UI ©Î¤WĽhµ{¦ˇĹŞ¨úłĚ«á¤@¦¸ĽŇ˛ŐľŢ§@Şş Win32 żů»~˝XˇC
'=====================================================================
Public Function HDR_GetLastError() As Long
    HDR_GetLastError = mLastError
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG¨ú±ołĚ«áżů»~¤ĺ¦r
' Ą\ŻŕˇG©Ň¦ł¦^¶Ç¤ş®e©T©w¨ĎĄÎÁcĹé¤¤¤ĺˇC
'=====================================================================
Public Function HDR_GetLastErrorText() As String
    If Len(mLastErrorText) > 0 Then
        HDR_GetLastErrorText = mLastErrorText
    ElseIf mLastError = ERROR_SUCCESS Then
        HDR_GetLastErrorText = "ľŢ§@¦¨Ą\ˇC"
    Else
        HDR_GetLastErrorText = HDR_FormatWin32ErrorChinese(mLastError)
    End If
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG¨ú±o Windows Build
' Ą\ŻŕˇG¨ĎĄÎ RtlGetVersion ¨ú±oąę»Ú Windows Build NumberˇC
'=====================================================================
Public Function HDR_GetWindowsBuild() As Long
    Dim osInfo() As Byte          ' RTL_OSVERSIONINFOEXW Şş­ě©l°OľĐĹéˇC
    Dim result As Long            ' RtlGetVersion ¦^¶Ç­ČˇC

    ReDim osInfo(0 To RTL_OSVERSIONINFOEXW_SIZE - 1)

    ' dwOSVersionInfoSize ¦ě©ó offset 0ˇAĄ˛¶·¶ń¤Jµ˛şc¤j¤pˇC
    HDR_WriteLong osInfo, 0, RTL_OSVERSIONINFOEXW_SIZE

    result = RtlGetVersion(osInfo(0))
    If result <> ERROR_SUCCESS Then
        HDR_SetLastError result, "ĹŞ¨ú Windows Ş©Ą»Ą˘±Ń"
        HDR_GetWindowsBuild = 0
        Exit Function
    End If

    ' RTL_OSVERSIONINFOEXWˇGdwBuildNumber ¦ě©ó offset 12ˇC
    HDR_GetWindowsBuild = HDR_ReadLong(osInfo, 12)
    HDR_ClearLastError
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG§PÂ_¬O§_¨ĎĄÎ Windows 11 24H2+ HDR API
' Ą\ŻŕˇGBuild >= 26100 ®É¨« GET_ADVANCED_COLOR_INFO_2 / SET_HDR_STATEˇC
'=====================================================================
Public Function HDR_IsWindows11_24H2() As Boolean
    Dim buildNumber As Long        ' ąę»Ú Windows Build NumberˇC

    buildNumber = HDR_GetWindowsBuild()
    HDR_IsWindows11_24H2 = (buildNumber >= 26100)
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG¨ú±oĄŘ«e API ĽŇ¦ˇŞşÁcĹé¤¤¤ĺ¦WşŮˇC
'=====================================================================
Public Function HDR_GetApiModeText() As String
    If HDR_IsWindows11_24H2() Then
        HDR_GetApiModeText = "Windows 11 24H2 ĄH¤WˇG­ěĄÍ HDR Ş¬şA¤¶­±"
    Else
        HDR_GetApiModeText = "ÂÂŞ© WindowsˇGAdvanced Color ¬Ű®e¤¶­±"
    End If
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG¨ú±oĄŘ«e±ŇĄÎŞşĹăĄÜľąĽĆ¶qˇC
'=====================================================================
Public Function HDR_GetDisplayCount() As Long
    Dim displays() As HDR_DISPLAY_INFO   ' ©Ň¦łĄŘ«e active display ŞşŞ¬şA°}¦CˇC

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
' ¤˝¶}¨çĽĆˇG¨ú±o©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş§ąľăŞ¬şAˇC
' Ą\ŻŕˇGĄý¨ú±o Display PathˇA¦A¨Ě¨C­Ó Target ¬d¸ß¦WşŮ»P HDR ¸ę°TˇC
'=====================================================================
Public Function HDR_GetDisplays(ByRef Displays() As HDR_DISPLAY_INFO) As Boolean
    Dim pathBuffer() As Byte             ' DISPLAYCONFIG_PATH_INFO °}¦C­ě©l BufferˇC
    Dim modeBuffer() As Byte             ' DISPLAYCONFIG_MODE_INFO °}¦C­ě©l BufferˇC
    Dim pathCount As Long                ' ąę»Ú¦^¶ÇŞş Display Path ĽĆ¶qˇC
    Dim modeCount As Long                ' ąę»Ú¦^¶ÇŞş Mode ĽĆ¶qˇC
    Dim i As Long                        ' ĹăĄÜľą°}¦CŻÁ¤ŢˇC

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
' ¤˝¶}¨çĽĆˇG¨ú±o«ü©wĹăĄÜľąŞş HDR Ş¬şAˇC
' Ą\ŻŕˇG¦^¶Ç HDR ¶}±ŇˇBĂöł¬ˇB¤Ł¤ä´©ˇBĄĽŞľ©Îżů»~ˇC
'=====================================================================
Public Function HDR_GetStatus(ByVal DisplayIndex As Long) As HDR_STATUS
    Dim displays() As HDR_DISPLAY_INFO   ' ©Ň¦łĄŘ«eĹăĄÜľą¸ę°TˇC

    HDR_GetStatus = HDR_ERROR

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "ĄŘ«e¨S¦łĄiĄÎŞş±ŇĄÎ¤¤ĹăĄÜľąˇC"
        Exit Function
    End If

    If DisplayIndex < LBound(displays) Or DisplayIndex > UBound(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "«ü©wŞşĹăĄÜľąŻÁ¤ŢµL®ÄˇC"
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
' ¤˝¶}¨çĽĆˇG±N HDR_STATUS Âŕ¦¨ÁcĹé¤¤¤ĺˇC
'=====================================================================
Public Function HDR_GetStatusText(ByVal Status As HDR_STATUS) As String
    Select Case Status
        Case HDR_UNSUPPORTED
            HDR_GetStatusText = "HDR ¤Ł¤ä´©"
        Case HDR_OFF
            HDR_GetStatusText = "HDR ¤wĂöł¬"
        Case HDR_ON
            HDR_GetStatusText = "HDR ¤w¶}±Ň"
        Case HDR_ERROR
            HDR_GetStatusText = "HDR µoĄÍżů»~"
        Case HDR_STATUS_UNKNOWN
            HDR_GetStatusText = "HDR Żŕ¤OĄĽŞľˇ]ÂÂŞ© API µLŞkşë˝T°Ď¤Ŕ HDR »P WCGˇ^"
        Case Else
            HDR_GetStatusText = "ĄĽŞľŞş HDR Ş¬şA"
    End Select
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG±N Windows Active Color Mode Âŕ¦¨ÁcĹé¤¤¤ĺˇC
' Ą\ŻŕˇG¨Ń UI ©Î¤WĽhµ{¦ˇ˛Î¤@ĹăĄÜ SDRˇBWCGˇBHDR ©ÎĄĽŞľĽŇ¦ˇˇC
'=====================================================================
Public Function HDR_GetColorModeText(ByVal ModeValue As Long) As String
    Select Case ModeValue
        Case DISPLAYCONFIG_ADVANCED_COLOR_MODE_SDR
            HDR_GetColorModeText = "SDRˇ]ĽĐ·Ç°ĘşA˝dłňˇ^"
        Case DISPLAYCONFIG_ADVANCED_COLOR_MODE_WCG
            HDR_GetColorModeText = "WCGˇ]Ľs¦â°ěˇ^"
        Case DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR
            HDR_GetColorModeText = "HDRˇ]°Ş°ĘşA˝dłňˇ^"
        Case Else
            HDR_GetColorModeText = "ĄĽŞľ¦â±mĽŇ¦ˇ"
    End Select
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG¶}±Ň©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDRˇC
'=====================================================================
Public Function HDR_Enable() As Boolean
    HDR_Enable = HDR_SetAllDisplays(True)
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇGĂöł¬©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDRˇC
'=====================================================================
Public Function HDR_Disable() As Boolean
    HDR_Disable = HDR_SetAllDisplays(False)
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇG¤Á´«©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDR Ş¬şAˇC
'=====================================================================
Public Function HDR_Toggle() As Boolean
    Dim displays() As HDR_DISPLAY_INFO   ' ©Ň¦łĄŘ«eĹăĄÜľąŞşłĚ·sŞ¬şAˇC
    Dim i As Long                        ' ĹăĄÜľą°}¦CŻÁ¤ŢˇC
    Dim okAll As Boolean                 ' ¬O§_©Ň¦łĹăĄÜľąłŁ¦¨Ą\ˇC
    Dim firstError As Long               ' ˛Ä¤@­ÓĄ˘±ŃŞş Win32 żů»~˝XˇC
    Dim targetState As Boolean            ' ¦ąĹăĄÜľą·ÇłĆ¤Á´«¨ěŞşŞ¬şAˇC

    HDR_Toggle = False

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "ĄŘ«e¨S¦łĄiĄÎŞş±ŇĄÎ¤¤ĹăĄÜľąˇC"
        Exit Function
    End If

    okAll = True
    firstError = ERROR_SUCCESS

    For i = LBound(displays) To UBound(displays)
        If HDR_IsWindows11_24H2() Then
            If Not displays(i).HDRCapabilityKnown Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = ERROR_NOT_SUPPORTED
                displays(i).LastError = ERROR_NOT_SUPPORTED
            ElseIf Not displays(i).HDRSupported Then
                okAll = False
                If firstError = ERROR_SUCCESS Then firstError = ERROR_NOT_SUPPORTED
                displays(i).LastError = ERROR_NOT_SUPPORTED
            Else
                targetState = Not displays(i).HDRActive
                If Not HDR_SetDisplay(displays(i), targetState) Then
                    okAll = False
                    If firstError = ERROR_SUCCESS Then firstError = mLastError
                End If
            End If
        Else
            ' ÂÂŞ©ĄuĄH Advanced Color Active §@¬°¤Á´«¨ĚľÚˇC
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
        HDR_SetLastError firstError, "łˇ¤Ŕ©ÎĄţłˇĹăĄÜľą¤Á´«Ą˘±ŃˇC"
    End If

    HDR_Toggle = okAll
End Function

'=====================================================================
' ¤˝¶}¨çĽĆˇGł]©w«ü©wĹăĄÜľąŞş HDRˇC
'=====================================================================
Public Function HDR_SetDisplayByIndex(ByVal DisplayIndex As Long, ByVal EnableHDR As Boolean) As Boolean
    Dim displays() As HDR_DISPLAY_INFO   ' ©Ň¦łĄŘ«eĹăĄÜľą¸ę°TˇC

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "ĄŘ«e¨S¦łĄiĄÎŞş±ŇĄÎ¤¤ĹăĄÜľąˇC"
        Exit Function
    End If

    If DisplayIndex < LBound(displays) Or DisplayIndex > UBound(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "«ü©wŞşĹăĄÜľąŻÁ¤ŢµL®ÄˇC"
        Exit Function
    End If

    HDR_SetDisplayByIndex = HDR_SetDisplay(displays(DisplayIndex), EnableHDR)
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇG¦wĄţ¨ú±oĄŘ«e active display configurationˇC
' Ą\ŻŕˇGłB˛z GetDisplayConfigBufferSizes »P QueryDisplayConfig ¤§¶ˇ
'       ĹăĄÜ˛ŐşA§ďĹÜły¦¨Şş ERROR_INSUFFICIENT_BUFFERˇAłĚ¦h­«¸Ő 3 ¦¸ˇC
'=====================================================================
Private Function HDR_QueryActiveDisplays( _
    ByRef pathBuffer() As Byte, _
    ByRef pathCount As Long, _
    ByRef modeBuffer() As Byte, _
    ByRef modeCount As Long) As Boolean

    Dim attempt As Long                  ' ˛Ä´X¦¸¬d¸ßˇAłĚ¦h 3 ¦¸ˇC
    Dim result As Long                   ' Win32 API ¦^¶Ç­ČˇC
    Dim requestedPaths As Long           ' API «ŘÄł°t¸mŞş Path ĽĆ¶qˇC
    Dim requestedModes As Long           ' API «ŘÄł°t¸mŞş Mode ĽĆ¶qˇC
    Dim pathBytes As Long                ' Path Buffer ąę»Ú°t¸m¤j¤pˇC
    Dim modeBytes As Long                ' Mode Buffer ąę»Ú°t¸m¤j¤pˇC

    Erase pathBuffer
    Erase modeBuffer
    pathCount = 0
    modeCount = 0

    For attempt = 1 To 3
        requestedPaths = 0
        requestedModes = 0

        result = GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, requestedPaths, requestedModes)
        If result <> ERROR_SUCCESS Then
            HDR_SetLastError result, "¨ú±oĹăĄÜ˛ŐşA Buffer ¤j¤pĄ˘±ŃˇC"
            Exit Function
        End If

        pathCount = requestedPaths
        modeCount = requestedModes

        ' §Y¨Ď API ¦^łř 0ˇA¤]°t¸m¦Ü¤Ö¤@µ§°OľĐĹéˇAÁ×§K VB6 ˛ŁĄÍŞĹ°}¦C°ÝĂDˇC
        If pathCount < 1 Then pathCount = 1
        If modeCount < 1 Then modeCount = 1

        pathBytes = pathCount * DISPLAYCONFIG_PATH_INFO_SIZE
        modeBytes = modeCount * DISPLAYCONFIG_MODE_INFO_SIZE

        ReDim pathBuffer(0 To pathBytes - 1)
        ReDim modeBuffer(0 To modeBytes - 1)

        ' QDC_ONLY_ACTIVE_PATHS ¤U currentTopologyId Ą˛¶·¶Ç NULLˇF32-bit VB6 ĄH ByVal 0& ¶Ç¤JˇC
        result = QueryDisplayConfig( _
                    QDC_ONLY_ACTIVE_PATHS, _
                    pathCount, _
                    pathBuffer(0), _
                    modeCount, _
                    modeBuffer(0), _
                    0&)

        If result = ERROR_SUCCESS Then
            If pathCount < 0 Then pathCount = 0
            If modeCount < 0 Then modeCount = 0
            HDR_ClearLastError
            HDR_QueryActiveDisplays = True
            Exit Function
        End If

        ' ĹăĄÜ˛ŐşAĄiŻŕ¦b¨â­Ó API ¤§¶ˇµoĄÍĹÜ¤ĆˇF­«·s¨ú±o Buffer ¤j¤p«á¦A¬d¤@¦¸ˇC
        If result <> ERROR_INSUFFICIENT_BUFFER Then
            HDR_SetLastError result, "ĹŞ¨úĄŘ«eĹăĄÜ˛ŐşAĄ˘±ŃˇC"
            Exit Function
        End If
    Next attempt

    HDR_SetLastError ERROR_INSUFFICIENT_BUFFER, "ĹŞ¨úĹăĄÜ˛ŐşAłsÄň 3 ¦¸łŁ¦] Buffer ¤j¤p¤Ł¨¬¦ÓĄ˘±ŃˇC"
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇG«ŘĄßłć¤@ĹăĄÜľą¸ę°TˇC
' Ą\ŻŕˇG±q DISPLAYCONFIG_PATH_INFO ¨úĄX Adapter LUID / Target IDˇA
'       ¦A¨ú±oĹăĄÜľą¦WşŮ»P HDR / Advanced Color Ş¬şAˇC
'=====================================================================
Private Function HDR_CreateDisplayInfo(ByVal Index As Long, ByRef pathBuffer() As Byte) As HDR_DISPLAY_INFO
    Dim info As HDR_DISPLAY_INFO            ' §Y±N¦^¶ÇŞşłć¤@ĹăĄÜľą¸ę®ĆˇC
    Dim offset As Long                     ' ¦ą Path ¦b Buffer ¤¤Şş°_©l offsetˇC
    Dim pathFlags As Long                  ' DISPLAYCONFIG_PATH_INFO.status ŞşşXĽĐˇC
    Dim nameError As Long                  ' ĹăĄÜľą¦WşŮ API Şşżů»~˝XˇF¤ŁĽvĹT HDR ®Ö¤ß¬d¸ßˇC

    info.Index = Index
    info.ActiveColorMode = DISPLAYCONFIG_ADVANCED_COLOR_MODE_SDR
    info.LastError = ERROR_SUCCESS

    offset = Index * DISPLAYCONFIG_PATH_INFO_SIZE

    ' DISPLAYCONFIG_PATH_INFO.targetInfoˇG
    ' adapterId.LowPart offset 20ˇBHighPart offset 24ˇBid offset 28ˇC
    info.AdapterLow = HDR_ReadLong(pathBuffer, offset + 20)
    info.AdapterHigh = HDR_ReadLong(pathBuffer, offset + 24)
    info.TargetId = HDR_ReadLong(pathBuffer, offset + 28)

    ' status ¦ě©ó path µ˛şcłĚ«á 4 bytesˇAoffset 68ˇC
    pathFlags = HDR_ReadLong(pathBuffer, offset + 68)

    If (pathFlags And DISPLAYCONFIG_PATH_ACTIVE) = 0 Then
        info.LastError = ERROR_INVALID_PARAMETER
        HDR_CreateDisplayInfo = info
        Exit Function
    End If

    ' Friendly Name ÄÝ©ó»˛§U¸ę°TˇF¦WşŮ¨ú±oĄ˘±Ń¤ŁŔłŞý¤î HDR Ş¬şA¬d¸ßˇC
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
' ¨p¦ł¨çĽĆˇG¬d¸ß Windows 11 24H2+ HDR Ş¬şAˇC
' Ą\ŻŕˇG¸ŃŞR GET_ADVANCED_COLOR_INFO_2 Şş bit field »P Active Color ModeˇC
'=====================================================================
Private Function HDR_ReadDisplayStatus2(ByRef info As HDR_DISPLAY_INFO) As Boolean
    Dim packet(0 To DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2_SIZE - 1) As Byte  ' API ¦^¶Çµ˛şcˇC
    Dim result As Long                        ' DisplayConfigGetDeviceInfo ¦^¶Ç­ČˇC
    Dim value As Long                         ' 20 bytes offset Şş bit fieldˇC
    Dim activeMode As Long                    ' 32 bytes offset Şş Active Color ModeˇC

    HDR_InitHeader packet, _
                   DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2, _
                   DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_2_SIZE, _
                   info.AdapterLow, info.AdapterHigh, info.TargetId

    result = DisplayConfigGetDeviceInfo(packet(0))
    If result <> ERROR_SUCCESS Then
        info.LastError = result
        HDR_SetLastError result, "ĹŞ¨ú·sŞ© HDR / Advanced Color Ş¬şAĄ˘±ŃˇC"
        Exit Function
    End If

    value = HDR_ReadLong(packet, 20)
    activeMode = HDR_ReadLong(packet, 32)

    info.AdvancedColorSupported = HDR_TestBit(value, AC2_ADVANCED_COLOR_SUPPORTED)
    info.AdvancedColorActive = HDR_TestBit(value, AC2_ADVANCED_COLOR_ACTIVE)
    info.AdvancedColorLimitedByPolicy = HDR_TestBit(value, AC2_ADVANCED_COLOR_LIMITED_BY_POLICY)

    ' ·sŞ© API ĄiŞ˝±µ§PÂ_ HDR capabilityˇC
    info.HDRSupported = HDR_TestBit(value, AC2_HDR_SUPPORTED)
    info.HDRCapabilityKnown = True
    info.HDRUserEnabled = HDR_TestBit(value, AC2_HDR_USER_ENABLED)

    ' ŻuĄżŞş HDR Active §PÂ_ĄH Active Color Mode ¬°·ÇˇA¦Ó¤Ł¬OĄu¬Ý User EnabledˇC
    info.HDRActive = (activeMode = DISPLAYCONFIG_ADVANCED_COLOR_MODE_HDR)

    info.WCGSupported = HDR_TestBit(value, AC2_WIDE_COLOR_SUPPORTED)
    info.WCGUserEnabled = HDR_TestBit(value, AC2_WIDE_COLOR_USER_ENABLED)

    ' ·sŞ©µ˛şc¨S¦łĄHÂÂŞ©¬Ű¦P¤č¦ˇ´Ł¨Ń Wide Color Enforced / Force DisabledˇC
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
' ¨p¦ł¨çĽĆˇG¬d¸ßÂÂŞ© Windows Advanced Color Ş¬şAˇC
' Ą\ŻŕˇGĄu¶ńĽgÂÂ API ŻŕĄiľa¨ú±oŞşÄć¦ěˇA¤Ł§â Advanced Color «_ĄR HDRˇC
'=====================================================================
Private Function HDR_ReadDisplayStatusLegacy(ByRef info As HDR_DISPLAY_INFO) As Boolean
    Dim packet(0 To DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_SIZE - 1) As Byte  ' ÂÂŞ© API µ˛şcˇC
    Dim result As Long                        ' API ¦^¶Ç­ČˇC
    Dim value As Long                         ' ÂÂŞ© bit fieldˇC

    HDR_InitHeader packet, _
                   DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO, _
                   DISPLAYCONFIG_GET_ADVANCED_COLOR_INFO_SIZE, _
                   info.AdapterLow, info.AdapterHigh, info.TargetId

    result = DisplayConfigGetDeviceInfo(packet(0))
    If result <> ERROR_SUCCESS Then
        info.LastError = result
        HDR_SetLastError result, "ĹŞ¨úÂÂŞ© Advanced Color Ş¬şAĄ˘±ŃˇC"
        Exit Function
    End If

    value = HDR_ReadLong(packet, 20)

    info.AdvancedColorSupported = HDR_TestBit(value, AC1_ADVANCED_COLOR_SUPPORTED)
    info.AdvancedColorActive = HDR_TestBit(value, AC1_ADVANCED_COLOR_ENABLED)
    info.AdvancedColorForceDisabled = HDR_TestBit(value, AC1_ADVANCED_COLOR_FORCE_DISABLED)
    info.WideColorEnforced = HDR_TestBit(value, AC1_WIDE_COLOR_ENFORCED)

    ' ÂÂ API µLŞkĄiľa°Ď¤Ŕ HDR »P WCGˇA©ŇĄH HDR capability ¨č·NĽĐ°O¬° UnknownˇC
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
' ¨p¦ł¨çĽĆˇGł]©włć¤@ĹăĄÜľąŞş HDR / Advanced ColorˇC
' Ą\ŻŕˇG24H2+ ÄY®ć¨ĎĄÎ SET_HDR_STATEˇFÂÂŞ©¨ĎĄÎ Advanced ColorˇC
'       ł]©w¦¨Ą\«á¤@©w­«·s¬d¸ß¨ĂĹçĂŇˇC
'=====================================================================
Private Function HDR_SetDisplay(ByRef info As HDR_DISPLAY_INFO, ByVal EnableHDR As Boolean) As Boolean
    Dim packet(0 To DISPLAYCONFIG_SET_HDR_STATE_SIZE - 1) As Byte              ' 24H2+ HDR ł]©wµ˛şcˇC
    Dim packetLegacy(0 To DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE_SIZE - 1) As Byte ' ÂÂŞ©ł]©wµ˛şcˇC
    Dim result As Long                       ' Set API ¦^¶Ç­ČˇC
    Dim expected As Boolean                  ' ĹçĂŇ®Éąw´ÁŞşłĚ˛×Ş¬şAˇC

    expected = EnableHDR

    If HDR_IsWindows11_24H2() Then
        ' ·sŞ©¸ô®|Ąu¤ął\ŻuĄżŞş HDR capabilityˇC
        If Not info.HDRCapabilityKnown Then
            HDR_SetLastError ERROR_NOT_SUPPORTED, "µLŞk¨ú±o¦ąĹăĄÜľąŞş HDR Żŕ¤O¸ę°TˇC"
            Exit Function
        End If

        If EnableHDR And Not info.HDRSupported Then
            HDR_SetLastError ERROR_NOT_SUPPORTED, "¦ąĹăĄÜľą¤Ł¤ä´© HDRˇC"
            Exit Function
        End If

        If Not EnableHDR And Not info.HDRSupported Then
            ' Ą»¨Ó´N¤Ł¤ä´© HDRˇAĂöł¬ HDR ¤w˛Ĺ¦XłĚ˛×Ş¬şAˇC
            HDR_SetDisplay = True
            HDR_ClearLastError
            Exit Function
        End If

        HDR_InitHeader packet, _
                       DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE, _
                       DISPLAYCONFIG_SET_HDR_STATE_SIZE, _
                       info.AdapterLow, info.AdapterHigh, info.TargetId

        If EnableHDR Then
            HDR_WriteLong packet, 20, 1
        Else
            HDR_WriteLong packet, 20, 0
        End If

        result = DisplayConfigSetDeviceInfo(packet(0))
        If result <> ERROR_SUCCESS Then
            info.LastError = result
            HDR_SetLastError result, "ł]©w·sŞ© HDR Ş¬şAĄ˘±ŃˇC"
            Exit Function
        End If
    Else
        ' ÂÂŞ© API Şş»y·N¬O Advanced ColorˇA¤ŁĄiµř¬°şë˝T HDR ¶}ĂöˇC
        If Not EnableHDR And Not info.AdvancedColorActive Then
            HDR_SetDisplay = True
            HDR_ClearLastError
            Exit Function
        End If

        HDR_InitHeader packetLegacy, _
                       DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE, _
                       DISPLAYCONFIG_SET_ADVANCED_COLOR_STATE_SIZE, _
                       info.AdapterLow, info.AdapterHigh, info.TargetId

        If EnableHDR Then
            HDR_WriteLong packetLegacy, 20, 1
        Else
            HDR_WriteLong packetLegacy, 20, 0
        End If

        result = DisplayConfigSetDeviceInfo(packetLegacy(0))
        If result <> ERROR_SUCCESS Then
            info.LastError = result
            HDR_SetLastError result, "ł]©wÂÂŞ© Advanced Color Ş¬şAĄ˘±ŃˇC"
            Exit Function
        End If
    End If

    ' Set ¦¨Ą\¤ŁĄNŞíąę»ÚŞ¬şA¤@©w¤w¸g¤Á´«ˇAĄ˛¶·­«·s Query ĹçĂŇˇC
    If HDR_VerifyDisplay(info, expected) Then
        HDR_SetDisplay = True
        HDR_ClearLastError
    End If
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇGł]©w©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąˇC
' Ą\ŻŕˇGłv¤@łB˛z¦hżĂąőˇA°Ożý˛Ä¤@­Óżů»~¨Ă«OŻd¨äĄLĹăĄÜľąŞş°ő¦ćµ˛ŞGˇC
'=====================================================================
Private Function HDR_SetAllDisplays(ByVal EnableHDR As Boolean) As Boolean
    Dim displays() As HDR_DISPLAY_INFO   ' ©Ň¦łĄŘ«e±ŇĄÎŞşĹăĄÜľąˇC
    Dim i As Long                        ' ĹăĄÜľą°}¦CŻÁ¤ŢˇC
    Dim okAll As Boolean                 ' ¬O§_©Ň¦łĹăĄÜľąłŁ¦¨Ą\ˇC
    Dim firstError As Long               ' ˛Ä¤@­Óżů»~˝XˇC

    If Not HDR_GetDisplays(displays) Then Exit Function

    If Not HDR_ArrayAllocated(displays) Then
        HDR_SetLastError ERROR_INVALID_PARAMETER, "ĄŘ«e¨S¦łĄiĄÎŞş±ŇĄÎ¤¤ĹăĄÜľąˇC"
        Exit Function
    End If

    okAll = True
    firstError = ERROR_SUCCESS

    For i = LBound(displays) To UBound(displays)
        If HDR_IsWindows11_24H2() Then
            If EnableHDR And Not displays(i).HDRSupported Then
                okAll = False
                displays(i).LastError = ERROR_NOT_SUPPORTED
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
        HDR_SetLastError firstError, "łˇ¤Ŕ©ÎĄţłˇĹăĄÜľął]©wĄ˘±ŃˇC"
    End If

    HDR_SetAllDisplays = okAll
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇGĹçĂŇłć¤@ĹăĄÜľąŞşłĚ˛×Ş¬şAˇC
' Ą\ŻŕˇG­«·s QueryˇA¨ĂĄH Adapter LUID + Target ID §ä¦^¦P¤@ĹăĄÜľąˇC
'       łĚ¦h­«¸Ő 3 ¦¸ˇAÁ×§K Windows ©|¦b®MĄÎŞ¬şA®É»~§PˇC
'=====================================================================
Private Function HDR_VerifyDisplay(ByRef OriginalInfo As HDR_DISPLAY_INFO, ByVal ExpectedHDR As Boolean) As Boolean
    Dim displays() As HDR_DISPLAY_INFO   ' ­«·s¬d¸ß«áŞşĹăĄÜľąŞ¬şAˇC
    Dim retry As Long                    ' ĹçĂŇ­«¸Ő¦¸ĽĆˇAłĚ¦h 3 ¦¸ˇC
    Dim i As Long                        ' ĹăĄÜľą°}¦CŻÁ¤ŢˇC
    Dim found As Boolean                 ' ¬O§_§ä¨ě­ěĄ»Şş TargetˇC
    Dim current As HDR_DISPLAY_INFO      ' §ä¨ěŞşłĚ·sĹăĄÜľąŞ¬şAˇC

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
                    If ExpectedHDR And current.HDRActive Then
                        HDR_VerifyDisplay = True
                        Exit Function
                    End If

                    If Not ExpectedHDR And Not current.HDRActive Then
                        HDR_VerifyDisplay = True
                        Exit Function
                    End If
                Else
                    ' ÂÂŞ©ĹçĂŇŞş¬O Advanced ColorˇA¦Ó¤Ł¬Oşë˝T HDRˇC
                    If current.AdvancedColorActive = ExpectedHDR Then
                        HDR_VerifyDisplay = True
                        Exit Function
                    End If
                End If
            Else
                HDR_SetLastError ERROR_NOT_FOUND, "ł]©w«á§ä¤Ł¨ě­ěĄ»ŞşĹăĄÜľą TargetˇC"
            End If
        End If

        Sleep 100
    Next retry

    If HDR_IsWindows11_24H2() Then
        HDR_SetLastError ERROR_GEN_FAILURE, "ł]©w«á­«·s¬d¸ßˇAHDR ąę»ÚŞ¬şAĄĽ˛Ĺ¦Xąw´ÁˇC"
    Else
        HDR_SetLastError ERROR_GEN_FAILURE, "ł]©w«á­«·s¬d¸ßˇAAdvanced Color ąę»ÚŞ¬şAĄĽ˛Ĺ¦Xąw´ÁˇC"
    End If
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇG¨ú±oĹăĄÜľą¤Íµ˝¦WşŮˇC
' Ą\ŻŕˇG¨ĎĄÎ DISPLAYCONFIG_TARGET_DEVICE_NAME ¨ú±o Friendly NameˇA
'       ¦P®É¦^¶Ç¦WşŮ¨Ó·˝şXĽĐˇC
'=====================================================================
Private Function HDR_GetDisplayName( _
    ByVal adapterLow As Long, _
    ByVal adapterHigh As Long, _
    ByVal targetId As Long, _
    ByRef FriendlyNameFromEdid As Boolean, _
    ByRef FriendlyNameForced As Boolean, _
    ByRef ErrorCode As Long) As String

    Dim packet(0 To DISPLAYCONFIG_TARGET_DEVICE_NAME_SIZE - 1) As Byte ' ¦WşŮ API Şş 420-byte µ˛şcˇC
    Dim result As Long                       ' API ¦^¶Ç­ČˇC
    Dim flags As Long                        ' Friendly Name ¨Ó·˝şXĽĐˇC
    Dim nameText As String                   ' łĚ˛×ĹăĄÜľą¦WşŮˇC

    FriendlyNameFromEdid = False
    FriendlyNameForced = False
    ErrorCode = ERROR_SUCCESS

    HDR_InitHeader packet, _
                   DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME, _
                   DISPLAYCONFIG_TARGET_DEVICE_NAME_SIZE, _
                   adapterLow, adapterHigh, targetId

    result = DisplayConfigGetDeviceInfo(packet(0))
    If result <> ERROR_SUCCESS Then
        ErrorCode = result
        HDR_GetDisplayName = "˛Ä " & CStr(targetId) & " ­ÓĹăĄÜĄŘĽĐ"
        Exit Function
    End If

    flags = HDR_ReadLong(packet, 20)

    ' DISPLAYCONFIG_TARGET_DEVICE_NAME_FLAGSˇGbit 0=EDIDˇBbit 1=±j¨î«ü©wˇC
    FriendlyNameFromEdid = HDR_TestBit(flags, 0)
    FriendlyNameForced = HDR_TestBit(flags, 1)

    ' 20-byte Header + 16-byte«e¸mÄć¦ě«áˇAFriendly Name ±q offset 36 ¶}©lˇC
    nameText = HDR_ReadUnicodeString(packet, 36, 64)

    If Len(nameText) = 0 Then
        nameText = "˛Ä " & CStr(targetId) & " ­ÓĹăĄÜĄŘĽĐ"
    End If

    HDR_GetDisplayName = nameText
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇGŞě©l¤Ć DISPLAYCONFIG_DEVICE_INFO_HEADERˇC
' Ą\ŻŕˇG©Ň¦ł Get / Set Device Info µ˛şc¦@ĄÎ¦P¤@­Ó Header «e 20 bytesˇC
'=====================================================================
Private Sub HDR_InitHeader( _
    ByRef packet() As Byte, _
    ByVal infoType As Long, _
    ByVal packetSize As Long, _
    ByVal adapterLow As Long, _
    ByVal adapterHigh As Long, _
    ByVal targetId As Long)

    HDR_WriteLong packet, 0, infoType         ' typeˇG¸ę®Ć­n¨DĂţ«¬ˇC
    HDR_WriteLong packet, 4, packetSize      ' sizeˇGľă­Ó«ĘĄ]¤j¤pˇC
    HDR_WriteLong packet, 8, adapterLow      ' adapterId.LowPartˇC
    HDR_WriteLong packet, 12, adapterHigh    ' adapterId.HighPartˇC
    HDR_WriteLong packet, 16, targetId       ' targetIdˇC
End Sub

'=====================================================================
' ¨p¦ł¨çĽĆˇG±q Byte Array «ü©w offset ĹŞ¨ú 32-bit LongˇC
' Ą\ŻŕˇG¶°¤¤łB˛z Windows API ­ě©lµ˛şcŞş¦ě¤¸˛ŐĹŞ¨úˇC
' ­«­n°ŃĽĆˇGOffset ¬Oµ˛şc¤ş¦ě˛ľ¶qˇAbuffer ¬O API ­ě©l¸ę®ĆˇC
'=====================================================================
Private Function HDR_ReadLong(ByRef buffer() As Byte, ByVal Offset As Long) As Long
    Dim value As Long                        ' ±q«ü©w offset ˝Ć»sĄXŞş 32-bit ľăĽĆˇC

    CopyMemory value, buffer(Offset), 4
    HDR_ReadLong = value
End Function

'=====================================================================
' ¨p¦ł¨çĽĆˇG±N 32-bit Long Ľg¤J Byte ArrayˇC
'=====================================================================
Private Sub HDR_WriteLong(ByRef buffer() As Byte, ByVal Offset As Long, ByVal value As Long)
    CopyMemory buffer(Offset), value, 4
End Sub

'=====================================================================
' ¨p¦ł¨çĽĆˇG´ú¸Ő 32-bit ¦ě¤¸Äć¦ě¤¤Şş«ü©w bitˇC
'=====================================================================
Private Function HDR_TestBit(ByVal value As Long, ByVal bitIndex As Long) As Boolean
    Select Case bitIndex
        Case 0: HDR_TestBit = ((value And 1) <> 0)
