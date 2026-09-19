VERSION 5.00
Begin VB.Form frmHDRTest
   Caption         =   "Windows HDR ­ěĄÍ API ´ú¸Ő"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "frmHDRTest"
   ScaleHeight     =   9480
   ScaleWidth      =   12480
   StartUpPosition =   2
   Begin VB.CommandButton cmdRead
      Caption         =   "ĹŞ¨úĄŘ«eŞ¬şA"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdEnable
      Caption         =   "¶}±Ň HDR"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdDisable
      Caption         =   "Ăöł¬ HDR"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdToggle
      Caption         =   "¤Á´« HDR"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   1800
   End
   Begin VB.TextBox txtOutput
      Height          =   7695
      Left            =   240
      MultiLine       =   -1
      ScrollBars      =   3
      TabIndex        =   4
      Top             =   960
      Width           =   12000
   End
End
Attribute VB_Name = "frmHDRTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Windows HDR ­ěĄÍ API ´ú¸ŐŞíłć"
Attribute VB_Ext_KEY = "SavedWithLastKnownToolVersion", "6.00.9782.0"

Option Explicit

' ĹŞ¨úĄŘ«e©Ň¦łĹăĄÜľąŞş HDR / Advanced Color Ş¬şAˇC
Private Sub cmdRead_Click()
    ShowHDRStatus False
End Sub

' ¶}±Ň©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDRˇC
Private Sub cmdEnable_Click()
    txtOutput.Text = "Ąż¦bł]©wˇG¶}±Ň©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDRˇKˇK" & vbCrLf
    txtOutput.Text = txtOutput.Text & "¶}©l®É¶ˇˇG" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txtOutput.Text = txtOutput.Text & String$(78, "-") & vbCrLf

    If HDR_Enable() Then
        txtOutput.Text = txtOutput.Text & "µ˛ŞGˇGHDR ł]©w¦¨Ą\ˇAĄB¤w§ą¦¨ąę»ÚŞ¬şAĹçĂŇˇC" & vbCrLf
    Else
        txtOutput.Text = txtOutput.Text & "µ˛ŞGˇGHDR ł]©wĄ˘±ŃˇC" & vbCrLf
        txtOutput.Text = txtOutput.Text & "¸Ô˛Ó¸ę°TˇG" & HDR_GetLastErrorText() & vbCrLf
    End If

    txtOutput.Text = txtOutput.Text & vbCrLf
    ShowHDRStatus True
End Sub

' Ăöł¬©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDRˇC
Private Sub cmdDisable_Click()
    txtOutput.Text = "Ąż¦bł]©wˇGĂöł¬©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDRˇKˇK" & vbCrLf
    txtOutput.Text = txtOutput.Text & "¶}©l®É¶ˇˇG" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txtOutput.Text = txtOutput.Text & String$(78, "-") & vbCrLf

    If HDR_Disable() Then
        txtOutput.Text = txtOutput.Text & "µ˛ŞGˇGHDR Ăöł¬¦¨Ą\ˇAĄB¤w§ą¦¨ąę»ÚŞ¬şAĹçĂŇˇC" & vbCrLf
    Else
        txtOutput.Text = txtOutput.Text & "µ˛ŞGˇGHDR Ăöł¬Ą˘±ŃˇC" & vbCrLf
        txtOutput.Text = txtOutput.Text & "¸Ô˛Ó¸ę°TˇG" & HDR_GetLastErrorText() & vbCrLf
    End If

    txtOutput.Text = txtOutput.Text & vbCrLf
    ShowHDRStatus True
End Sub

' ±N©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDR Ş¬şA¤ĎÂŕˇC
Private Sub cmdToggle_Click()
    txtOutput.Text = "Ąż¦b°ő¦ćˇG¤Á´«©Ň¦łĄŘ«e±ŇĄÎĹăĄÜľąŞş HDR Ş¬şAˇKˇK" & vbCrLf
    txtOutput.Text = txtOutput.Text & "¶}©l®É¶ˇˇG" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txtOutput.Text = txtOutput.Text & String$(78, "-") & vbCrLf

    If HDR_Toggle() Then
        txtOutput.Text = txtOutput.Text & "µ˛ŞGˇGHDR ¤Á´«¦¨Ą\ˇAĄB¤w§ą¦¨ąę»ÚŞ¬şAĹçĂŇˇC" & vbCrLf
    Else
        txtOutput.Text = txtOutput.Text & "µ˛ŞGˇGHDR ¤Á´«Ą˘±ŃˇC" & vbCrLf
        txtOutput.Text = txtOutput.Text & "¸Ô˛Ó¸ę°TˇG" & HDR_GetLastErrorText() & vbCrLf
    End If

    txtOutput.Text = txtOutput.Text & vbCrLf
    ShowHDRStatus True
End Sub

' Şíłć¸ü¤J®ÉĄß§YĹăĄÜ¤@¦¸ĄŘ«e HDR Ş¬şAˇC
Private Sub Form_Load()
    ShowHDRStatus False
End Sub

' ĹăĄÜ§ąľăŞ¬şA»PÁcĹé¤¤¤ĺ°OżýˇFAppend=True ®É«OŻd«e­±ŞşľŢ§@°OżýˇC
Private Sub ShowHDRStatus(Optional ByVal Append As Boolean = False)
    Dim displays() As HDR_DISPLAY_INFO   ' ĄŘ«e©Ň¦ł active display ŞşŞ¬şAˇC
    Dim i As Long                        ' ĹăĄÜľąŻÁ¤ŢˇC
    Dim buildNumber As Long              ' Windows Build NumberˇC
    Dim outputText As String             ' łĚ˛×ĹăĄÜ¦b¤ĺ¦r¤č¶ô¤¤ŞşÁcĹé¤¤¤ĺ°OżýˇC

    If Append Then
        outputText = txtOutput.Text
        If Len(outputText) > 0 Then outputText = outputText & vbCrLf
    Else
        outputText = vbNullString
    End If

    outputText = outputText & "ĄŘ«eŞ¬şA°Ożý®É¶ˇˇG" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf

    buildNumber = HDR_GetWindowsBuild()
    outputText = outputText & "Windows Ş©Ą»˛Ő«ŘˇG" & CStr(buildNumber) & vbCrLf
    outputText = outputText & "HDR ±±¨î¤¶­±ˇG" & HDR_GetApiModeText() & vbCrLf
    outputText = outputText & String$(78, "=") & vbCrLf

    If Not HDR_GetDisplays(displays) Then
        outputText = outputText & "ĹŞ¨úĹăĄÜľąŞ¬şAĄ˘±ŃˇC" & vbCrLf
        outputText = outputText & "¸Ô˛Ó¸ę°TˇG" & HDR_GetLastErrorText() & vbCrLf
        txtOutput.Text = outputText
        Exit Sub
    End If

    On Error GoTo NoDisplays

    For i = LBound(displays) To UBound(displays)
        outputText = outputText & "ĹăĄÜľą #" & CStr(displays(i).Index) & vbCrLf
        outputText = outputText & "ĹăĄÜľą¦WşŮˇG" & displays(i).Name & vbCrLf
        outputText = outputText & "°t±µľąĂŃ§O˝Xˇ]LUIDˇ^ˇG" & SignedHex8(displays(i).AdapterHigh) & SignedHex8(displays(i).AdapterLow) & vbCrLf
        outputText = outputText & "ĹăĄÜĄŘĽĐĂŃ§O˝XˇG" & CStr(displays(i).TargetId) & vbCrLf

        outputText = outputText & "°Ş°ĘşA˝dłňˇ]HDRˇ^¤ä´©ˇG" & BoolText(displays(i).HDRSupported) & vbCrLf
        outputText = outputText & "°Ş°ĘşA˝dłňŻŕ¤OĄi§P©wˇG" & BoolText(displays(i).HDRCapabilityKnown) & vbCrLf
        outputText = outputText & "¨ĎĄÎŞĚ HDR ł]©wˇG" & BoolText(displays(i).HDRUserEnabled) & vbCrLf
        outputText = outputText & "HDR ąę»Ú±ŇĄÎˇG" & BoolText(displays(i).HDRActive) & vbCrLf

        outputText = outputText & "¶i¶Ą¦â±m¤ä´©ˇG" & BoolText(displays(i).AdvancedColorSupported) & vbCrLf
        outputText = outputText & "¶i¶Ą¦â±mąę»Ú±ŇĄÎˇG" & BoolText(displays(i).AdvancedColorActive) & vbCrLf
        outputText = outputText & "¨ü¨ě¨t˛Î¬Fµ¦­­¨îˇG" & BoolText(displays(i).AdvancedColorLimitedByPolicy) & vbCrLf
        outputText = outputText & "łQ¨t˛Î±j¨î°±ĄÎˇG" & BoolText(displays(i).AdvancedColorForceDisabled) & vbCrLf

        outputText = outputText & "Ľs¦â°ěˇ]WCGˇ^¤ä´©ˇG" & BoolText(displays(i).WCGSupported) & vbCrLf
        outputText = outputText & "¨ĎĄÎŞĚĽs¦â°ěł]©wˇG" & BoolText(displays(i).WCGUserEnabled) & vbCrLf
        outputText = outputText & "Ľs¦â°ě±j¨î±ŇĄÎˇG" & BoolText(displays(i).WideColorEnforced) & vbCrLf

        outputText = outputText & "ĄŘ«eąę»Ú¦â±mĽŇ¦ˇˇG" & HDR_GetColorModeText(displays(i).ActiveColorMode) & vbCrLf
        outputText = outputText & "ĹăĄÜľą¦WşŮ¨Ó·˝¬° EDIDˇG" & BoolText(displays(i).FriendlyNameFromEdid) & vbCrLf
        outputText = outputText & "ĹăĄÜľą¦WşŮĄŃ¨t˛Î±j¨î«ü©wˇG" & BoolText(displays(i).FriendlyNameForced) & vbCrLf
        outputText = outputText & "¦ąĹăĄÜľąłĚ«áżů»~˝XˇG" & CStr(displays(i).LastError) & vbCrLf

        If displays(i).HDRCapabilityKnown Then
            outputText = outputText & "ĄŘ«e HDR Ş¬şAˇG" & HDR_GetStatusText(HDR_GetStatus(i)) & vbCrLf
        Else
            outputText = outputText & "ĄŘ«e HDR Ş¬şAˇGµLŞkĄŃÂÂŞ© API şë˝T§PÂ_" & vbCrLf
        End If

        outputText = outputText & String$(78, "-") & vbCrLf
    Next i

    outputText = outputText & "ĄŘ«eŞ¬şAĹŞ¨ú§ą¦¨ˇC" & vbCrLf
    txtOutput.Text = outputText
    Exit Sub

NoDisplays:
    txtOutput.Text = outputText & "ĄŘ«e¨S¦ł±ŇĄÎ¤¤ŞşĹăĄÜľą PathˇC" & vbCrLf
End Sub

' ±N Boolean Âŕ¦¨©T©wÁcĹé¤¤¤ĺˇAÁ×§KżéĄX True / False ­^¤ĺ¦rĽËˇC
Private Function BoolText(ByVal value As Boolean) As String
    If value Then
        BoolText = "¬O"
    Else
        BoolText = "§_"
    End If
End Function

' ±N 32-bit Long ĄH©T©w 8 ˝X¤Q¤»¶i¦ěĹăĄÜˇFło¬OĂŃ§O˝XˇA¤Ł¬O¦ŰµM»y¨Ą¤ĺ¦rˇC
Private Function SignedHex8(ByVal value As Long) As String
    SignedHex8 = "0x" & Right$("00000000" & Hex$(value), 8)
End Function
