VERSION 5.00
Begin VB.Form frmHDRTest
   Caption         =   "Windows HDR 原生 API 測試"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "frmHDRTest"
   ScaleHeight     =   9480
   ScaleWidth      =   12480
   StartUpPosition =   2
   Begin VB.CommandButton cmdRead
      Caption         =   "讀取狀態"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdEnable
      Caption         =   "開啟 HDR"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdDisable
      Caption         =   "關閉 HDR"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   1800
   End
   Begin VB.CommandButton cmdToggle
      Caption         =   "切換 HDR"
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
Attribute VB_Description = "Windows HDR 原生 API 測試表單"
Attribute VB_Ext_KEY = "SavedWithLastKnownToolVersion", "6.00.9782.0"

Option Explicit

' 讀取所有顯示器的目前狀態並顯示在記錄視窗。
Private Sub cmdRead_Click()
    ShowHDRStatus False
End Sub

' 開啟所有 active 顯示器 HDR，設定後重新查詢驗證。
Private Sub cmdEnable_Click()
    AppendLog "開始執行：開啟 HDR。"

    If HDR_Enable() Then
        AppendLog "結果：HDR 設定成功，且已完成實際狀態驗證。"
    Else
        AppendLog "結果：HDR 設定失敗。"
        AppendLog "詳細資訊：" & HDR_GetLastErrorText()
    End If

    ShowHDRStatus True
End Sub

' 關閉所有 active 顯示器 HDR，設定後重新查詢驗證。
Private Sub cmdDisable_Click()
    AppendLog "開始執行：關閉 HDR。"

    If HDR_Disable() Then
        AppendLog "結果：HDR 關閉成功，且已完成實際狀態驗證。"
    Else
        AppendLog "結果：HDR 關閉失敗。"
        AppendLog "詳細資訊：" & HDR_GetLastErrorText()
    End If

    ShowHDRStatus True
End Sub

' 依每一個顯示器目前狀態切換 HDR。
Private Sub cmdToggle_Click()
    AppendLog "開始執行：切換 HDR。"

    If HDR_Toggle() Then
        AppendLog "結果：HDR 切換成功，且已完成實際狀態驗證。"
    Else
        AppendLog "結果：HDR 切換失敗。"
        AppendLog "詳細資訊：" & HDR_GetLastErrorText()
    End If

    ShowHDRStatus True
End Sub

' 表單載入時立即顯示目前 HDR 狀態。
Private Sub Form_Load()
    ShowHDRStatus False
End Sub

' 將一行繁體中文記錄加入輸出視窗。
' messageText：要記錄的中文訊息。
Private Sub AppendLog(ByVal messageText As String)
    If Len(txtOutput.Text) > 0 Then
        txtOutput.Text = txtOutput.Text & vbCrLf
    End If

    txtOutput.Text = txtOutput.Text & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "　" & messageText
    txtOutput.SelStart = Len(txtOutput.Text)
End Sub

' 顯示 Windows Build、API 模式與所有顯示器的狀態。
' Append：True=接續現有記錄；False=清空後重新顯示。
Private Sub ShowHDRStatus(ByVal Append As Boolean)
    Dim displays() As HDR_DISPLAY_INFO       ' 所有 active 顯示器。
    Dim i As Long                            ' 顯示器索引。
    Dim majorVersion As Long                 ' Windows Major Version。
    Dim minorVersion As Long                 ' Windows Minor Version。
    Dim buildNumber As Long                  ' Windows Build。
    Dim outputText As String                 ' 本次產生的繁體中文輸出。

    If Append Then
        outputText = txtOutput.Text

        If Len(outputText) > 0 Then
            outputText = outputText & vbCrLf & vbCrLf
        End If
    Else
        outputText = vbNullString
    End If

    If HDR_GetWindowsVersionInfo(majorVersion, minorVersion, buildNumber) Then
        outputText = outputText & "Windows 版本：" & CStr(majorVersion) & "." & CStr(minorVersion) & "." & CStr(buildNumber) & vbCrLf
    Else
        outputText = outputText & "Windows 版本：無法取得" & vbCrLf
    End If

    outputText = outputText & "HDR API 模式：" & HDR_GetApiModeText() & vbCrLf
    outputText = outputText & "版本資訊來源：" & HDR_GetWindowsVersionSource() & vbCrLf
    outputText = outputText & String$(78, "=") & vbCrLf

    If Not HDR_GetDisplays(displays) Then
        outputText = outputText & "讀取顯示器失敗：" & HDR_GetLastErrorText() & vbCrLf
        txtOutput.Text = outputText
        txtOutput.SelStart = Len(txtOutput.Text)
        Exit Sub
    End If

    On Error GoTo NoDisplays

    For i = LBound(displays) To UBound(displays)
        outputText = outputText & "顯示器 #" & CStr(displays(i).Index) & vbCrLf
        outputText = outputText & "名稱：" & displays(i).Name & vbCrLf
        outputText = outputText & "Adapter LUID：0x" & SignedHex8(displays(i).AdapterHigh) & SignedHex8(displays(i).AdapterLow) & vbCrLf
        outputText = outputText & "Target ID：" & CStr(displays(i).TargetId) & vbCrLf

        outputText = outputText & "HDR 支援：" & BoolText(displays(i).HDRSupported) & vbCrLf
        outputText = outputText & "HDR 能力可判定：" & BoolText(displays(i).HDRCapabilityKnown) & vbCrLf
        outputText = outputText & "HDR 使用者設定：" & BoolText(displays(i).HDRUserEnabled) & vbCrLf
        outputText = outputText & "HDR 實際啟用：" & BoolText(displays(i).HDRActive) & vbCrLf

        outputText = outputText & "Advanced Color 支援：" & BoolText(displays(i).AdvancedColorSupported) & vbCrLf
        outputText = outputText & "Advanced Color 實際啟用：" & BoolText(displays(i).AdvancedColorActive) & vbCrLf
        outputText = outputText & "Advanced Color 受政策限制：" & BoolText(displays(i).AdvancedColorLimitedByPolicy) & vbCrLf
        outputText = outputText & "Advanced Color 強制停用：" & BoolText(displays(i).AdvancedColorForceDisabled) & vbCrLf

        outputText = outputText & "WCG 支援：" & BoolText(displays(i).WCGSupported) & vbCrLf
        outputText = outputText & "WCG 使用者設定：" & BoolText(displays(i).WCGUserEnabled) & vbCrLf
        outputText = outputText & "Wide Color 強制：" & BoolText(displays(i).WideColorEnforced) & vbCrLf

        outputText = outputText & "目前色彩模式：" & ColorModeText(displays(i).ActiveColorMode) & vbCrLf
        outputText = outputText & "名稱來自 EDID：" & BoolText(displays(i).FriendlyNameFromEdid) & vbCrLf
        outputText = outputText & "強制 Target：" & BoolText(displays(i).FriendlyNameForced) & vbCrLf
        outputText = outputText & "最後錯誤碼：" & CStr(displays(i).LastError) & vbCrLf

        If displays(i).HDRCapabilityKnown Then
            outputText = outputText & "HDR 狀態：" & HDR_GetStatusText(HDR_GetStatus(i)) & vbCrLf
        Else
            outputText = outputText & "HDR 狀態：舊版 API 無法精確拆分 HDR 與 WCG。" & vbCrLf
            outputText = outputText & "Advanced Color 狀態：" & BoolText(displays(i).AdvancedColorActive) & vbCrLf
        End If

        outputText = outputText & String$(78, "-") & vbCrLf
    Next i

    txtOutput.Text = outputText
    txtOutput.SelStart = Len(txtOutput.Text)
    Exit Sub

NoDisplays:
    txtOutput.Text = outputText & "目前沒有作用中的顯示器 Path。" & vbCrLf
    txtOutput.SelStart = Len(txtOutput.Text)
End Sub

' 將 Boolean 狀態轉成繁體中文「是 / 否」。
Private Function BoolText(ByVal value As Boolean) As String
    If value Then
        BoolText = "是"
    Else
        BoolText = "否"
    End If
End Function

' 將 Active Color Mode 轉成繁體中文。
Private Function ColorModeText(ByVal modeValue As Long) As String
    Select Case modeValue
        Case 0
            ColorModeText = "SDR（標準動態範圍）"
        Case 1
            ColorModeText = "WCG（廣色域）"
        Case 2
            ColorModeText = "HDR（高動態範圍）"
        Case Else
            ColorModeText = "未知（" & CStr(modeValue) & "）"
    End Select
End Function

' 將 32-bit Long 顯示成固定八碼十六進位。
Private Function SignedHex8(ByVal value As Long) As String
    SignedHex8 = Right$("00000000" & Hex$(value), 8)
End Function
