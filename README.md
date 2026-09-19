# WindowsHDR-OnOff-VB6

VB6 直接呼叫 Windows Display Configuration API 控制 Windows HDR。

## 目前版本

v1.2

## 設計目標

本專案完全不依賴 HDRTray、HDRCmd 或其他第三方 HDR 控制程式，直接使用 Windows 原生 API。

## 核心功能

- Windows 11 24H2 / Build 26100 以上使用新版 HDR API
- 舊版 Windows 使用 Advanced Color 相容 API
- 正確區分 SDR、WCG、HDR
- HDR capability、使用者設定、實際 Active 狀態分離
- 多顯示器支援
- Adapter LUID + Target ID 識別
- QueryDisplayConfig 遇到 ERROR_INSUFFICIENT_BUFFER 最多重試 3 次
- 設定完成後重新查詢並驗證實際狀態
- 顯示器 Friendly Name 與 EDID / Forced 狀態
- Win32 錯誤碼轉為繁體中文訊息
- 函數、重要變數、結構與 API 結構均有繁體中文註解
- 測試程式的狀態與記錄訊息使用繁體中文

## Windows 11 24H2+

使用：

- DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2
- DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE

可精確取得 HDR capability，並以 Active Color Mode 判斷實際 SDR / WCG / HDR 狀態。

## 舊版 Windows

使用：

- DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO
- DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE

舊 API 主要描述 Advanced Color，因此程式不會把 Advanced Color Supported 直接當成 HDR Supported；HDR Capability 會標記為未知。

## 色彩模式

- 0：SDR（標準動態範圍）
- 1：WCG（廣色域）
- 2：HDR（高動態範圍）

## 檔案

- `modWindowsHDR.bas)：核心 API 模組
- `frmHDRTest.frm)：VB6 測試表單

## 測試

加入 VB6 專案後執行 `frmHDRTest`。

測試功能：

- 讀取狀態
- 開啟 HDR
- 關閉 HDR
- HDR 切換

目前尚需在實際 VB6 IDE、Windows 11 24H2 / 25H2 與實際多螢幕環境完成編譯及實機驗證。

## 後續

底層模組刻意不包含 Timer 與 UI 監控。若 VMini6 需要常駐偵測，可在上層加入 WM_DISPLAYCHANGE、WM_DEVICECHANGE 或週期性重新查詢。
