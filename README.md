# WindowsHDR-OnOff-VB6

VB6 直接呼叫 Windows Display Configuration API 控制 Windows HDR。

## 版本

v1.4

## 特色

- 完全不依賴 HDRTray 或 HDRCmd
- Windows 11 24H2 / Build 26100 以上使用新版 HDR API
- 舊版 Windows 使用 Advanced Color 相容 API
- 正確區分 SDR、WCG、HDR
- HDR 能力、使用者設定、實際啟用狀態分離
- 多顯示器支援
- Adapter LUID + Target ID 識別
- QueryDisplayConfig 遇到 ERROR_INSUFFICIENT_BUFFER 最多重試 3 次
- Set 完成後重新查詢並驗證實際狀態
- 取得 Friendly Name、EDID 來源與 Forced Target 狀態
- 錯誤與狀態訊息統一使用繁體中文
- 使用 20-byte RTL_OSVERSIONINFOW 取得 Major / Minor / Build，並輸出完整版本
- 函數、重要變數與結構均加入繁體中文註解
- 可在 VB6 IDE 啟用 XP 相容模式時，仍以 Registry 取得實際 Windows Build

## Windows 11 24H2 以上

使用：

- DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO_2
- DISPLAYCONFIG_DEVICE_INFO_SET_HDR_STATE

新版 API 可以精確取得 HDR capability，並使用 Active Color Mode 判斷：

- SDR：標準動態範圍
- WCG：廣色域
- HDR：高動態範圍

## 舊版 Windows

使用：

- DISPLAYCONFIG_DEVICE_INFO_GET_ADVANCED_COLOR_INFO
- DISPLAYCONFIG_DEVICE_INFO_SET_ADVANCED_COLOR_STATE

舊 API 無法可靠區分 WCG 與 HDR capability，因此程式會將 HDR capability 標示為未知，避免把 Advanced Color 誤報為 HDR。

## 檔案

- `modWindowsHDR.bas`：Windows HDR / Advanced Color 核心 API 模組
- `frmHDRTest.frm`：VB6 測試表單

## 使用

1. 建立 VB6 Standard EXE 專案。
2. 匯入 `modWindowsHDR.bas`。
3. 匯入 `frmHDRTest.frm`。
4. 將 `frmHDRTest` 設定為啟動表單。
5. 執行後先按「讀取狀態」。
6. 測試「開啟 HDR」、「關閉 HDR」、「切換 HDR」。

## 注意事項

本專案針對 VB6 32-bit。Windows Display Configuration 結構以 Byte Array + RtlMoveMemory 解析，降低 VB6 對 Union、Bit Field 與結構對齊的風險。

版本資訊優先從 Windows Registry 取得，避免 VB6 IDE 的 XP 相容模式造成 5.1.2600 偽版本；HDR API 路徑則直接探測新版 API。仍需在實際 VB6 IDE 與 Windows 11 24H2 / 25H2 環境完成編譯及實機驗證。

底層 API 模組刻意不加入 Timer、WM_DISPLAYCHANGE 或 WM_DEVICECHANGE；
