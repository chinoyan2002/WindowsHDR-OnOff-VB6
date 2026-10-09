<#
.SYNOPSIS
    純原生 PowerShell DDC/CI 實體螢幕電源檢測腳本 (不依賴任何外部工具)
#>

$ErrorActionPreference = "Stop"

# 1. 宣告 C# P/Invoke 結構與 Win32 API 函式庫，直接調用 Windows 內建的 dxva2.dll
$csharpCode = @"
using System;
using System.Runtime.InteropServices;

public class NativeMonitorControl {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct PHYSICAL_MONITOR {
        public IntPtr hPhysicalMonitor;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szPhysicalMonitorDescription;
    }

    [DllImport("dxva2.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetPhysicalMonitorsFromHMONITOR(IntPtr hMonitor, uint dwPhysicalMonitorArraySize, [Out] PHYSICAL_MONITOR[] pPhysicalMonitorArray);

    [DllImport("dxva2.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DestroyPhysicalMonitors(uint dwPhysicalMonitorArraySize, [In, Out] PHYSICAL_MONITOR[] pPhysicalMonitorArray);

    [DllImport("dxva2.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetVCPFeatureAndVCPFeatureReply(IntPtr hMonitor, byte bVCPCode, ref uint pvct, ref uint pdwCurrentValue, ref uint pdwMaximumValue);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);

    public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref System.Drawing.Rectangle lprcMonitor, IntPtr dwData);
}
"@

# 載入 C# 類型
try {
    Add-Type -TypeDefinition $csharpCode -ReferencedAssemblies "System.Drawing.dll" -ErrorAction Stop
} catch {
    # 若已經載入過則略過
}

Write-Host "=========================================="
Write-Host "[DEBUG] 開始執行原生 Windows API 螢幕檢測"
Write-Host "=========================================="

$powerOnCount = 0
$monitorList = [System.Collections.Generic.List[IntPtr]]::new()

# 2. 列舉系統中所有實體顯示器的 HMONITOR
$callback = {
    param([IntPtr]$hMonitor, [IntPtr]$hdcMonitor, [ref]$lprcMonitor, [IntPtr]$dwData)
    $script:monitorList.Add($hMonitor)
    return $true
}

[NativeMonitorControl+MonitorEnumProc] $enumProc = $callback
[NativeMonitorControl]::EnumDisplayMonitors([IntPtr]::Zero, [IntPtr]::Zero, $enumProc, [IntPtr]::Zero) | Out-Null

Write-Host "[DEBUG] 偵測到掛載的螢幕視窗代碼數量: $($monitorList.Count)"

$index = 0
foreach ($hMon in $monitorList) {
    $index++
    $pcount = 0
    
    # 取得實體螢幕數量
    # 這裡我們簡化透過 Win32 取得實體控制代碼
    $phyMons = New-Object NativeMonitorControl+PHYSICAL_MONITOR[] 1
    
    # 嘗試從 HMONITOR 取得實體螢幕物件
    if ([NativeMonitorControl]::GetPhysicalMonitorsFromHMONITOR($hMon, 1, $phyMons)) {
        $desc = $phyMons[0].szPhysicalMonitorDescription
        $hPhysical = $phyMons[0].hPhysicalMonitor
        
        # 準備讀取 D6 VCP Code (電源模式)
        $vct = 0
        $currentVal = 0
        $maxVal = 0
        
        # 呼叫底層 API 讀取 VCP 狀態
        $success = [NativeMonitorControl]::GetVCPFeatureAndVCPFeatureReply($hPhysical, 0xD6, [ref]$vct, [ref]$currentVal, [ref]$maxVal)
        
        if ($success) {
            Write-Host "[CHECK] 實體螢幕 [$index] 描述: '$desc' | 電源狀態碼 (D6): $currentVal (最大值: $maxVal)"
            # 數值 1 通常代表開機通電 (On)，4 代表關閉/待機 (Off)
            if ($currentVal -eq 1) {
                $powerOnCount++
                Write-Host "   └── [結果] 判定為【通電中】"
            } else {
                Write-Host "   └── [結果] 判定為【已關閉或待機】(狀態碼: $currentVal)"
            }
        } else {
            Write-Host "[WARNING] 螢幕 [$index] ($desc) 支援實體監控，但無法透過 DDC/CI 讀取 D6 狀態（可能是虛擬螢幕或不支援 DDC/CI）"
        }
        
        # 釋放實體螢幕資源
        [NativeMonitorControl]::DestroyPhysicalMonitors(1, $phyMons) | Out-Null
    } else {
        Write-Host "[DEBUG] 螢幕 [$index] 無法對應到實體監視器結構（可能是虛擬顯示卡 / usbmmidd 虛擬螢幕）"
    }
}

Write-Host "=========================================="
Write-Host "[DEBUG] 最終統計: 真正通電中的實體螢幕數量 = $powerOnCount"
if ($powerOnCount -gt 0) {
    Write-Host "檢測結果：目前有 $powerOnCount 台實體螢幕通電中，【不執行】後續動作。"
} else {
    Write-Host "檢測結果：所有實體螢幕皆已斷電，【可觸發】後續虛擬顯卡邏輯。"
}
Write-Host "=========================================="