$TestMode = $false # 宣告測試模式開關，設為 $false 以進行真實部署
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Warning "請以系統管理員身分執行"; exit } # 檢查是否具備系統管理員權限

$pnpPath = if (Test-Path "$env:windir\sysnative\pnputil.exe") { "$env:windir\sysnative\pnputil.exe" } else { "$env:windir\System32\pnputil.exe" } # 動態判斷並取得繞過32位元重新導向的正確系統工具路徑
$targetGpus = 'AMD|NVIDIA|Intel|Microsoft Basic' # 排除實體與系統基本顯示卡
$allVirtual = @(Get-PnpDevice -Class Display | Where-Object { $_.FriendlyName -notmatch $targetGpus }) # 取得所有虛擬顯示卡
$workingVirtual = @($allVirtual | Where-Object { $_.Status -eq 'OK' -and $_.ConfigManagerErrorCode -eq 0 }) # 篩選正常運作無驚嘆號的裝置

if ($workingVirtual.Count -eq 0) {
    Write-Host "未偵測到正常虛擬顯卡，執行安裝批次檔..." # 觸發安裝條件
    $batPath = "$PSScriptRoot\usbmmidd2.bat" # 請將此處替換為您寫好的安裝批次檔路徑
    if (-not $TestMode -and (Test-Path $batPath)) { Start-Process -FilePath $batPath -Wait } else { Write-Host "[TEST] 模擬執行: $batPath" } # 執行並等待安裝
    
    $timer = [Diagnostics.Stopwatch]::StartNew() # 啟動安裝驗證計時器
    while ($timer.Elapsed.TotalSeconds -lt 45) {
        $checkVirtual = @(Get-PnpDevice -Class Display | Where-Object { $_.FriendlyName -notmatch $targetGpus -and $_.Status -eq 'OK' -and $_.ConfigManagerErrorCode -eq 0 }) # 重新檢查狀態
        if ($checkVirtual.Count -gt 0 -or $TestMode) { Write-Host "驗證成功：虛擬顯示卡已成功安裝且啟用正常"; break } # 驗證通過即跳出
        Start-Sleep -Seconds 2 # 每2秒複查
    }
} else {
    $keepDevice = $workingVirtual[0] # 保留第一張正常運作的顯卡
    $devicesToRemove = @($allVirtual | Where-Object { $_.InstanceId -ne $keepDevice.InstanceId }) # 找出其餘需移除的顯卡(包含故障與多餘的)
    
    if ($devicesToRemove.Count -gt 0) {
        Write-Host "保留裝置: $($keepDevice.FriendlyName)，移除其餘 $($devicesToRemove.Count) 張..." # 提示移除數量
        foreach ($dev in $devicesToRemove) {
            if ($TestMode) { Write-Host "[TEST] 模擬移除: $($dev.FriendlyName)" } else { & $pnpPath /remove-device $dev.InstanceId } # 執行卸載
            Start-Sleep -Seconds 1 # 間隔1秒避免衝突
        }
        
        $timer = [Diagnostics.Stopwatch]::StartNew() # 啟動移除驗證計時器
        while ($timer.Elapsed.TotalSeconds -lt 45) {
            $checkRemaining = @(Get-PnpDevice -Class Display | Where-Object { $_.FriendlyName -notmatch $targetGpus }) # 重新取得目前所有虛擬顯卡
            if (($checkRemaining.Count -eq 1 -and $checkRemaining[0].InstanceId -eq $keepDevice.InstanceId) -or $TestMode) { Write-Host "驗證成功：僅保留一張正常運作的虛擬顯卡"; break } # 驗證通過
            Start-Sleep -Seconds 2 # 每2秒複查
        }
    } else {
        Write-Host "已有一張正常運作的虛擬顯卡，無須處理。" # 狀態已達預期標準
    }
}