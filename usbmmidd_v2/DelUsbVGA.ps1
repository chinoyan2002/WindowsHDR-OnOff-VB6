$TestMode = $false # 宣告測試模式開關，設為 $false 以進行真實部署
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Warning "請以系統管理員身分執行"; exit } # 檢查是否具備系統管理員權限

$pnpPath = if (Test-Path "$env:windir\sysnative\pnputil.exe") { "$env:windir\sysnative\pnputil.exe" } else { "$env:windir\System32\pnputil.exe" } # 動態判斷並取得繞過32位元重新導向的正確系統工具路徑

$targetGpus = 'AMD|NVIDIA|Intel' # 定義要保留的顯示卡廠牌關鍵字
$virtualGpus = Get-PnpDevice -Class Display | Where-Object { $_.FriendlyName -notmatch $targetGpus } # 尋找所有不含這三大廠牌的顯示裝置

if ($virtualGpus) {
    $count = @($virtualGpus).Count # 計算總共有幾張虛擬顯卡需要處理
    Write-Host "偵測到 $count 張虛擬或額外顯卡，準備處理..." # 提示偵測到的數量
    
    foreach ($dev in $virtualGpus) {
        Write-Host "目標裝置: $($dev.FriendlyName)" # 顯示目前正在處理哪一張虛擬顯卡
        if ($TestMode) { Write-Host "[TEST] 模擬執行: & `"$pnpPath`" /remove-device $($dev.InstanceId)" } else { & $pnpPath /remove-device $dev.InstanceId } # 使用動態安全路徑呼叫工具確保不被重新導向干擾
        Start-Sleep -Seconds 1 # 每次處理後間隔1秒，避免系統核心衝突
    }
    
    $verified = $false # 初始化驗證變數
    $timer = [Diagnostics.Stopwatch]::StartNew() # 啟動計時器限制最長等待時間
    while ($timer.Elapsed.TotalSeconds -lt 45) {
        $remaining = Get-PnpDevice -Class Display | Where-Object { $_.FriendlyName -notmatch $targetGpus } # 重新檢查是否還有漏網之魚
        if ($TestMode) { $verified = $true; break } # 若為測試模式則直接強制視為驗證通過並跳出迴圈避免死鎖
        if (-not $remaining) { $verified = $true; break } # 真實模式下確認已卸載則跳出迴圈
        Start-Sleep -Seconds 2 # 尚未完全卸載則每2秒複查一次
    }
    
    if (-not $verified) { Write-Warning "等待多張虛擬顯卡完全移除超時，已終止後續動作"; exit } # 超時警示
    if ($TestMode) { Write-Host "[TEST] 測試模式模擬完成，硬體未實際移除" } else { Write-Host "已確認系統中的虛擬顯卡已全數清除，僅剩實體顯示卡" } # 依據模式顯示對應的完成提示
} else {
    Write-Host "系統中僅有實體顯示卡 (AMD/NVIDIA/Intel)，跳過移除程序" # 無虛擬顯卡則直接跳過
}

