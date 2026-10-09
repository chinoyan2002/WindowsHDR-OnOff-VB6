$ErrorActionPreference = "Stop"
$tempFile = "$env:TEMP\monitors_status.csv"
$powerOnCount = 0

Write-Host "=========================================="
Write-Host "[DEBUG] 開始執行實體螢幕 DDC/CI 檢測"
Write-Host "[DEBUG] 目前時間: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "[DEBUG] 暫存檔預定路徑: $tempFile"
Write-Host "=========================================="

$exePath = ".\ControlMyMonitor.exe"
if (-not (Test-Path $exePath)) {
    Write-Error "找不到執行檔：$exePath，請確認是否與腳本放在同一個目錄下。"
    exit
}
Write-Host "[DEBUG] 確認執行檔存在: $(Resolve-Path$exePath)"

Write-Host "[DEBUG] 正在呼叫 ControlMyMonitor 匯出狀態..."
$cmdLine = ".\ControlMyMonitor.exe /scomma `"$tempFile`""
Write-Host "[DEBUG] 執行命令: $cmdLine"

$null = cmd.exe /c$cmdLine
Write-Host "[DEBUG] cmd.exe 執行完畢, 返回碼: $LASTEXITCODE"

if (-not (Test-Path $tempFile)) {
    Write-Error "匯出失敗：找不到暫存檔案 $tempFile"
    exit
}

$fileInfo = Get-Item$tempFile
Write-Host "[DEBUG] 暫存檔產生成功！檔案大小: $($fileInfo.Length) bytes"

Write-Host "[DEBUG] 正在解析 CSV 內容..."
$csvData = Import-Csv -Path$tempFile
Write-Host "[DEBUG] CSV 總行數 (含標題): $($csvData.Count)"

if ($csvData.Count -eq 0) {
    Write-Warning "[WARNING] CSV 檔案內容為空，沒有讀取到任何螢幕資料。"
} else {
    $headers = ($csvData[0].psobject.properties.name) -join ", "
    Write-Host "[DEBUG] CSV 欄位清單: [$headers]"

    $index = 0
    foreach ($row in $csvData) {$index++
        Write-Host "[DEBUG-ROW $index] 螢幕名稱: '$($row.'Monitor Name')' | VCP Code: '$($row.'VCP Code')' | 數值 (Value): '$($row.'Value')' | 描述: '$($row.'Description')'"
    }

    $d6Rows =$csvData | Where-Object { $_.'VCP Code' -eq 'D6' -or$_.'VCP Code' -eq '0xD6' }
    Write-Host "[DEBUG] 篩選出 D6 (電源模式) 的資料筆數: $(@($d6Rows).Count)"

    foreach ($row in$d6Rows) {
        $monitorName =$row.'Monitor Name'
        $value =$row.'Value'
        
        Write-Host "[CHECK] 檢查螢幕: $monitorName | VCP Code: D6 | 數值: '$value'"
        
        if ($value -eq "1") {
            $powerOnCount++
            Write-Host "   └── [結果] 狀態為 1 -> 判定為【通電中】"
        } else {
            Write-Host "   └── [結果] 狀態為 '$value' -> 判定為【已關閉或待機】"
        }
    }
}

Remove-Item -Force $tempFile -ErrorAction SilentlyContinue
Write-Host "[DEBUG] 已順利清除暫存檔。"

Write-Host "=========================================="
Write-Host "[DEBUG] 最終統計: 通電中的實體螢幕數量 = $powerOnCount"
if ($powerOnCount -gt 0) {
    Write-Host "檢測結果：目前有 $powerOnCount 台實體螢幕通電中，【不執行】後續動作。"
} else {
    Write-Host "檢測結果：所有實體螢幕皆已斷電，【可觸發】後續虛擬顯卡邏輯。"
}
Write-Host "=========================================="