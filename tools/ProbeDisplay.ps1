<#
.SYNOPSIS
  Read-only display/HDR probe for AutoHdrMode. Makes NO changes.
  Captures: Pnp displays, WMI physical monitor presence, HDRCmd status.
#>
param(
  [string]$HdrCmdPath = "",
  [int]$Samples = 1,
  [int]$IntervalSeconds = 10
)

# Default: HDRCmd.exe next to the program (single portable folder),
# with legacy fallback to E:\HDRTray-v0.5.5.
if ([string]::IsNullOrWhiteSpace($HdrCmdPath)) {
  $here = Split-Path -Parent $PSCommandPath
  $cand = Join-Path (Join-Path $here "..\\dist") "HDRCmd.exe"
  if (Test-Path -LiteralPath $cand) { $HdrCmdPath = $cand }
  else { $HdrCmdPath = "E:\\HDRTray-v0.5.5\\HDRCmd.exe" }
}

for ($i = 1; $i -le $Samples; $i++) {
  Write-Output ("===== SAMPLE {0}/{1} {2} =====" -f $i, $Samples, (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))

  Write-Output "--- PnpDisplay ---"
  Get-PnpDevice -Class Display -ErrorAction SilentlyContinue |
    Select-Object FriendlyName, InstanceId, Status |
    Format-Table -AutoSize | Out-String | Write-Output

  Write-Output "--- WmiPhysical (count + Active) ---"
  $wmi = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue |
    Select-Object InstanceName, Active
  $count = @($wmi).Count
  Write-Output ("PhysicalCount={0}" -f $count)
  $wmi | Format-Table -AutoSize | Out-String | Write-Output

  Write-Output "--- HdrCmd status -m long ---"
  if (Test-Path -LiteralPath $HdrCmdPath) {
    & $HdrCmdPath status -m long
    Write-Output ("HdrCmdExit={0}" -f $LASTEXITCODE)
  } else {
    Write-Warning ("HdrCmd not found: {0}" -f $HdrCmdPath)
  }

  if ($i -lt $Samples) { Start-Sleep -Seconds $IntervalSeconds }
}
