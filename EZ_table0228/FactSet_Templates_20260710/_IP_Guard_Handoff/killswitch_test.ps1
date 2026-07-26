# killswitch_test.ps1 -- observe whether NordVPN App Kill Switch kills FactSetVpnProxy
# on a VPN drop. Fast 0.5s poll; logs proxy PID, FDS-family count, listen ports 3128/8765,
# and NordLynx adapter status. Run this, then manually Disconnect NordVPN and watch the
# proxy PID flip to DEAD + FDS count drop to 0 -> proves App KS terminates them.
# Self-locating log next to script. Stop via killswitch_test.STOP. ASCII only.
$sp   = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$log  = Join-Path $sp 'killswitch_test.log'
$stop = Join-Path $sp 'killswitch_test.STOP'
if (Test-Path $stop) { Remove-Item $stop -Force }
$fdsNames = @('FactSetVpnProxy','FDSPipe','FDSConduit','FdsOfcExec','FDSRealTime','FDSEmsRealTime','fdsw32','FDSWorkstation_x64','FDSTray','ApiBox','ApiBox_x64','FDSBrowser','FDSBrowser_x64','FDSOCMon','FDSDllHost','fdswFixExcel','FDSDiagnostics','FDSDiagnostics_x64','FDSChartCopier','FDSStatBar','MarqueeRestart','FDSUpdateDialog','fdsup','fdsup_x64','FDSSdm')

"=== killswitch_test start $(Get-Date -Format 'HH:mm:ss.fff') ===" | Out-File $log -Encoding utf8
$prev = ''
for ($k = 0; $k -lt 600; $k++) {   # ~300s max
  if (Test-Path $stop) { "=== STOP $(Get-Date -Format 'HH:mm:ss.fff') ===" | Add-Content $log; break }
  $t = Get-Date -Format 'HH:mm:ss.fff'
  $pp = Get-Process FactSetVpnProxy -EA SilentlyContinue
  $proxyPid = if ($pp) { $pp.Id } else { 'DEAD' }
  $fdsCount = @(Get-Process -EA SilentlyContinue | Where-Object { ($_.Name -replace '\.exe$','') -in $fdsNames }).Count
  $p3128 = @(Get-NetTCPConnection -LocalPort 3128 -State Listen -EA SilentlyContinue).Count
  $p8765 = @(Get-NetTCPConnection -LocalPort 8765 -State Listen -EA SilentlyContinue).Count
  $nl = (Get-NetAdapter -Name NordLynx -EA SilentlyContinue).Status
  if (-not $nl) { $nl = 'absent' }
  $line = "proxy=$proxyPid fds=$fdsCount p3128=$p3128 p8765=$p8765 NordLynx=$nl"
  # log every sample (compact) so the kill + recovery transition is fully captured
  "[$t] $line" | Add-Content $log
  Start-Sleep -Milliseconds 500
}
"=== killswitch_test exit $(Get-Date -Format 'HH:mm:ss.fff') ===" | Add-Content $log
