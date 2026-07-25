# capture_a2.ps1 -- fine-grained (0.5s) leak capture for a single A2 fetch.
# Complements leak_monitor2.ps1 (2s cadence) by catching short-lived gateway connections.
# Watches every EXTERNAL (public remote) established connection owned by:
#   - EXCEL.EXE and its msedgewebview2 descendants (the FactSet Add-in task pane), and
#   - the FDS process family (FactSetVpnProxy, FDSPipe, FdsOfcExec, ...).
# Classifies by LOCAL address:
#   TUNNEL   = local in NordLynx tunnel (100.126.x live adapter IP, or legacy 10.5.*)  -> OK
#   LOOPBACK = remote 127.0.0.1:3128 (WV2 correctly using the proxy)                   -> OK
#   PHYS-LEAK= local = physical Taiwan IP 111.185.192.56 (or other non-tunnel public)  -> LEAK
# Logs to capture_a2.log next to this script. Runs ~150s or until capture_a2.STOP appears.
$sp   = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$log  = Join-Path $sp 'capture_a2.log'
$stop = Join-Path $sp 'capture_a2.STOP'
if (Test-Path $stop) { Remove-Item $stop -Force }

$tunIPs = @()
try { $tunIPs = @(Get-NetIPAddress -InterfaceAlias 'NordLynx' -AddressFamily IPv4 -EA SilentlyContinue | Select-Object -ExpandProperty IPAddress) } catch {}
$physIP = '111.185.192.56'
$fdsNames = @('FactSetVpnProxy','FDSPipe','FDSConduit','FdsOfcExec','FDSRealTime','FDSEmsRealTime','fdsw32','FDSWorkstation_x64','FDSTray','ApiBox','ApiBox_x64','FDSBrowser','FDSBrowser_x64')
function Test-Tun([string]$la) { if ($la -like '10.5.*') { return $true }; if ($tunIPs -contains $la) { return $true }; return $false }

"=== capture_a2 start $(Get-Date -Format 'HH:mm:ss.fff') ===" | Out-File $log -Encoding utf8
"tunnel IP(s): $([string]::Join(', ',$tunIPs)) ; physical(leak) IP: $physIP" | Add-Content $log

$gwPrefix = '192.234.235.'   # known FactSet gateway /24 (runbook Part E)
$gwSet = @{}                 # learned gateway IPs = remotes FactSetVpnProxy reaches via tunnel
$leakSeen = 0; $tunSeen = 0; $proxySeen = 0
for ($k = 0; $k -lt 300; $k++) {
  if (Test-Path $stop) { "=== STOP $(Get-Date -Format 'HH:mm:ss.fff') ===" | Add-Content $log; break }
  $t = Get-Date -Format 'HH:mm:ss.fff'

  $snap = @{}
  foreach ($p in (Get-CimInstance Win32_Process -Property ProcessId,ParentProcessId,Name -EA SilentlyContinue)) {
    $snap[[int]$p.ProcessId] = @{ n = $p.Name; pp = [int]$p.ParentProcessId }
  }
  $xlPids = @($snap.Keys | Where-Object { $snap[$_].n -eq 'EXCEL.EXE' })
  $wv2All = @($snap.Keys | Where-Object { $snap[$_].n -eq 'msedgewebview2.exe' })
  $xlWv2 = @{}
  foreach ($w in $wv2All) {
    $cur = $w
    for ($h = 0; $h -lt 14; $h++) {
      if (-not $snap.ContainsKey($cur)) { break }
      $pp = $snap[$cur].pp
      if ($xlPids -contains $pp) { $xlWv2[$w] = $true; break }
      if ($pp -eq 0 -or $pp -eq $cur) { break }
      $cur = $pp
    }
  }
  $fdsPids = @{}
  foreach ($id in $snap.Keys) {
    $bn = $snap[$id].n -replace '\.exe$',''
    if ($fdsNames -contains $bn) { $fdsPids[$id] = $bn }
  }

  $est = Get-NetTCPConnection -State Established -EA SilentlyContinue
  # Pass 1: learn gateway IPs = external remotes that FactSetVpnProxy reaches over the tunnel.
  foreach ($c in $est) {
    $op = [int]$c.OwningProcess
    if (-not $fdsPids.ContainsKey($op)) { continue }
    if ($fdsPids[$op] -ne 'FactSetVpnProxy') { continue }
    $ra = $c.RemoteAddress
    if ($ra -match '^(127\.|::1|::$|fe80)') { continue }
    if (Test-Tun $c.LocalAddress) { $gwSet[$ra] = $true }
  }
  # Pass 2: classify. Hard leak = FactSet-bound traffic (gateway range OR learned gateway IP,
  # OR any FDS-family external) that is NOT on the tunnel. WV2/EXCEL to other public = benign
  # (Office/MS/CDN task-pane assets or TW crawl) -- only FactSet must hide the physical IP.
  foreach ($c in $est) {
    $op = [int]$c.OwningProcess
    $isXl = ($xlPids -contains $op); $isWv2 = $xlWv2.ContainsKey($op); $isFds = $fdsPids.ContainsKey($op)
    if (-not ($isXl -or $isWv2 -or $isFds)) { continue }
    $ra = $c.RemoteAddress; $la = $c.LocalAddress
    if ($ra -eq '127.0.0.1' -and $c.RemotePort -eq 3128) { $proxySeen++; continue }   # WV2->proxy OK
    if ($ra -match '^(127\.|::1|::$|fe80)') { continue }
    $who = if ($isFds) { $fdsPids[$op] } elseif ($isWv2) { 'wv2(xl)' } else { 'EXCEL' }
    $isFactset = $ra.StartsWith($gwPrefix) -or $gwSet.ContainsKey($ra)
    if ($isFactset) {
      if (Test-Tun $la) { $tunSeen++; "[$t] GW-TUNNEL $who $la -> $ra`:$($c.RemotePort) [OK]" | Add-Content $log }
      else { $leakSeen++; "[$t] *** GW-LEAK $who $la -> $ra`:$($c.RemotePort) [NON-TUNNEL LEAK] ***" | Add-Content $log }
      continue
    }
    if ($isFds) {
      if (Test-Tun $la) { $tunSeen++; "[$t] FDS-TUNNEL $who $la -> $ra`:$($c.RemotePort) [OK]" | Add-Content $log }
      else { $leakSeen++; "[$t] *** FDS-DIRECT $who $la -> $ra`:$($c.RemotePort) [NON-TUNNEL LEAK] ***" | Add-Content $log }
      continue
    }
    # WV2/EXCEL to non-FactSet public: benign background (Office/MS/CDN) or TW crawl. Not a leak.
    # (logged sparsely: only note distinct remotes once via $gwSet-style would be noisy; skip)
  }
  Start-Sleep -Milliseconds 500
}
"=== learned gateway IPs: $([string]::Join(', ', @($gwSet.Keys))) ===" | Add-Content $log
"=== capture_a2 summary: tunSeen=$tunSeen proxySeen=$proxySeen leakSeen=$leakSeen ===" | Add-Content $log
"=== capture_a2 exit $(Get-Date -Format 'HH:mm:ss.fff') ===" | Add-Content $log
