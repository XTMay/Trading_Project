# leak_monitor2.ps1 -- v2 leak monitor for the TW 9-stock test.
# v1 only watched the nativecloud gateway IP (192.234.235.45). The TW fetch
# proved to use OTHER FactSet endpoints (gw=0 during a full 2330 fetch), so v2
# watches by PROCESS as well:
#   (a) gateway connections (kept from v1): local 10.5.* TUNNEL / 127.* LOOPBACK
#       / else DIRECT-LEAK!
#   (b) FDS process family: ANY external conn must have local 10.5.* (NordVPN
#       tunnel via split-tunnel allowlist). local 192.168.* etc = FDS-DIRECT!
#   (c) EXCEL-descendant msedgewebview2: factset traffic must go via PAC ->
#       127.0.0.1:3128 (remote forward port; local box uses 18080). Any direct external conn is logged; remote in
#       192.234.235.0/24 = WV2-GW-LEAK! (hard leak), else WV2-XL-OUT (suspect,
#       logged for review -- task pane may load MS/CDN assets legitimately).
#   (d) EXCEL itself direct external = by design for TW crawls (TWSE/yahoo);
#       only flagged if it reaches the factset gateway range (XL-GW-LEAK!).
# Pure ASCII. Stop via leak_monitor2.STOP in the scratchpad.
# self-locating: logs/STOP go NEXT TO this script (portable to remote; the old hardcoded
# per-user/per-session scratchpad path did not exist on other machines -> Out-File threw).
$sp   = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { Join-Path $env:TEMP 'leakmon' }
New-Item -ItemType Directory -Force -Path $sp | Out-Null
$log  = Join-Path $sp 'leak_monitor2.log'
$stop = Join-Path $sp 'leak_monitor2.STOP'
$gwPrefix = '192.234.235.'
$fdsNames = @('FactSetVpnProxy','FDSPipe','FDSConduit','FdsOfcExec','FDSRealTime','FDSEmsRealTime','fdsw32','FDSWorkstation_x64','FDSTray','ApiBox','ApiBox_x64','FDSBrowser','FDSBrowser_x64')

# --- NordVPN tunnel local-IP classifier (portable across machines) ---
# The local golden box tunnels on 10.5.0.2; the remote GENE_AI-LAB tunnels on 100.126.x
# (NordLynx uses the 100.64.0.0/10 CGNAT range). Hardcoding '10.5.*' would falsely flag
# every correctly-tunneled FactSet connection on the remote as a leak. So read the live
# NordLynx adapter IPv4(s) at startup and treat those (plus the legacy 10.5.* subnet) as
# the tunnel. Test-Tun is the single source of truth for TUNNEL vs LEAK.
$tunIPs = @()
try { $tunIPs = @(Get-NetIPAddress -InterfaceAlias 'NordLynx' -AddressFamily IPv4 -EA SilentlyContinue | Select-Object -ExpandProperty IPAddress) } catch {}
function Test-Tun([string]$la) { if ($la -like '10.5.*') { return $true }; if ($script:tunIPs -contains $la) { return $true }; return $false }

if (Test-Path $stop) { Remove-Item $stop -Force }
"=== leak monitor v2 start $(Get-Date -Format 'HH:mm:ss') ===" | Out-File $log -Encoding utf8
"tunnel local IP(s) [NordLynx]: $([string]::Join(', ', $tunIPs)) (also matching legacy 10.5.*)" | Add-Content $log

for ($k = 0; $k -lt 2700; $k++) {
  if (Test-Path $stop) { "=== STOP $(Get-Date -Format 'HH:mm:ss') ===" | Add-Content $log; break }
  $t = Get-Date -Format 'HH:mm:ss'

  # process snapshot: pid -> (name, ppid)
  $snap = @{}
  foreach ($p in (Get-CimInstance Win32_Process -Property ProcessId,ParentProcessId,Name -EA SilentlyContinue)) {
    $snap[[int]$p.ProcessId] = @{ n = $p.Name; pp = [int]$p.ParentProcessId }
  }
  $xlPids  = @($snap.Keys | Where-Object { $snap[$_].n -eq 'EXCEL.EXE' })
  $fvp     = @($snap.Keys | Where-Object { $snap[$_].n -eq 'FactSetVpnProxy.exe' }).Count
  $wv2All  = @($snap.Keys | Where-Object { $snap[$_].n -eq 'msedgewebview2.exe' })
  # EXCEL-descendant webview2 set (walk ancestors, max 12 hops)
  $xlWv2 = @{}
  foreach ($w in $wv2All) {
    $cur = $w
    for ($h = 0; $h -lt 12; $h++) {
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
  $proxyListen = @(Get-NetTCPConnection -LocalPort 3128 -State Listen -EA SilentlyContinue).Count

  $leak = $false; $suspect = 0
  $detail = @()
  $gwN = 0; $fdsTun = 0; $wv2Proxy = 0
  foreach ($c in $est) {
    $ra = $c.RemoteAddress; $la = $c.LocalAddress
    $op = [int]$c.OwningProcess
    if ($ra -eq '127.0.0.1' -and $c.RemotePort -eq 3128) { $wv2Proxy++; continue }
    if ($ra -match '^(127\.|::1|fe80)') { continue }
    $isGw = $ra.StartsWith($gwPrefix)
    if ($isGw) {
      $gwN++
      $pn = if ($snap.ContainsKey($op)) { $snap[$op].n } else { '?' }
      $cls = if (Test-Tun $la) { 'TUNNEL' } elseif ($la -like '127.*') { 'LOOPBACK' } else { 'DIRECT-LEAK!' }
      if ($cls -eq 'DIRECT-LEAK!') { $leak = $true }
      $detail += ("      GW   {0,-20} {1,-15} -> {2,-21} [{3}]" -f $pn, $la, ($ra + ':' + $c.RemotePort), $cls)
      continue
    }
    if ($fdsPids.ContainsKey($op)) {
      $cls = if (Test-Tun $la) { $fdsTun++; 'TUNNEL-OK' } else { $leak = $true; 'FDS-DIRECT!' }
      if ($cls -eq 'FDS-DIRECT!') {
        $detail += ("      FDS  {0,-20} {1,-15} -> {2,-21} [{3}]" -f $fdsPids[$op], $la, ($ra + ':' + $c.RemotePort), $cls)
      }
      continue
    }
    if ($xlWv2.ContainsKey($op)) {
      $suspect++
      $detail += ("      WV2X msedgewebview2      {0,-15} -> {1,-21} [WV2-XL-OUT]" -f $la, ($ra + ':' + $c.RemotePort))
      continue
    }
    if ($xlPids -contains $op) {
      # EXCEL direct external: fine for TW crawls; only gateway range is a leak (handled above)
      continue
    }
  }
  $verdict = if ($leak) { '*** LEAK ***' } elseif ($suspect) { 'SUSPECT' } elseif ($gwN -or $fdsTun) { 'FDS-ACTIVE-OK' } else { 'idle' }
  "[$t] proxy3128=$proxyListen fvp=$fvp excel=$($xlPids.Count) wv2xl=$($xlWv2.Count) gw=$gwN fdsTun=$fdsTun wv2->proxy=$wv2Proxy => $verdict" | Add-Content $log
  if ($detail.Count) { $detail | Select-Object -First 12 | Add-Content $log }
  Start-Sleep -Seconds 2
}
"=== monitor v2 exit $(Get-Date -Format 'HH:mm:ss') ===" | Add-Content $log
