# fix_guard_port_remote.ps1 -- REMOTE-ONLY guard proxy-port fix (18080 -> 3128).
#
# WHY: the remote GENE_AI-LAB runs FactSetVpnProxy with the *forward* proxy on port
# 3128 (port 8765 only serves proxy.pac). The 2 books were copied from the LOCAL box
# where FactSetVpnProxy is on 18080, so their modVpnGuard hardcodes 18080 in two places:
#   (1) Private Const PROXY_PORT As String = "18080"   (used by EnsureProxy_Ready's
#       Get-NetTCPConnection -LocalPort listen check)
#   (2) the EnsureVpn_US TU probe: Invoke-WebRequest ... -Proxy 'http://127.0.0.1:18080'
# On the remote those target a dead port, so EnsureProxy_Ready returns PX_CANT and EVERY
# A2 fetch aborts with "Proxy 起不来" (fail-closed = no leak, but no data). This rewrites
# every "18080" -> "3128" inside the modVpnGuard module of BOTH remote books.
#
# Proven remote fact (2026-07-26): probe to 127.0.0.1:3128 -> FactSet gateway = 404
# (!=502 => tunnel up, works); to :8765 = no response (PAC file only). So 3128 is right.
#
# SAFE-BY-DESIGN:
#   * Refuses unless $env:COMPUTERNAME = GENE_AI-LAB (never touches the LOCAL 18080 books).
#   * Opens each book in an ISOLATED Excel; aborts if a book opens READ-ONLY (still open
#     elsewhere) so it can never fight the user's running Excel.
#   * Idempotent (already-3128 books report "changed 0").
#   * Pure ASCII (Chinese filename via [char] codes) so PS5.1 ANSI reading can't corrupt it.
#
# PRECONDITIONS (new session): CLOSE both 盈再表260722 books first; AccessVBOM=1
#   (HKCU\...\Excel\Security\AccessVBOM). AFTER: reopen a book, Alt+F11 Debug->Compile
#   (expect green), then A2 test (should now pass the guard instead of "Proxy 起不来").
$ErrorActionPreference = 'Stop'
if ($env:COMPUTERNAME -ne 'GENE_AI-LAB') {
  Write-Host "REFUSING: remote-only fix (18080->3128). Host=$env:COMPUTERNAME, expected GENE_AI-LAB."
  Write-Host "  (The LOCAL books must stay on 18080. If the remote hostname changed, edit this guard.)"
  exit 1
}
$dir = 'C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710'
$ho  = Join-Path $dir '_IP_Guard_Handoff'
$OLD = '18080'; $NEW = '3128'
function Get-Comp($wb,$name){ for($t=1;$t -le 6;$t++){ foreach($c in @($wb.VBProject.VBComponents)){ if($c.Name -eq $name){return $c} }; Start-Sleep -Milliseconds 200 }; throw "comp $name not found" }

# 盈再表 = U+76C8 U+518D U+8868  (ASCII source, no Chinese literals)
$books = @(
  ([char]0x76C8 + [char]0x518D + [char]0x8868 + '260722(FDSUS).xlsm'),
  ([char]0x76C8 + [char]0x518D + [char]0x8868 + '260722(TW).xlsm')
)

foreach ($bn in $books) {
  $file = Join-Path $dir $bn
  if (-not (Test-Path -LiteralPath $file)) { throw "not found: $file" }
  $before = @(Get-Process EXCEL -EA SilentlyContinue | Select-Object -ExpandProperty Id)
  $xl = $null; $wb = $null; $myPid = $null
  try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AskToUpdateLinks = $false
    try { $xl.AutomationSecurity = 3 } catch {}
    try { $xl.Calculation = -4135 } catch {}
    $after = @(Get-Process EXCEL -EA SilentlyContinue | Select-Object -ExpandProperty Id)
    $myPid = ($after | ? { $_ -notin $before }) | Select-Object -First 1
    Write-Host "[$bn] isolated Excel PID $myPid"
    $wb = $xl.Workbooks.Open($file, 0, $false)
    if ($wb.ReadOnly) { throw "$bn opened READ-ONLY -- it is still open in another Excel. CLOSE it first, then rerun." }

    $cm = $null
    try { $cm = (Get-Comp $wb 'modVpnGuard').CodeModule } catch { throw "${bn}: modVpnGuard module not found ($_). AccessVBOM=1?" }

    $changed = 0
    for ($i = 1; $i -le $cm.CountOfLines; $i++) {
      $ln = $cm.Lines($i, 1)
      if ($ln -match $OLD) { $cm.ReplaceLine($i, ($ln -replace $OLD, $NEW)); $changed++ }
    }

    # verify: no 18080 left; PROXY_PORT const now 3128; TU probe now 3128
    $rem = 0; for ($i = 1; $i -le $cm.CountOfLines; $i++) { if ($cm.Lines($i, 1) -match $OLD) { $rem++ } }
    $hasConst = $false; $hasTU = $false
    for ($i = 1; $i -le $cm.CountOfLines; $i++) {
      $ln = $cm.Lines($i, 1)
      if ($ln -match 'Private Const PROXY_PORT.*"3128"') { $hasConst = $true }
      if ($ln -match '127\.0\.0\.1:3128') { $hasTU = $true }
    }
    if ($rem -ne 0)     { throw "${bn}: still $rem line(s) contain $OLD after edit" }
    if (-not $hasConst) { throw "${bn}: PROXY_PORT const is not 3128 after edit" }
    if (-not $hasTU)    { throw "${bn}: TU probe (127.0.0.1:3128) not present after edit" }
    Write-Host "  [$bn] changed $changed line(s); PROXY_PORT=3128 OK; TU=3128 OK; remaining-18080=0"

    $wb.Save()
    $exp = Join-Path $ho 'verify_portfix'; New-Item -ItemType Directory -Force $exp | Out-Null
    $tag = ($bn -replace '[^0-9A-Za-z]', '_')
    (Get-Comp $wb 'modVpnGuard').Export((Join-Path $exp ("modVpnGuard_$tag.bas")))
    $wb.Close($false); $wb = $null; $xl.Quit(); $xl = $null
    Write-Host "  [$bn] saved + exported to verify_portfix\modVpnGuard_$tag.bas"
  }
  finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    try { if ($xl) { $xl.Quit() } } catch {}
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 400
    if ($myPid) { $s = Get-Process -Id $myPid -EA SilentlyContinue; if ($s) { Stop-Process -Id $myPid -Force -EA SilentlyContinue } }
  }
}
Write-Host 'FIX_GUARD_PORT_REMOTE_DONE -- both books 18080->3128. Now: reopen a book, Alt+F11 Debug->Compile (green), then A2 test (guard should PASS, not "Proxy 起不来").'
