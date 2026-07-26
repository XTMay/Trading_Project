Attribute VB_Name = "modVpnGuard"
Option Explicit

'============================================================================
' modVpnGuard  --  FactSet VPN Proxy + US-tunnel guard  (rev4, 2026-07-24, for 260722)
'
' rev4: all guard popups now show TOP-CENTER of the screen for 6 seconds
' (mshta HTA window; WScript.Popup could not be positioned). Logic unchanged.
'
' rev3 adds a SECOND gate after the proxy gate: the NordVPN tunnel must be UP and
' exiting the US before any FactSet fetch. It probes the FactSet gateway through the
' proxy (502 = tunnel down) and reads the NordVPN log for "Connected - United States".
' If down, it fires nordvpn://connect (auto-connect, follows the app's location, e.g.
' New York) and polls; if it still cannot reach a US tunnel it fronts NordVPN and
' aborts the fetch. This closes the gap where the boot auto-connect failed and FactSet
' would otherwise run with no VPN (the App Kill Switch does NOT catch a never-connected
' state). A live US session the user picked is preserved (connect is a no-op when up).
'
' Why: FactSetVpnProxy.exe keeps the FactSet Add-in WebView2 (msedgewebview2,
' nativecloud) traffic on the Houston tunnel. If port 127.0.0.1:3128 is not
' listening, WebView2 falls back to a DIRECT connection = Taiwan real IP leaked
' to FactSet gateway 192.234.235.45. This guard confirms the port is up BEFORE
' any FactSet fetch; if not, it (re)starts the proxy and restarts Excel once for
' a fresh WebView2, or aborts the fetch so no leak can occur.
'
' rev2 changes vs rev1:
'   * REMOVED case-4 (post-fetch leak detection + force-quit) per user request;
'     the pre-fetch port gate below is the sole protection.
'   * Port check uses Get-NetTCPConnection -State Listen (~138ms table lookup),
'     NOT TcpClient.Connect (which took ~2s per refused attempt on this box and
'     froze Excel ~69s -> Esc-interrupt). The blocking PS wait now also runs with
'     Application.EnableCancelKey = xlDisabled so a stray Esc cannot break it.
'   * Added a once-per-day check: if FactSetVpnProxy.exe becomes digitally
'     SIGNED, notify the user they may restore Smart App Control.
'
' NOTE: (re)starting the proxy requires Smart App Control to be OFF (SAC blocks
' the unsigned exe from any launch). With SAC on, a down proxy yields PX_CANT ->
' fetch aborted (no leak), never a restart loop.
'
' Public entry points (wired into Macro2 / ThisWorkbook.Workbook_Open):
'   VpnGuard_Open      - Workbook_Open gate (+ daily signature check)
'   VpnGuard_Start()   - Macro2 start gate; returns False => caller Exit Sub
'
' All source is ASCII.
'============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private Const PROXY_PORT  As String = "3128"
Private Const POPUP_TITLE As String = "FactSet VPN Guard"
'rev4 (2026-07-24): popups are top-center HTA windows (mshta), 6s auto-close
'(was WScript.Popup screen-center 3s). RestartExcelFresh watchdog delay follows
'POPUP_SECS so the kill still fires right as the popup closes.
Private Const POPUP_SECS  As Long = 6

Private Const PX_PASS    As Long = 0   'port already listening (clean)
Private Const PX_RESTART As Long = 1   'was down, (re)started + port now listening -> fresh WebView2
Private Const PX_CANT    As Long = 2   'not listening / cannot start / inconclusive -> abort, NO loop

Private Const VPN_OK   As Long = 0   'VPN connected to US + tunnel actually passes traffic to gateway
Private Const VPN_FAIL As Long = 1   'VPN down / non-US / cannot auto-reconnect -> abort (no leak)

'============================ PUBLIC ENTRY POINTS ============================

'Very start of Macro2 (before FDS fetch). True => Macro2 continues; False => Exit Sub.
Public Function VpnGuard_Start() As Boolean
    On Error Resume Next
    VpnGuard_Start = ProxyGate()
End Function

'B6 (2026-07-24) TW cross-book guard: mirror of the workbook's own
'FDSUS_EnsureUSSheetActive, but for the TaiwanStock sheet. Macro1's legacy body is
'ActiveSheet/ActiveWorkbook-bound; with two 260722 books open, a DoEvents/Wait during
'a fetch lets the user click the other book, and unqualified writes then hit the wrong
'book. This forces focus back to ThisWorkbook's TaiwanStock sheet. Returns False if it
'cannot (caller should Exit rather than write). Sheet name via ChrW keeps the .bas ASCII.
Public Function TW_EnsureTaiwanSheetActive() As Boolean
    Dim tws As Worksheet
    On Error Resume Next
    Set tws = ThisWorkbook.Worksheets(ChrW(&H53F0) & ChrW(&H80A1)) 'TaiwanStock
    If tws Is Nothing Then Exit Function
    If Not ActiveWorkbook Is ThisWorkbook Then ThisWorkbook.Activate
    If Not ActiveSheet Is tws Then tws.Activate
    TW_EnsureTaiwanSheetActive = (ActiveSheet Is tws)
    On Error GoTo 0
End Function

'Start of ThisWorkbook.Workbook_Open: once-a-day signature check, then the gate.
Public Sub VpnGuard_Open()
    On Error Resume Next
    DailySignatureCheck
    ProxyGate
End Sub

'=============================== CORE GATE =================================

'Returns True only when the proxy port is confirmed listening.
Private Function ProxyGate() As Boolean
    Dim st As Long
    st = EnsureProxy_Ready()
    Select Case st
        Case PX_PASS
            'proxy is up -> now the second gate: VPN tunnel must be up and exiting the US.
            If EnsureVpn_US() = VPN_OK Then
                ClearRestartMarker
                ShowPopupAsync "Proxy Passing", 64
                ProxyGate = True
            Else
                'VPN down/non-US and could not auto-connect to US -> bring NordVPN to front
                'and abort the fetch (never fetch without a US tunnel = never leak Taiwan IP).
                OpenNordVpnWindow
                '"VPN wei lian shang mei guo, yi zhong zhi qu shu (wu IP bao hu)" (ChrW = pure ASCII .bas)
                ShowPopupAsync "VPN " & ChrW(26410) & ChrW(36899) & ChrW(19978) & ChrW(32654) & ChrW(22283) & "," & ChrW(24050) & ChrW(20013) & ChrW(27490) & ChrW(21462) & ChrW(25976) & "(" & ChrW(28961) & "IP" & ChrW(20445) & ChrW(35703) & ")", 48
                ProxyGate = False
            End If
        Case PX_RESTART
            If RestartRecently() Then
                ShowPopupAsync "FactSetVpnProxy keeps failing - restart loop aborted. Fix the VPN, then retry.", 16
                ProxyGate = False
            ElseIf SaveAllForRestart() Then
                'every open workbook is safely saved -> OK to force-restart the Excel process
                WriteRestartMarker
                ShowPopupAsync "Proxy Failed & Need to Re-start", 48
                RestartExcelFresh                                   'kill this Excel tree + reopen (fresh WebView2)
                ProxyGate = False
            Else
                'a workbook could not be saved (read-only / never-saved / locked) -> do NOT
                'force-kill (would lose that data); abort the fetch instead.
                ShowPopupAsync "Proxy down but an open workbook could not be saved - fetch aborted (no restart).", 16
                ProxyGate = False
            End If
        Case Else   'PX_CANT / inconclusive (proxy cannot be revived - e.g. SAC still blocking, VPN down, exe gone)
            '"Proxy qi bu lai, yi zhong zhi qu shu (wu IP bao hu)" -- ChrW keeps this .bas pure ASCII
            ShowPopupAsync "Proxy " & ChrW(36215) & ChrW(19981) & ChrW(20358) & "," & ChrW(24050) & ChrW(20013) & ChrW(27490) & ChrW(21462) & ChrW(25976) & "(" & ChrW(28961) & "IP" & ChrW(20445) & ChrW(35703) & ")", 16
            ProxyGate = False
    End Select
End Function

'One PowerShell call: fast Listen-table check on 127.0.0.1:3128; if down, start
'the proxy and poll the PORT (not the process) up to ~5s. Emits PASS/RESTART/CANT.
Private Function EnsureProxy_Ready() As Long
    On Error GoTo unknown
    Dim tmp As String, ps As String, cmd As String, r As String, dq As String
    dq = Chr$(34)
    tmp = Environ$("TEMP") & "\vpnguard_proxy_" & StampStr() & ".txt"

    'try/catch on Start-Process: if SAC blocks the launch it throws -> immediate CANT
    '(no 20-iter poll), so the "cannot start" freeze is ~1s not ~11s.
    ps = "$p=" & PROXY_PORT & ";" & _
         "function L{$c=Get-NetTCPConnection -LocalPort $p -State Listen -EA SilentlyContinue;if($c){$pr=Get-Process -Id (@($c.OwningProcess)[0]) -EA SilentlyContinue;if($pr -and $pr.ProcessName -eq 'FactSetVpnProxy'){'Y'}else{'N'}}else{'N'}};" & _
         "if((L) -eq 'Y'){'PASS';exit};" & _
         "$e=Join-Path $env:LOCALAPPDATA 'FactSetVpnProxy\FactSetVpnProxy.exe';" & _
         "if(-not(Test-Path $e)){'CANT';exit};" & _
         "if(-not(Get-Process FactSetVpnProxy -EA SilentlyContinue)){try{Start-Process $e -WindowStyle Hidden -EA Stop}catch{'CANT';exit}};" & _
         "for($i=0;$i -lt 20;$i++){Start-Sleep -Milliseconds 250;if((L) -eq 'Y'){'RESTART';exit}};'CANT'"

    cmd = "cmd /c powershell -NoProfile -ExecutionPolicy Bypass -Command " & dq & ps & dq & " > " & dq & tmp & dq & " 2>nul"
    RunHiddenWait cmd

    r = UCase$(Trim$(ReadAnsi(tmp)))
    On Error Resume Next
    Kill tmp
    On Error GoTo unknown

    If InStr(r, "PASS") > 0 Then EnsureProxy_Ready = PX_PASS: Exit Function
    If InStr(r, "RESTART") > 0 Then EnsureProxy_Ready = PX_RESTART: Exit Function
    EnsureProxy_Ready = PX_CANT     'CANT or empty/inconclusive -> safe abort (no false PASS, no loop)
    Exit Function
unknown:
    EnsureProxy_Ready = PX_CANT
End Function

'============================ VPN (US tunnel) GATE =========================

'One PowerShell call: is NordVPN connected to a US server AND does the tunnel
'actually carry traffic to the FactSet gateway (probe via the proxy: 502 = tunnel
'down, any other response = up)? If disconnected, fire nordvpn://connect (which
'follows the app's auto-connect location) and poll up to ~20s for a US tunnel.
'Emits US_OK / NOTUS / FAIL. (nordvpn://connect is a no-op when already connected,
'so a live US session the user picked, e.g. New York, is preserved, not overridden.)
Private Function EnsureVpn_US() As Long
    On Error GoTo fail
    Dim tmp As String, ps As String, cmd As String, r As String, dq As String
    dq = Chr$(34)
    tmp = Environ$("TEMP") & "\vpnguard_vpn_" & StampStr() & ".txt"

    'TU catch uses [int](...) so a response-LESS WebException (timeout/reset) -> $c=0 -> tunnel
    'DOWN (fail-closed), not $null -> wrongly UP. If the log already says US, probe TU ONCE and
    'decide (no poll) -- nordvpn://connect is a no-op while "connected" so a stale-US+dead-tunnel
    'can never self-heal by looping. The reconnect poll carries a 25s deadline. TU timeout = 3s.
    'NOTE: TU proves the tunnel reaches the gateway; the US-exit fact comes from the NordVPN log.
    'B13 fix: match ONLY app-YYYYMMDD.log (exclude app-norddrop-/app-errors-/app-modules-, which
    'never carry 'VpnConnectionState change:'); re-fetch the newest log on EVERY LC call so the
    'reconnect poll and a cross-midnight new log are seen (old code cached $ld once = stale filter).
    ps = "function GD{Get-ChildItem (Join-Path $env:LOCALAPPDATA 'NordVPN\logs') -EA SilentlyContinue|Where-Object{$_.Name -match '^app-\d{8}\.log$'}|Sort-Object LastWriteTime -Descending|Select-Object -First 1};" & _
         "function LC{$ld=GD;if($ld){(Select-String -Path $ld.FullName -Pattern 'VpnConnectionState change:' -EA SilentlyContinue|Select-Object -Last 1).Line}else{''}};" & _
         "function TU{try{Invoke-WebRequest 'https://nativecloud-gateway-va.factset.com' -Proxy 'http://127.0.0.1:3128' -TimeoutSec 3 -UseBasicParsing -EA Stop|Out-Null;$true}catch{$c=0;try{$c=[int]($_.Exception.Response.StatusCode.value__)}catch{};($c -ne 0 -and $c -ne 502)}};" & _
         "$l=LC;if($l -match 'change: Connected - United States'){if(TU){'US_OK'}else{'NOTUS'};exit};" & _
         "Start-Process 'nordvpn://connect' -EA SilentlyContinue;" & _
         "$sw=[Diagnostics.Stopwatch]::StartNew();" & _
         "for($i=0;$i -lt 8 -and $sw.Elapsed.TotalSeconds -lt 25;$i++){Start-Sleep -Seconds 2;$l=LC;if(($l -match 'change: Connected - United States') -and (TU)){'US_OK';exit}};" & _
         "'FAIL'"

    cmd = "cmd /c powershell -NoProfile -ExecutionPolicy Bypass -Command " & dq & ps & dq & " > " & dq & tmp & dq & " 2>nul"
    RunHiddenWait cmd

    r = UCase$(Trim$(ReadAnsi(tmp)))
    On Error Resume Next
    Kill tmp
    On Error GoTo fail

    If InStr(r, "US_OK") > 0 Then EnsureVpn_US = VPN_OK Else EnsureVpn_US = VPN_FAIL
    Exit Function
fail:
    EnsureVpn_US = VPN_FAIL
End Function

'Bring the NordVPN app to the front so the user can pick/confirm a US server
'(e.g. New York, nearest to FactSet HQ in Norwalk CT). Fire-and-forget.
Private Sub OpenNordVpnWindow()
    On Error Resume Next
    CreateObject("WScript.Shell").Run """C:\Program Files\NordVPN\NordVPN.exe""", 1, False
End Sub

'====================== ONCE-A-DAY SIGNATURE NOTICE =======================

'If FactSetVpnProxy.exe has become digitally signed, tell the user they may
'restore Smart App Control. Runs at most once per calendar day.
Private Sub DailySignatureCheck()
    On Error Resume Next
    Dim flag As String, ff As Integer
    flag = Environ$("TEMP") & "\vpnguard_sig_" & Format$(Now, "yyyymmdd") & ".flag"
    If Len(Dir$(flag)) > 0 Then Exit Sub       'already checked today
    ff = FreeFile
    Open flag For Output As #ff                'stamp first so an error won't re-trigger
    Print #ff, "1"
    Close #ff
    If InStr(1, UCase$(SignatureStatus()), "VALID") > 0 Then
        ShowPopupAsync "FactSetVpnProxy is now digitally SIGNED. You may restore Smart App Control (requires a Windows reset).", 64
    End If
End Sub

Private Function SignatureStatus() As String
    On Error GoTo done
    Dim tmp As String, ps As String, cmd As String, dq As String
    dq = Chr$(34)
    tmp = Environ$("TEMP") & "\vpnguard_sigst_" & StampStr() & ".txt"
    ps = "$e=Join-Path $env:LOCALAPPDATA 'FactSetVpnProxy\FactSetVpnProxy.exe';" & _
         "if(Test-Path $e){(Get-AuthenticodeSignature $e).Status}else{'NoFile'}"
    cmd = "cmd /c powershell -NoProfile -ExecutionPolicy Bypass -Command " & dq & ps & dq & " > " & dq & tmp & dq & " 2>nul"
    RunHiddenWait cmd
    SignatureStatus = Trim$(ReadAnsi(tmp))
    On Error Resume Next
    Kill tmp
done:
End Function

'=========================== RESTART (kill/reopen) =========================

'Save every dirty open workbook BEFORE the force-restart so the taskkill /F cannot
'destroy unsaved data. Returns True only if nothing dirty remains (safe to kill).
'A never-saved (path-less) or unsaveable workbook -> False -> caller aborts instead.
Private Function SaveAllForRestart() As Boolean
    On Error Resume Next
    Dim wb As Object
    SaveAllForRestart = True
    For Each wb In Application.Workbooks
        If Not wb.Saved Then
            If Len(wb.path) = 0 Then
                SaveAllForRestart = False        'never saved: cannot save silently -> do not kill
            Else
                wb.Save
                If Not wb.Saved Then SaveAllForRestart = False   'save failed (read-only/locked) -> do not kill
            End If
        End If
    Next wb
End Function

'Spawn a DETACHED wscript that, after POPUP_SECS, taskkill /F /T /PID kills THIS
'Excel process tree (proxy-poisoned WebView2 dies with it), waits for the PID to
'vanish, then re-opens the workbook (fresh WebView2 re-evaluates the live proxy).
'Detached via "start" so the kill cannot hit the watchdog itself.
Private Sub RestartExcelFresh()
    On Error Resume Next
    Dim stamp As String, paramTxt As String, vbsPath As String
    Dim exePath As String, wbPath As String, pid As Long, ff As Integer, dq As String
    dq = Chr$(34)
    stamp = StampStr()
    paramTxt = Environ$("TEMP") & "\vpnguard_wd_" & stamp & ".txt"
    vbsPath = Environ$("TEMP") & "\vpnguard_wd_" & stamp & ".vbs"
    exePath = Application.path & "\EXCEL.EXE"
    wbPath = ThisWorkbook.FullName
    pid = GetCurrentProcessId()

    WriteUtf16 paramTxt, exePath & vbCrLf & wbPath & vbCrLf & CStr(pid)

    ff = FreeFile
    Open vbsPath For Output As #ff
    Print #ff, "q = Chr(34)"
    Print #ff, "Set fso = CreateObject(" & dq & "Scripting.FileSystemObject" & dq & ")"
    Print #ff, "Set f = fso.OpenTextFile(" & dq & paramTxt & dq & ", 1, False, -1)"
    Print #ff, "exe = Trim(f.ReadLine)"
    Print #ff, "wbp = Trim(f.ReadLine)"
    Print #ff, "pid = Trim(f.ReadLine)"
    Print #ff, "f.Close"
    Print #ff, "WScript.Sleep " & (POPUP_SECS * 1000)
    Print #ff, "CreateObject(" & dq & "WScript.Shell" & dq & ").Run " & dq & "taskkill /F /T /PID " & dq & " & pid, 0, True"
    Print #ff, "Set svc = GetObject(" & dq & "winmgmts:\\.\root\cimv2" & dq & ")"
    Print #ff, "tries = 0"
    Print #ff, "Do"
    Print #ff, "  If svc.ExecQuery(" & dq & "SELECT ProcessId FROM Win32_Process WHERE ProcessId=" & dq & " & pid).Count = 0 Then Exit Do"
    Print #ff, "  WScript.Sleep 500"
    Print #ff, "  tries = tries + 1"
    Print #ff, "Loop While tries < 120"
    'clear Excel crash-resiliency so the forced-kill relaunch is clean: no Safe-Mode
    'prompt (StartupItems) and the FactSet add-in is not pushed to Disabled Items.
    Print #ff, "CreateObject(" & dq & "WScript.Shell" & dq & ").Run " & dq & "reg delete HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems /f" & dq & ", 0, True"
    Print #ff, "CreateObject(" & dq & "WScript.Shell" & dq & ").Run " & dq & "reg delete HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\StartupItems /f" & dq & ", 0, True"
    Print #ff, "CreateObject(" & dq & "WScript.Shell" & dq & ").Run q & exe & q & " & dq & " " & dq & " & q & wbp & q, 1, False"
    Print #ff, "On Error Resume Next"
    Print #ff, "fso.DeleteFile " & dq & paramTxt & dq
    Print #ff, "fso.DeleteFile WScript.ScriptFullName"
    Close #ff

    'launch DETACHED so taskkill /T on THIS Excel (3s later) cannot kill the watchdog
    CreateObject("WScript.Shell").Run "cmd /c start " & dq & dq & " wscript.exe //B " & dq & vbsPath & dq, 0, False
End Sub

'=========================== ANTI-LOOP MARKER =============================

Private Function MarkerPath() As String
    MarkerPath = Environ$("TEMP") & "\vpnguard_restart.flag"
End Function

Private Function RestartRecently() As Boolean
    On Error GoTo no
    If Len(Dir$(MarkerPath())) = 0 Then Exit Function
    'window must exceed a worst-case cold reopen (FactSet add-in load can be 15-40s+)
    'so a genuine reopen never looks "not recent" and re-triggers a restart.
    RestartRecently = (DateDiff("s", FileDateTime(MarkerPath()), Now) < 240)
    Exit Function
no:
End Function

Private Sub WriteRestartMarker()
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open MarkerPath() For Output As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #ff
End Sub

Private Sub ClearRestartMarker()
    On Error Resume Next
    If Len(Dir$(MarkerPath())) > 0 Then Kill MarkerPath()
End Sub

'============================== HELPERS ====================================

'Run a hidden shell command and WAIT, with the Esc/Ctrl-Break cancel key disabled
'so the (bounded) PS wait cannot raise "code execution interrupted".
Private Sub RunHiddenWait(ByVal cmd As String)
    On Error Resume Next
    Dim prev As Long
    prev = Application.EnableCancelKey
    Application.EnableCancelKey = 0          'xlDisabled
    CreateObject("WScript.Shell").Run cmd, 0, True
    Application.EnableCancelKey = prev
End Sub

'Guard popup: write msg (may be Chinese) to a UTF-16 file, then hand off to the
'shared top-center renderer with POPUP_SECS (6s).
Private Sub ShowPopupAsync(ByVal msg As String, ByVal icon As Long)
    On Error Resume Next
    Dim msgPath As String
    msgPath = Environ$("TEMP") & "\vpnguard_msg_" & StampStr() & ".txt"
    WriteUtf16 msgPath, msg
    VpnGuard_PopupFile msgPath, POPUP_SECS, icon
End Sub

'PUBLIC shared renderer (rev4). Shows the UTF-16 message file `msgPath` at the
'TOP-CENTER of the screen for `secs` seconds, then deletes the message file.
'Used by the guard AND by Module10.CheckExcelVPNIP (its ShowFailPopup calls this
'so every A2-entry popup is top-center; the caller owns the secs value).
'Mechanism: an ASCII .hta (an mshta window we CAN position - WScript.Popup could
'not be moved) reads the UTF-16 file and self-closes; an ASCII .vbs runs mshta
'then deletes msgPath + the hta + itself. Fire-and-forget (Excel never waits).
'Passing a Chinese String in memory is fine; only the .bas SOURCE stays ASCII.
'Color strip by icon: 16=red, 48=orange, else blue.
Public Sub VpnGuard_PopupFile(ByVal msgPath As String, ByVal secs As Long, ByVal icon As Long)
    On Error Resume Next
    Dim stamp As String, htaPath As String, vbsPath As String, ff As Integer, dq As String, bg As String
    Dim wnd As Object, ax As Long, ay As Long, haveAnchor As Boolean, moveJs As String
    dq = Chr$(34)
    stamp = StampStr()
    htaPath = Environ$("TEMP") & "\vpnguard_popup_" & stamp & ".hta"
    vbsPath = Environ$("TEMP") & "\vpnguard_popup_" & stamp & ".vbs"

    Select Case icon
        Case 16: bg = "#c62828"     'error: red strip
        Case 48: bg = "#ef6c00"     'warning: orange strip
        Case Else: bg = "#1565c0"   'info: blue strip
    End Select

    'Anchor at the FRONT-CENTER-TOP of the Excel window in SCREEN pixels (via
    'PointsToScreenPixelsX/Y) so the popup lands on whatever monitor Excel is on.
    'Size 320x85 = 1/4 area of the old 640x170 (half each dimension). Fallback =
    'top-center of the primary screen if there is no active window.
    err.Clear
    Set wnd = Application.ActiveWindow
    If Not wnd Is Nothing Then
        ax = wnd.PointsToScreenPixelsX(wnd.UsableWidth / 2)
        ay = wnd.PointsToScreenPixelsY(0)
        If err.Number = 0 Then haveAnchor = True
        err.Clear
    End If
    If haveAnchor Then
        moveJs = "window.moveTo " & CStr(ax - 160) & ", " & CStr(ay + 24)
    Else
        moveJs = "window.moveTo CInt((screen.width - 320) / 2), 8"
    End If

    'ASCII .hta: caption OFF (so the title can be centered in-body), sized 320x85,
    'positioned at the Excel window, reads the UTF-16 message, self-closes.
    ff = FreeFile
    Open htaPath For Output As #ff
    Print #ff, "<html><head><title>" & POPUP_TITLE & "</title>"
    Print #ff, "<hta:application caption=" & dq & "no" & dq & " border=" & dq & "thin" & dq & " sysmenu=" & dq & "no" & dq & " showintaskbar=" & dq & "no" & dq & " scroll=" & dq & "no" & dq & " contextmenu=" & dq & "no" & dq & " selection=" & dq & "no" & dq & "/>"
    Print #ff, "<script language=" & dq & "VBScript" & dq & ">"
    Print #ff, "Sub Window_OnLoad"
    Print #ff, "On Error Resume Next"
    Print #ff, "window.resizeTo 320, 85"
    Print #ff, moveJs
    Print #ff, "Set fso = CreateObject(" & dq & "Scripting.FileSystemObject" & dq & ")"
    Print #ff, "Set f = fso.OpenTextFile(" & dq & msgPath & dq & ", 1, False, -1)"
    'NB: read into 'vtxt' and target id 'vpnbody' - an element with id 'm' would be
    'exposed by HTA as a global named 'm', colliding with a 'm' variable so innerText
    'ends up the DOM object -> "[object]". Distinct names avoid that.
    Print #ff, "vtxt = f.ReadAll : f.Close"
    Print #ff, "document.getElementById(" & dq & "vpnbody" & dq & ").innerText = vtxt"
    Print #ff, "window.setTimeout " & dq & "window.close" & dq & ", " & (secs * 1000)
    Print #ff, "End Sub"
    Print #ff, "</script></head>"
    Print #ff, "<body style=" & dq & "margin:0;background:#ffffff;border:1px solid #888888;border-top:5px solid " & bg & ";font-family:Microsoft JhengHei,Segoe UI,sans-serif;overflow:hidden;" & dq & ">"
    Print #ff, "<div style=" & dq & "text-align:center;font-size:8pt;color:#777777;padding:2px 0 0 0;" & dq & ">" & POPUP_TITLE & "</div>"
    Print #ff, "<div id=" & dq & "vpnbody" & dq & " style=" & dq & "text-align:center;padding:2px 8px;font-size:10pt;font-weight:600;color:#111111;" & dq & "></div>"
    Print #ff, "</body></html>"
    Close #ff

    'ASCII .vbs runner: mshta (wait) -> delete temps -> self-delete
    ff = FreeFile
    Open vbsPath For Output As #ff
    Print #ff, "Set sh = CreateObject(" & dq & "WScript.Shell" & dq & ")"
    Print #ff, "sh.Run " & dq & "mshta.exe " & dq & dq & htaPath & dq & dq & dq & ", 1, True"
    Print #ff, "Set fso = CreateObject(" & dq & "Scripting.FileSystemObject" & dq & ")"
    Print #ff, "On Error Resume Next"
    Print #ff, "fso.DeleteFile " & dq & msgPath & dq
    Print #ff, "fso.DeleteFile " & dq & htaPath & dq
    Print #ff, "fso.DeleteFile WScript.ScriptFullName"
    Close #ff

    CreateObject("WScript.Shell").Run "wscript.exe " & dq & vbsPath & dq, 0, False
End Sub

'Write text as UTF-16 LE + BOM (a .vbs OpenTextFile(...,-1) reads it, Chinese-safe).
Private Sub WriteUtf16(ByVal path As String, ByVal text As String)
    On Error Resume Next
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "unicode"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2
    stm.Close
End Sub

'Read a small ANSI file.
Private Function ReadAnsi(ByVal path As String) As String
    On Error GoTo done
    Dim ff As Integer
    If Len(Dir$(path)) = 0 Then Exit Function
    ff = FreeFile
    Open path For Input As #ff
    If LOF(ff) > 0 Then ReadAnsi = Input$(LOF(ff), ff)
    Close #ff
done:
End Function

'Unique-ish temp stamp (ASCII only).
Private Function StampStr() As String
    Randomize
    StampStr = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd * 100000))
End Function
