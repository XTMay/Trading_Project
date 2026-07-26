# 遠端主機 111.185.192.56 — FactSet IP 防護環境「複製本機最佳化」Runbook

> **本機 = golden reference（黃金基準）**。把本機定案的「OS 環境 + 瀏覽器 + NordVPN + 2 本活頁簿」最佳設定，套到遠端主機 `111.185.192.56`（主機名 `GENE_AI-LAB`），使**兩台操作環境一致**。
>
> **目標一句話**：FactSet 取數流量（含 Add-in 的 WebView2 nativecloud 路徑）**完全經由美國 NordVPN 出口**，真實台灣 IP 永不洩漏給 FactSet；同時 yfinance/TWSE 爬蟲 + 台灣網銀走台灣家用 ISP 直出。
>
> 相關記憶：`factset-260722-vbaguard-vpnproxy`、`factset-nordvpn-splittunnel-killswitch`、`nordvpn-dns-hijack-chrome-doh`、`factset-network-path-vpnproxy`。
> 本機值 = 2026-07-25 實機驗證；遠端值（SAC/主機名）= 2026-07-25 於 RDP 內實測。
> **本版已併入對抗式驗證的必修項 + 遠端 SAC 實測結果。**

---

## Part 0 — 給新 session 的執行說明（先讀）

### 0-A. 執行模型（工具在哪裡跑）
本機的 PowerShell/檔案工具**只作用在本機**，打不到遠端（遠端 WinRM 5985/5986 已實測關閉，只有 RDP 3389 開）。兩條路：
1. **（首選）在遠端主機上直接開一個新的 Claude Code**：透過 RDP 進遠端桌面 → 開 Terminal → `claude`。**前提**：遠端要已裝 Claude Code 執行環境；若沒有，改走路 2。
2. **（備援）在本機用 computer-use 操作螢幕4的 RDP 視窗**：GUI 步驟（NordVPN/瀏覽器）用點擊；PowerShell 要在 RDP 內的 PowerShell 視窗手動輸入。RDP 憑證/連線由使用者提供。
> 本 runbook 內所有 PowerShell 都假設「**在遠端主機上執行**」。

### 0-B. ★RDP 自我斷線安全警語（務必遵守）
整個 NordVPN 設定是在「RDP 連到遠端公網 IP」上做的。以下任一動作會**切斷你自己的 RDP、且 VPN 若重連失敗會把遠端鎖死連不回去**：
- **絕不開「網路 Kill Switch」**（只開「App Kill Switch」）——網路 KS 在 VPN 斷線瞬間會擋掉**所有**流量含 RDP。
- **絕不選反向/inverse 分割通道模式**——會把 RDP 也推進美國隧道→路由不對稱→RDP 立刻斷。
- **分割通道別誤加系統/網路程序**（曾差點誤加 Edge）——誤加會把該程序推上隧道。
- 動 NordVPN 前，**先確保有另一條 out-of-band 救援路徑**（主機商 console / VNC / TeamViewer 等）。

### 0-C. 安全紅線（延續本機政策）
- **`excel.exe` / `python` / 瀏覽器絕不可加進 NordVPN 分割通道**。
- **SAC 開關是使用者決策**，AI 不代按（本機、遠端皆使用者自己關）。**遠端已於 2026-07-26 關閉(=0)**（見 A-5；對 FactSet 無影響，且 0→1 需重設 Windows）。
- **App Kill Switch 保持開；網路 Kill Switch 保持關**。

---

## ★★ 最關鍵差異（新 session 務必最先處理）：proxy 埠不同 → 守衛埠寫死不符
2026-07-26 兩台實測：
- 本機 `GENE-4MONITOR`：FactSetVpnProxy 監聽 **18080**；PAC=`http://127.0.0.1:18080/factset.pac`。
- 遠端 `GENE_AI-LAB`：FactSetVpnProxy 監聽 **8765 + 3128**（**非 18080**；實測 `P18080=0`、`FVP_PORTS=8765,3128`）；PAC=`http://127.0.0.1:8765/proxy.pac`。
- 每台 PAC 各自指向自己的 proxy 埠 → **PAC 本身兩台都對**（FactSet 安裝時各自寫入的埠，兩台值不同）。
- **但 `modVpnGuard.bas` 把埠寫死**：`Private Const PROXY_PORT = "18080"`，且 VPN 閘 `EnsureVpn_US` 的 TU 探測寫死 `-Proxy 'http://127.0.0.1:18080'`。
- **後果**：遠端 proxy 在 8765，守衛卻查 18080→查無→`EnsureProxy_Ready` 回 **PX_CANT**→**每次 A2 取數都彈「Proxy 起不來，已中止取數」並中止**（fail-closed＝不洩漏，但也取不到數）；VPN 閘 TU 也會因 18080 無人而誤判隧道斷→中止。
- **⇒ 2026-07-26 實測確認：遠端真正的「轉發代理埠」= `3128`（不是 8765！）**。探測結果：透過 `127.0.0.1:3128` 打 FactSet 閘道回 **404**（≠502＝隧道通、可用）；透過 `8765` **無回應**（8765 只供 `proxy.pac` 檔，不是轉發代理）。本機 18080 是「PAC 伺服＋轉發」同一埠；**遠端把兩者拆開：8765 供 PAC、3128 才轉發。**
- **⇒ 新 session 必做（跑 Phase 前先修）**：把遠端兩本（FDSUS+TW）守衛的 `Private Const PROXY_PORT` 與 `EnsureVpn_US` 的 TU 探測 `-Proxy 'http://127.0.0.1:____'` **都改成 `3128`**。
  - **★已備好腳本：`_IP_Guard_Handoff\fix_guard_port_remote.ps1`**（純 ASCII；**只在主機名 GENE_AI-LAB 跑**、書若開著唯讀會拒絕、idempotent、把 modVpnGuard 內所有 18080→3128、改完自動 export＋驗 PROXY_PORT/TU/殘留=0）。用法：**先關閉兩本 260722**（COM 需獨佔開）、確認 `AccessVBOM=1`，遠端跑
    ```
    powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710\_IP_Guard_Handoff\fix_guard_port_remote.ps1"
    ```
    印 `FIX_GUARD_PORT_REMOTE_DONE` 後 → 重開一本 → Alt+F11 Debug→Compile（綠）→ A2 實測（守衛應 PASS，不再彈「Proxy 起不來」）。
  - ⚠️ **不要用「讀 AutoConfigURL 埠自適應」**——那會取到 **8765**（PAC 伺服埠），對遠端是**錯的**。真正可攜的自適應要嘛**解析 PAC 內容**抓 `PROXY 127\.0\.0\.1:(\d+)`、要嘛**自動探測** FactSetVpnProxy 各監聽埠挑「打閘道非 502」那個。**本腳本走「直接改 3128」最省事、已驗證。**
  - **狀態 2026-07-26（全部完成 ✅）**：✅ 交接夾已複製到遠端（75 檔）｜✅ Chrome DoH 政策已設（secure+Cloudflare）｜✅ **本埠修已完成**（兩本各改 4 行、`PROXY_PORT="3128"`、TU=3128、殘留 18080=0，已 export 到 `verify_portfix\`）｜✅ **NordVPN 清單回讀**（allowlist；FactSetVpnProxy + 其餘 FDS 在內；瀏覽器/excel/python 皆不在；App KS ON、Network KS OFF）｜✅ **MOTW/信任位置**（兩本 MOTW-clean、`VBAWarnings=1`、資料夾在信任位置 `C:\Github\Trading_Project\`；無需動作）｜✅ **A2 實測零洩漏**（見下）。
  - ✅ **A2 實測結果（2026-07-26）**：兩本開檔均彈「Proxy Passing」；美股 `MSFT`＋台股 `2330` 皆取數成功。細粒度擷取（`capture_a2.ps1`，0.5s）＋ `leak_monitor2.ps1`（2s）雙證：`leakSeen=0`、零 `LEAK/GW-LEAK/FDS-DIRECT`。WebView2 的 FactSet 流量全走 loopback 代理 `127.0.0.1:3128`（`proxySeen` 高），FactSetVpnProxy 對閘道連線一律走隧道（local `100.126.x`）。`SUSPECT`＝WV2 對 `150.171.27.11`/`52.96.*`（微軟/Office 資產）＝良性、非 FactSet。
  - ⚠️ **腳本修正紀錄（2026-07-26）**：`fix_guard_port_remote.ps1` 原本有 PS5.1 剖析錯誤——throw 字串裡的 `"$bn: ..."` 被當成 scope 變數（`$bn:`）→ 整支腳本**從未成功執行過**。已把 4 處 `$bn:` 改為 `${bn}:`，`Parser::ParseFile` 驗證 PARSE OK 後才成功跑完。日後若從本機重抄此腳本，記得該修正。
  - ⚠️ **隧道 IP 與閘道 IP 差異（新 session 必知；已修進 `leak_monitor2.ps1`）**：本機隧道 local IP＝`10.5.*`；**遠端 GENE_AI-LAB 隧道 local IP＝`100.126.x`（NordLynx 100.64.0.0/10 CGNAT 段）**。`leak_monitor2.ps1` 原寫死 `10.5.*` 判隧道→在遠端會把**每一條正確走隧道的 FactSet 連線誤判成 `DIRECT-LEAK!`**。已改為啟動時讀 NordLynx 介面實際 IPv4（新 `Test-Tun`，仍相容 `10.5.*`）。另：遠端實測 FactSet 閘道 IP＝`64.209.89.46` 與 `192.234.235.x`（`.45`/`.121`）——**`64.209.89.46` 不在 runbook 寫死的 `192.234.235.0/24`**，故 `leak_monitor2` 的閘道前綴法會漏抓此 IP，靠其 FDS-family 檢查補上（FactSetVpnProxy 走隧道＝TUNNEL-OK）；`capture_a2.ps1` 則會**從 FactSetVpnProxy 隧道連線動態學習閘道 IP**，較穩健。
  - 驗證改對沒（遠端跑）：`(Invoke-WebRequest 'https://nativecloud-gateway-va.factset.com' -Proxy 'http://127.0.0.1:3128' -TimeoutSec 5 -UseBasicParsing).StatusCode` 或 catch 到的狀態碼＝**非 502** 即通。
- ✅ **已解除（2026-07-26 實測）：遠端 proxy 埠＝「每次啟動固定」，非動態**。證據＝`%LOCALAPPDATA%\FactSetVpnProxy\FactSetVpnProxy.log` 記錄 07-19～07-25 共 **10+ 次真實重啟，每次都綁 `CONNECT proxy 127.0.0.1:3128` + `PAC server 127.0.0.1:8765`**（埠寫死在 binary，`CommandLine` 無埠參數、資料夾無 JSON/ini 埠設定檔）。→ **守衛寫死 `3128`、PAC 指 `8765` 長久安全，重開機/proxy 重啟不會換埠，不需自適應。** 免做手動強制重啟（只會多冒 A-1 洩漏窗風險）。日後只有 FactSet 改版換埠時才需重驗——徵兆＝守衛彈「Proxy 起不來」。查法：`Select-String "$env:LOCALAPPDATA\FactSetVpnProxy\FactSetVpnProxy.log" 'CONNECT proxy listening'` 看最後一行的埠。
- 註：這也解釋了「遠端 Excel 看得到舊資料」與「守衛應中止」的矛盾——那批資料多半是 proxy 還在 18080（或守衛尚未攔到）時取的；proxy 今天 22:38 重啟後若換到 8765，下次 A2 取數守衛就會攔。**新 session 先修埠、再實測一次 A2 才算數。**

---

## Part A — 本機最終最佳化設定（目標狀態，逐項含驗證）

### A-1. FactSetVpnProxy（命門程序）★proxy 必須「一直在跑」，不只是可用性
| 項目 | 值 |
|---|---|
| 路徑 | `%LOCALAPPDATA%\FactSetVpnProxy\FactSetVpnProxy.exe` |
| 簽章 | **NotSigned**（兩台一致） |
| 監聽埠 | `127.0.0.1:18080`（同時對外服務 `factset.pac`） |

**★為何 proxy 存活本身就是「防洩漏」而非只是「能不能取數」**：`factset.pac` 是 **proxy 自己在 18080 提供的**。proxy 一旦沒在跑，AutoConfigURL 連不到→**WinINET fail-open 退回 DIRECT**→凡是走 WinINET/Chromium 打 `*.factset.com` 的都直連洩漏：包括 **FactSet Add-in 開機時 msedgewebview2 的背景 auth/ping**、以及**瀏覽器開 factset.com**。VBA 守衛只擋「那 2 本活頁簿的取數」，擋不到 add-in 背景 WebView2 或瀏覽器。**所以「18080 有在聽」是硬性防洩漏前提，必須在 Excel/FactSet add-in 載入前就成立。**

**驗證**：
```powershell
$e = Join-Path $env:LOCALAPPDATA 'FactSetVpnProxy\FactSetVpnProxy.exe'
(Get-AuthenticodeSignature $e).Status
Get-NetTCPConnection -LocalPort 18080 -State Listen -EA SilentlyContinue |
  ForEach-Object { (Get-Process -Id $_.OwningProcess).ProcessName }   # 預期 FactSetVpnProxy
```

### A-2. WinINET PAC（路由核心）★「環境變數」實際對應這條
本機**沒有任何自訂 proxy 環境變數**（已查 process/HKCU/HKLM）；真正做路由的是這條 PAC。
| 項目 | 值 |
|---|---|
| `AutoConfigURL`（HKCU） | `http://127.0.0.1:18080/factset.pac` |
| `ProxyEnable`（HKCU） | `0`（只用 PAC，不設固定 proxy） |

**★設定順序鐵律**：PAC 指到 18080，若設 PAC 時 proxy 沒在聽、又開了瀏覽器/Excel → Chromium/WebView2 會把 18080 記成「壞代理」快取 ~5 分鐘退回 DIRECT → 洩漏，且守衛只看「埠有沒有聽」會誤報 PASS。**務必：proxy 先起來確認 18080 在聽 → 才寫 PAC → 才開瀏覽器/Excel。**

**設定（HKCU，一般使用者身分；若 FactSet 安裝已設好就跳過）**：
```powershell
# 先確認 proxy 在聽（見 A-1），再寫 PAC
$k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
Set-ItemProperty $k -Name AutoConfigURL -Value 'http://127.0.0.1:18080/factset.pac'
Set-ItemProperty $k -Name ProxyEnable -Value 0
(Get-ItemProperty $k).AutoConfigURL; (Get-ItemProperty $k).ProxyEnable   # 驗：...factset.pac ; 0
```

### A-3. 瀏覽器 DoH（繞過被 NordVPN 挾持的系統 DNS）
NordVPN 全域挾持系統 DNS，對 Google 系網域回假 IP `192.0.0.88`→Chrome/Gmail/reCAPTCHA 壞。修法＝瀏覽器 DoH（走 443，繞過系統 DNS），不動 FactSet/VPN。
| 瀏覽器 | 本機 | 遠端 |
|---|---|---|
| Edge | HKLM 政策 `secure`+Cloudflare | 政策寫（見下） |
| Chrome | UI 設 Cloudflare | **改用政策寫**（更確定） |
| Comet | **無 DoH 引擎→捨棄** | 遠端**已裝**（見 F-8）但預設瀏覽器＝Edge；**不用 Comet 開 Google/NordVPN 登入** |

**★HKLM 需「系統管理員身分」；PAC(A-2) 是 HKCU 一般身分。兩者範圍不同不可混**：非提權 session 寫 HKLM 會 access-denied 靜默失敗；若整個改用「另一個 admin 帳號」的提權 shell 跑，A-2 的 HKCU PAC 會落在 admin 的 hive、不是 Excel 使用者的 hive→factset.com 不被導向→洩漏。**正解：HKLM DoH 用「與 Excel 同一使用者」的提權 PowerShell 寫；HKCU PAC 用一般身分寫。**
```powershell
# 提權 PowerShell（與 Excel 同帳號）
reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v DnsOverHttpsMode /t REG_SZ /d secure /f
reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v DnsOverHttpsTemplates /t REG_SZ /d "https://cloudflare-dns.com/dns-query" /f
reg add "HKLM\SOFTWARE\Policies\Google\Chrome" /v DnsOverHttpsMode /t REG_SZ /d secure /f
reg add "HKLM\SOFTWARE\Policies\Google\Chrome" /v DnsOverHttpsTemplates /t REG_SZ /d "https://cloudflare-dns.com/dns-query" /f
```
重啟 Edge(`edge://restart`)、Chrome。驗：開 `https://www.google.com` 能通。

### A-4. NordVPN（分割通道 + Kill Switch + 自動連線）★GUI-only
> 遠端已裝 NordVPN **且已登入**（使用者 2026-07-25 確認）。以下是要核對/補齊的設定。
| 設定 | 值 |
|---|---|
| 連線類型 | **allowlist（針對指定應用程式使用 VPN）**，絕不反向 |
| **分割通道**清單 | 25 支 FactSet 程序（見清單） |
| **App** Kill Switch | **開**，**且其自己的清單也要放同樣 25 支**（NordVPN 有兩份獨立清單！） |
| **網路** Kill Switch | **關**（RDP 安全，見 0-B） |
| 自動連線 | 使用任何網路時 + **United States → New York** |

**25 支清單**：`FactSetVpnProxy`、`FDSPipe`、`FdsOfcExec`、`FDSConduit`、`FDSRealTime`、`FDSEmsRealTime`、`fdsw32`、`FDSWorkstation_x64`、`FDSBrowser`、`FDSBrowser_x64`、`FDSOCMon`、`FDSTray`、`ApiBox`、`ApiBox_x64`、`FDSDllHost`、`fdswFixExcel`、`FDSDiagnostics`、`FDSDiagnostics_x64`、`FDSChartCopier`、`FDSStatBar`、`MarqueeRestart`、`FDSUpdateDialog`、`fdsup`、`fdsup_x64`、`FDSSdm`。
- 根：`C:\Program Files (x86)\FactSet\`；例外 `FactSetVpnProxy` 在 `%LOCALAPPDATA%\FactSetVpnProxy\`。
- **★命門只有 `FactSetVpnProxy` 一支是必須**；遠端若某支 exe 不存在就跳過那支。
- **加入方法**（NordVPN 勾選清單會卡死）：用「瀏覽應用程式」開檔案對話框、**貼完整路徑逐一加**。先跑這行**列出遠端實際的完整路徑**再逐一貼：
```powershell
$names='FactSetVpnProxy','FDSPipe','FdsOfcExec','FDSConduit','FDSRealTime','FDSEmsRealTime','fdsw32','FDSWorkstation_x64','FDSBrowser','FDSBrowser_x64','FDSOCMon','FDSTray','ApiBox','ApiBox_x64','FDSDllHost','fdswFixExcel','FDSDiagnostics','FDSDiagnostics_x64','FDSChartCopier','FDSStatBar','MarqueeRestart','FDSUpdateDialog','fdsup','fdsup_x64','FDSSdm'
$roots='C:\Program Files (x86)\FactSet',(Join-Path $env:LOCALAPPDATA 'FactSetVpnProxy')
foreach($r in $roots){ if(Test-Path $r){ Get-ChildItem $r -Recurse -Filter *.exe -EA SilentlyContinue } } |
  Where-Object { $_.BaseName -in $names } | Select-Object -ExpandProperty FullName | Sort-Object -Unique
```
- **★回讀驗證（必做）**：分割通道清單 + Kill Switch 清單各開來看，確認 **`FactSetVpnProxy` 在內**、且 **`excel.exe`/`python*`/`chrome`/`msedge`/`msedgewebview2`/瀏覽器一律不在內**；各截一張圖存證。

### A-5. Smart App Control（SAC）★狀態變更：使用者於 2026-07-26 自行關閉(=0)
**⚠️ 現況（2026-07-26）：遠端 SAC = `0`（關閉）——使用者自行關閉。** 對 FactSet 取數／零洩漏**無任何影響**（proxy/隧道/守衛/PAC 與 SAC 無關，SAC=0 下 proxy 更是必然能跑）。
- **⚠️ 單程票**：SAC `1→0` 隨時可關，但 **`0→1` 需重設 Windows** 才能重開。故遠端目前**回不去 SAC=1**（除非重灌）。
- **歷史（2026-07-25 實測，供參）**：當時 `SAC=1（強制）`，proxy 雖 NotSigned 仍於 07-25 22:38 在 SAC=1 下成功啟動 → 證明遠端 SAC 本來就**放行**這支未簽章 proxy（ISG 信譽足夠）。**當時結論是「遠端維持 SAC=1 較佳（防護+功能兼得）」**；惟使用者 07-26 另有考量自行關閉，屬 0-C 的使用者決策，AI 不代按、僅記錄。→ **所以遠端關 SAC 對 FactSet 並非必要（本來就放行），但既已關且無害，維持現狀。**
- 對照本機：本機 07-22 切 SAC 強制時信譽未足→proxy 被擋→才關本機 SAC(=0)。兩台現皆 SAC=0。
- 查 SAC（現應回 `0`）：
```powershell
(Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy' -EA SilentlyContinue).VerifiedAndReputablePolicyState  # 0關/1強制/2評估
```
> ℹ️ 此前 SAC=1 時的顧慮「FactSet 更新 proxy exe→新 hash 一時無信譽被 SAC 擋」**現已不適用**（SAC=0，不再有簽章強制）。留作歷史參考。

### A-6. 系統 DNS（不要動）
NordLynx 網卡 DNS 維持 Nord 原廠 `103.86.96.100/99.100`；Google 網站靠瀏覽器 DoH(A-3)。改系統/網卡 DNS 在此環境不穩，**遠端一律不碰**。

### A-7. 兩本活頁簿（已含守衛，遠端已複製）★但要處理「巨集能不能跑」
- 遠端路徑：`C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710\`（2 本已複製）。
- 2 本已內建 `modVpnGuard`(rev4) + 各批強化。**遠端不需再注入 VBA**，只做環境複製。
- **★關鍵前提：守衛只有在「巨集真的執行」時才保護**。複製過去的 .xlsm 有兩個常見殺手：
  1. **Mark-of-the-Web（MOTW）**：檔案若經 zip/下載/網路磁碟複製會帶「來源不受信任」標記→現代 Excel **硬擋巨集**→`Workbook_Open`→`VpnGuard_Open` 不跑，但 **FactSet add-in 任務窗格仍能取數=正是守衛要擋的 WebView2 洩漏**。
  2. **信任中心**預設「停用所有巨集(含通知)」→守衛靜默、要按「啟用內容」才跑。
  → **處理**：
```powershell
Get-ChildItem 'C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710\*.xlsm' | Unblock-File   # 去 MOTW
```
  並把該資料夾設為 Excel **信任位置**（信任中心→信任位置→新增），或至少開檔時「啟用內容」。
- **★Phase 4 診斷提示**：若開檔**沒彈「Proxy Passing」**，先查是不是巨集被 MOTW/信任中心擋（不是 VPN/proxy 問題）。

---

## Part B — 這幾天怎麼修到定案（濃縮脈絡）
- **B-1 SAC**：本機未簽章 proxy 被強制 SAC 擋→關本機 SAC。**遠端當時 SAC=1 反而放行**；後於 07-26 由使用者自行關閉=0（現況見 A-5）。
- **B-2 WebView2 nativecloud 洩漏**：proxy 沒跑時 WebView2/瀏覽器直連 FactSet 閘道洩漏台灣 IP（FDSPipe 那半仍走美國＝「一半漏」）→守衛取數前把關 + proxy 必須存活(A-1)。
- **B-3 VPN vs 瀏覽器 DNS**：Nord 挾持 DNS 害 Google→瀏覽器 DoH（Comet 無解捨棄）。改系統 DNS 不穩已放棄。
- **B-4 NordVPN**：無 CLI；`nordvpn://connect`（連某美國）；登入走 OS 預設瀏覽器，**Comet 會吃掉 OAuth 回呼**→登入前把預設瀏覽器設 Edge，或把 `nordvpn://login...` 用 Edge/`Start-Process` 開；登入成功但 UI 卡＝重啟 NordVPN.exe(UI，非 service)。
- **B-5 城市標籤**：VPN 機房 IP 的 city geo 不可靠，可靠層是「國家=US」。
- **B-6 開機自動連線縫**：開機初期可能 `Failed to get recommended servers`，App Kill Switch 對「從頭沒連上」不觸發→守衛 rev3 VPN 閘補這縫。

---

## Part C — 簽章偵測 & SAC（回答使用者）
- **每日簽章偵測「已內建」**：`VpnGuard_Open`→`DailySignatureCheck`，每日至多一次（`%TEMP%\vpnguard_sig_YYYYMMDD.flag`），`Get-AuthenticodeSignature` 變 `Valid` 就彈「可恢復 SAC（需重設 Windows）」。遠端同一份程式碼，開檔即自動偵測，無需另裝。
- **注意「簽章 Valid」≠「SAC 放行」是兩件事**：proxy 仍 NotSigned；此前 SAC=1 時因 ISG 信譽放行，**現 SAC=0 更無簽章強制**(A-5)。簽章偵測是等 FactSet 真的出**數位簽章版**（更徹底、跨機一致）。
- **恢復 SAC 是使用者手動決策**，且是重設 Windows 等級動作，巨集只通知不代做。
- **遠端結論（已更新 2026-07-26）**：SAC 原本放行 proxy、當時建議保持開；但使用者已自行關閉(=0)。**FactSet 不受影響**；因 0→1 需重設 Windows，遠端維持現狀 SAC=0、不再動。

---

## Part D — 遠端套用步驟（新 session 照做；已按「防洩漏 + RDP 安全」重排順序）

### Phase 0 — 前提盤點（遠端 PowerShell）
```powershell
"== proxy =="; $e="$env:LOCALAPPDATA\FactSetVpnProxy\FactSetVpnProxy.exe"; Test-Path $e; if(Test-Path $e){(Get-AuthenticodeSignature $e).Status}
"== 18080 =="; Get-NetTCPConnection -LocalPort 18080 -State Listen -EA SilentlyContinue | % { (Get-Process -Id $_.OwningProcess).ProcessName }
"== FactSet add-in loaded =="; Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Excel\Addins','HKCU:\Software\Microsoft\Office\Excel\Addins' -EA SilentlyContinue | Select PSChildName
"== NordVPN =="; Get-Process nordvpn-service,NordVPN -EA SilentlyContinue | Select Name
"== SAC =="; (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy' -EA SilentlyContinue).VerifiedAndReputablePolicyState
"== PAC =="; (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').AutoConfigURL
"== Edge DoH =="; (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -EA SilentlyContinue).DnsOverHttpsMode
"== Chrome DoH =="; (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Google\Chrome' -EA SilentlyContinue).DnsOverHttpsMode
"== 預設瀏覽器 =="; (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice' -EA SilentlyContinue).ProgId
"== 2 本活頁簿 =="; Get-ChildItem 'C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710' -Filter *.xlsm | Select Name
"== 交接夾 + 腳本 =="; $h='C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710\_IP_Guard_Handoff'; Test-Path $h; 'leak_monitor.ps1','leak_monitor2.ps1','defender_harden.ps1' | % { "$_ : $(Test-Path (Join-Path $h $_))" }
"== AccessVBOM =="; (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -EA SilentlyContinue).AccessVBOM
```
**判讀 / 已知狀態**：
- FactSet + NordVPN(已登入)：**使用者已確認裝好**。
- SAC：**已於 2026-07-26 由使用者關閉=0**(A-5)；對 FactSet 無影響，不用動。
- **交接夾若不存在**：`_IP_Guard_Handoff` 內的 `leak_monitor*.ps1`/`defender_harden.ps1`（Part E 要用）可能沒複製過去 → **請使用者把整個 `_IP_Guard_Handoff` 資料夾複製到遠端同路徑**（不是只複製 2 本）。

### Phase 1 — 先把 proxy 拉起來並確認在聽（防洩漏第一關）
- 確認 `FactSetVpnProxy` 在跑、8765+3128 在聽（Phase 0 已看）。遠端 SAC=0（已關），proxy 必能跑，若沒跑可直接：
```powershell
if(-not(Get-Process FactSetVpnProxy -EA SilentlyContinue)){ Start-Process "$env:LOCALAPPDATA\FactSetVpnProxy\FactSetVpnProxy.exe" -WindowStyle Hidden }
Start-Sleep 2; Get-NetTCPConnection -LocalPort 18080 -State Listen -EA SilentlyContinue | % { (Get-Process -Id $_.OwningProcess).ProcessName }
```

### Phase 2 — NordVPN（已登入；核對設定，遵守 0-B RDP 安全）
1. 確認**已登入**（Phase 0）。若需登入：預設瀏覽器設 Edge（別 Comet）或用 Edge 開 `nordvpn://login...`；UI 卡就重啟 NordVPN.exe(UI)。
2. 連線類型 = **allowlist**（非反向）。**先確認「網路 Kill Switch = 關」**(0-B)。
3. 分割通道加 25 支（用 A-4 的探路指令列出完整路徑逐一貼）→ **回讀驗證**（FactSetVpnProxy 在、瀏覽器/excel/python 不在）。
4. **App Kill Switch 的清單也放同樣 25 支**（獨立清單！）。
5. 自動連線 = United States → New York。
6. **手動連到美國、確認 Connected**：
```powershell
$ld=Get-ChildItem (Join-Path $env:LOCALAPPDATA 'NordVPN\logs')|?{$_.Name -match '^app-\d{8}\.log$'}|sort LastWriteTime -desc|select -First 1
(Select-String $ld.FullName -Pattern 'VpnConnectionState change:'|select -Last 1).Line   # 預期 Connected - United States
```
7. **最後才開 App Kill Switch**（此時 VPN 已連，基準正確）。網路 KS 維持關。

### Phase 3 — OS 路由（proxy 已在跑才做）
1. **PAC**（A-2，HKCU 一般身分）：**立即重確認 18080 在聽** → 再寫 PAC。
2. **DoH**（A-3，HKLM 提權、與 Excel 同帳號）：Edge + Chrome 4 條 reg → 重啟瀏覽器。
3. 不碰系統 DNS(A-6)。

### Phase 4 — 活頁簿實測（乾淨開啟，避免壞代理快取）
1. **去 MOTW + 信任位置**(A-7)，確認 `AccessVBOM=1`。
2. **關掉所有 Excel**（`taskkill /IM EXCEL.EXE` 若需要），重確認 proxy+VPN 都在 → **才開活頁簿**（讓 WebView2 對「活的 proxy」重新評估，無壞代理殘留）。
3. 開 `盈再表260722(FDSUS).xlsm` → 應彈「**Proxy Passing**」。**沒彈就先查巨集被 MOTW/信任中心擋**(A-7)，不是 VPN 問題。
4. 美股 A2 輸入 1 檔（如 `AAPL`）→ Enter → 跑 Part E 監視器驗零洩漏。
5. 開 `盈再表260722(TW).xlsm` → 台股 A2 輸入 1 檔（如 `2330`）→ 同樣驗。

---

## Part E — 遠端驗證（做完才算「兩台一致」）
- `leak_monitor.ps1`（美股/閘道式）、`leak_monitor2.ps1`（台股/按進程式）**已改為自定位**（log/STOP 寫在腳本所在資料夾，不再是本機專屬路徑）；閘道改用 `192.234.235.` /24 前綴比對（US 節點解析到鄰近 IP 也抓得到）。背景跑：
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Github\Trading_Project\EZ_table0228\FactSet_Templates_20260710\_IP_Guard_Handoff\leak_monitor2.ps1"
```
- **通過標準**：監視器 log **確實有樣本**（不是空檔）、且美股/台股各至少 1 檔 `DIRECT-LEAK=0`、`SUSPECT=0`；開檔彈「Proxy Passing」；NordVPN log `Connected - United States`。
- 建議再跑一次 Phase 0 全表，逐項對齊 Part A。

---

## Part F — 遠端專屬注意
1. **執行模型**：PowerShell 在遠端跑(0-A)。
2. **SAC**：遠端原可維持開(=1，放行 proxy)，惟使用者已於 2026-07-26 自行關閉(=0)——對 FactSet 無影響；0→1 需重設 Windows(A-5)。兩台現皆 SAC=0。
3. **RDP 自我斷線**：網路 KS 關、不反向、別誤加程序(0-B)。
4. **巨集能不能跑**：MOTW/信任中心是「開檔沒彈窗」的頭號嫌疑(A-7)。
5. **交接夾要整個複製**（不只 2 本），否則 Part E 監視器不在。
6. **proxy 存活是硬性防洩漏前提**（A-1），不只是能不能取數。
7. books 已複製，**不重新注入 VBA**；日後本機守衛升級再同步。
8. **Comet 已安裝但預設仍為 Edge（2026-07-26 實測，安全）**：主程式在 `C:\Program Files\Perplexity\Comet\Application\Comet.exe`（版本 150.0.7871.230；`LocalAppData\Perplexity\Comet\User Data` 僅設定檔，07-25 仍有用），已註冊 HKLM `App Paths\comet.exe`。但 `https`/`http`/`.html` 的 UserChoice **全綁 `MSEdgeHTM`（Edge）**，故系統開連結、NordVPN OAuth 回呼、reCAPTCHA 都走 Edge，Comet 無 DoH 的風險**不會觸發**。
   - ⚠️ **兩條紅線**：(a) **別把預設瀏覽器改成 Comet**（改了→Google/reCAPTCHA 因無 DoH 壞、NordVPN 登入回呼被 Comet 吃）；(b) **別用 Comet 開 NordVPN/Google 登入**。一般網頁用 Comet 無妨（走台灣 ISP 直出，與 FactSet 無關）；不需移除。
   - 查法：`(Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice').ProgId` 應回 `MSEdgeHTM`。

---

## 附錄 — 新 session 啟動提示
> 我要把本機的 FactSet IP 防護環境複製到遠端 111.185.192.56（主機名 GENE_AI-LAB），使兩台一致。請照
> `...\EZ_table0228\FactSet_Templates_20260710\_IP_Guard_Handoff\REMOTE_SETUP_111.185.192.56.md`
> 執行：先跑 Part D Phase 0 前提盤點回報，再依 Phase 1→4 設定，最後 Part E 驗零洩漏。
> 已知：FactSet+NordVPN 已裝已登入；**遠端 SAC 已由使用者關閉=0（2026-07-26），對 FactSet 無影響、不用再動**。
> 安全紅線：excel/python/瀏覽器不加分割通道；網路 Kill Switch 保持關、不選反向模式（會斷我 RDP）；SAC 不動。
