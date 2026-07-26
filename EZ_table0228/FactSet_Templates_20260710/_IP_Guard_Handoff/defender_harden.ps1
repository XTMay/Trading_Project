# =====================================================================
#  Defender 加固 —— 關閉 SAC 後補回防護缺口
#  請以「系統管理員」開啟 PowerShell 執行。不會關閉任何防護、不影響 FactSet。
#  設計原則:只加「不會打到會呼叫 Win32 API / 生子程序 / 寫 .vbs 的 Office 巨集」的規則,
#           因為盈再表 260722 本身正是這類巨集(modVpnGuard + Module10 都要 shell 出去)。
# =====================================================================
$ErrorActionPreference = 'Continue'
$proxy = "$env:LOCALAPPDATA\FactSetVpnProxy\FactSetVpnProxy.exe"
$excel = 'C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE'

Write-Host '== 0. 基準確認(這幾項本來就該開) ==' -ForegroundColor Cyan
Set-MpPreference -PUAProtection Enabled                       # 潛在垃圾/捆綁軟體
Set-MpPreference -EnableControlledFolderAccess Enabled        # 勒索防護(受控資料夾)
Set-MpPreference -EnableNetworkProtection Enabled             # 封鎖已知惡意網域/IP(補 SmartScreen)
# 竄改防護(Tamper)只能在 Windows 安全性 UI 開/關,已是開啟狀態,無需動作。

Write-Host '== 1. 勒索防護白名單:放行 Excel 與 FactSetVpnProxy(避免受控資料夾誤擋) ==' -ForegroundColor Cyan
Add-MpPreference -ControlledFolderAccessAllowedApplications $excel
Add-MpPreference -ControlledFolderAccessAllowedApplications $proxy

Write-Host '== 2. 「類 SAC」主力:封鎖低信譽/未簽章執行檔 —— 但先把 FactSetVpnProxy 加進例外 ==' -ForegroundColor Cyan
# 這條最接近 SAC(擋不流行、無簽章、來路不明的 exe),差別是「可以設例外」。
# 先加例外再啟用,避免有任何一瞬間 proxy 被擋。
Add-MpPreference -AttackSurfaceReductionOnlyExclusions $proxy
# 建議先用「稽核模式(AuditMode)」跑幾天,看 Defender 事件記錄它「會擋什麼」,確認不誤傷你日常工具,再改成 Enabled。
Add-MpPreference -AttackSurfaceReductionRules_Ids '01443614-cd74-433a-b99e-2ecdc07bfc25' -AttackSurfaceReductionRules_Actions AuditMode
# 觀察無誤後,把上一行的 AuditMode 改成 Enabled 重跑一次即可正式啟用。

Write-Host '== 3. 高價值、且不會打到本活頁簿的 ASR 規則(直接啟用) ==' -ForegroundColor Cyan
$safe = @(
  '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2',  # 封鎖從 LSASS 竊取憑證
  'e6db77e5-3df2-4cf1-b95a-636979351e5b',  # 封鎖經 WMI 事件訂閱做持久化
  '56a863a9-875e-4185-98a7-b882c64b5ce5',  # 封鎖濫用已知漏洞的簽章驅動
  'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4',  # 封鎖 USB 來的未簽章/不信任程序
  'd3e037e1-3eb8-44c8-a917-57927947596d'   # 封鎖 JS/VBScript 啟動「下載來的」執行檔(注意:不影響你本機自產的 .vbs)
)
foreach ($id in $safe) { Add-MpPreference -AttackSurfaceReductionRules_Ids $id -AttackSurfaceReductionRules_Actions Enabled }

Write-Host '== 4. 驗證 ==' -ForegroundColor Cyan
$p = Get-MpPreference
Write-Host "  PUA=$($p.PUAProtection)  受控資料夾=$($p.EnableControlledFolderAccess)  網路防護=$($p.EnableNetworkProtection)"
Write-Host '  已設定的 ASR 規則:'
for ($i=0; $i -lt $p.AttackSurfaceReductionRules_Ids.Count; $i++) {
  '    {0} = {1}' -f $p.AttackSurfaceReductionRules_Ids[$i], $p.AttackSurfaceReductionRules_Actions[$i]
}
Write-Host '  ASR 例外:'; $p.AttackSurfaceReductionOnlyExclusions | ForEach-Object { "    $_" }

# =====================================================================
#  ❌ 絕對不要開的 ASR 規則(會直接搞死盈再表 260722 / FactSet Add-in):
#    d4f940ab-401b-4efc-aadc-ad5f3c50688a  封鎖 Office 建立子程序
#         → 守衛要生 powershell/cmd/wscript;Module10 也要;開了守衛與取數全掛。
#    92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b  封鎖 Office 巨集呼叫 Win32 API
#         → modVpnGuard 的 Declare GetCurrentProcessId 會被擋;整個守衛失效。
#    3b576869-a4ec-4529-8536-b80a7769e899  封鎖 Office 建立可執行內容
#         → 守衛/Module10 要寫 .vbs 到 TEMP;開了會被擋。
#    5beb7efe-fd9a-4556-801d-275e5ffc04cc  封鎖「可能混淆」的腳本
#         → 本活頁簿跑自產 .vbs,易誤判;不要開。
# =====================================================================
