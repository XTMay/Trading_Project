' ============================================================
' VBA 示例: 调用 yfinance_excel.py 的各个功能
' ============================================================
' 使用方式:
'   1. 在 Excel 中按 Alt+F11 打开 VBA 编辑器
'   2. 插入 -> 模块，将此代码粘贴进去
'   3. 修改下方 PYTHON_PATH 和 SCRIPT_PATH 为你的实际路径
'   4. 在 Excel 中按 Alt+F8 运行对应的 Sub
' ============================================================

' *** 请修改以下两个路径为你的实际路径 ***
Const PYTHON_PATH As String = "/Users/xiaotingzhou/Downloads/Trading Project/venv/bin/python"
Const SCRIPT_PATH As String = "/Users/xiaotingzhou/Downloads/Trading Project/yfinance_excel.py"
Const OUTPUT_FILE As String = "/Users/xiaotingzhou/Downloads/Trading Project/output.xlsx"

' ------------------------------------------------------------
' 通用调用函数
' ------------------------------------------------------------
Private Sub RunYFinance(cmdArgs As String)
    Dim cmd As String
    cmd = """" & PYTHON_PATH & """ """ & SCRIPT_PATH & """ " & cmdArgs

    ' macOS 用 MacScript 或 AppleScript; Windows 用 Shell
    #If Mac Then
        Dim script As String
        script = "do shell script """ & Replace(cmd, """", "\""") & """"
        MacScript (script)
    #Else
        Shell "cmd /c " & cmd, vbHide
    #End If

    MsgBox "完成! 请打开 " & OUTPUT_FILE & " 查看结果。", vbInformation
End Sub

' ------------------------------------------------------------
' 功能 1: 获取股票基本信息
' ------------------------------------------------------------
Sub GetStockInfo()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "基本信息", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "info " & symbol & " """ & OUTPUT_FILE & """ 基本信息 A1"
End Sub

' ------------------------------------------------------------
' 功能 2: 获取历史价格
' ------------------------------------------------------------
Sub GetHistory()
    Dim symbol As String, period As String
    symbol = InputBox("请输入股票代码:", "历史价格", "AAPL")
    If symbol = "" Then Exit Sub
    period = InputBox("请输入时间范围 (1d/5d/1mo/3mo/6mo/1y/2y/5y/max):", "历史价格", "1mo")
    If period = "" Then period = "1mo"
    RunYFinance "history " & symbol & " """ & OUTPUT_FILE & """ 历史价格 A1 --period " & period
End Sub

' ------------------------------------------------------------
' 功能 3: 获取财务报表 (损益表)
' ------------------------------------------------------------
Sub GetFinancials_Income()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "损益表", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "financials " & symbol & " """ & OUTPUT_FILE & """ 损益表 A1 --report income"
End Sub

' ------------------------------------------------------------
' 功能 3b: 获取财务报表 (资产负债表)
' ------------------------------------------------------------
Sub GetFinancials_Balance()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "资产负债表", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "financials " & symbol & " """ & OUTPUT_FILE & """ 资产负债表 A1 --report balance"
End Sub

' ------------------------------------------------------------
' 功能 3c: 获取财务报表 (现金流量表)
' ------------------------------------------------------------
Sub GetFinancials_Cashflow()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "现金流量表", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "financials " & symbol & " """ & OUTPUT_FILE & """ 现金流量表 A1 --report cashflow"
End Sub

' ------------------------------------------------------------
' 功能 4: 获取股息与拆股
' ------------------------------------------------------------
Sub GetDividends()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "股息", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "dividends " & symbol & " """ & OUTPUT_FILE & """ 股息 A1 --rows 20"
End Sub

' ------------------------------------------------------------
' 功能 5: 获取持有人信息
' ------------------------------------------------------------
Sub GetHolders()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "持有人", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "holders " & symbol & " """ & OUTPUT_FILE & """ 持有人 A1"
End Sub

' ------------------------------------------------------------
' 功能 6: 获取分析师建议
' ------------------------------------------------------------
Sub GetRecommendations()
    Dim symbol As String
    symbol = InputBox("请输入股票代码:", "分析师建议", "AAPL")
    If symbol = "" Then Exit Sub
    RunYFinance "recommend " & symbol & " """ & OUTPUT_FILE & """ 分析师建议 A1"
End Sub

' ------------------------------------------------------------
' 功能 7: 批量下载多只股票
' ------------------------------------------------------------
Sub DownloadMultiple()
    Dim symbols As String
    symbols = InputBox("请输入股票代码 (逗号分隔):", "批量下载", "AAPL,MSFT,GOOGL")
    If symbols = "" Then Exit Sub
    RunYFinance "download " & symbols & " """ & OUTPUT_FILE & """ 批量下载 A1 --period 1mo"
End Sub

' ------------------------------------------------------------
' 功能 8: 获取期权数据
' ------------------------------------------------------------
Sub GetOptions()
    Dim symbol As String, optType As String
    symbol = InputBox("请输入股票代码:", "期权", "AAPL")
    If symbol = "" Then Exit Sub
    optType = InputBox("请输入期权类型 (calls/puts):", "期权", "calls")
    If optType = "" Then optType = "calls"
    RunYFinance "options " & symbol & " """ & OUTPUT_FILE & """ 期权 A1 --type " & optType
End Sub
