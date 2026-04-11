"""
yfinance API 常用功能示例
安装: pip install yfinance
"""

import yfinance as yf

# ============================================================
# 1. 获取单只股票信息
# ============================================================
ticker = yf.Ticker("AAPL")

# 基本信息（公司名称、行业、市值等）
info = ticker.info
print("=== 基本信息 ===")
print(f"公司名称: {info.get('shortName')}")
print(f"行业:     {info.get('industry')}")
print(f"市值:     {info.get('marketCap')}")
print(f"PE 比率:  {info.get('trailingPE')}")
print(f"股息率:   {info.get('dividendYield')}")
print()

# ============================================================
# 2. 获取历史价格数据
# ============================================================
# period: 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max
# interval: 1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo
print("=== 最近 1 个月日线数据 ===")
hist = ticker.history(period="1mo", interval="1d")
print(hist.head(10))
print()

# 指定日期范围
print("=== 指定日期范围 ===")
hist_range = ticker.history(start="2025-01-01", end="2025-06-30")
print(hist_range.head())
print()

# ============================================================
# 3. 财务报表
# ============================================================
print("=== 损益表 (年度) ===")
print(ticker.financials.head())
print()

print("=== 资产负债表 (年度) ===")
print(ticker.balance_sheet.head())
print()

print("=== 现金流量表 (年度) ===")
print(ticker.cashflow.head())
print()

# 季度报表
# ticker.quarterly_financials
# ticker.quarterly_balance_sheet
# ticker.quarterly_cashflow

# ============================================================
# 4. 股息与拆股历史
# ============================================================
print("=== 股息历史 ===")
print(ticker.dividends.tail(10))
print()

print("=== 拆股历史 ===")
print(ticker.splits)
print()

# ============================================================
# 5. 机构持仓 & 分析师建议
# ============================================================
print("=== 主要持有人 ===")
print(ticker.major_holders)
print()

print("=== 机构持有人 (前5) ===")
print(ticker.institutional_holders.head() if ticker.institutional_holders is not None else "N/A")
print()

print("=== 分析师建议 ===")
print(ticker.recommendations.tail(5) if ticker.recommendations is not None else "N/A")
print()

# ============================================================
# 6. 批量下载多只股票数据
# ============================================================
print("=== 批量下载 ===")
data = yf.download(
    tickers=["AAPL", "MSFT", "GOOGL"],
    period="5d",
    interval="1d",
    group_by="ticker",  # 按股票分组
)
print(data)
print()

# ============================================================
# 7. 台湾 / 港股 / A股 代码格式
# ============================================================
# 台股: "2330.TW" (台积电)
# 港股: "0700.HK" (腾讯)
# A股 (不直接支持，但可尝试): 用 .SS(上海) 或 .SZ(深圳)
#   例如 "600519.SS" (贵州茅台), "000001.SZ" (平安银行)

print("=== 台积电 (2330.TW) ===")
tsmc = yf.Ticker("2330.TW")
print(tsmc.history(period="5d"))
print()

# ============================================================
# 8. 期权数据
# ============================================================
print("=== 期权到期日 ===")
expirations = ticker.options
print(expirations[:5] if len(expirations) > 5 else expirations)

if expirations:
    # 获取第一个到期日的期权链
    opt = ticker.option_chain(expirations[0])
    print(f"\n=== Call 期权 ({expirations[0]}) ===")
    print(opt.calls.head())
    print(f"\n=== Put 期权 ({expirations[0]}) ===")
    print(opt.puts.head())
