# FMP API 替代 yfinance — 完整迁移分析

> 目标：用 Financial Modeling Prep (FMP) API 完全替代 yfinance，生成一模一样的 `report_summary.xlsx`
>
> Base URL: `https://financialmodelingprep.com/api`
>
> 所有请求需带 `apikey=YOUR_API_KEY` 参数

---

## 一、当前 yfinance 调用总览

`generate_report_summary.py` 共有 **10 个数据抓取步骤**，使用以下 yfinance API：

| 步骤 | 函数名 | yfinance 调用 | 数据用途 | Excel 目标区域 |
|------|--------|--------------|---------|---------------|
| 1/10 | `fetch_quarterly_income` | `ticker.quarterly_financials` | 季度损益表 (5Q + TTM) | AE (col 31) |
| 2/10 | `fetch_quarterly_balance_sheet` | `ticker.quarterly_balance_sheet` | 季度资产负债表 (5Q + TTM) | AN (col 40) |
| 3/10 | `fetch_annual_income` | `ticker.financials` | 年度损益表 (5Y + TTM) | AW (col 49) |
| 4/10 | `fetch_annual_balance_sheet` | `ticker.balance_sheet` | 年度资产负债表 (5Y + TTM) | BF (col 58) |
| 5/10 | `fetch_annual_cashflow` | `ticker.cashflow` + `ticker.quarterly_cashflow` | 年度现金流 (5Y + TTM) | CE (col 83) |
| 6/10 | `fetch_company_profile` | `ticker.info` | 公司概况 | A1, Y16, Y23 等 |
| 7/10 | `fetch_share_capital` | `ticker.info` | 股本/持股结构 | K12, CN 区域 |
| 8/10 | `fetch_eps_earnings` | `ticker.income_stmt` + `ticker.quarterly_income_stmt` + `ticker.info` | EPS 数据 | K3, CN 区域 |
| 9/10 | `fetch_historical_prices_adj` | `ticker.history(auto_adjust=False)` + `ticker.actions` | 历史股价 + 公司行动 | BO (col 67) |
| 10/10 | `fetch_exchange_rate` | `yf.Ticker("USDJPY=X").history()` | 外汇汇率 | F15, CA (col 79) |

---

## 二、逐步 FMP API 替代方案

---

### 1/10 季度损益表 — Quarterly Income Statement

**yfinance 调用**：
```python
ticker.quarterly_financials          # 返回 DataFrame，行=科目，列=季度日期
```

**FMP 替代**：
```
GET /v3/income-statement/{symbol}?period=quarter&limit=5&apikey=KEY
```

**对照表**：

| yfinance 字段 | FMP 字段 | 说明 |
|--------------|----------|------|
| Total Revenue | `revenue` | 营收 |
| Cost Of Revenue | `costOfRevenue` | 营业成本 |
| Gross Profit | `grossProfit` | 毛利 |
| Operating Income | `operatingIncome` | 营业利润 |
| Net Income | `netIncome` | 净利润 |
| Net Income Common Stockholders | `netIncomeDeducted` 或 `netIncome` | 归属普通股东净利 |
| EBITDA | `ebitda` | EBITDA |
| EBIT | `operatingIncome` | 近似值 |
| Basic EPS | `eps` | 基本每股收益 |
| Diluted EPS | `epsdiluted` | 稀释每股收益 |
| Interest Expense | `interestExpense` | 利息费用 |
| Income Tax Expense | `incomeTaxExpense` | 所得税费用 |
| Research And Development | `researchAndDevelopmentExpenses` | 研发费用 |
| Selling General And Administrative | `sellingGeneralAndAdministrativeExpenses` | 销管费用 |

**TTM 计算**：取返回的 4 个最近季度求和（与 yfinance 逻辑一致）

**Unusual Items 提取**：FMP 返回数据中无直接 "unusual" 标签，需从以下字段识别：
- `otherExpenses` — 其他费用（可能包含非经常性）
- 或使用 `/v3/income-statement-as-reported/{symbol}` 获取 SEC 原始报表，搜索 restructuring 等关键词

**注意**：季度数据需付费账户

---

### 2/10 季度资产负债表 — Quarterly Balance Sheet

**yfinance 调用**：
```python
ticker.quarterly_balance_sheet       # DataFrame，行=科目，列=季度日期
```

**FMP 替代**：
```
GET /v3/balance-sheet-statement/{symbol}?period=quarter&limit=5&apikey=KEY
```

**对照表**：

| yfinance 字段 | FMP 字段 | 说明 |
|--------------|----------|------|
| Total Assets | `totalAssets` | 资产总额 |
| Current Assets | `totalCurrentAssets` | 流动资产 |
| Total Liabilities Net Minority Interest | `totalLiabilities` | 负债总额 |
| Current Liabilities | `totalCurrentLiabilities` | 流动负债 |
| Stockholders Equity | `totalStockholdersEquity` | 股东权益 |
| Cash And Cash Equivalents | `cashAndCashEquivalents` | 现金 |
| Short Term Investments | `shortTermInvestments` | 短期投资 |
| Net Receivables | `netReceivables` | 应收账款净额 |
| Inventory | `inventory` | 存货 |
| Long Term Debt | `longTermDebt` | 长期负债 |
| Common Stock | `commonStock` | 普通股 |
| Retained Earnings | `retainedEarnings` | 保留盈余 |
| Goodwill | `goodwill` | 商誉 |
| Intangible Assets | `intangibleAssets` | 无形资产 |

**TTM 计算**：BS 为时点数据 → TTM = 最新一季快照（与 yfinance 逻辑一致）

---

### 3/10 年度损益表 — Annual Income Statement

**yfinance 调用**：
```python
ticker.financials                    # DataFrame，年度数据
```

**FMP 替代**：
```
GET /v3/income-statement/{symbol}?limit=5&apikey=KEY
```

（不加 `period=quarter` 即为年度，默认行为）

**字段映射**：同 1/10 季度损益表

**TTM 计算**：取最近年度数据（与 yfinance 逻辑一致）

---

### 4/10 年度资产负债表 — Annual Balance Sheet

**yfinance 调用**：
```python
ticker.balance_sheet                 # DataFrame，年度数据
```

**FMP 替代**：
```
GET /v3/balance-sheet-statement/{symbol}?limit=5&apikey=KEY
```

**字段映射**：同 2/10 季度资产负债表

---

### 5/10 年度现金流量表 — Annual Cash Flow Statement

**yfinance 调用**：
```python
ticker.cashflow                      # 年度 CFS
ticker.quarterly_cashflow            # 季度 CFS（用于计算 TTM）
```

**FMP 替代**：
```
# 年度
GET /v3/cash-flow-statement/{symbol}?limit=5&apikey=KEY

# 季度（用于 TTM 计算）
GET /v3/cash-flow-statement/{symbol}?period=quarter&limit=4&apikey=KEY
```

**对照表**：

| yfinance 字段 | FMP 字段 | 说明 |
|--------------|----------|------|
| Operating Cash Flow | `operatingCashFlow` | 经营活动现金流 |
| Capital Expenditure | `capitalExpenditure` | 资本支出 |
| Free Cash Flow | `freeCashFlow` | 自由现金流 |
| Investing Cash Flow | `netCashUsedForInvestingActivites` | 投资活动现金流 |
| Financing Cash Flow | `netCashUsedProvidedByFinancingActivities` | 筹资活动现金流 |
| Depreciation And Amortization | `depreciationAndAmortization` | 折旧摊销 |
| Stock Based Compensation | `stockBasedCompensation` | 股票薪酬 |
| Change In Working Capital | `changeInWorkingCapital` | 营运资金变动 |
| Dividends Paid | `dividendsPaid` | 股利支付 |
| Common Stock Repurchased | `commonStockRepurchased` | 股票回购 |

**TTM 计算**：取最近 4 季求和（与 yfinance 逻辑一致）

---

### 6/10 公司概况 — Company Profile

**yfinance 调用**：
```python
ticker.info                          # dict，包含所有公司信息
```

**FMP 替代**：
```
# 主要 Profile
GET /v3/profile/{symbol}?apikey=KEY

# 补充：实时报价（PE、EPS、52 周高低等）
GET /v3/quote/{symbol}?apikey=KEY

# 补充：关键指标 TTM
GET /v3/key-metrics-ttm/{symbol}?apikey=KEY
```

**对照表**：

| yfinance 字段 | FMP 端点 | FMP 字段 | 说明 |
|--------------|---------|----------|------|
| `longName` | profile | `companyName` | 公司全称 |
| `symbol` | profile | `symbol` | 股票代号 |
| `exchange` | profile | `exchangeShortName` | 交易所 |
| `sector` | profile | `sector` | 行业大类 |
| `industry` | profile | `industry` | 细分行业 |
| `country` | profile | `country` | 国家 |
| `currency` | profile | `currency` | 交易币别 |
| `financialCurrency` | — | 需从财报 `reportedCurrency` 取 | 财报计价币别 |
| `currentPrice` | profile | `price` | 当前股价 |
| `marketCap` | profile | `mktCap` | 市值 |
| `enterpriseValue` | key-metrics-ttm | `enterpriseValueTTM` | 企业价值 |
| `trailingPE` | quote | `pe` | 市盈率 (TTM) |
| `forwardPE` | key-metrics-ttm | `peRatioTTM`（近似） | 前瞻市盈率 |
| `trailingEps` | quote | `eps` | EPS (TTM) |
| `dividendYield` | profile | `lastDiv` / 自行计算 | 股息率 |
| `payoutRatio` | key-metrics-ttm | `payoutRatioTTM` | 派息率 |
| `beta` | profile | `beta` | Beta 值 |
| `fiftyTwoWeekHigh` | quote | `yearHigh` | 52 周最高 |
| `fiftyTwoWeekLow` | quote | `yearLow` | 52 周最低 |
| `longBusinessSummary` | profile | `description` | 公司描述 |
| `website` | profile | `website` | 公司网站 |
| `totalRevenue` | key-metrics-ttm | 从 IS TTM 取 `revenue` | 营收 (TTM) |
| `netIncomeToCommon` | key-metrics-ttm | 从 IS TTM 取 `netIncome` | 净利 (TTM) |

**注意**：yfinance 的 `ticker.info` 是一次返回所有数据，FMP 需要组合 2~3 个端点才能覆盖全部字段。

---

### 7/10 股本/持股结构 — Share Capital

**yfinance 调用**：
```python
ticker.info  # 提取 sharesOutstanding, floatShares, heldPercentInsiders 等
```

**FMP 替代**：
```
# 流通股 / 浮动股
GET /v4/shares_float?symbol={symbol}&apikey=KEY

# 机构持股明细
GET /v3/institutional-holder/{symbol}?apikey=KEY

# 内部人交易统计
GET /v4/insider-trading?symbol={symbol}&apikey=KEY

# 补充：实时报价中的 sharesOutstanding
GET /v3/quote/{symbol}?apikey=KEY
```

**对照表**：

| yfinance 字段 | FMP 端点 | FMP 字段 | 说明 |
|--------------|---------|----------|------|
| `sharesOutstanding` | quote | `sharesOutstanding` | 流通股数 |
| `floatShares` | shares_float | `floatShares` | 浮动股数 |
| `impliedSharesOutstanding` | — | 无直接对应 | 隐含流通股 |
| `heldPercentInsiders` | — | 需从 insider-trading 汇总计算 | 内部人持股% |
| `heldPercentInstitutions` | — | 需从 institutional-holder 汇总计算 | 机构持股% |
| `sharesShort` | — | **FMP 无此数据** | 做空股数 |
| `sharesShortPriorMonth` | — | **FMP 无此数据** | 上月做空股数 |
| `shortPercentOfFloat` | — | **FMP 无此数据** | 做空占流通股% |

**缺失项**：FMP **不提供 Short Interest（做空）数据**，如需此数据需保留 yfinance 或使用其他数据源（如 FINRA）。

---

### 8/10 EPS / 盈余数据 — EPS Earnings

**yfinance 调用**：
```python
ticker.income_stmt                   # 年度 IS → 提取 Net Income
ticker.quarterly_income_stmt         # 季度 IS → 提取 Net Income
ticker.info["trailingEps"]           # TTM EPS
```

**FMP 替代**：
```
# 年度 Net Income（从年度 IS 取）
GET /v3/income-statement/{symbol}?limit=5&apikey=KEY
→ 取每年的 netIncome 字段

# 季度 Net Income（从季度 IS 取）
GET /v3/income-statement/{symbol}?period=quarter&limit=8&apikey=KEY
→ 取每季的 netIncome 字段

# TTM EPS
GET /v3/quote/{symbol}?apikey=KEY
→ 取 eps 字段

# 或用 Key Metrics TTM
GET /v3/key-metrics-ttm/{symbol}?apikey=KEY
→ 取 netIncomePerShareTTM 字段
```

**FMP 额外可用端点**：
```
# 历史 EPS（含预估 vs 实际）
GET /v3/historical/earning_calendar/{symbol}?apikey=KEY
```

---

### 9/10 历史股价（调整价）+ 公司行动 — Historical Prices (adj) + Actions

**yfinance 调用**：
```python
# 日线数据（含 Adj Close）
ticker.history(start=start_date, end=end_date, auto_adjust=False)

# 公司行动（股息 + 拆股）
ticker.actions
```

**FMP 替代**：
```
# 历史日线价格（含 adjClose）
GET /v3/historical-price-full/{symbol}?from=2014-01-01&to=2026-03-08&apikey=KEY

# 历史股息
GET /v3/historical-price-full/stock_dividend/{symbol}?apikey=KEY

# 历史拆股
GET /v3/historical-price-full/stock_split/{symbol}?apikey=KEY
```

**价格数据对照表**：

| yfinance 字段 | FMP 字段 | 说明 |
|--------------|----------|------|
| `Open` | `open` | 开盘价 |
| `High` | `high` | 最高价 |
| `Low` | `low` | 最低价 |
| `Close` | `close` | 收盘价（未调整） |
| `Adj Close` | `adjClose` | 调整后收盘价 |
| `Volume` | `volume` | 成交量 |

**股息数据对照表**：

| yfinance 字段 | FMP 字段 | 说明 |
|--------------|----------|------|
| `Dividends` | `dividend` | 每股股息金额 |
| — | `adjDividend` | 调整后股息 |
| — | `recordDate` | 登记日 |
| — | `paymentDate` | 付息日 |
| — | `declarationDate` | 宣布日 |

**拆股数据对照表**：

| yfinance 字段 | FMP 字段 | 说明 |
|--------------|----------|------|
| `Stock Splits` | `numerator / denominator` | 拆股比例 |
| — | `date` | 拆股日期 |

**月频重采样**：FMP 无月频端点 → 需在 Python 端抓取日线后用 `resample('MS')` 聚合（与当前逻辑一致）

**历史数据范围限制**：FMP 单次请求最长 5 年 → 2014-2026 需分 **3 次请求**：
```
2014-01-01 ~ 2018-12-31
2019-01-01 ~ 2023-12-31
2024-01-01 ~ 2026-03-08
```

---

### 10/10 外汇汇率 — Foreign Exchange Rate

**yfinance 调用**：
```python
yf.Ticker("USDJPY=X").history(period="5d")
```

**FMP 替代**：
```
# 实时汇率
GET /v3/fx/USDJPY?apikey=KEY

# 历史汇率（取最近几天）
GET /v3/historical-price-full/forex/USDJPY?from=2026-03-03&to=2026-03-08&apikey=KEY

# 或直接取报价
GET /v3/quote/USDJPY?apikey=KEY
```

**对照表**：

| yfinance 字段 | FMP 字段（quote 端点） | 说明 |
|--------------|---------------------|------|
| `Close` | `price` | 收盘汇率 |
| `Open` | `open` | 开盘汇率 |
| `High` | `dayHigh` | 日内最高 |
| `Low` | `dayLow` | 日内最低 |

**注意**：FMP 的外汇对符号格式为 `USDJPY`（不带 `=X` 后缀）。

---

## 三、FMP API 调用次数汇总

生成一份完整 `report_summary.xlsx` 所需的 FMP API 请求：

| 数据类别 | API 请求数 | 端点 |
|---------|-----------|------|
| 季度 IS (5Q) | 1 | `/v3/income-statement?period=quarter&limit=5` |
| 季度 BS (5Q) | 1 | `/v3/balance-sheet-statement?period=quarter&limit=5` |
| 年度 IS (5Y) | 1 | `/v3/income-statement?limit=5` |
| 年度 BS (5Y) | 1 | `/v3/balance-sheet-statement?limit=5` |
| 年度 CFS (5Y) | 1 | `/v3/cash-flow-statement?limit=5` |
| 季度 CFS (TTM 用) | 1 | `/v3/cash-flow-statement?period=quarter&limit=4` |
| 公司 Profile | 1 | `/v3/profile` |
| 实时报价 | 1 | `/v3/quote` |
| Key Metrics TTM | 1 | `/v3/key-metrics-ttm` |
| 股本浮动股 | 1 | `/v4/shares_float` |
| 历史日线价格 | 3 | `/v3/historical-price-full`（分 3 次，每次 ≤5 年） |
| 历史股息 | 1 | `/v3/historical-price-full/stock_dividend` |
| 历史拆股 | 1 | `/v3/historical-price-full/stock_split` |
| 外汇汇率 | 1 | `/v3/quote/USDJPY`（或 `/v3/fx`） |
| **合计** | **~16 次** | |

对比 yfinance：yfinance 在底层会发出更多 HTTP 请求（每个属性访问都会触发网络调用），实际请求数 > 16 次。FMP 的请求次数相当。

---

## 四、FMP vs yfinance 关键差异

| 对比维度 | yfinance | FMP API |
|---------|----------|---------|
| **认证** | 免费，无需 API Key | 需注册获取 API Key |
| **费用** | 免费（依赖 Yahoo Finance） | 免费层有限，季度数据需付费 |
| **稳定性** | 不稳定（Yahoo 可能随时更改接口） | 稳定的 REST API，有版本管理 |
| **速率限制** | 有但不透明，可能被 IP 封禁 | 明确的速率限制，按订阅层级 |
| **数据格式** | pandas DataFrame | JSON（需自行转 DataFrame） |
| **财报字段名** | 原始名（如 "Net Income Common Stockholders"） | camelCase（如 `netIncome`） |
| **Adj Close** | 需 `auto_adjust=False` 才保留 | 始终返回 `adjClose` |
| **月频数据** | `interval="1mo"` 直接支持 | 无月频端点，需日线 → resample |
| **做空数据** | 有（sharesShort 等） | **无** |
| **数据范围** | 理论上无限制 | 历史价格每次最多 5 年 |
| **外汇符号** | `USDJPY=X` | `USDJPY` |
| **台股/日股** | 支持（.TW / .T 后缀） | 部分支持，需验证各市场覆盖度 |

---

## 五、迁移风险与注意事项

### 5.1 数据缺失风险

| 风险项 | 严重度 | 应对方案 |
|--------|--------|---------|
| Short Interest 数据缺失 | 中 | 保留 yfinance 作为补充源，或使用 FINRA 数据 |
| `financialCurrency` 无直接字段 | 低 | 从财报的 `reportedCurrency` 字段取得 |
| `impliedSharesOutstanding` 无对应 | 低 | 可省略或从 diluted EPS 反算 |
| Unusual Items 无标签 | 中 | 使用 `as-reported` 端点搜索关键词 |

### 5.2 字段名映射风险

yfinance 返回的财报行名（如 `"Total Revenue"`、`"Net Income Common Stockholders"`）是**不稳定**的，不同公司/时期可能变化。FMP 使用**固定的 camelCase 字段名**（如 `revenue`、`netIncome`），这反而是一个**优势** — 迁移后代码更健壮。

### 5.3 历史价格分页

FMP 历史价格端点每次最多返回 5 年数据。当前脚本默认从 2014 年开始，需要分 3 次请求并合并。建议封装一个 `fetch_all_history(symbol, start, end)` 函数自动分页。

### 5.4 台股/日股/港股支持

FMP 对非美股市场的覆盖度需验证：
- 美股 (`AAPL`, `DIOD`)：完全支持
- 日股 (`9022.T`)：FMP 使用格式如 `9022.T`，需测试
- 台股 (`2211.TW`)：需测试 FMP 是否覆盖
- 港股 (`0005.HK`)：需测试

### 5.5 费用考量

| FMP 订阅层级 | 价格 | 关键限制 |
|-------------|------|---------|
| Free | $0 | 250 次/天，仅年度财报，5 年历史 |
| Starter | $9.99/月 | 300 次/天，含季度财报 |
| Professional | ~$29/月 | 更高限制 |

**最低要求**：需 **Starter 以上**（因为要季度财报数据）

---

## 六、建议实施路线

### Phase 1：基础替换（最小改动）
1. 新增 `fmp_client.py` — 封装所有 FMP API 调用（requests + JSON → DataFrame）
2. 保持 `generate_report_summary.py` 的函数签名不变
3. 新增 `generate_report_summary_fmp.py` — 替换数据抓取层，写入逻辑不变

### Phase 2：字段映射层
1. 建立 yfinance → FMP 字段名映射 dict
2. 将 FMP 返回的 camelCase 字段转为 yfinance 风格行名（确保 Excel 输出一致）

### Phase 3：测试验证
1. 同一股票（如 AAPL）分别用 yfinance 和 FMP 生成 Excel
2. 逐 cell 比对两份 Excel 的数值差异
3. 记录并处理差异

### Phase 4：做空数据补充
1. 如需 Short Interest → 保留 yfinance 单独调用或接入 FINRA

---

## 七、FMP Python SDK 参考

FMP 提供官方 Python SDK：

```bash
pip install financialmodelingprep
```

```python
from financialmodelingprep import FMP

fmp = FMP(api_key="YOUR_API_KEY")

# 年度损益表
fmp.get_income_statement("AAPL", period="annual", limit=5)

# 季度资产负债表
fmp.get_balance_sheet("AAPL", period="quarter", limit=5)

# 公司 Profile
fmp.get_company_profile("AAPL")

# 历史价格
fmp.get_historical_price("AAPL", from_date="2014-01-01", to_date="2026-03-08")
```

也可以直接用 `requests`：
```python
import requests

API_KEY = "YOUR_API_KEY"
BASE_URL = "https://financialmodelingprep.com/api"

def fmp_get(endpoint, params=None):
    params = params or {}
    params["apikey"] = API_KEY
    resp = requests.get(f"{BASE_URL}{endpoint}", params=params)
    resp.raise_for_status()
    return resp.json()

# 示例
data = fmp_get("/v3/income-statement/AAPL", {"period": "quarter", "limit": 5})
```
