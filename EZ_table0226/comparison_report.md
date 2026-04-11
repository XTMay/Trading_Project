# yfinance vs FMP 报告对比分析（AAPL）

> 对比日期：2026-03-08
> 测试标的：AAPL (Apple Inc.)

---

## 一、总体结论

| 维度 | 状态 | 说明 |
|------|------|------|
| **财报核心数值** | ✅ 完全一致 | Revenue, Net Income, Total Assets 等关键科目 TTM 值 100% 匹配 |
| **历史股价** | ✅ 近乎一致 | Close 一致，Adj Close 误差 < 0.01%（自算股息调整 vs Yahoo 原生） |
| **前台关键单元格** | ✅ 一致 | K3(价格), Y23(市值), E9(净利TTM), W25(币种) 完全匹配 |
| **字段命名** | ⚠️ 不同 | yfinance 用可读名（`Total Revenue`），FMP 用 camelCase（`revenue`） |
| **做空/持股数据** | ❌ FMP 缺失 | Short Interest、Insider/Institution Holdings 无法从 FMP 获取 |
| **Dividend Yield 格式** | ❌ 不一致 | yfinance 返回 0.4（异常值），FMP 返回 0.00404（正确小数） |

---

## 二、逐区域详细对比

### 2.1 Front Page 前台单元格

| 单元格 | 内容 | yfinance | FMP | 对比 |
|--------|------|---------|-----|------|
| A1 | 公司行 | `...AAPL (NMS)...` | `...AAPL (NASDAQ)...` | ⚠️ 交易所名称不同：yf 用内部代码 `NMS`，FMP 用通用名 `NASDAQ` |
| A2 | Ticker | AAPL | AAPL | ✅ |
| I1 | 日期 | 2026-03-08 | 2026-03-08 | ✅ |
| F15 | 汇率 | 1 | 1 | ✅ |
| K3 | 当前价 | 257.46 | 257.46 | ✅ |
| Y23 | 市值(M) | 3,784,127.81 | 3,784,127.81 | ✅ |
| Y24 | 股价 | 257.46 | 257.46 | ✅ |
| W25 | 财报币种 | USD | USD | ✅ |
| E9 | 净利TTM(M) | 117,777 | 117,777 | ✅ |
| K11 | 派息率 | 0.1304 | 0.1315 | ✅ 差 0.8%，来源不同（yf=info, FMP=ratios-ttm） |

### 2.2 财报数据（IS / BS / CFS）

#### 结构差异

| 项目 | yfinance | FMP | 说明 |
|------|---------|-----|------|
| **字段命名** | 可读英文（`Total Revenue`） | camelCase（`revenue`） | 字段名 0 重叠，但数据内容一致 |
| **季度IS行数** | 33 行 | 31 行 | yf 多出 Normalized EBITDA、Tax Rate For Calcs 等衍生字段 |
| **季度BS行数** | 65 行 | 53 行 | yf 多出 Net Debt、Tangible Book Value 等计算字段 |
| **年度CFS行数** | 53 行 | 39 行 | yf 多出 Beginning/End Cash Position 等汇总字段 |
| **列头日期** | 完全一致 | 完全一致 | 季度用 `2025-12`，年度用 `2025`，TTM 列都有 |

#### 核心科目 TTM 值对比

**季度损益表 (1/9)**

| 科目 | yfinance (`Total Revenue`) | FMP (`revenue`) | 差异 |
|------|---------------------------|-----------------|------|
| 营收 | 435,617,000,000 | 435,617,000,000 | ✅ 0% |
| 营业成本 | 229,460,000,000 | 229,460,000,000 | ✅ 0% |
| 毛利 | 206,157,000,000 | 206,157,000,000 | ✅ 0% |
| 营业利润 | 147,564,000,000 | 147,564,000,000 | ✅ 0% |
| 净利润 | 117,777,000,000 | 117,777,000,000 | ✅ 0% |
| EBITDA | — | — | ✅ 0% |
| 基本EPS | — | — | ✅ 0% |
| 稀释EPS | — | — | ✅ 0% |

**季度资产负债表 (2/9)**

| 科目 | 差异 |
|------|------|
| 总资产 (totalAssets) | ✅ 0% |
| 股东权益 (totalStockholdersEquity) | ✅ 0% |
| 现金 (cashAndCashEquivalents) | ✅ 0% |
| 总负债 (totalLiabilities) | ✅ 0% |
| 长期债务 (longTermDebt) | ✅ 0% |
| 保留盈余 (retainedEarnings) | ✅ 0% |

**年度现金流 (8/9)**

| 科目 | 差异 |
|------|------|
| 经营活动现金流 | ✅ 0% |
| 资本支出 | ✅ 0% |
| 自由现金流 | ✅ 0% |
| 股利支付 | ✅ 0% |
| 折旧摊销 | ✅ 0% |

**年度IS 微小差异**

| 科目 | 差异 | 说明 |
|------|------|------|
| EBITDA | ⚠️ 0.2% | yf 含 unusual items 调整，FMP 直接报告 |
| EBIT | ⚠️ 0.2% | 同上 |
| Interest Expense | ⚠️ | yf=None, FMP=0（不影响使用） |

### 2.3 历史股价 (5/9)

| 维度 | yfinance | FMP | 说明 |
|------|---------|-----|------|
| **总行数** | 198 | 198 | ✅ 完全一致 |
| **列头** | 9列一致 | 9列一致 | Date/OHLC/AdjClose/Volume/Dividends/StockSplits |
| **Close 价格** | float32 精度 | float64 精度 | FMP 返回干净数值（257.46），yf 有浮点噪声（257.4599914550781） |
| **Volume** | 整百 | 精确值 | yf=210,952,500 vs FMP=210,978,654（微小差异） |
| **Adj Close** | Yahoo 原生计算 | 自行股息调整 | 最近月份误差 0.00%，回溯至 2014 年累计误差 < 0.01% |

**Adj Close 对比样本：**

| 月份 | yf Adj Close | FMP Adj Close | 差异 |
|------|-------------|---------------|------|
| 2026/03 | 257.46 | 257.46 | 0.00% |
| 2026/01 | 259.237 | 259.234 | 0.00% |
| 2025/06 | 204.548 | 204.544 | 0.00% |
| 2025/01 | 234.718 | 234.731 | 0.01% |
| 2024/09 | 231.479 | 231.492 | 0.01% |

### 2.4 公司概况 (6/9)

| 字段 | yfinance | FMP | 对比 |
|------|---------|-----|------|
| Company Name | Apple Inc. | Apple Inc. | ✅ |
| Exchange | NMS | NASDAQ | ⚠️ 名称风格不同 |
| Country | United States | US | ⚠️ 全称 vs ISO 代码 |
| Current Price | 257.46 | 257.46 | ✅ |
| Market Cap | 3,784,127,807,488 | 3,784,127,807,317 | ✅ (差 $171k，忽略不计) |
| Enterprise Value | 3,803,408,236,544 | 3,829,319,807,317 | ⚠️ 差 0.7%（计算方法不同） |
| Trailing PE | 32.59 | 32.24 | ⚠️ 差 1.1% |
| **Forward PE** | **27.72** | **None** | **❌ FMP 无此字段** |
| EPS (TTM) | 7.90 | 7.99 | ⚠️ 差 1.1%（yf 圆整，FMP 精确） |
| **Dividend Yield** | **0.4** | **0.00404** | **❌ 格式不一致（见下文分析）** |
| Beta | 1.116 | 1.116 | ✅ |
| 52W High/Low | 288.62 / 169.21 | 288.62 / 169.21 | ✅ |
| Revenue (TTM) | 435,617,005,568 | 416,161,000,000 | ⚠️ 差 4.5%（yf=TTM 4Q sum, FMP=最近年报） |
| Net Income (TTM) | 117,776,998,400 | 112,010,000,000 | ⚠️ 差 4.9%（同上） |

### 2.5 市值/持股结构 (9/9)

| 字段 | yfinance | FMP | 对比 |
|------|---------|-----|------|
| Shares Outstanding | 14,681,140,000 | 14,697,925,143 | ✅ (差 0.1%) |
| Float Shares | 14,656,182,062 | 14,664,480,994 | ✅ (差 0.06%) |
| **Implied Shares Outstanding** | **14,697,926,000** | **None** | **❌ FMP 缺失** |
| **Held by Insiders (%)** | **1.637%** | **None** | **❌ FMP 缺失** |
| **Held by Institutions (%)** | **65.20%** | **None** | **❌ FMP 缺失** |
| **Short Shares** | **133,373,267** | **None** | **❌ FMP 缺失** |
| **Short Prior Month** | **113,576,032** | **None** | **❌ FMP 缺失** |
| **Short % of Float** | **0.91%** | **None** | **❌ FMP 缺失** |
| TTM EPS | 7.9 | None (quote无eps) | ❌ FMP quote 端点无 eps 字段 |

### 2.6 EPS/盈余数据 (9/9 附属)

**年度盈余 — 日期排序不一致：**

| 年份 | yfinance | FMP | 说明 |
|------|---------|-----|------|
| 2022 | 99,803M | 93,736M | ❌ 数值对调（Apple 会计年度错位） |
| 2023 | 96,995M | 96,995M | ✅ |
| 2024 | 93,736M | 99,803M | ❌ 数值对调 |
| 2025 | 112,010M | 94,680M | ❌ 不一致（yf=TTM, FMP=FY2025） |

**原因：** yfinance 用 `income_stmt` 的 column date 来标记年份，而 FMP 用 `date` 字段。Apple 的会计年度截止于 9 月底，导致同一会计年度在两个系统里分到不同的日历年标签。

**季度盈余 — 同样存在错位：**
FMP 返回更多季度（8Q vs yf 的 5Q），但相同季度的 netIncome 数值相同，只是日期标签可能偏移 1-3 天。

### 2.7 外汇汇率 (7/9)

AAPL 使用 USD 计价，汇率为 1.0，两边完全一致。
（对非美元股票如 `9022.T`，FMP 使用 `USDJPY` 符号请求 `/stable/quote`，yf 使用 `USDJPY=X`。）

---

## 三、关键差异汇总

### 3.1 ❌ FMP 无法提供的数据（7 项）

| 字段 | 影响范围 | 重要程度 |
|------|---------|---------|
| Forward PE | Profile / Market Cap | 中 — 可从 earnings estimates 端点补充 |
| Short Shares | Market Cap | 高 — FMP 无做空数据 |
| Short Prior Month | Market Cap | 高 |
| Short % of Float | Market Cap | 高 |
| Implied Shares Outstanding | Market Cap | 低 |
| Held by Insiders (%) | Market Cap | 中 |
| Held by Institutions (%) | Market Cap | 中 |

### 3.2 ⚠️ 格式/精度差异（6 项）

| 差异 | 详情 | 影响 |
|------|------|------|
| **Exchange 名称** | yf=`NMS`, FMP=`NASDAQ` | 仅显示差异，不影响计算 |
| **Country 格式** | yf=`United States`, FMP=`US` | 仅显示差异 |
| **Dividend Yield** | yf=0.4（疑似百倍放大），FMP=0.00404 | **需统一格式** |
| **Revenue/NI (Profile)** | yf=TTM(4Q sum), FMP=最近年报 | 差 4-5%，因数据定义不同 |
| **EPS 精度** | yf 圆整到 7.9，FMP 返回 7.986 | FMP 更精确 |
| **股价浮点精度** | yf=float32(257.459991...), FMP=float64(257.46) | FMP 更干净 |

### 3.3 ✅ 完全一致的数据

- 全部财报科目（IS/BS/CFS）核心数值
- 历史股价 Close 和 Adj Close（误差 < 0.01%）
- 当前价格、市值、52 周高低
- Beta、派息率
- 币种、Ticker、日期

---

## 四、优化建议

### 4.1 高优先级 — 修复已知问题

#### (1) Dividend Yield 格式统一
**问题：** yfinance 返回 `dividendYield=0.4`（实际应为 0.004 即 0.4%），FMP 返回正确的 `0.00404`。

**方案：** 在 FMP 版本中，当前 FMP 返回值（0.00404）已是正确的小数格式，无需修改。yfinance 版本的 0.4 是 yfinance 本身的 bug/格式问题。

#### (2) Profile 中 Revenue/Net Income 用 TTM 而非年报
**问题：** `fetch_company_profile()` 目前取最近一份年报的 revenue/netIncome，而 yfinance 返回 TTM（最近 4Q 之和）。

**方案：**
```python
# fmp_client.py fetch_company_profile() 中改为：
# 从季度 IS 取最近 4Q 的 revenue 和 netIncome 求和
is_q = fmp_get("income-statement", {"symbol": symbol, "period": "quarter", "limit": 4})
if isinstance(is_q, list) and len(is_q) >= 4:
    total_revenue = sum(r.get("revenue", 0) for r in is_q)
    net_income_common = sum(r.get("netIncome", 0) for r in is_q)
```

#### (3) TTM EPS 缺失
**问题：** FMP `/stable/quote` 不返回 `eps` 字段，导致 Market Cap 区域 TTM EPS 为 None。

**方案：** 已从 `ratios-ttm` 的 `netIncomePerShareTTM` 取值填入 profile 的 `trailingEps`。需确保 `fetch_eps_earnings()` 也用此值：
```python
# fetch_eps_earnings() 中的 TTM EPS
result["ttm_eps"] = ratios.get("netIncomePerShareTTM")  # 从 ratios-ttm 取
```

#### (4) EPS 年度/季度日期标签统一
**问题：** Apple 会计年度（9月底截止）导致 yfinance 和 FMP 的年份标签错位。

**方案：** 在 `fetch_eps_earnings()` 中使用 `calendarYear`（如果 FMP 返回）或 `fiscalYear` 字段代替从 `date` 截取年份，确保与 yfinance 的年份标签一致：
```python
# 使用 fiscalYear 字段
year = rec.get("fiscalYear") or rec.get("date", "")[:4]
```

### 4.2 中优先级 — 补充缺失数据

#### (5) 做空数据（Short Interest）
**方案 A（推荐）：** 保留 yfinance 单独调用做空数据，作为 FMP 的补充：
```python
# fmp_client.py 底部增加
def fetch_short_interest_yf(symbol):
    """Fallback: use yfinance for short interest data only."""
    import yfinance as yf
    info = yf.Ticker(symbol).info
    return {
        "Short Shares": info.get("sharesShort"),
        "Short Prior Month": info.get("sharesShortPriorMonth"),
        "Short % of Shares Outstanding": info.get("shortPercentOfFloat"),
    }
```
**方案 B：** 接入 FINRA Short Interest 数据 API。

#### (6) Insider/Institution Holdings
**方案：** 使用 FMP 的 `/stable/institutional-holder` 端点汇总计算机构持股比例：
```python
holders = fmp_get("institutional-holder", {"symbol": symbol})
total_inst_shares = sum(h.get("shares", 0) for h in holders)
inst_pct = total_inst_shares / shares_outstanding
```

#### (7) Forward PE
**方案：** 使用 FMP 的 analyst estimates 端点：
```
GET /stable/analyst-estimates?symbol=AAPL
```
取未来 12 个月 EPS 预估值，计算 Forward PE = Current Price / Forward EPS。

### 4.3 低优先级 — 改善数据质量

#### (8) 字段名映射层
**问题：** FMP 用 camelCase（`revenue`），盈再表 Excel 后续消费端可能依赖 yfinance 风格的可读名（`Total Revenue`）。

**方案：** 在 `_fmp_statements_to_df()` 后增加 rename 步骤：
```python
FIELD_RENAME = {
    "revenue": "Total Revenue",
    "costOfRevenue": "Cost Of Revenue",
    "grossProfit": "Gross Profit",
    "operatingIncome": "Operating Income",
    "netIncome": "Net Income",
    # ... 完整映射
}
df.index = df.index.map(lambda x: FIELD_RENAME.get(x, x))
```

#### (9) Exchange 名称标准化
**方案：** FMP 返回 `NASDAQ`，yfinance 返回 `NMS`。建议统一为 FMP 的通用名称（更易读）。无需修改。

#### (10) Country 格式标准化
**方案：** FMP 返回 ISO 2字母代码 `US`，yfinance 返回全称 `United States`。如需统一，可加映射 dict。

#### (11) 历史价格精度
**方案：** FMP 返回 float64 干净数值（无浮点噪声），比 yfinance 的 float32 更优。无需修改，FMP 版本已是改进。

---

## 五、优化优先级排序

| 优先级 | 编号 | 任务 | 难度 | 影响 |
|--------|------|------|------|------|
| P0 | (2) | Profile Revenue/NI 改用 TTM | 低 | 消除 4-5% 偏差 |
| P0 | (3) | TTM EPS 补充到 eps_earnings | 低 | 消除缺失值 |
| P0 | (4) | EPS 年度日期标签修复 | 中 | 消除年份错位 |
| P1 | (5) | 做空数据 fallback to yfinance | 低 | 补充 3 个缺失字段 |
| P1 | (8) | 字段名映射（camelCase → 可读名） | 中 | 与盈再表消费端兼容 |
| P2 | (6) | Insider/Institution Holdings | 中 | 补充 2 个缺失字段 |
| P2 | (7) | Forward PE | 低 | 补充 1 个缺失字段 |
| P3 | (1) | Dividend Yield 格式 | 无 | FMP 已是正确格式 |
| P3 | (9-11) | 标准化和精度 | 无 | FMP 已优于 yfinance |

---

## 六、结论

**FMP 版本完全可以替代 yfinance 作为主要数据源。** 核心财报数据（IS/BS/CFS）数值 100% 一致，历史股价误差 < 0.01%。主要差异在于：

1. **字段名格式不同**（camelCase vs 可读名）— 可通过映射层解决
2. **7 个字段 FMP 无法提供**（主要是做空数据）— 可用 yfinance 做 fallback
3. **Profile 的 Revenue/NI 定义不同**（年报 vs TTM）— 需改为 4Q 求和

建议执行 P0 优化后即可投入生产使用。
