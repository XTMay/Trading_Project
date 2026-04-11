# EZ_table0226 Python 文件分析报告

## 总体概述

所有文件共用相同的技术栈和模式：

- **数据源**：`yfinance` (Yahoo Finance API)
- **数据处理**：`pandas` + `numpy`
- **输出格式**：Excel (`.xlsx`)，使用 `xlsxwriter` 引擎并禁用公式/URL 解析（防止 Excel 注入）
- **共通功能**：台股代号自动判断（`.TW` / `.TWO` 后缀）、数据清洗（`sanitize_excel`）

---

## 逐文件分析

### 1. BS_Financials_5Q_TTM.py — 季度资产负债表

| 项目 | 内容 |
|------|------|
| **股票** | `9022.T`（日股） |
| **数据** | **季度** Balance Sheet |
| **逻辑** | 检查最近5季是否连续（70-120天间隔）→ 完整输出5季 / 不足输出全部 / 不连续仅输出最近1季 |
| **TTM** | 直接取最新一季（资产负债表为时点数据，不需累加） |
| **输出** | `{ticker}_BS_autoSafe.xlsx` |

### 2. BS_Financials_5Y_TTM.py — 年度资产负债表

| 项目 | 内容 |
|------|------|
| **股票** | `8031.T`（日股） |
| **数据** | **年度** Balance Sheet |
| **逻辑** | 检查最近5年是否连续（330-400天间隔）→ 同上三种情况 |
| **TTM** | 取最新年度 |
| **输出** | `{ticker}_BS_5Y_plus_TTM.xlsx` |

### 3. CFS_Financials_5Y_TTM.py — 年度现金流量表

| 项目 | 内容 |
|------|------|
| **股票** | `2211.TW`（台股） |
| **数据** | **年度** Cash Flow + **季度** Cash Flow（用于 TTM） |
| **逻辑** | 年度完整性检查 → 输出5年/全部/最近1年 |
| **TTM** | 若有 ≥4 季季度数据 → 最近4季求和；否则取最近季度；无季度数据则取最新年度 |
| **输出** | `{ticker}_CFS_5Y_plus_TTM.xlsx` |

### 4. Company_Profile.py — 公司基本资料

| 项目 | 内容 |
|------|------|
| **股票** | `MITSY`（美股） |
| **数据** | `ticker.info` 中的公司概况 |
| **抓取字段** | 公司名、行业、国家、币别、市值、企业价值、PE（trailing/forward）、EPS、股息率、Beta、52周高低、流通股数、营收、净利、网站 |
| **输出** | `{ticker}_Basic_Info.xlsx`（两栏 Item/Value 格式） |

### 5. Company_Action.py — 公司行动（股息/拆股）

| 项目 | 内容 |
|------|------|
| **股票** | `MITSY` |
| **数据** | `ticker.actions`（历史除息与拆股记录） |
| **逻辑** | 分别提取 Dividends（非零）和 Stock Splits（非零），按日期倒序排列 |
| **输出** | `{ticker}_Actions.xlsx`（两栏格式，分段标注） |

### 6. EPS_Earnings.py — EPS / 盈余数据

| 项目 | 内容 |
|------|------|
| **股票** | `MITSY` |
| **数据** | 年度 EPS (`ticker.earnings`)、季度 EPS (`ticker.quarterly_earnings`)、TTM EPS (`ticker.info["trailingEps"]`) |
| **逻辑** | 依次尝试抓取三类数据，任一失败跳过 |
| **输出** | `{ticker}_EPS_Earnings.xlsx` |

### 7. Share_capital_data.py — 股本/持股结构

| 项目 | 内容 |
|------|------|
| **股票** | `MITSY` |
| **数据** | `ticker.info` 中的股本相关字段 |
| **抓取字段** | 流通股数、浮动股数、内部人持股比例、机构持股比例、做空股数、做空占比、市值、企业价值 |
| **输出** | `{ticker}_Share_Capital.xlsx` |

### 8. Foreign_exchange_rates.py — 外汇汇率

| 项目 | 内容 |
|------|------|
| **币别** | `JPY`（USD/JPY） |
| **数据** | 通过 `yfinance` 抓取 `USDJPY=X` 最近5天行情 |
| **抓取字段** | 收盘价、开盘价、最高价、最低价 |
| **输出** | `USD_to_{currency}_Exchange_Rate.xlsx` |

### 9. IS_Financials_5Y_plus_TTM_with_Unusual.py — 年度利润表

| 项目 | 内容 |
|------|------|
| **股票** | `MD`（美股） |
| **数据** | **年度** Income Statement (`ticker.financials`) |
| **逻辑** | 过滤掉当前年度（仅保留已结束年度）→ 取最近5年 + TTM（最新年度） |
| **特殊处理** | 文件名含 "Unusual" 但此版本未单独提取异常项目 |
| **输出** | `{ticker}_IS_5Y_plus_TTM_with_Unusual.xlsx` |

### 10. IS_Financials_5Q_plus_TTM_with_Unusual.py — 季度利润表 + 异常项

| 项目 | 内容 |
|------|------|
| **股票** | `DIOD`（美股） |
| **数据** | **季度** Income Statement (`ticker.quarterly_financials`) |
| **逻辑** | 取最近5季 + TTM（最近4季求和） |
| **异常项提取** | 搜索行名中包含 `unusual`、`special`、`restructuring`、`non recurring` 的项目，单独附在表格末尾（分隔符标注） |
| **输出** | `{ticker}_IS_5Q_plus_TTM_with_Unusual.xlsx` |

---

### 11. `Historical_stock_price_adj.py` — 历史股价（含调整价与公司行动）

| 项目 | 内容 |
|------|------|
| **股票** | `DIOD`（美股） |
| **数据** | 日线 OHLCV + Adj Close（`auto_adjust=False`）+ Corporate Actions（Dividends / Stock Splits） |
| **逻辑** | 日线数据 → 月频重采样（`MS` 月初基准）→ Open=first, High=max, Low=min, Close=last, Adj Close=last, Volume=sum → 合并 Corporate Actions → 按日期倒序排列 |
| **特殊处理** | `auto_adjust=False` 保留原始 Close 与 Adj Close 双栏；Corporate Action 以独立行插入（Entry_Type 标记为 `Corporate_Action`） |
| **输出** | `{ticker}_Historical_stock_price_adj.xlsx` |
| **集成位置** | `report_summary.xlsx` → BO1（标题）、BO4:BW4（列标题：Date, Open, High, Low, Close, Adj Close, Volume, Dividends, Stock Splits）、Row 5+（数据） |

---

## 文件分类总结

| 分类 | 文件 |
|------|------|
| **三大财务报表** | BS x2（季度/年度）、CFS x1（年度）、IS x2（季度/年度） |
| **公司信息** | Company_Profile、Company_Action、Share_capital_data |
| **盈余数据** | EPS_Earnings |
| **汇率** | Foreign_exchange_rates |
| **历史股价** | Historical_stock_price_adj |

这套脚本本质上是一个**股票基本面数据批量抓取工具集**，覆盖了财务报表、公司概况、股本结构、EPS、汇率和历史股价数据，支持美股、台股、日股，输出为标准化的 Excel 文件。
