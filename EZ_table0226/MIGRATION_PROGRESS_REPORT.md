# 盈再表 VBA → Python 迁移进度报告

> 生成日期: 2026-03-14 项目路径: `EZ_table0226/`

------------------------------------------------------------------------

## 一、项目总览

**盈再表**（Profit Reinvestment Rate Analysis）是一个多市场价值投资分析 Excel 工具（作者：洪瑞泰），原始 VBA 代码分布在 9 个模块 + 1 个共享模块，共 **11,646 行 VBA**。

Python 迁移目标：将 VBA 中的数据获取 + 核心计算逻辑迁移为可独立运行的 Python 脚本，摆脱对 Excel/VBA 的依赖。

------------------------------------------------------------------------

## 二、迁移进度总表

| 阶段 | 状态 | 文件数 | 代码行数 | 覆盖的 VBA 模块 |
|---------------|---------------|---------------|---------------|---------------|
| **Phase 1: 数据获取** | ✅ 已完成 | 14 个 .py | 2,811 行 | Module1-5 (Step 1-9) |
| **Phase 2: 核心计算** | ✅ 已完成 | 5 个 .py | 2,213 行 | Module1-5 (Step 10-11), Module6-9 |
| **总计** | — | **19 个 .py** | **5,024 行** | — |

### VBA 模块覆盖率

| VBA 模块 | 行数 | 功能 | Python 覆盖 | 状态 |
|---------------|---------------|---------------|---------------|---------------|
| Module1 (台股) | 930 | 台股数据获取+计算 | generate_report_summary.py + financial_calculator.py | ✅ 计算部分已迁移 |
| Module2 (美股) | 2,196 | 美股数据获取+ROE/盈再率计算 | generate_report_summary.py + financial_calculator.py | ✅ 已迁移 |
| Module3 (港股) | 952 | 港股数据获取+计算 | generate_report_summary.py + financial_calculator.py | ✅ 计算部分已迁移 |
| Module4 (中股) | 858 | A股数据获取+计算 | generate_report_summary.py + financial_calculator.py | ✅ 计算部分已迁移 |
| Module5 (全球) | 1,337 | 全球市场获取+计算 | generate_report_summary.py + financial_calculator.py | ✅ 计算部分已迁移 |
| Module6 (指数成分) | 1,000 | S&P 500/TWSE 成分获取 | index_components.py | ✅ 已迁移 |
| Module7 (投资组合) | 2,547 | 持仓管理/排名/更新 | portfolio_manager.py | ✅ 已迁移 |
| Module8 (绩效) | 1,205 | XIRR/巴菲特距离/年化报酬 | performance_calculator.py | ✅ 已迁移 |
| Module9 (税务) | 370 | 台湾股利税计算 | tax_calculator.py | ✅ 已迁移 |
| shared_modules | 251 | HTTP/HTML 解析工具 | yfinance + pandas 替代 | ✅ 不再需要 |

------------------------------------------------------------------------

## 三、Phase 1 — 数据获取层（已完成）

12 个独立数据获取脚本 + 1 个统一汇总脚本 + 1 个 FMP 替代方案：

| 文件名 | 行数 | 对应功能 | 数据源 |
|------------------|------------------|-------------------|------------------|
| `IS_Financials_5Q_plus_TTM_with_Unusual.py` | 113 | 季度利润表 + 异常项 | yfinance |
| `IS_Financials_5Y_plus_TTM_with_Unusual.py` | 84 | 年度利润表 + 异常项 | yfinance |
| `BS_Financials_5Q_TTM.py` | 109 | 季度资产负债表 | yfinance |
| `BS_Financials_5Y_TTM.py` | 109 | 年度资产负债表 | yfinance |
| `CFS_Financials_5Y_TTM.py` | 134 | 年度现金流量表 | yfinance |
| `Company_Profile.py` | 94 | 公司基本资料 | yfinance |
| `Company_Action.py` | 97 | 股息/拆股记录 | yfinance |
| `EPS_Earnings.py` | 102 | EPS/盈余数据 | yfinance |
| `Share_capital_data.py` | 89 | 股本/持股结构 | yfinance |
| `Foreign_exchange_rates.py` | 72 | 外汇汇率 | yfinance |
| `Historical_stock_price_adj.py` | 84 | 历史股价（含还原） | yfinance |
| `generate_report_summary.py` | 724 | **统一汇总**：一键生成完整报告 | yfinance |
| `generate_report_summary_fmp.py` | 416 | FMP API 替代方案 | FMP |
| `fmp_client.py` | 584 | FMP API 客户端 | FMP |

**运行方式**: `python generate_report_summary.py AAPL` → 输出 `report_summary.xlsx`

------------------------------------------------------------------------

## 四、Phase 2 — 核心计算层（本次新增 ✅）

5 个新增 Python 文件，共 **2,213 行**，**43 个函数**：

------------------------------------------------------------------------

### 4.1 `financial_calculator.py` — 核心财务指标引擎 (576 行, 12 函数)

**对应 VBA**: Module1-5 Step 10-11 前端计算区

| 函数 | 公式 | VBA 来源 |
|---------------------|---------------------|------------------------------|
| `calc_roe(is_y, bs_y)` | Net Income / Stockholders Equity | Module2 ROE% 计算 |
| `calc_reinvestment_rate(cfs_y, is_y)` | \|CapEx\| / Net Income | Module2 盈再率% |
| `calc_recurring_profit(is_y)` | Net Income - Unusual Items | Module2 常利 |
| `calc_recurring_eps(profit, shares)` | Recurring Profit / Shares Outstanding | Module2 常EPS |
| `calc_payout_ratio(ticker, is_y)` | Annual Dividends / Net Income | Module2 配息% |
| `calc_expected_return(roe, reinv)` | ROE × (1 - 盈再率) | Module2 預期報酬 |
| `calc_price_zones(eps, return)` | 便宜价(×0.6) / 合理价 / 昂贵价(×1.4) | Module2 价格区间 |
| `calc_restored_price(prices_df)` | 累计还原除息除权 | Module2 還原股價 |
| `generate_analysis(ticker)` | **主函数**: 获取数据 + 计算全部指标 | — |
| `export_analysis_to_excel(result)` | 输出分析结果到 Excel | — |

**运行方式**: `python financial_calculator.py AAPL` → 输出 `financial_analysis_AAPL.xlsx`

**数据流**:

```         
yfinance → fetch_annual_income/BS/CFS (from generate_report_summary.py)
         → calc_roe → calc_reinvestment_rate → calc_expected_return
         → calc_recurring_profit → calc_recurring_eps → calc_price_zones
         → generate_analysis() → export_analysis_to_excel()
```

------------------------------------------------------------------------

### 4.2 `performance_calculator.py` — XIRR & 巴菲特距离引擎 (429 行, 10 函数)

**对应 VBA**: Module8

| 函数 | 公式 | VBA 来源 |
|---------------------|---------------------|------------------------------|
| `calc_xirr(cashflows, dates)` | Σ cf_i / (1+r)\^((d_i-d_0)/365) = 0 | Module8 XIRR |
| `calc_annual_return(current, initial, years)` | (current/initial)\^(1/years) - 1 | Module8 年化报酬 |
| `calc_buffett_distance(years, return)` | 100 × (1.2^(50-yr))^0.5 × log(9100)/log(1+r) / 3701 | Module8 `pf(f)` |
| `calc_holding_ratio(stock, cash)` | stock / (stock + cash) | Module8 持股比 |
| `evaluate_portfolio_performance(df)` | 解析交易记录 → XIRR + 巴菲特距离 | Module8 综合 |
| `evaluate_from_excel(path)` | 从 Excel 加载交易记录 | — |

**巴菲特距离评分标准** (from VBA):

| 分数      | 评级          | 说明           |
|-----------|---------------|----------------|
| ≤ 21      | Buffett-level | 巴菲特等级！   |
| 22 - 100  | Excellent     | 非常优秀！     |
| 101 - 168 | Good          | 加油！         |
| \> 168    | Needs review  | 常回来看看讲义 |

**验证结果** ✅:

| 测试案例                     | 预期             | 实际           |
|------------------------------|------------------|----------------|
| 投入\$1000 → 1年后收回\$1100 | XIRR = 10%       | ✅ 10.0000%    |
| 基准线: 8年, 12%年化         | 巴菲特距离 ≈ 100 | ✅ score = 100 |
| 20年, 20%年化                | 巴菲特距离 ≤ 21  | ✅ score = 20  |

**运行方式**: `python performance_calculator.py portfolio.xlsx`

------------------------------------------------------------------------

### 4.3 `portfolio_manager.py` — 投资组合管理引擎 (479 行, 9 函数)

**对应 VBA**: Module7

| 函数 | 功能 | VBA 来源 |
|---------------------|---------------------|------------------------------|
| `load_holdings(excel_path)` | 从 Excel 加载持仓（弹性列名匹配） | Module7 `lst()` |
| `update_prices(holdings)` | yfinance 批量获取最新价格 | Module7 `Macro21()` |
| `calc_expected_returns(holdings)` | 调用 financial_calculator 计算每只股票期望报酬 | Module7 Col 23 |
| `rank_holdings(holdings)` | 按期望报酬降序排列 | Module7 `Macro21()` 排序 |
| `check_duplicates(holdings)` | 跨市场重复检测 (BABA/9988.HK 等) | Module7 `Macro20()` |
| `generate_portfolio_summary(holdings)` | 总市值/成本/报酬率 + 分市场统计 | Module7 汇总 |
| `export_portfolio(holdings, summary)` | 输出投组报告 (Holdings + Summary 两个 sheet) | — |

**支持的 Excel 列名** (自动匹配): - 代号: `code`, `ticker`, `symbol`, `股票代号` - 股数: `shares`, `股数`, `持股`, `quantity` - 成本: `buy_price`, `cost`, `成本价`, `买入价` - 市场: `market`, `市场` (可选, 自动检测 .TW/.HK/.SS)

**运行方式**: `python portfolio_manager.py portfolio.xlsx`

------------------------------------------------------------------------

### 4.4 `tax_calculator.py` — 台湾股利税引擎 (393 行, 6 函数)

**对应 VBA**: Module9

| 函数 | 功能 | VBA 来源 |
|---------------------|---------------------|------------------------------|
| `calc_dividend_tax(div, shares, is_ky)` | 股利扣繳稅款计算 | Module9 `exdiv()` |
| `is_ky_stock(code)` | KY 股判定 | Module9 `InStr(,"KY")` |
| `match_ex_dividends(holdings, year)` | 配對除息日期+金額 | Module9 Line 77-119 |
| `generate_tax_summary(year, holdings)` | 年度稅務彙總 (KY/非KY 分開) | Module9 Line 159-170 |
| `export_tax_report(summary)` | 输出稅務報告 Excel | — |

**税务常数** (from VBA \[i10\], \[i11\], \[i12\]):

| 常数                 | 值         | 说明              |
|----------------------|------------|-------------------|
| `DIVIDEND_THRESHOLD` | 20,000 NTD | 大/小额股利分界   |
| `SUPPLEMENT_RATE_KY` | 2.11%      | KY 股补充保费费率 |
| `FLAT_FEE`           | 10 NTD     | 固定手续费        |
| `DEFAULT_TAX_RATE`   | 25%        | 预设扣缴税率      |

**验证结果** ✅:

| 测试案例 | 预期 | 实际 |
|------------------------------|---------------------|---------------------|
| 非KY, 1000股×\$5 = \$5,000 (\< 阈值) | tax = 5000×0.25 - 10 = \$1,240 | ✅ \$1,240.0 |
| KY 检测: "6547-KY" | True | ✅ |
| KY 检测: "2330" | False | ✅ |

**运行方式**: `python tax_calculator.py holdings.xlsx 2024`

------------------------------------------------------------------------

### 4.5 `index_components.py` — 市场指数成分引擎 (336 行, 6 函数)

**对应 VBA**: Module6 (高可行性部分)

| 函数 | 功能 | VBA 来源 |
|---------------------|---------------------|------------------------------|
| `fetch_sp500_components()` | 从 Wikipedia 获取 S&P 500 成分 | Module6 `Macro11()` (原用 FinViz) |
| `fetch_twse_components()` | 从 TWSE JSON API 获取台股成分 | Module6 `Macro10()` |
| `sort_by_pe(df)` | 按 P/E 排序 | Module6 `stwpe()` |
| `enrich_with_metrics(df, max)` | 批量添加 P/E, 殖利率, 市值 | Module6 数据增强 |
| `export_components(df)` | 输出成分列表 Excel | — |

**数据源改进**:

| 数据         | VBA 原始来源              | Python 新来源        | 改进       |
|--------------|---------------------------|----------------------|------------|
| S&P 500 成分 | FinViz (需分页爬取 26 页) | Wikipedia (单次请求) | 更简洁可靠 |
| TWSE 成分    | TWSE HTML 表格            | TWSE JSON API        | 结构化数据 |
| 财务指标     | FinViz 表格               | yfinance batch       | 统一数据源 |

**运行方式**: `python index_components.py sp500 --enrich --top 50`

------------------------------------------------------------------------

## 五、文件依赖关系

```         
generate_report_summary.py          ← 基础数据获取层 (Phase 1)
  ├── _sanitize(), _safe_val()      ← 共享工具函数
  ├── fetch_annual_income()
  ├── fetch_annual_balance_sheet()
  ├── fetch_annual_cashflow()
  ├── fetch_company_profile()
  ├── fetch_share_capital()
  ├── fetch_historical_prices_adj()
  └── fetch_exchange_rate()
        │
        ▼
financial_calculator.py             ← 导入上述 fetch 函数
  ├── calc_roe()
  ├── calc_reinvestment_rate()
  ├── calc_expected_return()
  └── generate_analysis()
        │
        ▼
portfolio_manager.py                ← 导入 financial_calculator
  ├── calc_expected_returns()       ← 调用 generate_analysis()
  └── rank_holdings()

performance_calculator.py           ← 独立（仅导入 _safe_val）
  ├── calc_xirr()
  └── calc_buffett_distance()

tax_calculator.py                   ← 独立（仅导入 _safe_val）
  └── calc_dividend_tax()

index_components.py                 ← 独立（仅导入 _safe_val）
  ├── fetch_sp500_components()
  └── fetch_twse_components()
```

------------------------------------------------------------------------

## 六、技术栈

| 依赖     | 版本   | 用途               |
|----------|--------|--------------------|
| Python   | 3.x    | 运行环境           |
| yfinance | latest | 统一金融数据 API   |
| pandas   | latest | 数据处理           |
| numpy    | latest | 数值计算           |
| openpyxl | latest | Excel 读写         |
| scipy    | latest | XIRR 求解 (brentq) |

------------------------------------------------------------------------

## 七、运行指南

``` bash
cd EZ_table0226/

# Phase 1: 数据获取 — 生成原始报告
python generate_report_summary.py AAPL

# Phase 2: 核心计算
python financial_calculator.py AAPL              # ROE/盈再率/价格区间分析
python performance_calculator.py                 # XIRR/巴菲特距离 Demo
python performance_calculator.py txns.xlsx       # 从交易记录计算绩效
python portfolio_manager.py portfolio.xlsx       # 投资组合管理
python tax_calculator.py holdings.xlsx 2024      # 台湾股利税计算
python index_components.py sp500 --enrich        # S&P 500 成分 + 财务指标
python index_components.py twse                  # TWSE 台股成分
```

------------------------------------------------------------------------

## 八、后续待完成项目（低可行性 / 未迁移）

| 项目                      | 原因                       | 优先级    |
|---------------------------|----------------------------|-----------|
| 台股 MoneyDJ 数据源       | 需要解析中文 HTML, 无 API  | 低        |
| 港股 AAStocks 数据源      | 需要复杂网页爬虫           | 低        |
| 中股 Sina Finance 数据源  | 数据源不稳定               | 低        |
| VBA GUI 交互 (Macro 触发) | 需要前端替代方案 (Web/TUI) | 中        |
| Excel 公式层 (VLOOKUP/IF) | 已由 Python 计算替代       | ✅ 不需要 |