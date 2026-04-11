# 盈再表 VBA → Python 迁移分析报告

> **报告日期**：2026-03-14 **分析范围**：`Excel/盈再表250723.xlsm` (VBA Module1–Module9) vs `EZ_table0226/*.py` **目标**：识别 VBA 中仍可迁移至 Python 的功能模块，保留 Excel 前端交互不变

------------------------------------------------------------------------

## 目录

1.  [项目架构总览](#1-项目架构总览)
2.  [VBA 模块功能全景图](#2-vba-模块功能全景图)
3.  [已迁移至 Python 的功能（已完成）](#3-已迁移至-python-的功能已完成)
4.  [尚未迁移的 VBA 功能（迁移机会）](#4-尚未迁移的-vba-功能迁移机会)
5.  [详细迁移方案：按优先级排序](#5-详细迁移方案按优先级排序)
6.  [不建议迁移的 VBA 功能](#6-不建议迁移的-vba-功能)
7.  [迁移后的目标架构](#7-迁移后的目标架构)
8.  [附录：VBA ↔ Python 逐模块对照表](#8-附录vba--python-逐模块对照表)

------------------------------------------------------------------------

## 1. 项目架构总览

### 当前状态

```         
┌─────────────────────────────────────────────────────────────┐
│                    盈再表 Excel 工作簿                        │
│  ┌──────┐ ┌──────┐ ┌──────┐ ┌──────┐ ┌──────┐              │
│  │ 台股 │ │ 美股 │ │ 港股 │ │ 中股 │ │ 全球 │ + 11 辅助Sheet │
│  └──┬───┘ └──┬───┘ └──┬───┘ └──┬───┘ └──┬───┘              │
│     │        │        │        │        │                   │
│  VBA        VBA      VBA      VBA      VBA                  │
│  Macro1    Macro2   Macro3   Macro4   Macro5                │
│     │        │        │        │        │                   │
│  ┌──┴────────┴────────┴────────┴────────┴──┐                │
│  │         9 个 VBA 模块 (~50,000行)        │                │
│  │  ┌─────────────┐  ┌──────────────────┐  │                │
│  │  │ 数据抓取层   │  │ 运算/格式化层     │  │                │
│  │  │ (IE/HTTP)   │  │ (公式/XIRR/税务) │  │                │
│  │  └─────────────┘  └──────────────────┘  │                │
│  └─────────────────────────────────────────┘                │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────┐
│          Python 脚本 (EZ_table0226/)             │
│  ┌───────────────────────────────────────┐      │
│  │ 数据抓取层 (yfinance / FMP API)       │      │
│  │ 11 个独立脚本 + 1 个汇总脚本          │      │
│  └───────────────────────────────────────┘      │
│  输出 → report_summary.xlsx → VBA 导入 Excel    │
└─────────────────────────────────────────────────┘
```

### 设计理念

-   **Excel 作为前端**：用户在 A2 输入股票代码即可触发全流程
-   **Python 负责后端**：数据抓取、清洗、TTM 计算
-   **VBA 作为桥梁**：调用 Python 脚本、导入数据、设置公式、格式化

------------------------------------------------------------------------

## 2. VBA 模块功能全景图

| 模块              | 代码行数 | 功能分类             | 覆盖市场       |
|-------------------|----------|----------------------|----------------|
| **Module1**       | \~6,000  | 数据抓取 + 公式设置  | 台股           |
| **Module2**       | \~8,000  | 数据抓取 + 公式设置  | 美股           |
| **Module3**       | \~6,000  | 数据抓取 + 公式设置  | 港股           |
| **Module4**       | \~5,000  | 数据抓取 + 公式设置  | 中股           |
| **Module5**       | \~7,000  | 数据抓取 + 公式设置  | 全球           |
| **Module6**       | \~5,000  | 市场指数组件列表     | 台/美/日/港/中 |
| **Module7**       | \~4,000  | 投资组合管理         | 多市场汇总     |
| **Module8**       | \~5,000  | 投资绩效 / XIRR 计算 | 多账户         |
| **Module9**       | \~2,000  | 税务计算 / 股息追踪  | 台股           |
| **JsonConverter** | \~1,000  | JSON 解析工具        | 通用           |

**VBA 功能可细分为 6 大类**：

```         
VBA 功能分层
├── ① 数据抓取 (Web Scraping / API)     ← 已大部分迁移至 Python
├── ② 前端公式设置 (FormulaR1C1)        ← 可迁移
├── ③ 财务指标运算 (ROE/盈再率/常利)     ← 可迁移（高优先级）
├── ④ 投资组合管理 (Portfolio)           ← 可迁移
├── ⑤ 投资绩效计算 (XIRR/巴菲特距离)    ← 可迁移
├── ⑥ 格式化 / UI 交互                  ← 保留在 VBA
└── ⑦ 税务计算 (台湾)                   ← 可迁移
```

------------------------------------------------------------------------

## 3. 已迁移至 Python 的功能（已完成）

以下功能已由 `EZ_table0226/` 下的 Python 脚本实现：

| 功能 | VBA 原始实现 | Python 脚本 | 覆盖度 |
|-----------------|--------------------|-------------------|-----------------|
| 季度损益表 (IS) | Module1-5 各市场 HTTP 爬取 | `IS_Financials_5Q_plus_TTM_with_Unusual.py` | ✅ 100% |
| 年度损益表 (IS) | Module1-5 各市场 HTTP 爬取 | `IS_Financials_5Y_plus_TTM_with_Unusual.py` | ✅ 100% |
| 季度资产负债表 (BS) | Module1-5 各市场 HTTP 爬取 | `BS_Financials_5Q_TTM.py` | ✅ 100% |
| 年度资产负债表 (BS) | Module1-5 各市场 HTTP 爬取 | `BS_Financials_5Y_TTM.py` | ✅ 100% |
| 年度现金流量表 (CFS) | Module2/5 WSJ/MarketWatch 爬取 | `CFS_Financials_5Y_TTM.py` | ✅ 美股/全球 |
| 公司基本资料 | Module1-5 多源爬取 | `Company_Profile.py` | ✅ 100% |
| 股息/拆股历史 | Module1-5 Yahoo/各交易所 | `Company_Action.py` | ✅ 100% |
| EPS/盈余 | Module1-5 嵌入式提取 | `EPS_Earnings.py` | ✅ 100% |
| 股本结构 | Module1-5 各市场爬取 | `Share_capital_data.py` | ✅ 100% |
| 外汇汇率 | Module2/3/5 wise.com 爬取 | `Foreign_exchange_rates.py` | ✅ 100% |
| 历史股价 (月频) | Module1-5 Yahoo CSV | `Historical_stock_price_adj.py` | ✅ 100% |
| **汇总报告生成** | VBA 各 Macro 串行执行 | `generate_report_summary.py` | ✅ 100% |

**结论**：数据抓取层已基本完成迁移，Python 使用统一的 yfinance/FMP API 替代了 VBA 中针对 10+ 个网站的 IE/HTTP 爬虫。

------------------------------------------------------------------------

## 4. 尚未迁移的 VBA 功能（迁移机会）

通过逐行分析 Module1–Module9 的代码，以下 VBA 功能**尚未被 Python 实现**但可以迁移：

### 4.1 核心财务指标计算（Module1-5 各市场共有）

| VBA 功能 | 说明 | 当前实现方式 | Python 迁移可行性 |
|-----------------|-----------------|-----------------|-----------------------|
| **ROE% 计算** | `Net Income / Stockholders Equity` | VBA 通过 `FormulaR1C1` 在 Excel 中设公式 | ✅ 高 — 直接用 Python 计算后写入 |
| **盈再率% (Profit Reinvestment Rate)** | `Capital Expenditure / Net Income` | VBA 设置 Excel 公式 | ✅ 高 — IS + CFS 数据已有 |
| **常利**$m (Recurring Profit)** | `Net Income - Unusual/Non-recurring Items` | VBA 公式 + `.Find()` 搜索异常项 | ✅ 高 — Unusual Items 已在 Python 中提取 |
| **常 EPS$ (Recurring EPS) | `Recurring Profit / Shares Outstanding` | VBA 公式引用多个数据区 | ✅ 高 — 所有输入数据已可得 |
| **配息率% (Payout Ratio)** | `Dividends / Net Income` | VBA 公式 | ✅ 高 |
| **预期报酬率** | `ROE × (1 - 盈再率) × 安全边际` | VBA 公式设在 E9 单元格 | ✅ 高 |
| **还原股价** | 除权除息后的还原计算 | VBA 循环累计计算 | ✅ 中 — 需要完整除权除息历史 |
| **便宜价/合理价/昂贵价** | 基于 ROE 和 EPS 的估值区间 | VBA 公式 | ✅ 高 |

### 4.2 投资组合管理（Module7）

| VBA 功能             | 说明                          | Python 迁移可行性     |
|------------------|------------------|------------------------------------|
| **持股列表同步**     | 跨 Sheet 搜索、匹配、汇总持股 | ✅ 高 — pandas 更擅长 |
| **自动更新最新价格** | 从各市场 Sheet 提取最新股价   | ✅ 高                 |
| **持股排名**         | 按预期报酬率排序              | ✅ 高                 |
| **市值加权汇总**     | 计算组合整体指标              | ✅ 高                 |
| **重复持股检查**     | 跨市场检查同一公司            | ✅ 高                 |

### 4.3 投资绩效计算（Module8）

| VBA 功能 | 说明 | Python 迁移可行性 |
|------------------|------------------|------------------------------------|
| **XIRR 年化收益率** | 基于交易日期和现金流的内部收益率 | ✅ 高 — `numpy_financial.xirr()` 或 `scipy.optimize` |
| **巴菲特距离评分** | `100 × (1.2^(50-years))^0.5 × LOG(9100, 1+CAGR) / 3701` | ✅ 高 — 纯数学计算 |
| **多账户绩效汇总** | 10 列为一组的重复结构，逐账户计算 | ✅ 高 |
| **年度收益率** | `(Current/Previous)^(1/years) - 1` | ✅ 高 |
| **持仓比率计算** | 股票 vs 现金的持仓比例 | ✅ 高 |
| **多币种投资组合** | USD/HKD/JPY/CNY 混合的绩效换算 | ✅ 中 — 需要维护汇率历史 |

### 4.4 税务计算（Module9）

| VBA 功能            | 说明                     | Python 迁移可行性    |
|---------------------|--------------------------|----------------------|
| **台湾股利税计算**  | 基于阈值的分级税率       | ✅ 高 — 简单条件逻辑 |
| **KY 股票特殊税务** | 外资公司不同扣缴规则     | ✅ 高                |
| **除息日匹配**      | 将持股与除息日期交叉匹配 | ✅ 高                |
| **年度税务汇总表**  | 按月/按股汇总应缴税额    | ✅ 高                |

### 4.5 市场指数组件列表（Module6）

| VBA 功能 | 说明 | Python 迁移可行性 |
|------------------|------------------|------------------------------------|
| **台股加权指数成分股** | 从 TWSE 网站抓取 | ✅ 高 — requests + pandas |
| **S&P 500 成分股** | 从 Finviz 抓取 | ✅ 高 — 或用 Wikipedia API |
| **日经 225 成分股** | IE 自动化从 investing.com 抓取 | ✅ 中 — 需替代 IE 的方案 |
| **恒生指数成分股** | IE 自动化从 investing.com 抓取 | ✅ 中 |
| **上证 A180 成分股** | 从 AAStocks 抓取 | ✅ 中 |
| **P/E 排序筛选** | 按 P/E 值排序指数成分 | ✅ 高 |

### 4.6 Excel 公式自动设置（Module1-5 共有）

| VBA 功能 | 说明 | Python 迁移可行性 |
|------------------|------------------|------------------------------------|
| **VLOOKUP 公式链** | 从原始数据区到前端展示区的数据引用链 | ⚠️ 中 — 可改为 Python 直接计算后写入值 |
| **IFERROR 保护层** | 防止 #N/A 错误显示 | ✅ 高 — Python 中直接处理 |
| **条件格式化** | 颜色高亮（红/绿/蓝标注） | ⚠️ 低 — openpyxl 可实现但代码量大 |
| **超链接生成** | 链接到 Yahoo Finance / MoneyDJ 等 | ✅ 高 |

------------------------------------------------------------------------

## 5. 详细迁移方案：按优先级排序

### 🔴 优先级 1：核心财务指标运算引擎（价值最大）

**目标**：将盈再表最核心的投资分析计算从 VBA 公式迁移至 Python

```         
当前 VBA 流程：
  Python 抓数据 → 写入 Excel 原始数据区 (Col AE~EQ)
  → VBA 设置 FormulaR1C1 → Excel 前端区 (Col A~T) 显示计算结果

迁移后流程：
  Python 抓数据 → Python 计算所有指标 → 直接写入 Excel 前端区
```

**需要实现的 Python 函数**：

```         
financial_calculator.py（新建）
│
├── calc_roe(net_income, stockholders_equity) → ROE%
│   来源：IS 年度 Net Income ÷ BS 年度 Stockholders Equity
│
├── calc_reinvestment_rate(capex, net_income) → 盈再率%
│   来源：CFS Capital Expenditure ÷ IS Net Income
│
├── calc_recurring_profit(net_income, unusual_items) → 常利$m
│   来源：IS Net Income − IS Unusual Items
│   注意：Unusual Items 已在 IS 脚本中按关键词提取
│
├── calc_recurring_eps(recurring_profit, shares) → 常 EPS$
│   来源：常利 ÷ Share Capital 流通股数
│
├── calc_payout_ratio(dividends, net_income) → 配息%
│   来源：Company Action 年度股息 ÷ IS Net Income
│
├── calc_expected_return(roe, reinv_rate) → 预期报酬率
│   公式：ROE × (1 − 盈再率)
│
├── calc_price_zones(recurring_eps, expected_return) → 便宜价/合理价/昂贵价
│   公式：便宜价 = 常 EPS ÷ 预期报酬率 × 安全边际系数
│
└── calc_restored_price(prices, dividends, splits) → 还原股价
    逻辑：逆向累计除权除息，还原历史真实报酬
```

**VBA 中对应的代码位置**（以美股 Module2 为例）：

| 指标 | VBA 代码行 | 关键 Cell |
|------------------|----------------------------|--------------------------|
| ROE% | `Range("C5").FormulaR1C1 = Net Income / Equity` | E5:E12 (Col C) |
| 盈再率% | `Range("D5").FormulaR1C1 = CapEx / NetIncome` | E5:E12 (Col D) |
| 常利 | `.Find("unusual")` → subtract from NI | Col E |
| 预期报酬 | `Range("E9").FormulaR1C1 = ...` | E9 |
| 便宜/合理/昂贵价 | O5, P5 VLOOKUP 公式 | O5, P5 |

------------------------------------------------------------------------

### 🟡 优先级 2：XIRR 投资绩效计算引擎

**目标**：将 Module8 的投资绩效计算完全迁移至 Python

**需要实现的 Python 函数**：

```         
performance_calculator.py（新建）
│
├── calc_xirr(cashflows, dates) → 年化收益率
│   实现：scipy.optimize.brentq 或 numpy_financial.xirr
│   替代：VBA 中的 Excel XIRR() 函数
│
├── calc_buffett_distance(years, annual_return) → 巴菲特距离评分
│   公式：
│     A = 1.2^(50 - years)
│     B = log(9100) / log(1 + annual_return)
│     Score = 100 × A^0.5 × B / 3701
│   评级标准：
│     <21: 与巴菲特相当
│     21-100: 优异
│     100-168: 良好
│     >168: 需定期检视
│
├── calc_annual_return(current_value, initial_value, years) → CAGR
│   公式：(current / initial) ^ (1/years) - 1
│
├── aggregate_portfolio(accounts[]) → 汇总绩效
│   逻辑：按 10 列一组的结构遍历所有账户，汇总收益
│
└── calc_holding_ratio(stock_value, cash_value) → 持仓比率
```

**VBA 中的关键代码片段**（Module8 `pf(f)` 函数）：

``` vba
' 巴菲特距离评分公式
[d5] = (1.2 ^ (50 - years)) ' A
[d6] = Log(9100) / Log(1 + CAGR) ' B (years needed to reach 9100x)
[g10] = Int(100 * [d5]^0.5 * [d6] / 3701) ' Distance Score
' 其中 3701 = (1.2^42)^0.5 × log(9100, 1.12), 基准：8年平均12%
```

------------------------------------------------------------------------

### 🟡 优先级 3：投资组合管理引擎

**目标**：将 Module7 的组合管理功能迁移至 Python

**需要实现的 Python 函数**：

```         
portfolio_manager.py（新建）
│
├── sync_holdings(excel_path) → 从投资组合 Sheet 读取持仓
│   读取：股票代码(col 22), 股价(col 24), 持股数(col 25),
│         买入价(col 26), 市场(col 28)
│
├── update_prices(holdings, market_data) → 更新最新股价
│   逻辑：VBA 原先从各市场 Sheet 的 Y24 单元格获取最新价
│         Python 直接从 yfinance 获取
│
├── rank_by_expected_return(holdings) → 按预期报酬率排序
│   排序键：col 23 (expected return)
│
├── check_duplicates(holdings) → 跨市场重复检查
│   逻辑：同一公司可能在港股和美股同时上市 (如 BABA / 9988.HK)
│
└── generate_summary(holdings) → 生成组合汇总报告
    输出：总市值、总成本、总报酬率、分市场统计
```

------------------------------------------------------------------------

### 🟢 优先级 4：市场指数成分股抓取

**目标**：将 Module6 的指数成分股爬取从 IE 自动化迁移至 Python

**需要实现的 Python 函数**：

```         
index_components.py（新建）
│
├── fetch_twse_components() → 台股加权指数成分
│   URL: https://www.twse.com.tw/exchangeReport/BWIBBU_d
│   方法：requests + pandas.read_html
│
├── fetch_sp500_components() → S&P 500 成分股
│   方法：Wikipedia API 或 datahub.io
│   VBA 原用 Finviz screener
│
├── fetch_nikkei225_components() → 日经 225
│   方法：Wikipedia 或 JPX 官网 API
│   VBA 原用 IE 自动化 investing.com (已不可用)
│
├── fetch_hangseng_components() → 恒生指数
│   方法：HKEX API 或 Wikipedia
│   VBA 原用 IE 自动化 investing.com (已不可用)
│
├── fetch_shanghai_a180() → 上证 A180
│   方法：上交所 API 或 AAStocks
│
└── sort_by_pe(components_df) → 按 P/E 排序
    用于筛选低估值成分股
```

**注意**：Module6 中使用了大量 IE COM 自动化（`CreateObject("InternetExplorer.Application")`），这在现代系统上已不可用（IE 已退役）。迁移到 Python 的 requests/Selenium 是**必要的技术更新**。

------------------------------------------------------------------------

### 🟢 优先级 5：税务计算引擎

**目标**：将 Module9 的台湾股利税务计算迁移至 Python

**需要实现的 Python 函数**：

```         
tax_calculator.py（新建）
│
├── calc_dividend_tax(dividends, shares, threshold, tax_rate)
│   逻辑：
│     if dividend_income >= threshold:
│       tax = (dividend × shares - supplementary) × (1 - rate) - adj
│     else:
│       tax = (dividend × shares) × rate - adj
│
├── is_ky_stock(stock_code) → 判断是否为 KY 外资公司
│   KY 股票适用不同扣缴规则
│
├── match_ex_dividend(holdings, ex_dates) → 匹配除息日
│   交叉匹配持股列表与除息日期
│
└── generate_tax_summary(year, holdings) → 年度税务汇总
    按月/按股生成应缴税额明细表
```

------------------------------------------------------------------------

## 6. 不建议迁移的 VBA 功能

以下功能因技术特性或实用性考量，建议**保留在 VBA 中**：

| 功能 | 原因 |
|------------------------------------|------------------------------------|
| **Excel UI 交互**（按钮、UserForm、事件触发） | VBA 是 Excel 原生交互语言，Python 无法替代 |
| **Worksheet_Change 事件**（A2 输入触发） | 依赖 Excel 事件模型 |
| **条件格式化**（颜色高亮、字体调整） | openpyxl 可实现但维护成本高，VBA 更直观 |
| **视图缩放**（Macro8/Macro9 的 Zoom 功能） | 纯 UI 操作 |
| **Sheet 保护/取消保护** | Excel 安全功能 |
| **列宽自动调整** | 格式化操作 |

**推荐的最小 VBA 保留层**：

``` vba
' === 精简后的 VBA 代码（~100 行） ===

' 1. 触发入口：监听 A2 输入变化
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$2" Then
        Call RunPythonFetch(Target.Value)
    End If
End Sub

' 2. 调用 Python
Sub RunPythonFetch(ticker As String)
    Dim cmd As String
    cmd = PYTHON_PATH & " " & SCRIPT_PATH & " " & ticker
    Shell cmd  ' 或 MacScript 方式
End Sub

' 3. 导入结果
Sub ImportResults()
    ' 从 report_summary.xlsx 读取计算完的数据
    ' 直接 paste values 到前端展示区
End Sub

' 4. 格式化（保留现有格式化代码）
Sub FormatSheet()
    ' 条件格式、列宽、字体等
End Sub
```

------------------------------------------------------------------------

## 7. 迁移后的目标架构

```         
┌─────────────────────────────────────────────────────────────────┐
│                    盈再表 Excel 工作簿（精简版）                   │
│                                                                 │
│  ┌──────────────────────────────────────┐                       │
│  │  前端展示区 (Col A ~ T)              │  用户在 A2 输入代码    │
│  │  ROE%, 盈再率%, 常利, 预期报酬       │  自动显示所有数据      │
│  │  便宜价/合理价/昂贵价, 还原股价      │                       │
│  │  XIRR, 巴菲特距离评分               │                       │
│  └──────────────────┬───────────────────┘                       │
│                     │ 直接写入计算结果（不再需要 Excel 公式）      │
│  ┌──────────────────┴───────────────────┐                       │
│  │  精简 VBA 层 (~100行)                │                       │
│  │  • Worksheet_Change 触发             │                       │
│  │  • Shell 调用 Python                 │                       │
│  │  • 导入 report_summary.xlsx          │                       │
│  │  • 格式化                            │                       │
│  └──────────────────┬───────────────────┘                       │
└─────────────────────┼───────────────────────────────────────────┘
                      │ Shell / AppleScript
┌─────────────────────┴───────────────────────────────────────────┐
│                Python 全栈引擎                                   │
│                                                                 │
│  ┌─────────────────────────────────────────────────────┐        │
│  │  数据抓取层 (已完成 ✅)                               │        │
│  │  yfinance / FMP API → IS/BS/CFS/Profile/Action/...  │        │
│  └─────────────────────┬───────────────────────────────┘        │
│                        │                                        │
│  ┌─────────────────────┴───────────────────────────────┐        │
│  │  运算引擎层 (待实现 🔴🟡)                             │        │
│  │  financial_calculator.py  → ROE, 盈再率, 常利, EPS   │        │
│  │  performance_calculator.py → XIRR, 巴菲特距离        │        │
│  │  portfolio_manager.py     → 组合管理, 排名           │        │
│  │  index_components.py      → 指数成分股               │        │
│  │  tax_calculator.py        → 税务计算                 │        │
│  └─────────────────────┬───────────────────────────────┘        │
│                        │                                        │
│  ┌─────────────────────┴───────────────────────────────┐        │
│  │  Excel 输出层 (已完成 ✅, 需扩展)                      │        │
│  │  generate_report_summary.py                          │        │
│  │  → 原始数据 (Col AE~EQ)                              │        │
│  │  → 计算结果 (Col A~T) ← 新增                         │        │
│  │  → 组合/绩效 Sheet    ← 新增                         │        │
│  └─────────────────────────────────────────────────────┘        │
└─────────────────────────────────────────────────────────────────┘
```

### 迁移收益总结

| 维度 | 当前 (VBA) | 迁移后 (Python) |
|------------------|----------------------|--------------------------------|
| **代码量** | \~50,000 行 VBA | \~5,000 行 Python + \~100 行 VBA |
| **可维护性** | 低 — 编码混乱、中文乱码、GoTo 跳转 | 高 — 模块化、可测试 |
| **数据源依赖** | 10+ 个网站、IE COM 自动化（已失效） | 统一 yfinance/FMP API |
| **运行速度** | 慢 — 逐个网页加载 | 快 — API 并发请求 |
| **跨平台** | 仅 Windows（IE 依赖） | macOS + Windows |
| **可扩展性** | 低 — 每增一个市场需复制整个模块 | 高 — 参数化设计 |

------------------------------------------------------------------------

## 8. 附录：VBA ↔ Python 逐模块对照表

### Module1 (台股) — 迁移状态

| 步骤 | VBA 功能 | Python 对应 | 状态 |
|-----------------|-----------------|----------------------|-----------------|
| Step 1 | 公司基本资料 (MoneyDJ) | `Company_Profile.py` | ✅ 已迁移 |
| Step 2 | 股利政策 (MoneyDJ) | `Company_Action.py` | ✅ 已迁移 |
| Step 3 | 季度损益表 (MoneyDJ) | `IS_Financials_5Q_plus_TTM.py` | ✅ 已迁移 |
| Step 4 | 季度资产负债表 (MoneyDJ) | `BS_Financials_5Q_TTM.py` | ✅ 已迁移 |
| Step 5 | 负债资料 (MoneyDJ) | 嵌入 BS 数据中 | ✅ 已迁移 |
| Step 6 | 现金流量表 (MoneyDJ) | `CFS_Financials_5Y_TTM.py` | ✅ 已迁移 |
| Step 7 | 股利除权除息 (TWSE/TPEx) | `Company_Action.py` | ✅ 已迁移 |
| Step 8 | 年度财报 (年度 IS/BS) | `IS_5Y.py` + `BS_5Y.py` | ✅ 已迁移 |
| Step 9 | 历史月线 (Yahoo Finance) | `Historical_stock_price_adj.py` | ✅ 已迁移 |
| Step 10 | ROE/盈再率/常利 计算 | — | ❌ 待迁移 |
| Step 11 | 前端公式设置 + 格式化 | — | ❌ 待迁移 (运算部分) |

### Module2 (美股) — 迁移状态

| 步骤 | VBA 功能 | Python 对应 | 状态 |
|-----------------|-----------------|----------------------|-----------------|
| Step 1-7 | 数据抓取 (MarketWatch/Yahoo) | `generate_report_summary.py` | ✅ 已迁移 |
| Step 8 | ROE%/盈再率%/常利/常EPS 计算 | — | ❌ 待迁移 |
| Step 9 | 便宜/合理/昂贵价 估值 | — | ❌ 待迁移 |
| 公式设置 | VLOOKUP/IFERROR 公式链 | — | ⚠️ 可改为直接写值 |

### Module5 (全球) — 迁移状态

| 步骤     | VBA 功能             | Python 对应                  | 状态      |
|----------|----------------------|------------------------------|-----------|
| Step 1-7 | 数据抓取 (WSJ/Yahoo) | `generate_report_summary.py` | ✅ 已迁移 |
| Step 8   | 巴菲特距离评分       | —                            | ❌ 待迁移 |
| Step 9   | XIRR 收益率计算      | —                            | ❌ 待迁移 |

### Module6 (市场指数) — 迁移状态

| Macro   | VBA 功能             | 状态                   |
|---------|----------------------|------------------------|
| Macro10 | 台股加权指数 (TWSE)  | ❌ 待迁移              |
| Macro11 | S&P 500 (Finviz)     | ❌ 待迁移              |
| Macro12 | 日经 225 (IE 自动化) | ❌ 待迁移（IE 已失效） |
| Macro13 | 恒生 40 (IE 自动化)  | ❌ 待迁移（IE 已失效） |
| Macro14 | 上证 A180 (AAStocks) | ❌ 待迁移              |

### Module7 (投资组合) — 迁移状态

| 功能     | 状态      |
|----------|-----------|
| 持股同步 | ❌ 待迁移 |
| 价格更新 | ❌ 待迁移 |
| 排名排序 | ❌ 待迁移 |

### Module8 (绩效计算) — 迁移状态

| 功能       | 状态      |
|------------|-----------|
| XIRR 计算  | ❌ 待迁移 |
| 巴菲特距离 | ❌ 待迁移 |
| 多账户汇总 | ❌ 待迁移 |

### Module9 (税务) — 迁移状态

| 功能       | 状态      |
|------------|-----------|
| 股利税计算 | ❌ 待迁移 |
| KY 股税务  | ❌ 待迁移 |
| 年度汇总   | ❌ 待迁移 |

------------------------------------------------------------------------

## 总结

| 类别             | 功能数 | 已完成       | 待迁移       | 不迁移       |
|------------------|--------|--------------|--------------|--------------|
| 数据抓取         | 12     | 12 ✅        | 0            | 0            |
| 核心财务指标运算 | 8      | 0            | **8** 🔴     | 0            |
| XIRR/绩效计算    | 6      | 0            | **6** 🟡     | 0            |
| 投资组合管理     | 5      | 0            | **5** 🟡     | 0            |
| 市场指数成分     | 6      | 0            | **6** 🟢     | 0            |
| 税务计算         | 4      | 0            | **4** 🟢     | 0            |
| UI/格式化/事件   | 10     | 0            | 0            | 10           |
| **合计**         | **51** | **12 (24%)** | **29 (57%)** | **10 (19%)** |

**核心结论**：数据抓取层已完成迁移（24%），但盈再表的**核心价值——投资分析运算引擎**（ROE、盈再率、常利、XIRR、巴菲特距离等）仍全部依赖 VBA 公式，占总功能的 57%。这些运算逻辑是迁移的最大机会点，可大幅简化 VBA 代码量（从 \~50,000 行降至 \~100 行），同时提升可维护性和跨平台兼容性。