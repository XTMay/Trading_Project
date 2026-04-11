# Python 脚本 ↔ 盈再表 Excel 数据映射分析

> 分析对象：`EZ_table0226/*.py` → `Excel/盈再表250723.xlsm`

---

## Excel 工作簿概述

**盈再表**是一个多市场价值投资分析工具（作者：洪瑞泰），共 16 个 sheet。其中 5 个 sheet 与 Python 脚本直接相关：

| Sheet 名称 | 对应市场 | VBA 入口 | 触发方式 |
|------------|---------|---------|---------|
| **台股** | 台湾股市 | `Macro1` | 修改 `A2` 单元格 |
| **美股** | 美国股市 | `Macro2` | 修改 `A2` 单元格 |
| **港股** | 香港股市 | `Macro3` | 修改 `A2` 单元格 |
| **中股** | 中国 A 股 | `Macro4` | 修改 `A2` 单元格 |
| **全球** | 全球市场 | `Macro5` | 修改 `A2` 单元格 |

每个 sheet 的**原始数据存储区**在高编号列（通常 Column AE ~ EQ），前端区域（Column A ~ T）为计算和展示区域。

---

## 逐文件映射

---

### 1. `BS_Financials_5Q_TTM.py` — 季度资产负债表 (Quarterly Balance Sheet)

**Python 输出**：最近 5 季 Balance Sheet + TTM（最新季度）

| Sheet | 存储列区域 | VBA 偏移变量 | 列号范围 | 行数 |
|-------|-----------|-------------|---------|------|
| **台股** | Column AQ ~ AY | `dc2 + 11 = 43` | 第 43~51 列 | ~50 行财务项目 |
| **美股** | Column AN ~ AT | `dc1 + 9 = 40` | 第 40~46 列 | ~50 行 |
| **港股** | Column AM ~ AT | `dc2 = 39` | 第 39~46 列 | ~50 行 |
| **全球** | Column AT ~ AZ | `dc1 + 9 = 46` | 第 46~52 列 | ~50 行 |

**Excel 中的具体呈现**（以美股 sheet 为例）：

| 行位置 | 内容 |
|--------|------|
| Row 3 | 表头："Quarterly Balance Sheet" |
| Row 4 | 日期列标签（如 2024-09, 2024-06, 2024-03...） |
| Row 5+ | 项目行：Total Assets, Total Liabilities, Stockholders Equity, Cash And Cash Equivalents 等 |

**前端引用单元格**：
- `[K12]` — 引用股本数据（来自 Balance Sheet 的 Shares Outstanding）
- 各种 VLOOKUP 公式从原始数据区提取值到前端展示区

---

### 2. `BS_Financials_5Y_TTM.py` — 年度资产负债表 (Annual Balance Sheet)

**Python 输出**：最近 5 年 Balance Sheet + TTM

| Sheet | 存储列区域 | VBA 偏移变量 | 列号范围 |
|-------|-----------|-------------|---------|
| **台股** | Column BM ~ BU | `dc2 + 33 = 65` | 第 65~73 列 |
| **美股** | Column BF ~ BL | `dc1 + 27 = 58` | 第 58~64 列 |
| **港股** | Column BC ~ BI | `dc4 = 55` | 第 55~61 列 |
| **全球** | Column BL ~ BR | `dc1 + 27 = 64` | 第 64~70 列 |

**Excel 中的具体呈现**（以台股 sheet 为例）：

| 行位置 | 内容 |
|--------|------|
| Row 3 | 表头："合併資產負債表-年" |
| Row 4 | 年度列标签（如 2024, 2023, 2022, 2021, 2020） |
| Row 5+ | 项目行：流動資產、非流動資產、資產總額、流動負債、非流動負債、負債總額、權益總額 等 |

**前端引用**：
- 年度 ROE 计算：`ROE% = Net Income / Stockholders Equity`，从 BS 年度数据和 IS 年度数据交叉引用
- `[K12]` — 股本/流通股数

---

### 3. `CFS_Financials_5Y_TTM.py` — 年度现金流量表 (Annual Cash Flow Statement)

**Python 输出**：最近 5 年 Cash Flow + TTM（最近 4 季加总）

| Sheet | 存储列区域 | VBA 偏移变量 | 列号范围 |
|-------|-----------|-------------|---------|
| **美股** | Column CE ~ CK | `dccf = 83` | 第 83~89 列 |
| **全球** | Column CK ~ CQ | `dccf = 89` | 第 89~95 列 |
| **台股** | 无独立 CFS 区域 | — | 台股财报数据来自 MoneyDJ，CFS 嵌入其他区域 |
| **港股** | 无独立 CFS 区域 | — | 港股使用 AAStocks 数据 |

**Excel 中的具体呈现**（以美股 sheet 为例）：

| 行位置 | 内容 |
|--------|------|
| Row 3 | 表头："Annual Cash Flow" |
| Row 4 | 年度标签（2024, 2023, 2022, 2021, 2020） |
| Row 5+ | Operating Cash Flow, Capital Expenditure, Free Cash Flow, Investing Cash Flow, Financing Cash Flow 等 |

**前端引用**：
- 盈再率（Profit Reinvestment Rate）计算需要 Capital Expenditure 数据
- 常利（Recurring Profit）计算可能引用 Operating Cash Flow

---

### 4. `Company_Profile.py` — 公司基本资料

**Python 输出**：公司名、行业、市值、PE、EPS、Beta、52 周高低等

| 数据字段 | 台股 Cell | 美股 Cell | 港股 Cell | 全球 Cell |
|---------|----------|----------|----------|----------|
| **公司名称** | `A1` | `B1` | `A1` | `B1` |
| **行业/描述** | `Y16` | `Y16` | — | `AE16` |
| **市值 (Market Cap)** | `Y23` | `Y23` | — | `AE23` |
| **股价** | `Y24` | `Y24` | `Q6` | `AE24` |
| **币别** | — | `W25` / `Y25` | `Y24`~`Y25` | `AC25` / `AE25` |
| **日期** | `I1` | `I1` | `I1` | `I1` |

**原始数据存储区**：

| Sheet | 存储列区域 | VBA 偏移变量 | 列号 |
|-------|-----------|-------------|------|
| **台股** | Column BX ~ CF | `dc1 = 75` | 第 75~84 列 |
| **美股** | Column BX ~ CB | `dc6 = 76` | 第 76~80 列 |
| **港股** | Column BZ ~ CD | `dc6 = 78` | 第 78~82 列 |
| **全球** | Column CD ~ CH | `dc6 = 82` | 第 82~86 列 |

**VBA 原始数据源对照**：
- 台股：`pscnetinvest.moneydj.com.tw`
- 美股：`marketwatch.com` + `company-people` page
- 港股：`aastocks.com`
- 全球：`wsj.com` + `marketwatch.com`

---

### 5. `Company_Action.py` — 公司行动（股息 Dividends / 拆股 Stock Splits）

**Python 输出**：历史除息记录 + 拆股记录

| Sheet | 存储列区域 | VBA 偏移变量 | 说明 |
|-------|-----------|-------------|------|
| **台股** | Column CY ~ DE | `dc8 = 103` | 第 103~109 列，含除息日、配息金额 |
| **美股** | 附加在历史股价区域之后 | `dc5 = 67` (BO 列) | Split 和 Dividend 数据追加在价格数据下方 |
| **港股** | Column BL ~ BR | `dc5 = 64` | 第 64~70 列，包含现金股利、股票股利、特别股利 |
| **中股** | Column BW ~ CE | `dc13 = 75` | 第 75~83 列，分红送股数据 |

**前端引用单元格**（以台股 sheet 为例）：

| Cell | 内容 |
|------|------|
| `W32` | 除息日（Ex-dividend date） |
| `W33` | 除权日（Record date） |
| `X36` | 上市日（Listing date） |
| Column G (Row 7+) | EPS/配息历史显示在前端分析区 |

---

### 5b. `Historical_stock_price_adj.py` — 历史股价（含调整价与公司行动）

**Python 输出**：月频 OHLCV + Adj Close + Corporate Actions（Dividends / Stock Splits），合并为单一 DataFrame

| Sheet | 存储列区域 | 列号范围 | 说明 |
|-------|-----------|---------|------|
| **美股** | Column BO ~ BW | 第 67~75 列 | 月频重采样 + Corporate Actions 内嵌行 |

**与原有 5/9 Historical stock price 的区别**：

| 对比项 | 原版（fetch_historical_prices） | 新版（Historical_stock_price_adj） |
|--------|-------------------------------|----------------------------------|
| API 调用 | `ticker.history(period="max", interval="1mo")` | `ticker.history(start, end, auto_adjust=False)` |
| 月频处理 | 直接用 yfinance 的 `interval="1mo"` | 日线 → `resample('MS')` 手动聚合 |
| Adj Close | 无 | 有（`auto_adjust=False` 保留） |
| Corporate Actions | 分开存储（价格数据下方追加 Dividends/Splits） | 合并为同一 DataFrame（Entry_Type 标记） |
| 排序 | 升序（最早在上） | 降序（最新在上） |

**Excel 中的具体呈现**（美股 sheet）：

| 行位置 | 列范围 | 内容 |
|--------|--------|------|
| Row 1 | BO1 | 标题："5 / 9 Historical stock price (adj)" |
| Row 4 | BO4:BW4 | 列标题：Date, Open, High, Low, Close, Adj Close, Volume, Dividends, Stock Splits |
| Row 5+ | BO5:BW5+ | 数据行（Month_Start 行含 OHLCV，Corporate_Action 行含 Dividends/Stock Splits） |

---

### 6. `EPS_Earnings.py` — EPS / 盈余数据

**Python 输出**：年度 EPS、季度 EPS、TTM EPS

| Sheet | 存储列区域 | VBA 偏移变量 | 说明 |
|-------|-----------|-------------|------|
| **台股** | Column CI ~ CR | `dc6 = 87` | 第 87~96 列，含年度/季度 EPS 和配息率 |
| **美股** | 从 IS 数据区提取 | `dc1 = 31` 区域 | EPS 行嵌入在 Income Statement 内 |
| **港股** | 从 IS 数据区提取 | `dc1 = 31` 区域 | 同上 |
| **全球** | 从 IS 数据区提取 | `dc1 = 37` 区域 | 同上 |

**前端引用单元格**：

| Cell | 内容 | 所在 Sheet |
|------|------|-----------|
| `K3` | 公式引用 EPS（Row 14, Col Q 相关） | 所有 sheet |
| `K10` | 月数（12） | 所有 sheet |
| `K11` | EPS 相关公式 | 所有 sheet |
| Column E~F (Row 5~12) | 年度 ROE%、盈再%、常利、配息%、常 EPS | 前端分析区 |

**盈再表前端分析区 EPS 展示**（Row 5 ~ Row 12，Column A ~ I）：

| 列 | 字段 |
|----|------|
| A | 年度 |
| B | 还原股价 |
| C | ROE% |
| D | 盈再% (Profit Reinvestment Rate) |
| E | 常利$m (Recurring Profit, millions) |
| F | 配息% (Payout Ratio) |
| G | 常 EPS$ (Recurring EPS) |
| H | 股息$ (Dividend) |
| I | 股子 (Stock Dividend / Split) |

---

### 7. `Share_capital_data.py` — 股本/持股结构

**Python 输出**：流通股数、浮动股、内部人/机构持股比例、做空数据

| Sheet | 存储列区域 | VBA 偏移变量 | 说明 |
|-------|-----------|-------------|------|
| **台股** | Column CT ~ CX | `dc7 = 98` | 第 98~102 列，含董监持股、外资持股 |
| **美股** | 嵌入 Company Profile 区域 | `dc6 = 76` | Market Cap / Shares Outstanding 在 profile 数据中 |
| **港股** | Column CN ~ CQ | `dc8 = 92` | 第 92~95 列，公司基本面数据 |

**前端引用单元格**：

| Cell | 内容 | 说明 |
|------|------|------|
| `K12` | 股本 / 流通股数 | 用于计算 EPS = Net Income / Shares Outstanding |
| `Q14` (台股) | 市值 | Market Cap = Price × Shares |
| `Y23` | 市值 | 同上 |

---

### 8. `Foreign_exchange_rates.py` — 外汇汇率

**Python 输出**：USD/JPY 汇率（收盘、开盘、最高、最低）

| Sheet | 存储列区域 | VBA 偏移变量 | 前端 Cell |
|-------|-----------|-------------|----------|
| **美股** | Column CA | `dc7 = 79` (第 79 列) | `F15` |
| **港股** | Column CR | `dc9 = 96` (第 96 列) | `F15` |
| **全球** | Column CG | `dc7 = 85` (第 85 列) | `F15` |
| **中股** | 从 AAStocks 取得 | — | `F15` |
| **台股** | 不需要（本币计价） | — | — |

**前端引用**：
- `[F15]` — **所有非台股 sheet 的核心汇率单元格**
- 用于将外币计价的财务数据换算为台币
- 默认值为 1（如果是美股且以 USD 计价）
- VBA 原始数据源：`wise.com`（原 TransferWise）或 Yahoo Finance

---

### 9. `IS_Financials_5Y_plus_TTM_with_Unusual.py` — 年度利润表

**Python 输出**：最近 5 年 Income Statement + TTM

| Sheet | 存储列区域 | VBA 偏移变量 | 列号范围 |
|-------|-----------|-------------|---------|
| **台股** | Column BB ~ BJ | `dc2 + 22 = 54` | 第 54~62 列 |
| **美股** | Column AW ~ BC | `dc1 + 18 = 49` | 第 49~55 列 |
| **港股** | Column AU ~ BA | `dc3 = 47` | 第 47~53 列 |
| **全球** | Column BC ~ BI | `dc1 + 18 = 55` | 第 55~61 列 |
| **中股** | Column AE (行 58+) | `dc1 = 31` | 第 31 列起，IS+BS 混合存储 |

**Excel 中的具体呈现**（以美股 sheet 为例）：

| 行位置 | 内容 |
|--------|------|
| Row 3 | 表头："Annual Income Statement" |
| Row 4 | 年度标签（2024, 2023, 2022, 2021, 2020） |
| Row 5+ | Revenue, Cost of Revenue, Gross Profit, Operating Income, Net Income, EPS 等 |

**前端引用**：
- ROE% = Net Income (IS) / Stockholders Equity (BS)
- 常利$m = Recurring Profit（剔除 unusual items 后的经常性利润）
- 配息% = Dividends / Net Income
- 常 EPS = Recurring Profit / Shares Outstanding

---

### 10. `IS_Financials_5Q_plus_TTM_with_Unusual.py` — 季度利润表 + 异常项

**Python 输出**：最近 5 季 Income Statement + TTM（4 季加总）+ Unusual Items

| Sheet | 存储列区域 | VBA 偏移变量 | 列号范围 |
|-------|-----------|-------------|---------|
| **台股** | Column AF ~ AO | `dc2 = 32` | 第 32~41 列 |
| **美股** | Column AE ~ AK | `dc1 = 31` | 第 31~37 列 |
| **港股** | Column AE ~ AK | `dc1 = 31` | 第 31~37 列 |
| **全球** | Column AK ~ AQ | `dc1 = 37` | 第 37~43 列 |
| **中股** | Column AE ~ BD | `dc1 = 31` | 第 31~56 列（5 组季度累计） |

**Unusual Items 的特殊处理**：
- Python 脚本搜索行名包含 `unusual`、`special`、`restructuring`、`non recurring` 的项目
- 盈再表中对应的概念是**常利**（Recurring Profit）= 总利润 - 非经常性项目
- VBA 通过 `.Find()` 在原始数据中搜索类似关键词，将非经常性项目标记后从 EPS 计算中剔除

**前端引用**：
- `[E12]` — IFERROR 公式，引用季度数据计算最新 TTM 指标
- `[M1]` — 财报检查日期（最新季报日期）

---

## 汇总映射表

| Python 脚本 | 数据类型 | 台股 列区域 | 美股 列区域 | 港股 列区域 | 全球 列区域 | 中股 列区域 | 前端关键 Cell |
|------------|---------|-----------|-----------|-----------|-----------|-----------|-------------|
| `BS_5Q_TTM` | 季度 BS | AQ~AY (43) | AN~AT (40) | AM~AT (39) | AT~AZ (46) | — | K12 |
| `BS_5Y_TTM` | 年度 BS | BM~BU (65) | BF~BL (58) | BC~BI (55) | BL~BR (64) | AE 行 58+ | K12 |
| `CFS_5Y_TTM` | 年度 CFS | — | CE~CK (83) | — | CK~CQ (89) | — | 盈再率计算 |
| `Company_Profile` | 公司资料 | BX~CF (75) | BX~CB (76) | BZ~CD (78) | CD~CH (82) | CH (86) | A1/B1, Y16, Y23, Y24, I1 |
| `Company_Action` | 股息/拆股 | CY~DE (103) | BO 下方 (67) | BL~BR (64) | BU 下方 (73) | BW~CE (75) | W32, W33, H 列 |
| `EPS_Earnings` | EPS | CI~CR (87) | IS 内嵌 | IS 内嵌 | IS 内嵌 | IS 内嵌 | K3, K10, K11, G 列 |
| `Share_Capital` | 股本 | CT~CX (98) | Profile 内嵌 | CN~CQ (92) | Profile 内嵌 | — | K12, Q14 |
| `Forex_Rates` | 汇率 | — | CA (79) | CR (96) | CG (85) | — | **F15** |
| `IS_5Y_TTM` | 年度 IS | BB~BJ (54) | AW~BC (49) | AU~BA (47) | BC~BI (55) | AE (31) | ROE%, 常利, 常 EPS |
| `IS_5Q_TTM` | 季度 IS | AF~AO (32) | AE~AK (31) | AE~AK (31) | AK~AQ (37) | AE~BD (31) | E12, M1 |
| `Historical_stock_price_adj` | 历史股价(adj) | — | BO~BW (67) | — | — | — | BO1, BO4:BW4 |

> 括号内数字为 VBA 中的列号偏移变量值（`dc` 变量）

---

## 数据流向图

```
yfinance API
    │
    ├── ticker.quarterly_balance_sheet ──→ BS_5Q_TTM.py ──→ [台股]AQ / [美股]AN / [港股]AM / [全球]AT
    ├── ticker.balance_sheet ───────────→ BS_5Y_TTM.py ──→ [台股]BM / [美股]BF / [港股]BC / [全球]BL
    ├── ticker.cashflow ────────────────→ CFS_5Y_TTM.py ─→ [美股]CE / [全球]CK
    ├── ticker.info ────────────────────→ Company_Profile ─→ A1/B1, Y16, Y23, Y24, I1
    │                                  ├→ Share_Capital ──→ [台股]CT / K12
    │                                  └→ EPS (TTM) ─────→ K3, K10, K11
    ├── ticker.actions ─────────────────→ Company_Action ──→ [台股]CY / W32, W33
    ├── ticker.earnings ────────────────→ EPS_Earnings ───→ [台股]CI / G列
    ├── ticker.financials ──────────────→ IS_5Y_TTM.py ──→ [台股]BB / [美股]AW / [港股]AU / [全球]BC
    ├── ticker.quarterly_financials ────→ IS_5Q_TTM.py ──→ [台股]AF / [美股]AE / [港股]AE / [全球]AK
    ├── ticker.history(auto_adjust=False)─→ Historical_stock_price_adj ─→ [美股]BO (月频OHLCV+Adj Close+Actions)
    └── USD/XXX=X ──────────────────────→ Forex_Rates ───→ [美股]CA / [港股]CR / [全球]CG → F15
```

---

## 关键发现

### 1. 数据源差异
Python 脚本统一使用 **yfinance** 作为数据源，而盈再表 VBA 使用**多个网站**：
- 台股：MoneyDJ (`pscnetinvest.moneydj.com.tw`) + TWSE/TPEx
- 美股：MarketWatch + Yahoo Finance + GuruFocus + SeekingAlpha
- 港股：AAStocks (`aastocks.com`) + Yahoo Finance + Wise
- 中股：Sina Finance (`money.finance.sina.com.cn`) + AAStocks
- 全球：WSJ + MarketWatch + Yahoo Finance + Wise + GuruFocus

### 2. TTM 计算方式差异
- **Balance Sheet**：Python 和 Excel 一致 → 取最新时点数据
- **Income Statement / Cash Flow**：Python 用最近 4 季求和；VBA 的 TTM 处理方式依数据源不同

### 3. 台股 CFS 缺失
Python `CFS_Financials_5Y_TTM.py` 输出现金流量表，但盈再表**台股 sheet 没有独立的 CFS 存储区域**（MoneyDJ 数据中未单独抓取现金流量表）。美股和全球 sheet 有 CFS 区域。

### 4. 前端计算公式的核心数据来源

| 前端指标 | 需要的原始数据 | 来源 Python 脚本 |
|---------|-------------|-----------------|
| ROE% | Net Income (IS) + Equity (BS) | IS_5Y + BS_5Y |
| 盈再率% | CapEx (CFS) / Net Income (IS) | CFS_5Y + IS_5Y |
| 常利$m | Net Income - Unusual Items | IS_5Y (with Unusual) |
| 常 EPS$ | 常利 / Shares Outstanding | IS + Share_Capital |
| 配息% | Dividends / Net Income | Company_Action + IS |
| 预期报酬 | ROE × (1 - 盈再率) / Price | 综合计算 |
