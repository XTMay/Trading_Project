# AI 自動交易系統 — 技術架構與可行性報告

> **報告日期**：2026-03-26
> **專案基礎**：盈再表（Profit Reinvestment Rate Analysis Tool）
> **目標**：營再表 → TradingView → AI 分析 → Webhook → IBKR 自動下單

---

## 目錄

1. [專案現狀](#1-專案現狀)
2. [目標架構總覽](#2-目標架構總覽)
3. [模組一：營再表基本面引擎](#3-模組一營再表基本面引擎)
4. [模組二：TradingView 技術面信號](#4-模組二tradingview-技術面信號)
5. [模組三：AI 雙引擎分析](#5-模組三ai-雙引擎分析)
6. [模組四：信號融合決策層](#6-模組四信號融合決策層)
7. [模組五：風控引擎](#7-模組五風控引擎)
8. [模組六：Webhook 中間層](#8-模組六webhook-中間層)
9. [模組七：IBKR 自動交易引擎](#9-模組七ibkr-自動交易引擎)
10. [模組八：績效回饋與回測](#10-模組八績效回饋與回測)
11. [技術堆疊總覽](#11-技術堆疊總覽)
12. [可行性評估矩陣](#12-可行性評估矩陣)
13. [風險分析與對策](#13-風險分析與對策)
14. [分階段實施路線圖](#14-分階段實施路線圖)
15. [核心程式碼範例](#15-核心程式碼範例)
16. [結論與建議](#16-結論與建議)

---

## 1. 專案現狀

### 1.1 已有能力盤點

當前盈再表專案是一個**成熟的財務分析工具**，已具備以下能力：

```
已完成模組                                 程式碼規模
─────────────────────────────────────────────────────
generate_report_summary.py   統一報表生成    724 行
financial_calculator.py      核心財務計算    576 行
performance_calculator.py    XIRR/巴菲特距離 429 行
portfolio_manager.py         投資組合管理    479 行
tax_calculator.py            台灣股利稅計算  393 行
index_components.py          指數成分股      336 行
fmp_client.py                FMP API 備用源  584 行
12 個數據抓取腳本             各市場數據      ~2,000 行
VBA 模組 (Module1-9)         Excel 前端      12,850 行
─────────────────────────────────────────────────────
合計                                        ~18,000+ 行
```

### 1.2 已有 vs 缺失

```
✅ 已有                          ❌ 缺失（需新建）
─────────────────────────────    ─────────────────────────────
多市場數據抓取(美/台/港/中/全球)   TradingView Webhook 接收
20+ 財務指標計算                  AI 分析模組 (GPT/Claude)
ROE / 盈再率 / 價格區間           信號融合評分系統
投資組合管理與排名                風控引擎
XIRR / 巴菲特距離績效評估         IBKR API 交易對接
yfinance + FMP 雙數據源          交易伺服器 (FastAPI)
Excel VBA 前端                   績效回饋與回測系統
```

### 1.3 當前數據流

```
使用者在 Excel A2 輸入股票代碼
        │
        ▼
VBA Worksheet_Change 事件觸發
        │
        ▼
Shell 呼叫 Python 腳本
  python generate_report_summary.py AAPL
        │
        ▼
yfinance API 抓取 10+ 端點數據
  ├── 季度/年度損益表
  ├── 資產負債表
  ├── 現金流量表
  ├── 公司基本資料
  ├── 歷史股價（月度）
  └── 匯率數據
        │
        ▼
Python 計算財務指標
  ├── ROE%
  ├── 盈再率%
  ├── 經常性利潤
  ├── 預期報酬率
  └── 價格區間（便宜/合理/昂貴）
        │
        ▼
輸出 report_summary.xlsx
        │
        ▼
VBA 導入數據 → Excel 前端展示
```

---

## 2. 目標架構總覽

### 2.1 完整系統架構圖

```
┌──────────────────────────────────────────────────────────────────────┐
│                     AI 自動交易系統 — 完整架構                         │
└──────────────────────────────────────────────────────────────────────┘

  ┌─────────────┐    ┌─────────────────┐    ┌──────────────────────┐
  │  營再表       │    │  TradingView     │    │   市場情緒數據        │
  │  基本面引擎   │    │  技術面信號       │    │   (新聞/社群)         │
  │  (Python)    │    │  (Pine Script)   │    │   (選配)             │
  └──────┬───────┘    └────────┬─────────┘    └──────────┬───────────┘
         │                     │                         │
         │ JSON                │ Webhook POST            │ API
         ▼                     ▼                         ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │                  模組三：AI 雙引擎分析層                           │
  │                                                                  │
  │   ┌─────────────┐         ┌─────────────┐                       │
  │   │   GPT-4o    │         │  Claude API  │                       │
  │   │  基本面分析   │         │  技術面驗證   │                       │
  │   └──────┬──────┘         └──────┬──────┘                        │
  │          │                       │                               │
  │          └───────────┬───────────┘                               │
  │                      ▼                                           │
  │          ┌───────────────────────┐                               │
  │          │  結構化 JSON 輸出      │                               │
  │          │  {action, score, ...} │                               │
  │          └───────────┬───────────┘                               │
  └──────────────────────┼───────────────────────────────────────────┘
                         │
                         ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │              模組四：信號融合決策層（多因子評分）                     │
  │                                                                  │
  │   技術面分數 (40%)  +  基本面分數 (30%)  +  AI 判斷 (30%)         │
  │                                                                  │
  │   總分 > 75 → 產生交易信號    總分 ≤ 75 → 放棄                    │
  └──────────────────────┬───────────────────────────────────────────┘
                         │
                         ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │                   模組五：風控引擎                                 │
  │                                                                  │
  │   ☐ 單筆風險 ≤ 2% 帳戶淨值                                       │
  │   ☐ 最大持倉 ≤ 5 檔                                              │
  │   ☐ 停損 -5%（ATR 動態調整）                                      │
  │   ☐ 停利 +10~20%                                                 │
  │   ☐ 最大回撤 ≤ 15% → 觸發全面暫停                                 │
  │   ☐ 每日最大下單次數 ≤ 3                                          │
  │                                                                  │
  │   全部通過 → 放行    任一不通過 → 攔截                              │
  └──────────────────────┬───────────────────────────────────────────┘
                         │
                         ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │              模組六：交易伺服器（FastAPI）                          │
  │                                                                  │
  │   POST /webhook/tradingview   ← TradingView Alert               │
  │   POST /api/execute-trade     ← 內部觸發                         │
  │   GET  /api/portfolio         ← 查詢持倉                         │
  │   GET  /api/performance       ← 績效報告                         │
  └──────────────────────┬───────────────────────────────────────────┘
                         │
                         ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │              模組七：IBKR 交易引擎 (ib_insync)                    │
  │                                                                  │
  │   TWS Gateway / IB Gateway (需持續運行)                           │
  │   ├── 連線管理（心跳/重連/Session）                                │
  │   ├── 下單執行（限價單/市價單）                                     │
  │   ├── 持倉查詢                                                    │
  │   └── 訂單狀態追蹤                                                │
  └──────────────────────┬───────────────────────────────────────────┘
                         │
                         ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │              模組八：績效回饋系統                                   │
  │                                                                  │
  │   交易記錄 → PostgreSQL                                           │
  │   ├── 勝率統計                                                    │
  │   ├── Sharpe Ratio                                               │
  │   ├── 最大回撤                                                    │
  │   └── 因子權重優化建議                                             │
  └──────────────────────────────────────────────────────────────────┘
```

---

## 3. 模組一：營再表基本面引擎

### 3.1 現有能力（可直接複用）

當前 `financial_calculator.py` 已實現完整的基本面評分能力：

```python
# 已有的核心計算函數 (financial_calculator.py)

calc_roe(is_y, bs_y)              # ROE% = 淨利 / 股東權益
calc_reinvestment_rate(is_y, cfs) # 盈再率% = 資本支出 / 淨利
calc_recurring_profit(is_y)       # 經常性利潤 = 淨利 - 非經常性項目
calc_expected_return(roe, reinv)  # 預期報酬 = ROE × (1 - 盈再率)
calc_price_zones(eps, exp_ret)    # 價格區間：便宜/合理/昂貴
```

### 3.2 需要改造的部分

將 Excel 輸出改為 **標準化 JSON API**，供下游模組調用：

```python
# 改造目標：輸出標準化評分 JSON
{
    "symbol": "AAPL",
    "timestamp": "2026-03-26T10:30:00Z",
    "fundamental_score": {
        "roe_5y_avg": 147.2,           # ROE 5年平均
        "reinvestment_rate": 22.1,      # 盈再率
        "recurring_profit_growth": 8.5, # 經常性利潤成長率
        "expected_return": 11.5,        # 預期報酬率
        "price_zone": "fair",           # 當前價格區間
        "price_vs_cheap": 1.15,         # 當前價 / 便宜價
        "buffett_distance": 0.82,       # 巴菲特距離
        "composite_score": 72           # 基本面綜合分 (0-100)
    }
}
```

### 3.3 評分規則設計

```
基本面評分 (0-100) 計算方式：

指標                    權重    滿分條件              零分條件
──────────────────────────────────────────────────────────
ROE 5年平均             25%    ≥ 20%                 ≤ 5%
盈再率                  15%    ≤ 30%（低資本支出）     ≥ 80%
經常性利潤成長率         20%    ≥ 15%                 ≤ 0%
預期報酬率              20%    ≥ 15%                 ≤ 5%
價格區間                20%    便宜價以下             昂貴價以上
──────────────────────────────────────────────────────────
```

---

## 4. 模組二：TradingView 技術面信號

### 4.1 技術指標配置

在 TradingView 中設定三個核心指標：

```
指標                配置參數                    買入信號               賣出信號
──────────────────────────────────────────────────────────────────────────
一目均衡表          轉換線=9, 基準線=26,        價格在雲帶上方          價格跌破雲帶
(Ichimoku)         先行帶=52                   轉換線 > 基準線         轉換線 < 基準線

RSI                 週期=14                     RSI 30-50（超賣回升）   RSI > 70（超買）

MACD               快線=12, 慢線=26, 信號=9     MACD 金叉              MACD 死叉
                                               (MACD線上穿信號線)      (MACD線下穿信號線)
```

### 4.2 TradingView Alert → Webhook

TradingView Pine Script 策略範例：

```pine
//@version=5
indicator("AI Trading Signal", overlay=true)

// === Ichimoku ===
conversionLine = ta.sma(close, 9)
baseLine       = ta.sma(close, 26)
leadLine1      = math.avg(conversionLine, baseLine)
leadLine2      = ta.sma(close, 52)
aboveCloud     = close > math.max(leadLine1[26], leadLine2[26])
belowCloud     = close < math.min(leadLine1[26], leadLine2[26])

// === RSI ===
rsiVal = ta.rsi(close, 14)

// === MACD ===
[macdLine, signalLine, _] = ta.macd(close, 12, 26, 9)
macdCross = ta.crossover(macdLine, signalLine)
macdDeathCross = ta.crossunder(macdLine, signalLine)

// === 綜合信號 ===
buySignal  = aboveCloud and rsiVal < 50 and macdCross
sellSignal = belowCloud and rsiVal > 70 and macdDeathCross

// === Alert ===
if buySignal
    alert('{"action":"BUY","symbol":"' + syminfo.ticker + '","price":' +
          str.tostring(close) + ',"rsi":' + str.tostring(rsiVal) +
          ',"ichimoku":"above_cloud","macd":"golden_cross"}', alert.freq_once_per_bar)

if sellSignal
    alert('{"action":"SELL","symbol":"' + syminfo.ticker + '","price":' +
          str.tostring(close) + ',"rsi":' + str.tostring(rsiVal) +
          ',"ichimoku":"below_cloud","macd":"death_cross"}', alert.freq_once_per_bar)
```

### 4.3 技術面評分設計

```
技術面評分 (0-100) 計算方式：

條件                          分數
──────────────────────────────────
Ichimoku 雲帶上方              +30
轉換線 > 基準線                +10
RSI 30-50（超賣回升區）         +20
RSI 50-70（正常區）            +10
MACD 金叉                     +20
MACD 柱狀體放大                +10
──────────────────────────────────
最高                          100
```

### 4.4 前置條件

```
需求                            說明
─────────────────────────────────────────────────
TradingView Pro 帳號            免費版不支援 Webhook Alert
Webhook URL                    指向你的交易伺服器
持續運行的伺服器                 接收 TradingView 推送（可用雲端 VPS）
```

---

## 5. 模組三：AI 雙引擎分析

### 5.1 架構設計

```
                    輸入數據打包
                    ├── 基本面數據 (來自模組一)
                    ├── 技術面信號 (來自模組二)
                    └── 近期新聞摘要 (選配)
                         │
              ┌──────────┴──────────┐
              ▼                     ▼
      ┌──────────────┐     ┌──────────────┐
      │    GPT-4o    │     │  Claude API  │
      │              │     │              │
      │  擅長：       │     │  擅長：       │
      │  市場情緒解讀  │     │  邏輯推理驗證  │
      │  新聞影響評估  │     │  數據一致性    │
      │  創意策略建議  │     │  風險識別      │
      └──────┬───────┘     └──────┬───────┘
             │                     │
             ▼                     ▼
      ┌──────────────┐     ┌──────────────┐
      │  JSON 輸出    │     │  JSON 輸出    │
      │  {            │     │  {            │
      │   action,     │     │   action,     │
      │   confidence, │     │   confidence, │
      │   reasoning   │     │   reasoning   │
      │  }            │     │  }            │
      └──────┬───────┘     └──────┬───────┘
             │                     │
             └──────────┬──────────┘
                        ▼
                 取兩者共識或加權平均
```

### 5.2 結構化 Prompt 設計

```python
ANALYSIS_PROMPT = """
你是一位專業的量化交易分析師。根據以下數據，給出交易建議。

## 基本面數據
- 股票：{symbol}
- ROE (5年平均)：{roe_5y}%
- 盈再率：{reinvestment_rate}%
- 預期報酬率：{expected_return}%
- 當前價格區間：{price_zone}
- 巴菲特距離：{buffett_distance}

## 技術面信號
- Ichimoku：{ichimoku_status}
- RSI (14)：{rsi_value}
- MACD：{macd_status}

## 要求
請以以下 JSON 格式回答，不要添加任何其他文字：
{
    "action": "BUY" | "SELL" | "HOLD",
    "confidence": 0-100,
    "target_price": float,
    "stop_loss": float,
    "reasoning": "一句話原因",
    "risk_factors": ["風險1", "風險2"]
}
"""
```

### 5.3 雙引擎共識機制

```
GPT 判斷     Claude 判斷     最終決策
─────────────────────────────────────
BUY          BUY             BUY（高信心）
BUY          HOLD            HOLD（觀望）
BUY          SELL            HOLD（矛盾，不動作）
SELL         SELL            SELL（高信心）
SELL         HOLD            HOLD（觀望）
HOLD         HOLD            HOLD

規則：兩個 AI 必須一致才執行交易，否則 HOLD。
這是最保守也最安全的策略。
```

### 5.4 API 成本估算

```
場景                       單次成本          月成本（假設每日 10 支股票）
───────────────────────────────────────────────────────────────────
GPT-4o (input+output)      ~$0.03           ~$9
Claude Sonnet              ~$0.02           ~$6
合計                                        ~$15/月

結論：成本可控，不會是瓶頸。
```

---

## 6. 模組四：信號融合決策層

### 6.1 多因子評分公式

```
最終分數 = 技術面分數 × 0.40
         + 基本面分數 × 0.30
         + AI 判斷分數 × 0.30

交易決策：
  分數 > 75 → 產生 BUY 信號
  分數 < 25 → 產生 SELL 信號
  其他      → HOLD（不動作）
```

### 6.2 實作範例

```python
def calculate_composite_score(technical, fundamental, ai_scores):
    """
    多因子信號融合。

    Args:
        technical:   dict  {"score": 0-100, "details": {...}}
        fundamental: dict  {"score": 0-100, "details": {...}}
        ai_scores:   dict  {"gpt": {"action": "BUY", "confidence": 80},
                            "claude": {"action": "BUY", "confidence": 75}}
    Returns:
        dict {"composite_score": float, "action": str, "breakdown": dict}
    """
    WEIGHTS = {"technical": 0.40, "fundamental": 0.30, "ai": 0.30}

    # AI 分數：兩個引擎的平均信心度
    ai_avg = (ai_scores["gpt"]["confidence"] + ai_scores["claude"]["confidence"]) / 2

    # AI 共識檢查：不一致則信心度減半
    if ai_scores["gpt"]["action"] != ai_scores["claude"]["action"]:
        ai_avg *= 0.5

    composite = (
        technical["score"]   * WEIGHTS["technical"]
        + fundamental["score"] * WEIGHTS["fundamental"]
        + ai_avg               * WEIGHTS["ai"]
    )

    if composite > 75:
        action = "BUY"
    elif composite < 25:
        action = "SELL"
    else:
        action = "HOLD"

    return {
        "composite_score": round(composite, 1),
        "action": action,
        "breakdown": {
            "technical": round(technical["score"] * WEIGHTS["technical"], 1),
            "fundamental": round(fundamental["score"] * WEIGHTS["fundamental"], 1),
            "ai": round(ai_avg * WEIGHTS["ai"], 1),
        },
    }
```

### 6.3 假信號過濾規則

```
過濾條件（任一觸發則放棄交易）：

1. 成交量過低     → 日均量 < 50萬股（美股）
2. 財報發布前     → 財報日前後 3 天不交易
3. 極端波動       → 當日振幅 > 5% 不追進
4. 連續信號       → 同一支股票 7天內不重複下單
5. 市場環境       → VIX > 30 時暫停所有買入
```

---

## 7. 模組五：風控引擎

### 7.1 風控參數表

```
參數                    設定值              說明
──────────────────────────────────────────────────────────
單筆最大風險            2% 帳戶淨值          虧損上限
最大同時持倉            5 檔                分散風險
單筆停損                -5% 或 2×ATR        取較大者
單筆停利                +15% 或 3×ATR       取較大者
帳戶最大回撤            -15%                觸發全面暫停
每日最大下單次數         3 次                防止過度交易
單一產業最大佔比         30%                 產業分散
保證金使用率上限         60%                 防止追繳
──────────────────────────────────────────────────────────
```

### 7.2 風控檢查流程

```python
def risk_check(signal, portfolio, account):
    """
    風控閘門：所有條件通過才放行。

    Returns:
        (bool, str) — (是否通過, 原因)
    """
    checks = []

    # 1. 帳戶回撤檢查
    drawdown = (account["peak_value"] - account["current_value"]) / account["peak_value"]
    if drawdown > 0.15:
        return False, f"帳戶回撤 {drawdown:.1%} 超過 15%，全面暫停交易"

    # 2. 持倉數量
    if len(portfolio["positions"]) >= 5 and signal["action"] == "BUY":
        return False, "持倉已達 5 檔上限"

    # 3. 單筆風險
    risk_amount = signal["shares"] * signal["price"] * 0.05  # 假設 5% 停損
    max_risk = account["current_value"] * 0.02
    if risk_amount > max_risk:
        return False, f"單筆風險 ${risk_amount:.0f} 超過帳戶 2% (${max_risk:.0f})"

    # 4. 每日下單次數
    if account["today_orders"] >= 3:
        return False, "今日已下單 3 次，達到上限"

    # 5. 重複下單
    if signal["symbol"] in [p["symbol"] for p in portfolio["positions"]]:
        if signal["action"] == "BUY":
            return False, f"{signal['symbol']} 已在持倉中，不重複買入"

    # 6. 產業集中度
    sector = signal.get("sector", "Unknown")
    sector_exposure = sum(
        p["market_value"] for p in portfolio["positions"]
        if p.get("sector") == sector
    )
    if sector_exposure / account["current_value"] > 0.30:
        return False, f"{sector} 產業佔比已超過 30%"

    return True, "風控檢查全部通過"
```

---

## 8. 模組六：Webhook 中間層

### 8.1 FastAPI 伺服器架構

```python
# server.py — 交易伺服器入口
from fastapi import FastAPI, Request, HTTPException
from pydantic import BaseModel
import logging

app = FastAPI(title="AI Trading Server")
logger = logging.getLogger("trading")

class TradingViewAlert(BaseModel):
    action: str        # "BUY" / "SELL"
    symbol: str
    price: float
    rsi: float | None = None
    ichimoku: str | None = None
    macd: str | None = None

@app.post("/webhook/tradingview")
async def receive_tradingview_alert(alert: TradingViewAlert, request: Request):
    """
    接收 TradingView Webhook 推送。
    流程：技術面信號 → 基本面查詢 → AI 分析 → 信號融合 → 風控 → 下單
    """
    logger.info(f"收到 TradingView 信號: {alert.symbol} {alert.action} @ {alert.price}")

    # Step 1: 計算技術面分數
    tech_score = calculate_technical_score(alert)

    # Step 2: 查詢基本面分數（從營再表引擎）
    fund_score = get_fundamental_score(alert.symbol)

    # Step 3: AI 雙引擎分析
    ai_scores = await run_ai_analysis(alert.symbol, tech_score, fund_score)

    # Step 4: 信號融合
    decision = calculate_composite_score(tech_score, fund_score, ai_scores)

    # Step 5: 風控檢查
    passed, reason = risk_check(decision, get_portfolio(), get_account())

    if not passed:
        logger.warning(f"風控攔截: {reason}")
        return {"status": "blocked", "reason": reason}

    # Step 6: 執行交易
    if decision["action"] in ("BUY", "SELL"):
        order = await execute_trade(decision)
        logger.info(f"下單成功: {order}")
        return {"status": "executed", "order": order}

    return {"status": "hold", "score": decision["composite_score"]}

@app.get("/api/portfolio")
async def get_portfolio_api():
    """查詢當前持倉。"""
    return get_portfolio()

@app.get("/api/performance")
async def get_performance_api():
    """查詢績效報告。"""
    return get_performance_report()
```

### 8.2 API 端點一覽

```
方法     路徑                         說明                觸發方式
────────────────────────────────────────────────────────────────────
POST    /webhook/tradingview         接收 TV 信號         TradingView Alert
POST    /api/manual-trade            手動觸發交易          前端/API
GET     /api/portfolio               查詢持倉              前端/API
GET     /api/performance             績效報告              前端/API
GET     /api/signals/history         歷史信號記錄          前端/API
POST    /api/risk/override           手動覆蓋風控（需密碼） 緊急情況
GET     /api/health                  健康檢查              監控系統
```

---

## 9. 模組七：IBKR 自動交易引擎

### 9.1 架構

```
交易伺服器 (FastAPI)
        │
        ▼
  ┌──────────────┐
  │  ib_insync   │  Python 庫，封裝 IBKR TWS API
  └──────┬───────┘
         │  Socket 連線 (port 7497: TWS / 4001: IB Gateway)
         ▼
  ┌──────────────┐
  │  TWS Gateway │  IBKR 客戶端，必須持續運行
  │  或           │  建議用 IB Gateway（無 GUI，更穩定）
  │  IB Gateway  │
  └──────┬───────┘
         │  IBKR 內部網路
         ▼
  ┌──────────────┐
  │  IBKR 交易所  │  實際執行下單
  └──────────────┘
```

### 9.2 核心交易程式碼

```python
from ib_insync import IB, Stock, MarketOrder, LimitOrder
import asyncio

class IBKRTrader:
    def __init__(self, host="127.0.0.1", port=7497, client_id=1):
        self.ib = IB()
        self.host = host
        self.port = port       # 7497=TWS, 4001=IB Gateway
        self.client_id = client_id

    async def connect(self):
        """連線到 TWS/IB Gateway。"""
        await self.ib.connectAsync(self.host, self.port, clientId=self.client_id)

    def get_portfolio(self):
        """取得當前持倉。"""
        positions = self.ib.positions()
        return [
            {
                "symbol": p.contract.symbol,
                "shares": p.position,
                "avg_cost": p.avgCost,
                "market_value": p.marketValue,
            }
            for p in positions
        ]

    def place_order(self, symbol, action, quantity, order_type="LMT", limit_price=None):
        """
        下單。

        Args:
            symbol:      股票代碼 (e.g., "AAPL")
            action:      "BUY" / "SELL"
            quantity:     股數
            order_type:  "MKT" (市價) / "LMT" (限價)
            limit_price: 限價單價格
        """
        contract = Stock(symbol, "SMART", "USD")
        self.ib.qualifyContracts(contract)

        if order_type == "MKT":
            order = MarketOrder(action, quantity)
        else:
            order = LimitOrder(action, quantity, limit_price)

        trade = self.ib.placeOrder(contract, order)
        return {
            "order_id": trade.order.orderId,
            "symbol": symbol,
            "action": action,
            "quantity": quantity,
            "status": trade.orderStatus.status,
        }

    def cancel_order(self, order_id):
        """取消訂單。"""
        for trade in self.ib.openTrades():
            if trade.order.orderId == order_id:
                self.ib.cancelOrder(trade.order)
                return True
        return False

    def disconnect(self):
        """斷開連線。"""
        self.ib.disconnect()
```

### 9.3 關鍵注意事項

```
注意事項                    說明                              嚴重程度
──────────────────────────────────────────────────────────────────────
TWS/Gateway 必須運行       沒有它就無法下單，需 24h 保持運行     致命
Paper Trading 先行         先用模擬帳戶測試至少 3 個月           致命
PDT 規則                  美股帳戶 < $25K 有日內交易限制        高
Session 超時               TWS 每天會自動重啟，需處理重連        高
2FA 認證                  首次登入需要手機驗證                  中
API 速率限制               IBKR 限制每秒請求數                  中
時區處理                   美股交易時間 EST，需正確換算           中
```

---

## 10. 模組八：績效回饋與回測

### 10.1 交易記錄結構

```sql
-- PostgreSQL 交易記錄表
CREATE TABLE trades (
    id              SERIAL PRIMARY KEY,
    symbol          VARCHAR(20) NOT NULL,
    action          VARCHAR(4)  NOT NULL,    -- BUY / SELL
    quantity        INTEGER     NOT NULL,
    price           DECIMAL(12,4) NOT NULL,
    timestamp       TIMESTAMPTZ NOT NULL DEFAULT NOW(),

    -- 信號來源
    technical_score  DECIMAL(5,1),
    fundamental_score DECIMAL(5,1),
    ai_gpt_score     DECIMAL(5,1),
    ai_claude_score  DECIMAL(5,1),
    composite_score  DECIMAL(5,1),

    -- 結果（平倉後填入）
    exit_price       DECIMAL(12,4),
    exit_timestamp   TIMESTAMPTZ,
    pnl              DECIMAL(12,2),          -- 盈虧金額
    pnl_pct          DECIMAL(6,2),           -- 盈虧百分比
    hold_days        INTEGER
);

-- 每日帳戶快照
CREATE TABLE account_snapshots (
    date             DATE PRIMARY KEY,
    total_value      DECIMAL(14,2),
    cash             DECIMAL(14,2),
    positions_value  DECIMAL(14,2),
    daily_pnl        DECIMAL(12,2),
    drawdown_pct     DECIMAL(6,2)
);
```

### 10.2 關鍵績效指標 (KPI)

```
指標                    公式                            目標值
──────────────────────────────────────────────────────────────
勝率                    獲利交易數 / 總交易數             > 55%
盈虧比                  平均獲利 / 平均虧損               > 2.0
Sharpe Ratio           (年化報酬 - 無風險利率) / 標準差    > 1.5
最大回撤                最高點到最低點的最大跌幅           < 15%
年化報酬率              複合年增長率 (CAGR)               > 15%
平均持倉天數            持倉總天數 / 交易次數              參考值
```

### 10.3 回測框架

```python
# 使用 backtrader 進行歷史回測
import backtrader as bt

class AITradingStrategy(bt.Strategy):
    """基於多因子評分的回測策略。"""

    params = (
        ("tech_weight", 0.40),
        ("fund_weight", 0.30),
        ("ai_weight",   0.30),
        ("buy_threshold",  75),
        ("sell_threshold", 25),
        ("stop_loss_pct",  0.05),
        ("take_profit_pct", 0.15),
    )

    def next(self):
        # 在歷史數據上模擬信號融合邏輯
        score = self.calculate_historical_score()
        if score > self.params.buy_threshold and not self.position:
            self.buy()
        elif score < self.params.sell_threshold and self.position:
            self.sell()
```

---

## 11. 技術堆疊總覽

```
層級              技術選型              說明
──────────────────────────────────────────────────────────────────
數據層            yfinance + FMP API    已有，財務數據抓取
基本面分析        Python (現有程式碼)    已有，financial_calculator.py
技術面信號        TradingView Pro       Pine Script 策略 + Alert
AI 分析          OpenAI API + Claude API  GPT-4o + Claude Sonnet
交易伺服器        FastAPI (Python)       Webhook 接收 + API 服務
交易執行          ib_insync (Python)     IBKR TWS API 封裝
資料庫            PostgreSQL             交易記錄 + 績效數據
任務佇列          (選配) Celery + Redis  非同步任務處理
前端監控          (選配) Streamlit       即時儀表板
部署              VPS (AWS/DigitalOcean) 需 24h 運行
版本控制          Git + GitHub           已有
──────────────────────────────────────────────────────────────────

Python 套件依賴：
  已有：yfinance, pandas, numpy, openpyxl, scipy, requests
  新增：fastapi, uvicorn, ib_insync, openai, anthropic,
        psycopg2, backtrader, pandas-ta
```

---

## 12. 可行性評估矩陣

```
模組                    技術可行性   複雜度   現有基礎    優先級    狀態
──────────────────────────────────────────────────────────────────────
① 營再表 API 化          ★★★★★      低       已有全部    P0       可直接改造
② TradingView 信號       ★★★★★      低       無         P0       成熟方案
③ AI 分析模組            ★★★★☆      中       無         P1       API 調用即可
④ 信號融合決策            ★★★★★      低       無         P1       純邏輯計算
⑤ 風控引擎               ★★★★★      中       無         P0       必須最先做
⑥ Webhook 伺服器         ★★★★★      低       無         P0       FastAPI 標準
⑦ IBKR 交易引擎          ★★★☆☆      高       無         P1       最大技術挑戰
⑧ 績效回饋/回測           ★★★★☆      中       部分有     P2       可後期加入
──────────────────────────────────────────────────────────────────────

★★★★★ = 完全可行    ★★★☆☆ = 可行但有挑戰    ★☆☆☆☆ = 困難
```

---

## 13. 風險分析與對策

### 13.1 風險矩陣

```
風險等級     風險描述                          對策
─────────────────────────────────────────────────────────────────────
🔴 致命     無風控直接上實盤                    必須先完成風控模組 + Paper Trading
🔴 致命     IBKR 連線中斷時掛單未撤             心跳檢測 + 異常自動撤單
🟠 高       AI 幻覺產生錯誤信號                 雙引擎共識機制 + 信心度閾值
🟠 高       TWS Gateway 每日自動重啟             自動重連邏輯 + 異常通知
🟡 中       API 費用失控                        設定月度上限 + 用量監控
🟡 中       數據延遲導致錯誤決策                 多數據源交叉驗證
🟢 低       TradingView Webhook 偶發丟失         重試機制 + 本地信號備份
🟢 低       網路波動                            VPS 部署 + 備用方案
```

### 13.2 安全紅線（不可違反）

```
1. 絕不在未經 Paper Trading 驗證前投入真金白銀
2. 絕不關閉風控模組
3. 絕不單靠 AI 判斷下單（AI 只是輔助，不是決策者）
4. 絕不使用超過帳戶 60% 的資金
5. 最大回撤達 15% 時必須全面暫停，人工介入
```

---

## 14. 分階段實施路線圖

```
Phase 1 — 基礎搭建（最小可行產品）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ✦ 營再表 Python API 化（JSON 輸出）
  ✦ FastAPI 伺服器搭建
  ✦ TradingView Alert → Webhook 接收（只記錄，不下單）
  ✦ IBKR Paper Trading 帳戶連接測試
  ✦ 交付物：能接收信號並記錄到日誌
  │
  ▼
Phase 2 — 核心功能（半自動）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ✦ AI 分析模組（GPT + Claude）
  ✦ 信號融合評分系統
  ✦ 風控引擎
  ✦ IBKR Paper Trading 自動下單
  ✦ 交付物：Paper Trading 環境完整自動化
  │
  ▼
Phase 3 — 實盤試跑（小資金）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ✦ Paper Trading 驗證至少 3 個月
  ✦ 績效達標後切換實盤（小資金）
  ✦ 即時監控儀表板
  ✦ 異常告警（Email/Telegram）
  ✦ 交付物：小資金實盤運行
  │
  ▼
Phase 4 — 優化迭代
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ✦ 回測系統（backtrader）
  ✦ 因子權重動態調整
  ✦ 多市場支持（台股 + 美股）
  ✦ 績效反饋循環
  ✦ 交付物：持續優化的成熟系統
```

---

## 15. 核心程式碼範例

### 15.1 專案目錄結構（建議）

```
ai-trading-system/
├── server/
│   ├── main.py                  # FastAPI 入口
│   ├── routers/
│   │   ├── webhook.py           # TradingView Webhook 端點
│   │   ├── portfolio.py         # 持倉查詢 API
│   │   └── performance.py       # 績效查詢 API
│   ├── services/
│   │   ├── fundamental.py       # 基本面分析（複用營再表）
│   │   ├── technical.py         # 技術面評分
│   │   ├── ai_analyzer.py       # GPT + Claude 分析
│   │   ├── signal_fusion.py     # 信號融合決策
│   │   ├── risk_manager.py      # 風控引擎
│   │   └── ibkr_trader.py       # IBKR 交易執行
│   ├── models/
│   │   ├── schemas.py           # Pydantic 數據模型
│   │   └── database.py          # 資料庫連線
│   └── config.py                # 設定檔
├── backtest/
│   ├── strategy.py              # 回測策略
│   └── run_backtest.py          # 回測執行
├── ez_table/                    # 營再表核心（現有程式碼）
│   ├── generate_report_summary.py
│   ├── financial_calculator.py
│   ├── performance_calculator.py
│   └── ...
├── tests/
│   ├── test_risk_manager.py
│   ├── test_signal_fusion.py
│   └── test_ibkr_trader.py
├── requirements.txt
├── docker-compose.yml           # PostgreSQL + Server
└── README.md
```

### 15.2 完整交易流程程式碼（簡化版）

```python
# server/services/trading_pipeline.py
"""完整交易流程：從信號到下單。"""

import asyncio
import logging
from datetime import datetime

from .fundamental import get_fundamental_score
from .technical import calculate_technical_score
from .ai_analyzer import analyze_with_gpt, analyze_with_claude
from .signal_fusion import calculate_composite_score
from .risk_manager import risk_check
from .ibkr_trader import IBKRTrader

logger = logging.getLogger("trading_pipeline")

async def process_trading_signal(alert: dict) -> dict:
    """
    完整交易流程。

    Input: TradingView alert JSON
    Output: 交易結果或攔截原因
    """
    symbol = alert["symbol"]
    logger.info(f"=== 開始處理 {symbol} 信號 ===")

    # ── Step 1: 技術面評分 ──
    tech = calculate_technical_score(alert)
    logger.info(f"技術面分數: {tech['score']}")

    # ── Step 2: 基本面評分（從營再表引擎） ──
    fund = get_fundamental_score(symbol)
    logger.info(f"基本面分數: {fund['score']}")

    # ── Step 3: AI 雙引擎分析（並行） ──
    gpt_result, claude_result = await asyncio.gather(
        analyze_with_gpt(symbol, tech, fund),
        analyze_with_claude(symbol, tech, fund),
    )
    ai_scores = {"gpt": gpt_result, "claude": claude_result}
    logger.info(f"GPT: {gpt_result['action']} ({gpt_result['confidence']})")
    logger.info(f"Claude: {claude_result['action']} ({claude_result['confidence']})")

    # ── Step 4: 信號融合 ──
    decision = calculate_composite_score(tech, fund, ai_scores)
    logger.info(f"綜合分數: {decision['composite_score']} → {decision['action']}")

    if decision["action"] == "HOLD":
        return {"status": "hold", "score": decision["composite_score"]}

    # ── Step 5: 風控檢查 ──
    passed, reason = risk_check(decision)
    if not passed:
        logger.warning(f"風控攔截: {reason}")
        return {"status": "blocked", "reason": reason}

    # ── Step 6: 計算下單參數 ──
    quantity = calculate_position_size(decision, alert["price"])

    # ── Step 7: 執行下單 ──
    trader = IBKRTrader()
    await trader.connect()
    order_result = trader.place_order(
        symbol=symbol,
        action=decision["action"],
        quantity=quantity,
        order_type="LMT",
        limit_price=alert["price"],
    )
    trader.disconnect()

    logger.info(f"下單完成: {order_result}")

    # ── Step 8: 記錄交易 ──
    save_trade_record(
        symbol=symbol,
        action=decision["action"],
        quantity=quantity,
        price=alert["price"],
        scores=decision["breakdown"],
        timestamp=datetime.utcnow(),
    )

    return {"status": "executed", "order": order_result, "score": decision["composite_score"]}
```

---

## 16. 結論與建議

### 16.1 總體可行性判定

```
結論：技術上完全可行，但需分階段實施、嚴格測試。

優勢：
  ✅ 營再表已有完整基本面分析引擎（~3,000 行 Python）
  ✅ TradingView Webhook 是成熟方案
  ✅ GPT/Claude API 調用簡單，成本可控（~$15/月）
  ✅ ib_insync 大幅降低 IBKR 對接難度

挑戰：
  ⚠️ IBKR 連線穩定性需要大量測試
  ⚠️ 風控邏輯必須嚴謹，一個漏洞可能造成重大損失
  ⚠️ AI 不可作為唯一決策者
```

### 16.2 給學生的三條核心建議

```
1. 先跑 Paper Trading，至少 3 個月
   不要急著投入真金白銀。系統穩定運行 + 績效達標後再切換實盤。

2. 風控是系統的生命線
   沒有風控的自動交易 = 定時炸彈。風控模組必須最先完成、最後關閉。

3. AI 是參謀，不是將軍
   GPT/Claude 擅長綜合分析，但會產生幻覺。
   永遠用多因子評分機制過濾，不要讓 AI 單獨做決策。
```

---

> **報告生成工具**：Claude Code (claude-opus-4-6)
> **基於專案**：盈再表 Trading Project
> **下一步**：選擇 Phase 1 開始實施，或選擇特定模組深入開發
