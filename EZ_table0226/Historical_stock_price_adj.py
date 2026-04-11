import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import os


# ========= 1️⃣ 單一股票輸入 =========
ticker = "DIOD"   # 🔥 在這裡改股票代號

start_date = "2014-01-01"
end_date = (datetime.today() + timedelta(days=1)).strftime("%Y-%m-%d")


print(f"正在處理 {ticker}...")

stock = yf.Ticker(ticker)

# ========= 2️⃣ 抓日線資料 =========
df = stock.history(start=start_date, end=end_date, auto_adjust=False)

if df.empty:
    raise ValueError("⚠ 無資料")

# ========= 3️⃣ 每月1日基準 =========
monthly_df = df.resample('MS').agg({
    'Open': 'first',
    'High': 'max',
    'Low': 'min',
    'Close': 'last',
    'Adj Close': 'last',
    'Volume': 'sum'
})

# ========= 4️⃣ 取得 Corporate Actions =========
actions = stock.actions

if not actions.empty:
    actions = actions.loc[start_date:end_date]
    actions = actions[
        (actions['Dividends'] != 0) |
        (actions['Stock Splits'] != 0)
    ]
else:
    actions = pd.DataFrame()

# ========= 5️⃣ 合併 =========
combined = pd.concat([monthly_df, actions], axis=0)

# ========= 6️⃣ 確保欄位存在 =========
for col in ['Dividends', 'Stock Splits']:
    if col not in combined.columns:
        combined[col] = 0
    else:
        combined[col] = combined[col].fillna(0)

# ========= 7️⃣ 標記類型 =========
combined['Entry_Type'] = "Month_Start"

if not actions.empty:
    combined.loc[actions.index, 'Entry_Type'] = "Corporate_Action"

# ========= 8️⃣ 最新排最上 =========
combined = combined.sort_index(ascending=False)

# ========= 9️⃣ 日期格式轉換 =========
combined.index = combined.index.strftime('%Y/%m/%d')
combined.index.name = "Date"

# ========= 🔟 輸出到程式目錄 =========
output_file = f"{ticker}_Historical_stock_price_adj.xlsx"

with pd.ExcelWriter(
    output_file,
    engine="xlsxwriter",
    engine_kwargs={
        "options": {
            "strings_to_formulas": False,
            "strings_to_urls": False
        }
    }
) as writer:
    combined.to_excel(writer)

print(f"✅ 已輸出：{os.path.abspath(output_file)}")
print("🎉 任務完成！")