import yfinance as yf
import pandas as pd
from datetime import datetime


# ========= 1️⃣ 輸入幣別 =========
currency_code = "JPY"  # 可改成 EUR / JPY / CNY / GBP 等

symbol = f"USD{currency_code}=X"


# ========= 2️⃣ 抓取匯率 =========
def get_exchange_rate(symbol, currency_code):

    ticker = yf.Ticker(symbol)

    try:
        data = ticker.history(period="5d")
    except:
        print("❌ 無法取得匯率資料")
        return None

    if data.empty:
        print("❌ 無匯率資料")
        return None

    latest = data.iloc[-1]

    today_str = datetime.today().strftime("%Y-%m-%d")

    exchange_rate = latest["Close"]
    open_price = latest["Open"]
    high_price = latest["High"]
    low_price = latest["Low"]

    result = {
        "Date": today_str,   # 🔥 第一列為日期
        "Currency Pair": f"USD/{currency_code}",
        "Exchange Rate (Close)": exchange_rate,
        "Open": open_price,
        "High": high_price,
        "Low": low_price
    }

    df = pd.DataFrame(list(result.items()), columns=["Item", "Value"])
    df = df.fillna("")

    return df


# ========= 3️⃣ 執行 =========
df = get_exchange_rate(symbol, currency_code)

if df is not None:

    output_file = f"USD_to_{currency_code}_Exchange_Rate.xlsx"

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
        df.to_excel(writer, index=False)

    print(f"\n✅ 已輸出：{output_file}")

else:
    print("❌ 無資料可輸出")