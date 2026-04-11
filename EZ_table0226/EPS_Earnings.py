import yfinance as yf
import pandas as pd
from datetime import datetime


# ========= 1️⃣ 台股自動判斷 =========
def smart_generate_ticker(code):

    if "." in code:
        return code

    for suffix in [".TW", ".TWO"]:
        test_symbol = f"{code}{suffix}"
        try:
            data = yf.Ticker(test_symbol).history(period="5d")
            if not data.empty:
                return test_symbol
        except:
            pass

    return code


# ========= 2️⃣ 抓 EPS / Earnings =========
def get_eps_data(symbol):

    ticker = yf.Ticker(symbol)

    rows = []

    today_str = datetime.today().strftime("%Y-%m-%d")
    rows.append(("Date", today_str))
    rows.append(("Ticker", symbol))
    rows.append(("", ""))

    # ========= 年度 EPS =========
    try:
        earnings = ticker.earnings  # 年度
        if not earnings.empty:
            rows.append(("--- Annual EPS ---", ""))

            for idx, row in earnings.iterrows():
                rows.append((str(idx), row["Earnings"]))

            rows.append(("", ""))
    except:
        pass

    # ========= 季度 EPS =========
    try:
        quarterly = ticker.quarterly_earnings
        if not quarterly.empty:
            rows.append(("--- Quarterly EPS ---", ""))

            for idx, row in quarterly.iterrows():
                rows.append((str(idx.date()), row["Earnings"]))

            rows.append(("", ""))
    except:
        pass

    # ========= TTM EPS =========
    try:
        info = ticker.info
        ttm_eps = info.get("trailingEps", None)

        if ttm_eps:
            rows.append(("--- TTM EPS ---", ttm_eps))
    except:
        pass

    return pd.DataFrame(rows, columns=["Item", "Value"])


# ========= 3️⃣ 主程式 =========
stock_code = "MITSY"  # 可改台股或美股

symbol = smart_generate_ticker(stock_code)

print(f"🔍 使用代號：{symbol}")

df = get_eps_data(symbol)

if df is not None and not df.empty:

    output_file = f"{symbol}_EPS_Earnings.xlsx"

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
    print("❌ 無 EPS 資料")