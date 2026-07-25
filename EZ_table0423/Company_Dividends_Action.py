import yfinance as yf
import pandas as pd
from datetime import datetime


# ========= 1️⃣ 台股自動判斷 =========
def smart_generate_ticker(code):

    if "." in code:
        return code

    for suffix in [".TW", ".TWO"]:
        test_symbol = f"{code}{suffix}"
        data = yf.Ticker(test_symbol).history(period="5d")
        if not data.empty:
            return test_symbol

    return code


# ========= 2️⃣ 抓 Actions =========
def get_actions_data(symbol):

    ticker = yf.Ticker(symbol)

    try:
        actions = ticker.actions
    except:
        print("❌ 無法取得 Actions")
        return None

    if actions.empty:
        print("❌ 無 Actions 資料")
        return None

    actions = actions.sort_index(ascending=False)

    output_rows = []

    # 加入標頭資訊
    today_str = datetime.today().strftime("%Y-%m-%d")
    output_rows.append(("Date", today_str))
    output_rows.append(("Ticker", symbol))
    output_rows.append(("", ""))

    # ========= Dividends =========
    dividends = actions[actions["Dividends"] != 0]

    if not dividends.empty:
        output_rows.append(("--- Dividends ---", ""))

        for idx, row in dividends.iterrows():
            date_str = idx.strftime("%Y-%m-%d")
            output_rows.append((date_str, row["Dividends"]))

        output_rows.append(("", ""))

    # ========= Stock Splits =========
    splits = actions[actions["Stock Splits"] != 0]

    if not splits.empty:
        output_rows.append(("--- Stock Splits ---", ""))

        for idx, row in splits.iterrows():
            date_str = idx.strftime("%Y-%m-%d")
            output_rows.append((date_str, row["Stock Splits"]))

    return pd.DataFrame(output_rows, columns=["Item", "Value"])


# ========= 3️⃣ 主程式 =========
stock_code = "DIOD"  # 可改成台股或美股
symbol = smart_generate_ticker(stock_code)

print(f"🔍 使用代號：{symbol}")

df = get_actions_data(symbol)

if df is not None:

    output_file = f"{symbol}_Dividends_Actions.xlsx"

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