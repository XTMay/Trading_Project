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


# ========= 2️⃣ 抓股本資料 =========
def get_share_capital_data(symbol):

    ticker = yf.Ticker(symbol)

    try:
        info = ticker.info
    except:
        print("❌ 無法取得資料")
        return None

    if not info:
        print("❌ info 為空")
        return None

    today_str = datetime.today().strftime("%Y-%m-%d")

    data = {
        "Date": today_str,
        "Ticker": symbol,
        "Shares Outstanding": info.get("sharesOutstanding"),
        "Float Shares": info.get("floatShares"),
        "Implied Shares Outstanding": info.get("impliedSharesOutstanding"),
        "Held by Insiders (%)": info.get("heldPercentInsiders"),
        "Held by Institutions (%)": info.get("heldPercentInstitutions"),
        "Short Shares": info.get("sharesShort"),
        "Short Prior Month": info.get("sharesShortPriorMonth"),
        "Short % of Shares Outstanding": info.get("shortPercentOfFloat"),
        "Market Cap": info.get("marketCap"),
        "Enterprise Value": info.get("enterpriseValue")
    }

    df = pd.DataFrame(list(data.items()), columns=["Item", "Value"])
    df = df.fillna("")

    return df


# ========= 3️⃣ 主程式 =========
stock_code = "MITSY"   # 可改為台股或美股

symbol = smart_generate_ticker(stock_code)

print(f"🔍 使用代號：{symbol}")

df = get_share_capital_data(symbol)

if df is not None:

    output_file = f"{symbol}_Share_Capital.xlsx"

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