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


# ========= 2️⃣ 抓基本公司資料 =========
def get_basic_info(symbol):

    ticker = yf.Ticker(symbol)

    try:
        info = ticker.info
    except:
        print(f"❌ {symbol} 無法取得資料")
        return None

    if not info:
        print(f"❌ {symbol} info 為空")
        return None

    today_str = datetime.today().strftime("%Y-%m-%d")

    data = {
        "Date": today_str,  # 🔥 加入日期
        "Ticker": symbol,
        "Company Name": info.get("longName"),
        "Sector": info.get("sector"),
        "Industry": info.get("industry"),
        "Country": info.get("country"),
        "Currency": info.get("currency"),
        "Market Cap": info.get("marketCap"),
        "Enterprise Value": info.get("enterpriseValue"),
        "Trailing PE": info.get("trailingPE"),
        "Forward PE": info.get("forwardPE"),
        "EPS (TTM)": info.get("trailingEps"),
        "Dividend Yield": info.get("dividendYield"),
        "Beta": info.get("beta"),
        "52 Week High": info.get("fiftyTwoWeekHigh"),
        "52 Week Low": info.get("fiftyTwoWeekLow"),
        "Shares Outstanding": info.get("sharesOutstanding"),
        "Revenue (TTM)": info.get("totalRevenue"),
        "Net Income (TTM)": info.get("netIncomeToCommon"),
        "Website": info.get("website"),
    }

    return data


# ========= 3️⃣ 主程式 =========
stock_code = "MITSY"  # 可改成任何台股或美股
symbol = smart_generate_ticker(stock_code)

print(f"🔍 使用代號：{symbol}")

data = get_basic_info(symbol)

if data:

    # 轉成兩欄格式
    df = pd.DataFrame(list(data.items()), columns=["Item", "Value"])
    df = df.fillna("")

    output_file = f"{symbol}_Basic_Info.xlsx"

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