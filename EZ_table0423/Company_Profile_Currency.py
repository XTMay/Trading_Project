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
        fast_info = ticker.fast_info
    except:
        print(f"❌ {symbol} 無法取得資料")
        return None

    if not info:
        print(f"❌ {symbol} info 為空")
        return None

    today_str = datetime.today().strftime("%Y-%m-%d")

    current_price = fast_info.get("lastPrice")

    data = {
        "Date": today_str,
        "Ticker": symbol,
        "Company Name": info.get("longName"),
        "Sector": info.get("sector"),
        "Industry": info.get("industry"),
        "Country": info.get("country"),
        "Currency": info.get("currency"),
        "Current Price": current_price,
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
#--------------------------------------------------------------------------
        #"Shares Outstanding": info.get("sharesOutstanding"),
        "Float Shares": info.get("floatShares"),
        "Implied Shares Outstanding": info.get("impliedSharesOutstanding"),
        "Held by Insiders (%)": info.get("heldPercentInsiders"),
        "Held by Institutions (%)": info.get("heldPercentInstitutions"),
        "Short Shares": info.get("sharesShort"),
        "Short Prior Month": info.get("sharesShortPriorMonth"),
        "Short % of Shares Outstanding": info.get("shortPercentOfFloat"),
        #"Market Cap": info.get("marketCap"),
        #"Enterprise Value": info.get("enterpriseValue")        
    }

    return data


# ========= 3️⃣ 主程式 =========
stock_code = "SWZNF"  # 可改成任何台股或美股
symbol = smart_generate_ticker(stock_code)

print(f"🔍 使用代號：{symbol}")

data = get_basic_info(symbol)

if data:

    df = pd.DataFrame(list(data.items()), columns=["Item", "Value"])
    df = df.fillna("")

    # ========= 🔥 取得 Country =========
    country_value = df.loc[df["Item"] == "Country", "Value"].values

    if len(country_value) > 0:
        country_name = country_value[0]
    else:
        country_name = ""

    # ========= 🔥 Country → Currency Code =========
    country_currency_map = {
        "Japan": "JPY",
        "Taiwan": "TWD",
        "United States": "USD",
        "USA": "USD",
        "China": "CNY",
        "Hong Kong": "HKD",
        "United Kingdom": "GBP",
        "South Korea": "KRW",
        "Germany": "EUR",
        "France": "EUR",
        "Canada": "CAD",
        "Australia": "AUD",
        "Singapore": "SGD",
        "Switzerland": "CHF",
        "Ireland": "EUR"
    }

    converted_currency = country_currency_map.get(country_name, "")

    # ========= 🔥 插入第7列 =========
    insert_position = 6  # 第7列 (index從0開始)

    new_row = pd.DataFrame(
        [["Country Converted Currency", converted_currency]],
        columns=["Item", "Value"]
    )

    df = pd.concat(
        [df.iloc[:insert_position], new_row, df.iloc[insert_position:]],
        ignore_index=True
    )

    # ========= 輸出 =========
    output_file = f"{symbol}_FULL_Profile_Share_Currency.xlsx"

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