import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime


# ========= 設定 =========
ticker_symbol = "MD"
output_file = f"{ticker_symbol}_IS_5Y_plus_TTM_with_Unusual.xlsx"

ticker = yf.Ticker(ticker_symbol)


# ========= 1️⃣ 抓原始資料 =========
income_raw = ticker.financials
income_raw.columns = pd.to_datetime(income_raw.columns)
income_raw = income_raw.sort_index(axis=1)

current_year = datetime.today().year
income_raw = income_raw.loc[:, income_raw.columns.year < current_year]


# ========= 2️⃣ 最近5年 =========
if income_raw.shape[1] >= 5:
    income_5y = income_raw.iloc[:, -5:]
else:
    income_5y = income_raw

income_5y.columns = income_5y.columns.strftime("%Y")


# ========= 3️⃣ TTM =========
income_full = ticker.financials
income_full.columns = pd.to_datetime(income_full.columns)
income_full = income_full.sort_index(axis=1)

income_ttm = income_full.iloc[:, -1:].copy()
income_ttm = income_ttm.apply(pd.to_numeric, errors="coerce")
income_ttm.columns = ["TTM"]


# ========= 4️⃣ 合併 =========
combined = income_5y.merge(
    income_ttm,
    left_index=True,
    right_index=True,
    how="left"
)


# ========= 🔥 關鍵清洗 =========
def sanitize_excel(df):

    df = df.copy()

    for col in df.columns:
        if col != "Item":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.fillna("")

    return df


combined.index.name = "Item"
combined_output = combined.reset_index()

combined_output = sanitize_excel(combined_output)


# ========= 5️⃣ 安全輸出 =========
with pd.ExcelWriter(
    output_file,
    engine="xlsxwriter",
    engine_kwargs={
        "options": {
            "strings_to_formulas": False,   # 🔥 關鍵
            "strings_to_urls": False
        }
    }
) as writer:
    combined_output.to_excel(writer, index=False)

print("✅ Excel 公式錯誤已完全 bypass")