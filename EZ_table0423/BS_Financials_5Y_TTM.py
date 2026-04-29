import yfinance as yf
import pandas as pd
import numpy as np


# ========= 設定 =========
ticker_symbol = "8031.T"
output_file = f"{ticker_symbol}_BS_5Y_plus_TTM.xlsx"

ticker = yf.Ticker(ticker_symbol)


# ========= 1️⃣ 抓年度 Balance Sheet =========
bs_y = ticker.balance_sheet

if bs_y.empty:
    raise ValueError("❌ 無法取得年度 Balance Sheet")

bs_y.columns = pd.to_datetime(bs_y.columns)
bs_y = bs_y.sort_index(axis=1)


# ========= 2️⃣ 檢查年度完整性 =========
def check_year_integrity(df):

    dates = df.columns.sort_values()

    if len(dates) < 5:
        print("⚠ 年度資料不足5年")
        return "LESS_THAN_5"

    diffs = np.diff(dates).astype("timedelta64[D]").astype(int)

    for d in diffs[-5:]:
        if not (330 <= d <= 400):  # 年度容忍範圍
            print("⚠ 偵測到年度不連續")
            return "BROKEN"

    return "VALID"


status = check_year_integrity(bs_y)


# ========= 3️⃣ 決定輸出範圍 =========
if status == "VALID":
    print("✅ 年度完整 → 輸出最近5年 + TTM")
    last_y = bs_y.iloc[:, -5:]

elif status == "LESS_THAN_5":
    print("🔄 不足5年 → 輸出所有年度 + TTM")
    last_y = bs_y

else:  # BROKEN
    print("🔄 年度不連續 → 輸出最近單年 + TTM")
    last_y = bs_y.iloc[:, -1:]


last_y.columns = last_y.columns.strftime("%Y")


# ========= 4️⃣ TTM（Balance Sheet = 最新年度） =========
bs_ttm = bs_y.iloc[:, -1:].copy()
bs_ttm.columns = ["TTM"]


# ========= 5️⃣ 合併 =========
combined = last_y.merge(
    bs_ttm,
    left_index=True,
    right_index=True,
    how="left"
)


# ========= 6️⃣ 清洗 =========
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


# ========= 7️⃣ 安全輸出 =========
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
    combined_output.to_excel(writer, index=False)


print(f"✅ 已輸出年度 Balance Sheet：{output_file}")