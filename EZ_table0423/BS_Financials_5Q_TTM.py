import yfinance as yf
import pandas as pd
import numpy as np


# ========= 設定 =========
ticker_symbol = "9022.T"
output_file = f"{ticker_symbol}_BS_autoSafe.xlsx"

ticker = yf.Ticker(ticker_symbol)


# ========= 1️⃣ 抓季度 Balance Sheet =========
bs_q = ticker.quarterly_balance_sheet

if bs_q.empty:
    raise ValueError("❌ 無法取得季度 Balance Sheet")

bs_q.columns = pd.to_datetime(bs_q.columns)
bs_q = bs_q.sort_index(axis=1)


# ========= 2️⃣ 檢查季度完整性 =========
def check_quarter_integrity(df):

    dates = df.columns.sort_values()

    if len(dates) < 5:
        print("⚠ 季度資料不足5季")
        return "LESS_THAN_5"

    diffs = np.diff(dates).astype("timedelta64[D]").astype(int)

    for d in diffs[-5:]:
        if not (70 <= d <= 120):
            print("⚠ 偵測到季度不連續")
            return "BROKEN"

    return "VALID"


status = check_quarter_integrity(bs_q)


# ========= 3️⃣ 決定輸出範圍 =========
if status == "VALID":
    print("✅ 季度完整 → 輸出最近5季 + TTM")
    last_q = bs_q.iloc[:, -5:]

elif status == "LESS_THAN_5":
    print("🔄 不足5季 → 輸出所有季度 + TTM")
    last_q = bs_q   # 🔥 全部季度

else:  # BROKEN
    print("🔄 季度不連續 → 輸出最近單季 + TTM")
    last_q = bs_q.iloc[:, -1:]


last_q.columns = last_q.columns.strftime("%Y-%m")


# ========= 4️⃣ TTM（Balance Sheet = 最新季度） =========
bs_ttm = bs_q.iloc[:, -1:].copy()
bs_ttm.columns = ["TTM"]


# ========= 5️⃣ 合併 =========
combined = last_q.merge(
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


print(f"✅ 已輸出 Balance Sheet：{output_file}")