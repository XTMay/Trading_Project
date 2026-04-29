import yfinance as yf
import pandas as pd
import numpy as np


# ========= 設定 =========
ticker_symbol = "2211.TW"
output_file = f"{ticker_symbol}_CFS_5Y_plus_TTM.xlsx"

ticker = yf.Ticker(ticker_symbol)


# ========= 1️⃣ 抓年度 Cash Flow =========
cf_y = ticker.cashflow

if cf_y.empty:
    raise ValueError("❌ 無法取得年度 Cash Flow")

cf_y.columns = pd.to_datetime(cf_y.columns)
cf_y = cf_y.sort_index(axis=1)


# ========= 2️⃣ 抓季度 Cash Flow（若有） =========
try:
    cf_q = ticker.quarterly_cashflow
except:
    cf_q = pd.DataFrame()

if not cf_q.empty:
    cf_q.columns = pd.to_datetime(cf_q.columns)
    cf_q = cf_q.sort_index(axis=1)


# ========= 3️⃣ 檢查年度完整性 =========
def check_year_integrity(df):

    dates = df.columns.sort_values()

    if len(dates) < 5:
        print("⚠ 年度資料不足5年")
        return "LESS_THAN_5"

    diffs = np.diff(dates).astype("timedelta64[D]").astype(int)

    for d in diffs[-5:]:
        if not (330 <= d <= 400):
            print("⚠ 偵測到年度不連續")
            return "BROKEN"

    return "VALID"


status = check_year_integrity(cf_y)


# ========= 4️⃣ 決定年度輸出範圍 =========
if status == "VALID":
    print("✅ 年度完整 → 輸出最近5年")
    last_y = cf_y.iloc[:, -5:]

elif status == "LESS_THAN_5":
    print("🔄 不足5年 → 輸出所有年度")
    last_y = cf_y

else:
    print("🔄 年度不連續 → 輸出最近單年")
    last_y = cf_y.iloc[:, -1:]


last_y.columns = last_y.columns.strftime("%Y")


# ========= 5️⃣ 計算 TTM =========
if not cf_q.empty:

    if cf_q.shape[1] >= 4:
        print("✅ 使用最近4季加總計算 TTM")
        last_4q = cf_q.iloc[:, -4:]
        cf_ttm = last_4q.sum(axis=1).to_frame(name="TTM")

    else:
        print("⚠ 季度不足4季 → 使用最近季度")
        cf_ttm = cf_q.iloc[:, -1:].copy()
        cf_ttm.columns = ["TTM"]

else:
    print("⚠ 無法取得季度資料 → 使用最新年度作為 TTM")
    cf_ttm = cf_y.iloc[:, -1:].copy()
    cf_ttm.columns = ["TTM"]


# ========= 6️⃣ 合併 =========
combined = last_y.merge(
    cf_ttm,
    left_index=True,
    right_index=True,
    how="left"
)


# ========= 7️⃣ 清洗 =========
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


# ========= 8️⃣ 安全輸出 =========
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


print(f"✅ 已輸出 Cash Flow 5Y + TTM（自動容錯版）：{output_file}")