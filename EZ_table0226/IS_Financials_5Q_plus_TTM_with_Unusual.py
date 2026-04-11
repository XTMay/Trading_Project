import yfinance as yf
import pandas as pd
import numpy as np


# ========= 設定 =========
ticker_symbol = "DIOD"
output_file = f"{ticker_symbol}_IS_5Q_plus_TTM_with_Unusual.xlsx"

ticker = yf.Ticker(ticker_symbol)


# ========= 1️⃣ 抓季度資料 =========
income_q = ticker.quarterly_financials

income_q.columns = pd.to_datetime(income_q.columns)
income_q = income_q.sort_index(axis=1)


# ========= 2️⃣ 最近5季 =========
if income_q.shape[1] >= 5:
    last_5q = income_q.iloc[:, -5:]
else:
    last_5q = income_q

last_5q.columns = last_5q.columns.strftime("%Y-%m")


# ========= 3️⃣ 計算 TTM =========
if income_q.shape[1] >= 4:
    last_4q = income_q.iloc[:, -4:]
    income_ttm = last_4q.sum(axis=1).to_frame(name="TTM")
else:
    raise ValueError("季度資料不足4季，無法計算TTM")


# ========= 4️⃣ 合併 =========
combined = last_5q.merge(
    income_ttm,
    left_index=True,
    right_index=True,
    how="left"
)


# ========= 🔥 關鍵清洗 =========
def sanitize_excel(df):

    df = df.copy()

    # 強制轉數值
    for col in df.columns:
        if col != "Item":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 移除 inf
    df = df.replace([np.inf, -np.inf], np.nan)

    # NaN 轉空白
    df = df.fillna("")

    return df


# ========= 5️⃣ 抓 Unusual =========
keywords = ["unusual", "special", "restructuring", "non recurring"]

unusual_rows = combined.loc[
    combined.index.str.lower().str.contains("|".join(keywords))
]


# ========= 6️⃣ 重設 index =========
combined.index.name = "Item"
combined_output = combined.reset_index()

combined_output = sanitize_excel(combined_output)


if not unusual_rows.empty:

    unusual_rows.index.name = "Item"
    unusual_output = unusual_rows.reset_index()
    unusual_output = sanitize_excel(unusual_output)

    separator = pd.DataFrame(
        [["=== Unusual / Special Items ==="] +
         [""] * (combined_output.shape[1] - 1)],
        columns=combined_output.columns
    )

    final_output = pd.concat(
        [combined_output, separator, unusual_output],
        ignore_index=True
    )
else:
    final_output = combined_output


# ========= 🔥 安全輸出 =========
with pd.ExcelWriter(
    output_file,
    engine="xlsxwriter",
    engine_kwargs={
        "options": {
            "strings_to_formulas": False,  # 🚀 核心修正
            "strings_to_urls": False
        }
    }
) as writer:
    final_output.to_excel(writer, index=False)


print(f"✅ 已安全輸出 5Q + TTM 合併版本：{output_file}")