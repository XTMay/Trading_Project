import yfinance as yf
import pandas as pd
import numpy as np


# ========= 設定 =========
ticker_symbol = "AER"
output_file = f"{ticker_symbol}_IS_Financials_5Y_plus_TTM_with_Unusual.xlsx"

ticker = yf.Ticker(ticker_symbol)


# ========= 1️⃣ 年度資料 =========
income_y = ticker.financials.copy()

# 若無資料直接報錯
if income_y.empty:
    raise ValueError("無年度資料")

# 日期處理
income_y.columns = pd.to_datetime(income_y.columns)
income_y = income_y.sort_index(axis=1)

# 🔥 直接取最後5欄（不要刪當年度）
last_5y = income_y.iloc[:, -5:]
last_5y.columns = last_5y.columns.strftime("%Y")


# ========= 2️⃣ TTM =========
income_q = ticker.quarterly_financials.copy()

if not income_q.empty and income_q.shape[1] >= 4:

    income_q.columns = pd.to_datetime(income_q.columns)
    income_q = income_q.sort_index(axis=1)

    last_4q = income_q.iloc[:, -4:]
    income_ttm = last_4q.sum(axis=1)
    income_ttm.name = "TTM"
    income_ttm = income_ttm.to_frame()

else:
    income_ttm = pd.DataFrame()


# ========= 3️⃣ 合併（保證不掉 index） =========
if not income_ttm.empty:
    combined = pd.concat([last_5y, income_ttm], axis=1)
else:
    combined = last_5y.copy()


# ========= 4️⃣ 🔥 絕對保留 Item =========
combined.index.name = "Item"
combined_output = combined.reset_index()


# ========= 5️⃣ 清洗 =========
for col in combined_output.columns[1:]:
    combined_output[col] = pd.to_numeric(combined_output[col], errors="coerce")

combined_output = combined_output.replace([np.inf, -np.inf], np.nan)
combined_output = combined_output.fillna("")


# ========= 6️⃣ 抓 Unusual =========
keywords = ["unusual", "special", "restructuring", "non recurring"]

mask = combined_output["Item"].str.lower().str.contains(
    "|".join(keywords),
    na=False
)

unusual_output = combined_output[mask]


# ========= 7️⃣ 加分隔 =========
if not unusual_output.empty:

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
    final_output.to_excel(writer, index=False)


print(f"✅ 已修正輸出 5Y + TTM（完整保留 Item）：{output_file}")