"""
fmp_client.py
FMP (Financial Modeling Prep) API wrapper.
All HTTP calls to FMP are centralized here, returning data in formats
compatible with the generate_report_summary Excel writer.

Uses the /stable/ API path with ?symbol=XXX parameter format.
"""

import requests
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# ========= FMP Configuration =========
FMP_API_KEY = "hwCzbwI2zXSu7DKQ8iiF0vtKTeSCQ19y"
BASE_URL = "https://financialmodelingprep.com/stable"


def fmp_get(endpoint, params=None):
    """Base HTTP GET for FMP stable API. Returns parsed JSON (list or dict)."""
    params = params or {}
    params["apikey"] = FMP_API_KEY
    url = f"{BASE_URL}/{endpoint.lstrip('/')}"
    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data


# ------------------------------------------------------------------
# Internal helpers
# ------------------------------------------------------------------

def _fmp_statements_to_df(records, date_field="date", date_fmt=None):
    """
    Convert a list of FMP statement dicts into a DataFrame with
    rows = field names and columns = period dates (sorted ascending).
    This matches the yfinance DataFrame orientation.
    """
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    if date_field not in df.columns:
        return pd.DataFrame()
    df[date_field] = pd.to_datetime(df[date_field])
    df = df.sort_values(date_field)
    df = df.set_index(date_field).T
    # Drop metadata rows that are not numeric financial data
    meta_rows = {"symbol", "reportedCurrency", "cik", "fillingDate",
                 "filingDate", "acceptedDate", "calendarYear", "fiscalYear",
                 "period", "link", "finalLink"}
    df = df.loc[~df.index.isin(meta_rows)]
    return df


def _compute_adj_close(daily_df, dividends_df):
    """
    Compute dividend-adjusted close prices.
    FMP's close is already split-adjusted but NOT dividend-adjusted.
    Walk backwards from newest date; every time we cross an ex-dividend
    date d, multiply all earlier prices by (1 - dividend / close_on_d).

    Parameters
    ----------
    daily_df : DataFrame with DatetimeIndex sorted ascending, must have 'close' column
    dividends_df : DataFrame with columns ['date', 'dividend'] (or 'amount')

    Returns
    -------
    Series of adjusted close prices, same index as daily_df
    """
    if daily_df.empty:
        return pd.Series(dtype=float)

    adj = daily_df["close"].copy().astype(float)

    if dividends_df is None or dividends_df.empty:
        return adj

    # Normalize dividend column name
    if "dividend" in dividends_df.columns:
        div_col = "dividend"
    elif "amount" in dividends_df.columns:
        div_col = "amount"
    elif "adjDividend" in dividends_df.columns:
        div_col = "adjDividend"
    else:
        return adj

    divs = dividends_df[["date", div_col]].copy()
    divs["date"] = pd.to_datetime(divs["date"])
    divs = divs.sort_values("date", ascending=False)

    for _, row in divs.iterrows():
        ex_date = row["date"]
        div_amt = float(row[div_col])
        if div_amt <= 0:
            continue
        # Find the close on ex-date (or the nearest trading day)
        mask_on = daily_df.index >= ex_date
        if not mask_on.any():
            continue
        close_on_ex = float(daily_df.loc[mask_on, "close"].iloc[0])
        if close_on_ex <= 0:
            continue
        factor = 1 - div_amt / close_on_ex
        if factor <= 0 or factor > 1:
            continue
        adj.loc[adj.index < ex_date] *= factor

    return adj


# ==================================================================
# Public fetch functions
# ==================================================================

def fetch_quarterly_income(symbol):
    """
    Fetch quarterly income statement: last 5 quarters + TTM (sum of last 4Q).
    Returns (combined_df, unusual_df) matching yfinance orientation.
    """
    records = fmp_get("income-statement", {"symbol": symbol, "period": "quarter", "limit": 8})
    df = _fmp_statements_to_df(records)
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Keep last 5 quarters
    last_q = df.iloc[:, -5:] if df.shape[1] >= 5 else df.copy()
    last_q.columns = pd.to_datetime(last_q.columns).strftime("%Y-%m")

    # TTM = sum of last 4 quarters
    if df.shape[1] >= 4:
        last_4q = df.iloc[:, -4:]
        ttm = last_4q.apply(pd.to_numeric, errors="coerce").sum(axis=1).to_frame(name="TTM")
    else:
        ttm = df.iloc[:, -1:].apply(pd.to_numeric, errors="coerce").copy()
        ttm.columns = ["TTM"]

    combined = last_q.merge(ttm, left_index=True, right_index=True, how="left")

    # Unusual items (FMP rarely has explicit unusual labels)
    keywords = ["unusual", "special", "restructuring", "non recurring"]
    unusual_mask = combined.index.str.lower().str.contains("|".join(keywords))
    unusual = combined.loc[unusual_mask] if unusual_mask.any() else pd.DataFrame()

    combined = _sanitize(combined)
    return combined, unusual


def fetch_quarterly_balance_sheet(symbol):
    """Fetch quarterly balance sheet: last 5 quarters + TTM (latest snapshot)."""
    records = fmp_get("balance-sheet-statement", {"symbol": symbol, "period": "quarter", "limit": 8})
    df = _fmp_statements_to_df(records)
    if df.empty:
        return pd.DataFrame()

    last_q = df.iloc[:, -5:] if df.shape[1] >= 5 else df.copy()
    last_q.columns = pd.to_datetime(last_q.columns).strftime("%Y-%m")

    ttm = df.iloc[:, -1:].copy()
    ttm.columns = ["TTM"]

    combined = last_q.merge(ttm, left_index=True, right_index=True, how="left")
    return _sanitize(combined)


def fetch_annual_income(symbol):
    """Fetch annual income statement: last 5 years + TTM."""
    records = fmp_get("income-statement", {"symbol": symbol, "limit": 8})
    df = _fmp_statements_to_df(records)
    if df.empty:
        return pd.DataFrame()

    # Exclude current year
    current_year = datetime.today().year
    cols_keep = [c for c in df.columns if pd.to_datetime(c).year < current_year]
    df = df[cols_keep]
    if df.empty:
        return pd.DataFrame()

    last_y = df.iloc[:, -5:] if df.shape[1] >= 5 else df.copy()
    last_y.columns = pd.to_datetime(last_y.columns).strftime("%Y")

    # TTM = most recent year
    ttm = df.iloc[:, -1:].apply(pd.to_numeric, errors="coerce").copy()
    ttm.columns = ["TTM"]

    combined = last_y.merge(ttm, left_index=True, right_index=True, how="left")
    return _sanitize(combined)


def fetch_annual_balance_sheet(symbol):
    """Fetch annual balance sheet: last 5 years + TTM (latest snapshot)."""
    records = fmp_get("balance-sheet-statement", {"symbol": symbol, "limit": 8})
    df = _fmp_statements_to_df(records)
    if df.empty:
        return pd.DataFrame()

    last_y = df.iloc[:, -5:] if df.shape[1] >= 5 else df.copy()
    last_y.columns = pd.to_datetime(last_y.columns).strftime("%Y")

    ttm = df.iloc[:, -1:].copy()
    ttm.columns = ["TTM"]

    combined = last_y.merge(ttm, left_index=True, right_index=True, how="left")
    return _sanitize(combined)


def fetch_annual_cashflow(symbol):
    """Fetch annual cash flow: last 5 years + TTM (4Q sum or fallback)."""
    records_y = fmp_get("cash-flow-statement", {"symbol": symbol, "limit": 8})
    df_y = _fmp_statements_to_df(records_y)
    if df_y.empty:
        return pd.DataFrame()

    last_y = df_y.iloc[:, -5:] if df_y.shape[1] >= 5 else df_y.copy()
    last_y.columns = pd.to_datetime(last_y.columns).strftime("%Y")

    # TTM from quarterly cash flow
    try:
        records_q = fmp_get("cash-flow-statement", {"symbol": symbol, "period": "quarter", "limit": 8})
        df_q = _fmp_statements_to_df(records_q)
        if not df_q.empty and df_q.shape[1] >= 4:
            ttm = df_q.iloc[:, -4:].apply(pd.to_numeric, errors="coerce").sum(axis=1).to_frame(name="TTM")
        elif not df_q.empty:
            ttm = df_q.iloc[:, -1:].apply(pd.to_numeric, errors="coerce").copy()
            ttm.columns = ["TTM"]
        else:
            ttm = df_y.iloc[:, -1:].apply(pd.to_numeric, errors="coerce").copy()
            ttm.columns = ["TTM"]
    except Exception:
        ttm = df_y.iloc[:, -1:].apply(pd.to_numeric, errors="coerce").copy()
        ttm.columns = ["TTM"]

    combined = last_y.merge(ttm, left_index=True, right_index=True, how="left")
    return _sanitize(combined)


def fetch_company_profile(symbol):
    """
    Fetch company profile by combining /stable/profile, /stable/quote,
    /stable/ratios-ttm, and /stable/key-metrics-ttm.
    Returns a dict with keys matching the yfinance-style names used by
    the Excel writer.
    """
    # Profile
    profile_data = fmp_get("profile", {"symbol": symbol})
    p = profile_data[0] if isinstance(profile_data, list) and profile_data else {}

    # Quote
    quote_data = fmp_get("quote", {"symbol": symbol})
    q = quote_data[0] if isinstance(quote_data, list) and quote_data else {}

    # Ratios TTM
    try:
        ratios_data = fmp_get("ratios-ttm", {"symbol": symbol})
        r = ratios_data[0] if isinstance(ratios_data, list) and ratios_data else {}
    except Exception:
        r = {}

    # Key Metrics TTM
    try:
        km_data = fmp_get("key-metrics-ttm", {"symbol": symbol})
        km = km_data[0] if isinstance(km_data, list) and km_data else {}
    except Exception:
        km = {}

    # Get financialCurrency from the latest income statement
    fin_currency = "USD"
    try:
        is_data = fmp_get("income-statement", {"symbol": symbol, "limit": 1})
        if isinstance(is_data, list) and is_data:
            fin_currency = is_data[0].get("reportedCurrency", "USD")
    except Exception:
        pass

    # Dividend yield: prefer ratios-ttm, fallback to lastDividend/price
    div_yield = r.get("dividendYieldTTM")
    if div_yield is None and p.get("lastDividend") and p.get("price"):
        try:
            div_yield = float(p["lastDividend"]) / float(p["price"])
        except (TypeError, ZeroDivisionError, ValueError):
            div_yield = None

    # EPS: prefer ratios-ttm netIncomePerShareTTM, fallback to quote eps
    trailing_eps = r.get("netIncomePerShareTTM") or q.get("eps")

    # Trailing PE: prefer ratios-ttm
    trailing_pe = r.get("priceToEarningsRatioTTM") or q.get("pe")

    # Payout ratio: dividendPayoutRatioTTM (not payoutRatioTTM)
    payout_ratio = r.get("dividendPayoutRatioTTM")

    # Revenue and Net Income from latest annual IS (most reliable)
    total_revenue = None
    net_income_common = None
    try:
        is_annual = fmp_get("income-statement", {"symbol": symbol, "limit": 1})
        if isinstance(is_annual, list) and is_annual:
            total_revenue = is_annual[0].get("revenue")
            net_income_common = is_annual[0].get("netIncome")
    except Exception:
        pass

    # Enterprise value from ratios-ttm (more reliable than key-metrics-ttm)
    ev = r.get("enterpriseValueTTM") or km.get("enterpriseValueTTM")

    return {
        "longName": p.get("companyName", ""),
        "symbol": p.get("symbol", symbol),
        "exchange": p.get("exchange", "") or q.get("exchange", ""),
        "sector": p.get("sector", ""),
        "industry": p.get("industry", ""),
        "country": p.get("country", ""),
        "currency": p.get("currency", ""),
        "financialCurrency": fin_currency,
        "currentPrice": p.get("price") or q.get("price"),
        "marketCap": p.get("marketCap") or q.get("marketCap"),
        "enterpriseValue": ev,
        "trailingPE": trailing_pe,
        "forwardPE": None,  # FMP stable API has no direct forward PE; PEG ratio is different
        "trailingEps": trailing_eps,
        "dividendYield": div_yield,
        "payoutRatio": payout_ratio,
        "beta": p.get("beta"),
        "fiftyTwoWeekHigh": q.get("yearHigh"),
        "fiftyTwoWeekLow": q.get("yearLow"),
        "longBusinessSummary": p.get("description", ""),
        "website": p.get("website", ""),
        "totalRevenue": total_revenue,
        "netIncomeToCommon": net_income_common,
    }


def fetch_share_capital(symbol):
    """
    Fetch share capital data from /stable/shares-float + /stable/quote.
    Some fields (short interest, insider/institution %) are not available
    from FMP and will be None.
    """
    # Shares float
    try:
        sf_data = fmp_get("shares-float", {"symbol": symbol})
        sf = sf_data[0] if isinstance(sf_data, list) and sf_data else {}
    except Exception:
        sf = {}

    # Quote for sharesOutstanding
    try:
        q_data = fmp_get("quote", {"symbol": symbol})
        q = q_data[0] if isinstance(q_data, list) and q_data else {}
    except Exception:
        q = {}

    return {
        "Shares Outstanding": sf.get("outstandingShares") or q.get("sharesOutstanding"),
        "Float Shares": sf.get("floatShares"),
        "Implied Shares Outstanding": None,  # Not available from FMP
        "Held by Insiders (%)": None,  # Would need separate endpoint + calculation
        "Held by Institutions (%)": None,  # Would need separate endpoint + calculation
        "Short Shares": None,  # FMP does not provide short interest
        "Short Prior Month": None,
        "Short % of Shares Outstanding": None,
    }


def fetch_eps_earnings(symbol):
    """
    Fetch EPS / earnings data from income statements + quote.
    Returns dict with 'annual', 'quarterly', 'ttm_eps' keys.
    """
    result = {"annual": {}, "quarterly": {}, "ttm_eps": None}

    # Annual net income from IS
    try:
        is_annual = fmp_get("income-statement", {"symbol": symbol, "limit": 5})
        if isinstance(is_annual, list):
            for rec in is_annual:
                date_str = rec.get("date", "")
                ni = rec.get("netIncome")
                if date_str and ni is not None:
                    year = date_str[:4]
                    result["annual"][year] = ni
    except Exception:
        pass

    # Quarterly net income from IS
    try:
        is_q = fmp_get("income-statement", {"symbol": symbol, "period": "quarter", "limit": 8})
        if isinstance(is_q, list):
            for rec in is_q:
                date_str = rec.get("date", "")
                ni = rec.get("netIncome")
                if date_str and ni is not None:
                    result["quarterly"][date_str] = ni
    except Exception:
        pass

    # TTM EPS from quote
    try:
        q_data = fmp_get("quote", {"symbol": symbol})
        if isinstance(q_data, list) and q_data:
            result["ttm_eps"] = q_data[0].get("eps")
    except Exception:
        pass

    return result


def fetch_historical_prices_adj(symbol, start_date="2014-01-01"):
    """
    Fetch daily historical prices, compute dividend-adjusted close,
    resample to monthly, and merge with corporate actions (dividends + splits).
    Returns a DataFrame sorted descending (newest first), matching yfinance output.
    """
    end_date = (datetime.today() + timedelta(days=1)).strftime("%Y-%m-%d")

    # Fetch daily prices (FMP /stable/ endpoint)
    try:
        price_data = fmp_get("historical-price-eod/full", {
            "symbol": symbol,
            "from": start_date,
            "to": end_date,
        })
    except Exception:
        return pd.DataFrame()

    # price_data can be a list directly or a dict with "historical" key
    if isinstance(price_data, dict):
        records = price_data.get("historical", [])
    elif isinstance(price_data, list):
        records = price_data
    else:
        return pd.DataFrame()

    if not records:
        return pd.DataFrame()

    daily = pd.DataFrame(records)
    daily["date"] = pd.to_datetime(daily["date"])
    daily = daily.sort_values("date").set_index("date")

    # Fetch dividends
    try:
        div_data = fmp_get("dividends", {"symbol": symbol})
        dividends_df = pd.DataFrame(div_data) if isinstance(div_data, list) and div_data else pd.DataFrame()
    except Exception:
        dividends_df = pd.DataFrame()

    # Fetch splits
    try:
        split_data = fmp_get("splits", {"symbol": symbol})
        splits_df = pd.DataFrame(split_data) if isinstance(split_data, list) and split_data else pd.DataFrame()
    except Exception:
        splits_df = pd.DataFrame()

    # Compute adjusted close (dividend-adjusted)
    adj_close = _compute_adj_close(daily, dividends_df)
    daily["Adj Close"] = adj_close

    # Rename columns to match yfinance style
    col_map = {"open": "Open", "high": "High", "low": "Low",
               "close": "Close", "volume": "Volume"}
    daily = daily.rename(columns=col_map)

    # Monthly resampling (month start)
    agg_cols = {}
    for c in ["Open", "High", "Low", "Close", "Adj Close", "Volume"]:
        if c in daily.columns:
            if c == "Open":
                agg_cols[c] = "first"
            elif c == "High":
                agg_cols[c] = "max"
            elif c == "Low":
                agg_cols[c] = "min"
            elif c in ("Close", "Adj Close"):
                agg_cols[c] = "last"
            elif c == "Volume":
                agg_cols[c] = "sum"

    monthly_df = daily.resample("MS").agg(agg_cols)

    # Build corporate actions DataFrame
    actions_list = []

    if not dividends_df.empty and "date" in dividends_df.columns:
        div_col = "dividend" if "dividend" in dividends_df.columns else (
            "amount" if "amount" in dividends_df.columns else
            "adjDividend" if "adjDividend" in dividends_df.columns else None)
        if div_col:
            for _, row in dividends_df.iterrows():
                d = pd.to_datetime(row["date"])
                if d >= pd.to_datetime(start_date):
                    actions_list.append({
                        "date": d,
                        "Dividends": float(row[div_col]),
                        "Stock Splits": 0,
                    })

    if not splits_df.empty and "date" in splits_df.columns:
        for _, row in splits_df.iterrows():
            d = pd.to_datetime(row["date"])
            if d >= pd.to_datetime(start_date):
                num = float(row.get("numerator", 0))
                den = float(row.get("denominator", 1))
                ratio = num / den if den != 0 else 0
                actions_list.append({
                    "date": d,
                    "Dividends": 0,
                    "Stock Splits": ratio,
                })

    if actions_list:
        actions_df = pd.DataFrame(actions_list)
        actions_df["date"] = pd.to_datetime(actions_df["date"])
        actions_df = actions_df.set_index("date")
        # Aggregate same-date actions
        actions_df = actions_df.groupby(level=0).agg({
            "Dividends": "sum",
            "Stock Splits": "sum",
        })
        actions_df = actions_df[
            (actions_df["Dividends"] != 0) | (actions_df["Stock Splits"] != 0)
        ]
    else:
        actions_df = pd.DataFrame()

    # Merge price data with corporate actions
    combined = pd.concat([monthly_df, actions_df], axis=0)

    # Ensure Dividends / Stock Splits columns exist
    for col in ["Dividends", "Stock Splits"]:
        if col not in combined.columns:
            combined[col] = 0
        else:
            combined[col] = combined[col].fillna(0)

    # Entry_Type label
    combined["Entry_Type"] = "Month_Start"
    if not actions_df.empty:
        combined.loc[actions_df.index, "Entry_Type"] = "Corporate_Action"

    # Sort descending (newest first)
    combined = combined.sort_index(ascending=False)

    return combined


def fetch_exchange_rate(currency_code):
    """
    Fetch USD/XXX exchange rate using /stable/quote?symbol=USDXXX.
    """
    if not currency_code or currency_code.upper() == "USD":
        return {"rate": 1.0, "pair": "USD/USD"}
    pair_symbol = f"USD{currency_code.upper()}"
    try:
        data = fmp_get("quote", {"symbol": pair_symbol})
        if isinstance(data, list) and data:
            q = data[0]
            return {
                "pair": f"USD/{currency_code.upper()}",
                "rate": q.get("price") or q.get("previousClose") or 1.0,
                "open": q.get("open"),
                "high": q.get("dayHigh"),
                "low": q.get("dayLow"),
            }
    except Exception:
        pass
    return {"rate": 1.0, "pair": f"USD/{currency_code.upper()}"}


# ------------------------------------------------------------------
# Shared helpers (same as generate_report_summary.py)
# ------------------------------------------------------------------

def _sanitize(df):
    """Clean DataFrame for Excel output."""
    df = df.copy()
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.replace([np.inf, -np.inf], np.nan)
    return df
