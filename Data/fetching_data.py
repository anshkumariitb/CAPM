"""
FRESH VERSION of the NIFTY 500 data-pull / portfolio-count script.
 
Same job as before -- download month-end ADJUSTED prices for every NIFTY 500
constituent, keep only stocks with COMPLETE monthly data, and report how many
10-stock portfolios you can form -- BUT both output files now carry an explicit
"Date" column as the first column (so the month/time is never blank).
 
HOW TO USE
  1) pip install yfinance pandas openpyxl
  2) Download ind_nifty500list.csv from niftyindices.com, put it next to this file.
  3) Set START / END below.
  4) python fetch_nifty500_monthly_v2.py
 
OUTPUTS (Date is the first column in every one):
  nifty500_monthly_ALL.csv / .xlsx       every downloaded stock
  nifty500_monthly_COMPLETE.csv / .xlsx  only stocks with no missing month
  SUMMARY.txt                            counts + portfolio number
"""
 
import time
import pandas as pd
import yfinance as yf
 
# ----------------------------- CONFIG -----------------------------
CONSTITUENTS_CSV = "ind_nifty500list.csv"
SYMBOL_COL       = "Symbol"
START            = "2016-01-01"
END              = "2025-12-31"
STOCKS_PER_PORTFOLIO = 10
PAUSE_SECONDS        = 0.4
# ------------------------------------------------------------------
 
 
def load_symbols(path):
    df = pd.read_csv(path)
    syms = df[SYMBOL_COL].dropna().astype(str).str.strip().tolist()
    return [(s, s + ".NS") for s in syms]
 
 
def fetch_one(yahoo_ticker):
    try:
        df = yf.download(yahoo_ticker, start=START, end=END, interval="1mo",
                         auto_adjust=True, progress=False, threads=False)
        if df is None or df.empty or "Close" not in df:
            return None
        s = df["Close"].dropna()
        if isinstance(s, pd.DataFrame):
            s = s.iloc[:, 0]
        return s if len(s) else None
    except Exception:
        return None
 
 
def add_date_column(df):
    """Move the month-end index into a real first column called 'Date' (YYYY-MM)."""
    out = df.copy()
    out.index = out.index.to_period("M").to_timestamp("M")
    out = out.reset_index()
    out = out.rename(columns={out.columns[0]: "Date"})
    out["Date"] = pd.to_datetime(out["Date"]).dt.strftime("%Y-%m")
    return out
 
 
def main():
    pairs = load_symbols(CONSTITUENTS_CSV)
    print(f"Loaded {len(pairs)} symbols\n")
 
    prices, failed = {}, []
    for i, (nse_sym, yf_sym) in enumerate(pairs, 1):
        s = fetch_one(yf_sym)
        if s is None:
            failed.append(nse_sym)
        else:
            prices[nse_sym] = s
        print(f"[{i:>3}/{len(pairs)}] {nse_sym:<15} {'FAILED' if s is None else f'{len(s)} months'}")
        time.sleep(PAUSE_SECONDS)
 
    price_matrix = pd.DataFrame(prices).sort_index()
    price_matrix.index = price_matrix.index.to_period("M").to_timestamp("M")
    price_matrix = price_matrix[~price_matrix.index.duplicated(keep="last")]
    price_matrix = price_matrix.reindex(pd.date_range(START, END, freq="M"))
 
    complete = price_matrix.dropna(axis=1, how="any")
 
    n_complete = complete.shape[1]
    n_ports    = n_complete // STOCKS_PER_PORTFOLIO
 
    # ---- write outputs WITH a Date column first ----
    all_out      = add_date_column(price_matrix)
    complete_out = add_date_column(complete)
 
    all_out.to_csv("nifty500_monthly_ALL.csv", index=False)
    complete_out.to_csv("nifty500_monthly_COMPLETE.csv", index=False)
    all_out.to_excel("nifty500_monthly_ALL.xlsx", index=False)
    complete_out.to_excel("nifty500_monthly_COMPLETE.xlsx", index=False)
 
    with open("SUMMARY.txt", "w") as f:
        f.write("\n".join([
            f"Window               : {START} to {END}",
            f"Symbols in list      : {len(pairs)}",
            f"Downloaded OK        : {price_matrix.shape[1]}",
            f"Failed on Yahoo      : {len(failed)}",
            f"Expected months      : {len(price_matrix)}",
            f"Stocks COMPLETE data : {n_complete}",
            f"Stocks per portfolio : {STOCKS_PER_PORTFOLIO}",
            f"==> PORTFOLIOS YOU CAN FORM : {n_ports}",
            "",
            "Failed tickers: " + (", ".join(failed) if failed else "(none)"),
        ]))
 
    print(f"\nComplete-data stocks: {n_complete}  ->  Portfolios: {n_ports}")
    print("Wrote ALL/COMPLETE .csv and .xlsx (Date is the first column) + SUMMARY.txt")
 
 
if __name__ == "__main__":
    main()
