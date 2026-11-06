"""
Build a verification-friendly workbook of RETURNS and EXCESS RETURNS.

Requirements honoured:
  1. Every sheet has a 'Date' (time) column as the first column.
  2. Inputs: nifty500_monthly_COMPLETE  (320 stocks)  and  rf  (risk-free rate).
  3. Each input may be .csv OR .xlsx/.xls -- the code finds it by NAME, any format.
  4. Sheet "Returns": each stock's monthly return, then 'T-bill annual %', then
     'RF monthly' as the final two columns. A formula row sits under the header.
  5. Sheet "Excess_Returns": each stock's excess return, with a formula row too.

The numbers are computed here; the formula row (row 2) documents, in plain text,
exactly how each column was produced, so you can re-check any cell by hand.

Run:  pip install pandas openpyxl
      python build_returns_workbook.py
"""

import glob
import numpy as np
import pandas as pd

PRICES_NAME = "nifty500_monthly_COMPLETE"   # 320 stocks
RF_NAME     = "rf"                          # Date, annual_rate_pct
OUT_XLSX    = "returns_and_excess.xlsx"


# ---------- helper: load a file by NAME whatever its format ----------
def load_any(basename):
    hits = (glob.glob(basename + ".csv") + glob.glob(basename + ".xlsx")
            + glob.glob(basename + ".xls"))
    if not hits:
        raise FileNotFoundError(f"No file named '{basename}' (.csv/.xlsx/.xls) here.")
    path = hits[0]
    if path.lower().endswith(".csv"):
        return pd.read_csv(path), path
    return pd.read_excel(path), path


def to_month_end(series):
    return pd.to_datetime(series).dt.to_period("M").dt.to_timestamp("M")


def main():
    # ---- read prices (first column = date, rest = 320 stocks) ----
    px_raw, px_path = load_any(PRICES_NAME)
    px_raw = px_raw.rename(columns={px_raw.columns[0]: "Date"})
    px_raw["Date"] = to_month_end(px_raw["Date"])
    prices = px_raw.set_index("Date").sort_index()

    # ---- read risk-free (Date, annual_rate_pct) ----
    rf_raw, rf_path = load_any(RF_NAME)
    rf_raw = rf_raw.rename(columns={rf_raw.columns[0]: "Date"})
    # find the annual-rate column (second column, or by name)
    rate_col = "annual_rate_pct" if "annual_rate_pct" in rf_raw.columns else rf_raw.columns[1]
    rf_raw["Date"] = to_month_end(rf_raw["Date"])
    rf = rf_raw.set_index("Date").sort_index()
    tbill_annual = rf[rate_col].rename("T-bill annual %")

    # ---- align on common months ----
    idx = prices.index.intersection(tbill_annual.index)
    prices, tbill_annual = prices.loc[idx], tbill_annual.loc[idx]
    stock_cols = list(prices.columns)

    # ---- calculations ----
    stock_ret  = np.log(prices / prices.shift(1))                 # monthly log return
    rf_monthly = (1 + tbill_annual / 100.0) ** (1 / 12) - 1       # annual% -> monthly
    rf_monthly = rf_monthly.rename("RF monthly")
    stock_excess = stock_ret.sub(rf_monthly, axis=0)             # return - rf

    # =====================================================================
    # Sheet 1: Returns  -> Date | <stocks returns> | T-bill annual % | RF monthly
    # =====================================================================
    ret_cols = ["Date"] + stock_cols + ["T-bill annual %", "RF monthly"]
    ret_formula = {c: "=LN(price_this_month / price_previous_month)" for c in stock_cols}
    ret_formula["Date"] = ""                              # raw dates, no formula
    ret_formula["T-bill annual %"] = "(raw input from rf file)"
    ret_formula["RF monthly"] = "=(1 + 'T-bill annual %'/100)^(1/12) - 1"

    ret_rows = []
    ret_rows.append([ret_formula[c] for c in ret_cols])  # formula row (row 2)
    for d in idx:
        row = [d.strftime("%Y-%m-%d")]
        row += [stock_ret.loc[d, s] for s in stock_cols]
        row += [tbill_annual.loc[d], rf_monthly.loc[d]]
        ret_rows.append(row)
    returns_sheet = pd.DataFrame(ret_rows, columns=ret_cols)

    # =====================================================================
    # Sheet 2: Excess_Returns -> Date | <stocks excess returns>
    # =====================================================================
    exc_cols = ["Date"] + stock_cols
    exc_formula = {c: "= stock_log_return - 'RF monthly'" for c in stock_cols}
    exc_formula["Date"] = ""

    exc_rows = []
    exc_rows.append([exc_formula[c] for c in exc_cols])  # formula row (row 2)
    for d in idx:
        row = [d.strftime("%Y-%m-%d")]
        row += [stock_excess.loc[d, s] for s in stock_cols]
        exc_rows.append(row)
    excess_sheet = pd.DataFrame(exc_rows, columns=exc_cols)

    # ---- write workbook ----
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xl:
        returns_sheet.to_excel(xl, sheet_name="Returns", index=False)
        excess_sheet.to_excel(xl, sheet_name="Excess_Returns", index=False)
        # light formatting: bold header, italic formula row, freeze panes
        from openpyxl.styles import Font
        for name in ("Returns", "Excess_Returns"):
            ws = xl.sheets[name]
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for cell in ws[2]:
                cell.font = Font(italic=True, color="888888")
            ws.freeze_panes = "B3"      # keep Date col + header/formula rows visible

    print(f"Prices read from : {px_path}  ({prices.shape[1]} stocks)")
    print(f"RF read from     : {rf_path}")
    print(f"Months           : {len(idx)}  ({idx.min().date()} -> {idx.max().date()})")
    print("Row 1 = column names, Row 2 = formula used, Row 3+ = monthly values.")
    print("First month's returns are blank (no earlier price to compare).")
    print(f"Wrote {OUT_XLSX}  (sheets: Returns, Excess_Returns)")


if __name__ == "__main__":
    main()
