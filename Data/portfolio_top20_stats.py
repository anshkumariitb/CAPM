# -*- coding: utf-8 -*-
# ===========================================================================
#  SUMMARY STATISTICS FOR THE TOP 20 (HIGHEST-BETA) PORTFOLIOS  P1..P20
# ===========================================================================
#  WHAT THIS PROGRAM DOES, IN ONE PARAGRAPH (for a non-coder):
#  Your rolling analysis already produced 32 beta-sorted portfolios, where P1
#  holds the highest-beta stocks and P32 the lowest. This program opens that
#  result file, keeps only the first 20 portfolios (P1..P20), and for each one
#  computes five everyday statistics of its 60 monthly numbers: the average,
#  the median (middle value), the maximum, the minimum, and the standard
#  deviation (spread). It does this for the RETURNS and again for the EXCESS
#  RETURNS, and saves both tables to a new Excel file.
#
#  HOW TO RUN:
#    1) pip install pandas openpyxl
#    2) Put this script in the SAME folder as capm_rolling_results.xlsx
#    3) python portfolio_top20_stats.py
#    OUTPUT: portfolio_top20_stats.xlsx
# ===========================================================================

import pandas as pd            # pandas = the tool that lets Python work with tables (like Excel)

# ---- SETTINGS you could change ----
IN_FILE  = "capm_rolling_results.xlsx"     # the file made by the rolling code (the INPUT)
OUT_FILE = "portfolio_top20_stats.xlsx"    # the file this script will create (the OUTPUT)
TOP_N    = 20                              # how many portfolios to keep: P1 .. P20


# ---------------------------------------------------------------------------
#  HELPER 1: clean up a messy column heading.
#  In the input file each portfolio column is named like  "P1\n(= average ...)".
#  The "\n(= ...)" part is just a note. This keeps only the "P1" part.
# ---------------------------------------------------------------------------
def short(colname):
    return str(colname).split("\n")[0].strip()   # take text before the line-break, trim spaces


# ---------------------------------------------------------------------------
#  HELPER 2: given a table of monthly numbers (columns = portfolios,
#  rows = the 60 months), return a small table of the five statistics per column.
# ---------------------------------------------------------------------------
def stats_table(df):
    out = pd.DataFrame({                    # build a new table, one column per statistic
        "Average":  df.mean(),              # mean of each portfolio's 60 monthly values
        "Median":   df.median(),            # middle value (half above, half below)
        "Maximum":  df.max(),               # the single best month
        "Minimum":  df.min(),               # the single worst month
        "Std Dev":  df.std(ddof=1),         # spread; ddof=1 = SAMPLE standard deviation (the usual one)
    })
    out.index.name = "Portfolio"            # label the row index (P1, P2, ...) as "Portfolio"
    return out.reset_index()                # turn that index into a normal first column for Excel


# ---------------------------------------------------------------------------
#  HELPER 3: open one sheet of the input file and keep only columns P1..P20.
# ---------------------------------------------------------------------------
def load_top20(sheet):
    raw = pd.read_excel(IN_FILE, sheet_name=sheet, header=0)   # read the sheet into a table
    raw.columns = [short(c) for c in raw.columns]              # simplify headings to "P1", "P2", "Date", ...
    wanted = [f"P{i}" for i in range(1, TOP_N + 1)]            # the list we want, in order: P1..P20
    keep   = [c for c in wanted if c in raw.columns]           # keep only those that exist in the sheet
    data   = raw[keep].apply(pd.to_numeric, errors="coerce")   # select them and force them to be numbers
    return data                                               # a 60-row x 20-column table of monthly values


# ---------------------------------------------------------------------------
#  MAIN: do the work and write the output file.
# ---------------------------------------------------------------------------
def main():
    ret = load_top20("2_Portfolio_Returns")     # 60 months x 20 portfolios of RETURNS
    exc = load_top20("3_Portfolio_Excess")       # 60 months x 20 portfolios of EXCESS returns

    ret_stats = stats_table(ret)                 # five statistics for each of P1..P20 (returns)
    exc_stats = stats_table(exc)                 # five statistics for each of P1..P20 (excess returns)

    # Open a new Excel file and write the two tables onto two separate sheets.
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as xl:
        ret_stats.to_excel(xl, sheet_name="Returns_Stats", index=False)        # first sheet
        exc_stats.to_excel(xl, sheet_name="ExcessReturns_Stats", index=False)  # second sheet
        # --- make the header row bold and widen the first column (just cosmetics) ---
        from openpyxl.styles import Font
        for name in ("Returns_Stats", "ExcessReturns_Stats"):
            ws = xl.sheets[name]                 # grab the sheet we just wrote
            for c in ws[1]:                      # ws[1] = the first (header) row
                c.font = Font(bold=True)         # make each header cell bold
            ws.column_dimensions["A"].width = 12 # widen column A (the Portfolio names)

    # Print a short confirmation in the terminal so you know it worked.
    print(f"Portfolios summarised : P1..P{TOP_N}")
    print(f"Months per portfolio  : {len(ret)}")
    print("Wrote", OUT_FILE, "(sheets: Returns_Stats, ExcessReturns_Stats)")


# This line means: only run main() when you execute this file directly.
if __name__ == "__main__":
    main()
