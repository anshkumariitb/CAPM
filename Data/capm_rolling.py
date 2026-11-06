# -*- coding: utf-8 -*-
# ===========================================================================
#  TESTING CAPM ON NIFTY 500 (2016-2025) -- ROLLING METHOD + TWO-PASS TEST
# ===========================================================================
#  WHAT THIS PROGRAM DOES, IN ONE PARAGRAPH (for a non-coder):
#  It reads three things you prepared in Excel: every stock's monthly return,
#  every stock's "excess return" (return minus the risk-free rate), and the
#  market's excess return. Then, year by year, it measures how sensitive each
#  stock is to the market (its "beta"), sorts the 320 stocks from highest to
#  lowest beta, and puts them into 32 baskets ("portfolios") of 10. It tracks
#  how those baskets perform the following year. Finally it runs the classic
#  CAPM test: do higher-beta baskets earn higher returns? It writes everything
#  to one Excel file with 6 sheets, including plain-English notes.
#
#  HOW TO RUN:
#    1) pip install pandas numpy openpyxl
#    2) Put these two files in the SAME folder as this script, named exactly:
#         returns_and_excess.xlsx   (sheets 'Returns' and 'Excess_Returns')
#         NIFTY500_CRSLDX_Historical_Data.csv  (has a 'Market Excess Return' column)
#    3) python capm_rolling.py
#    OUTPUT: capm_rolling_results.xlsx
# ===========================================================================

import numpy as np                       # numpy = fast math on lists of numbers
import pandas as pd                       # pandas = work with tables (like Excel) in code
from openpyxl.styles import Font, Alignment   # used only to make the Excel output pretty

# ---- SETTINGS you could change ----
RET_FILE = "returns_and_excess.xlsx"             # the workbook with stock returns + excess
MKT_FILE = "NIFTY500_CRSLDX_Historical_Data.csv" # the market (Nifty 500) file
N_PER_PF = 10                                     # 10 stocks in each portfolio (as in the paper)

# The rolling plan. Each line = (first year of beta window, last year, the year we then hold).
# Read line 1 as: "use 2016-2020 to measure beta, then hold those portfolios through 2021".
PLAN = [(2016,2020,2021), (2017,2021,2022), (2018,2022,2023),
        (2019,2023,2024), (2020,2024,2025)]


# ---------------------------------------------------------------------------
#  HELPER 1: read one sheet from the Excel workbook into a table
# ---------------------------------------------------------------------------
def read_sheet(path, sheet):
    # Read the sheet. header=0 -> first row is the column names.
    # skiprows=[1] -> SKIP the 2nd row, because in your file that row holds the
    #                 formula text (e.g. "=LN(...)"), not real numbers.
    df = pd.read_excel(path, sheet_name=sheet, header=0, skiprows=[1])
    # Rename the very first column to "Date" no matter what it was called.
    df = df.rename(columns={df.columns[0]: "Date"})
    # Turn the date text into a real monthly date (e.g. 2016-01) so months line up.
    df["Date"] = pd.to_datetime(df["Date"]).dt.to_period("M")
    # Make Date the row label and sort oldest-to-newest.
    return df.set_index("Date").sort_index()


# ---------------------------------------------------------------------------
#  HELPER 2: a small linear-regression (OLS). Given Y and inputs X, it returns
#  the coefficients (the lambdas), their t-values (how trustworthy each is),
#  and R-squared (how well the line fits, 0 to 1).
# ---------------------------------------------------------------------------
def ols(y, X):
    XtXi = np.linalg.inv(X.T @ X)          # the standard regression matrix maths
    coef = XtXi @ X.T @ y                   # the slope/intercept estimates (lambdas)
    resid = y - X @ coef                    # leftover error the line could not explain
    n, k = X.shape                          # n = number of points ( rows of x = 32), k = number of inputs ( column = no. of variables i.e, Lambda ) 
    s2 = (resid @ resid) / (n - k)          # variance of the leftover error
    #resid @ resid is the sum of squared residuals (RSS = Σ residual²). Dividing by n − k (not n) estimates the 
    #error variance σ². We subtract k because k numbers were "used up" fitting the line — that's the degrees of freedom.
    se = np.sqrt(np.diag(s2 * XtXi))        # standard error of each coefficient
    # this is where the ± comes from. The estimated variance-covariance matrix of the coefficients is s² · (XᵀX)⁻¹, a k×k matrix.
    # Its diagonal entries are the variances of each coefficient (np.diag picks them out), the square root gives each coefficient's standard error
    t = coef / se                           # t-value = coefficient / its standard error
    r2 = 1 - (resid @ resid) / (((y - y.mean())**2).sum())  # R-squared 1 − RSS/TSS, where TSS = Σ(y − mean(y))² is the total scatter to begin with
    return coef, t, r2


# ---------------------------------------------------------------------------
#  HELPER 3: turn a t-value into a word
# ---------------------------------------------------------------------------
def sig(tv):
    return "significant" if abs(tv) > 2 else "not significant"  # |t|>2 ~ 5% level


def main():
    # =======================================================================
    #  SECTION A: LOAD THE DATA
    # =======================================================================
    returns = read_sheet(RET_FILE, "Returns")          # raw monthly returns + RF monthly
    excess  = read_sheet(RET_FILE, "Excess_Returns")   # each stock's excess return
    rf      = returns["RF monthly"]                     # the monthly risk-free column
    stocks  = list(excess.columns)                      # the 320 stock names ( only names )
    sret    = returns[stocks]                           # table: months x 320 raw returns ( get table data )
    sexc    = excess[stocks]                            # table: months x 320 excess returns ( get table data )

    # Load the market file and pull its 'Market Excess Return' column.
    mkt = pd.read_csv(MKT_FILE)
    mkt["Date"] = pd.to_datetime(mkt["Date"]).dt.to_period("M")  # same monthly dates
    mkt = mkt.set_index("Date").sort_index()
    mexc = pd.to_numeric(mkt["Market Excess Return"], errors="coerce")  # market excess (X)

    # A tiny function: beta = how a stock's excess return moves with the market's.
    # Mathematically beta = covariance(stock, market) / variance(market) -- same as Excel SLOPE.
    def beta(yv, xv):
        d = pd.concat([yv, xv], axis=1).dropna()   # line up the two series, drop blanks
        x = d.iloc[:, 1].values; y = d.iloc[:, 0].values
        return np.cov(x, y, ddof=1)[0, 1] / np.var(x, ddof=1)

    # =======================================================================
    #  SECTION B: ROLLING -- for each window, measure betas, build 32 portfolios,
    #             and record how each portfolio performs the next year
    # =======================================================================
    betas = {}      # will store each window's 320 betas
    pfcols = {}     # will store each year's portfolio number for every stock
    # An empty table to fill: 60 months (2021-2025) as rows, 32 portfolios as columns.
    pr = pd.DataFrame(index=pd.period_range("2021-01", "2025-12", freq="M"),
                      columns=[f"P{i}" for i in range(1, 33)], dtype=float)

    for e0, e1, hold in PLAN:                 # loop over the 5 rolling windows
        # 'w' marks the months inside this 5-year estimation window.
        w = (sexc.index >= f"{e0}-01") & (sexc.index <= f"{e1}-12")
        mw = mexc[w]                          # market excess returns in that window
        # Measure beta for every stock over this window, then sort highest beta first.
        b = pd.Series({s: beta(sexc.loc[w, s], mw) for s in stocks}).sort_values(ascending=False)
        betas[f"Beta {e0}-{e1}"] = b          # save these betas for the output

        # Assign portfolios: top 10 betas -> P1, next 10 -> P2, ... last 10 -> P32.
        assign = {}
        for p in range(32):
            for s in b.index[p*N_PER_PF:(p+1)*N_PER_PF]:
                assign[s] = p + 1
        pfcols[f"Pf {hold}"] = pd.Series(assign)   # remember each stock's portfolio this year

        # Now compute each portfolio's monthly return for the HOLD year (the next year).
        hi = pd.period_range(f"{hold}-01", f"{hold}-12", freq="M")   # the 12 months to fill
        for p in range(1, 33):
            mem = [s for s in stocks if assign[s] == p]   # the 10 stocks in portfolio p
            # Portfolio return that month = simple average of its 10 members' returns.
            pr.loc[hi, f"P{p}"] = sret.loc[hi, mem].mean(axis=1).values

    # Portfolio EXCESS return = portfolio return minus that month's risk-free rate.
    pexc = pr.sub(rf.reindex(pr.index), axis=0)
    m = mexc.reindex(pr.index)                # market excess for the 60 test months

    # =======================================================================
    #  SECTION C: FIRST PASS -- one regression per portfolio (over 60 months)
    #             to get each portfolio's stable beta, alpha, residual variance.
    # =======================================================================
    meanexc = []; betal = []; rvl = []; rows = []      # lists to collect results
    for p in range(1, 33):
        # Pair this portfolio's excess return (Y) with the market excess (X), drop blanks.
        d = pd.concat([pexc[f"P{p}"], m], axis=1).dropna()
        yy = d.iloc[:, 0].values; xx = d.iloc[:, 1].values
        bta = np.cov(xx, yy, ddof=1)[0, 1] / np.var(xx, ddof=1)  # portfolio beta (slope)
        al  = yy.mean() - bta * xx.mean()                        # alpha (intercept)
        resid = yy - (al + bta*xx)                               # leftover error
        rv = np.var(resid, ddof=1)                               # residual variance
        r2 = np.corrcoef(xx, yy)[0, 1]**2                        # fit quality
        meanexc.append(yy.mean()); betal.append(bta); rvl.append(rv)  # save for 2nd pass
        rows.append([f"P{p}", yy.mean(), bta, bta**2, rv, al, r2])    # one output row
    # Put the first-pass results into a table; the headings carry the formulas.
    fp = pd.DataFrame(rows, columns=["Portfolio",
        "Mean Excess Return\n(= average of 60 monthly excess returns)",
        "FirstPass Beta β\n(= Cov(port excess, mkt excess) / Var(mkt excess))",
        "β²\n(= FirstPass Beta squared)",
        "Residual Variance\n(= Var of residuals; residual = port excess - (α + β·mkt excess))",
        "Alpha α\n(= mean(port excess) - β·mean(mkt excess))",
        "R²\n(= corr(port excess, mkt excess)²)"])

    # =======================================================================
    #  SECTION D: SECOND PASS -- ONE regression ACROSS the 32 portfolios.
    #             This is the actual CAPM test.
    # =======================================================================
    y = np.array(meanexc)          # Y = the 32 average excess returns
    bt = np.array(betal)           # X1 = the 32 betas
    bt2 = bt**2                    # X2 = beta squared (to test non-linearity)
    rv = np.array(rvl)             # X3 = residual variance (to test unsystematic risk)
    one = np.ones(32)              # a column of 1s, needed for the intercept
    mm = m.mean()                  # average market excess return (the risk premium)

    # Model 1: return = λ0 + λ1*beta            (the core CAPM line)
    c1, t1, r1 = ols(y, np.column_stack([one, bt]))
    # Model 2: add beta-squared                 (tests if the line is straight)
    c2, t2, r2 = ols(y, np.column_stack([one, bt, bt2]))
    # Model 3: add residual variance            (tests if non-market risk is paid for)
    c3, t3, r3 = ols(y, np.column_stack([one, bt, bt2, rv]))

    # Build plain-English verdicts from the numbers we just got.
    i1 = (f"λ₁={c1[1]:.4f} (t={t1[1]:.2f}, {sig(t1[1])}"
          f"{' & positive' if c1[1]>0 else ''}): "
          f"{'higher beta gives higher return -> supports CAPM core.' if (abs(t1[1])>2 and c1[1]>0) else 'beta link is weak.'} "
          f"λ₀ t={t1[0]:.2f}: {'differs from 0 (CAPM expects 0).' if abs(t1[0])>2 else 'about 0 (consistent with CAPM).'} "
          f"λ₁={c1[1]:.4f} vs avg market excess {mm:.4f}.")
    i2 = (f"λ₂ (β²) t={t2[2]:.2f}: "
          f"{'NON-linear -> violates CAPM.' if abs(t2[2])>2 else 'not significant -> beta-return relation is linear, consistent with CAPM.'}")
    i3 = (f"λ₃ (residual variance) t={t3[3]:.2f}: "
          f"{'significant -> unsystematic risk IS priced, a deviation from strict CAPM.' if abs(t3[3])>2 else 'not significant -> only systematic risk is priced, consistent with CAPM.'}")
    ov = (f"Overall: {'beta is significantly priced and positive' if (abs(t1[1])>2 and c1[1]>0) else 'beta not clearly priced'}; "
          f"relation is {'linear' if abs(t2[2])<2 else 'non-linear'}; "
          f"{'residual (unsystematic) risk is also priced, so strict CAPM is partly rejected.' if abs(t3[3])>2 else 'no residual-risk pricing, consistent with CAPM.'}")

    # Lay the second-pass results out as rows (blank cells where a term isn't in that model).
    sp_rows = [
      ["Model 1: MeanExcess = λ₀ + λ₁·β", c1[0],t1[0],c1[1],t1[1],"","","","",r1,i1],
      ["Model 2: + λ₂·β²", c2[0],t2[0],c2[1],t2[1],c2[2],t2[2],"","",r2,i2],
      ["Model 3: + λ₃·ResidVar", c3[0],t3[0],c3[1],t3[1],c3[2],t3[2],c3[3],t3[3],r3,i3],
      ["Reference: avg market excess return", mm,"","","","","","","","","CAPM expects λ₁ to be close to this value (the market risk premium)."],
      ["Legend", "","","","","","","","","","|t| > 2 means significant at about 5% (n = 32 portfolios)."],
      ["Overall conclusion", "","","","","","","","","",ov],
    ]
    sp = pd.DataFrame(sp_rows, columns=["Model",
        "λ₀ (intercept)","t(λ₀)","λ₁ (β)","t(λ₁)",
        "λ₂ (β²)","t(λ₂)","λ₃ (ResidVar)","t(λ₃)","R²",
        "Significance / Interpretation"])

    # =======================================================================
    #  SECTION E: ARRANGE THE FOUR DATA SHEETS (formulas shown in the headings)
    # =======================================================================
    # Sheet 1: each stock's beta per window + which portfolio it landed in.
    s1 = pd.DataFrame(index=stocks); s1.index.name = "Stock"
    for e0, e1, hold in PLAN:
        s1[f"Beta {e0}-{e1}\n(= Cov(stock excess, mkt excess)/Var(mkt excess))"] = betas[f"Beta {e0}-{e1}"]
        s1[f"Pf {hold}\n(rank beta high-low, 10 per pf)"] = pfcols[f"Pf {hold}"]
    s1 = s1.reset_index()

    # Sheet 2: the 32 portfolios' monthly returns (60 rows).
    s2 = pr.copy(); s2.columns = [f"{c}\n(= average of 10 members' returns)" for c in s2.columns]
    s2.insert(0, "Market Excess\n(from market file)", m.values)
    s2 = s2.reset_index().rename(columns={"index": "Date"}); s2["Date"] = s2["Date"].astype(str)

    # Sheet 3: the 32 portfolios' EXCESS returns (the previously-missing table).
    s3 = pexc.copy(); s3.columns = [f"{c}\n(= {c} return - RF monthly)" for c in s3.columns]
    s3.insert(0, "Market Excess\n(= mkt return - RF)", m.values)
    s3.insert(0, "RF monthly\n(from returns file)", rf.reindex(pr.index).values)
    s3 = s3.reset_index().rename(columns={"index": "Date"}); s3["Date"] = s3["Date"].astype(str)

    # =======================================================================
    #  SECTION F: THE README SHEET (definitions + every formula in words)
    # =======================================================================
    R = [
    ("HOW TO READ THIS WORKBOOK","6 sheets. Each computed column heading shows its formula in plain words. This sheet lists every term and formula."),
    ("",""),
    ("--- PLAIN-ENGLISH TERMS ---",""),
    ("Monthly return","How much a price moved in a month (log): LN(this month price / last month price). First month blank."),
    ("RF monthly","One-month risk-free return from the 91-day T-bill: (1 + annual_pct/100)^(1/12) minus 1."),
    ("Excess return","Return above the risk-free rate: return minus RF monthly. CAPM looks only at this."),
    ("Market excess return","Nifty 500 return minus RF monthly. The X (input) of the regressions."),
    ("Beta β","Sensitivity to the market: slope of (its excess return) on (market excess return). 1 moves with market, above 1 jumpier, below 1 calmer."),
    ("Portfolio","Basket of 10 stocks. 32 baskets (320/10). P1 = 10 highest-beta stocks ... P32 = 10 lowest."),
    ("",""),
    ("--- ALL FORMULAS (column heading: formula in words) ---",""),
    ("Sheet 1: Beta YYYY-YYYY","β = Cov(stock excess, market excess) / Var(market excess), over that 5-year window."),
    ("Sheet 1: Pf YYYY","Not a formula: sort the 320 betas highest-first, 10 per portfolio (P1 = highest)."),
    ("Sheet 2: P1..P32","Average of the 10 member stocks' monthly returns that month."),
    ("Sheet 2: Market Excess","Market return minus RF monthly (from the market file)."),
    ("Sheet 3: P1..P32","That portfolio's monthly return (Sheet 2) minus RF monthly. This is the portfolio EXCESS return series."),
    ("Sheet 4: Mean Excess Return","Average of that portfolio's 60 monthly excess returns."),
    ("Sheet 4: FirstPass Beta β","β = Cov(portfolio excess, market excess) / Var(market excess), over the 60 months."),
    ("Sheet 4: β²","FirstPass Beta squared (tests non-linearity in the second pass)."),
    ("Sheet 4: Residual Variance","Var of the first-pass residuals; residual = portfolio excess - (α + β·market excess). Tests unsystematic risk."),
    ("Sheet 4: Alpha α","α = mean(portfolio excess) - β·mean(market excess). The first-pass intercept; about zero under CAPM."),
    ("Sheet 4: R²","Corr(portfolio excess, market excess) squared."),
    ("Sheet 5: second pass","Across the 32 portfolios, regress Mean Excess Return on β (and β², Residual Variance). Models below."),
    ("",""),
    ("Portfolio EXCESS return","A monthly SERIES: 60 numbers per portfolio (Sheet 3). One value per month."),
    ("Mean Excess Return","A SINGLE number per portfolio (Sheet 4) = the average of those 60 monthly values."),
    ("",""),
    ("--- SECOND PASS: CROSS-SECTION REGRESSION (Sheet 5) ---",""),
    ("Model 1","MeanExcess = λ₀ + λ₁·β. CAPM holds if λ₀ is about 0 and λ₁ is positive. λ₁ estimates the market risk premium and should be close to the average market excess return."),
    ("Model 2","MeanExcess = λ₀ + λ₁·β + λ₂·β². Tests linearity: λ₂ should be about 0."),
    ("Model 3","MeanExcess = λ₀ + λ₁·β + λ₂·β² + λ₃·ResidualVariance. Tests whether unsystematic risk is priced: λ₃ should be about 0."),
    ("How to judge","Sheet 5 has a plain-English note per model. Rule of thumb: |t| > 2 means significant at about 5% (n = 32 portfolios)."),
    ]
    readme = pd.DataFrame(R, columns=["Topic", "Explanation"])

    # =======================================================================
    #  SECTION G: WRITE EVERYTHING TO ONE EXCEL FILE + tidy formatting
    # =======================================================================
    def style_headers(ws):                       # makes 2-line headings wrap & center
        ws.row_dimensions[1].height = 46
        for c in ws[1]:
            c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    with pd.ExcelWriter("capm_rolling_results.xlsx", engine="openpyxl") as xl:
        readme.to_excel(xl, sheet_name="ReadMe", index=False)
        s1.to_excel(xl, sheet_name="1_Betas_and_Portfolios", index=False)
        s2.to_excel(xl, sheet_name="2_Portfolio_Returns", index=False)
        s3.to_excel(xl, sheet_name="3_Portfolio_Excess", index=False)
        fp.to_excel(xl, sheet_name="4_FirstPass_Beta", index=False)
        sp.to_excel(xl, sheet_name="5_SecondPass_CrossSection", index=False)
        # --- pretty ReadMe ---
        ws = xl.sheets["ReadMe"]; ws.column_dimensions["A"].width = 40; ws.column_dimensions["B"].width = 95
        for c in ws[1]: c.font = Font(bold=True)
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top"); row[0].alignment = Alignment(vertical="top")
            if isinstance(row[0].value, str) and row[0].value.startswith("---"): row[0].font = Font(bold=True)
        # --- wrap the 2-line headings on the data sheets ---
        for name in ["1_Betas_and_Portfolios","2_Portfolio_Returns","3_Portfolio_Excess","4_FirstPass_Beta"]:
            style_headers(xl.sheets[name])
        # --- widen the interpretation column on sheet 5 ---
        w5 = xl.sheets["5_SecondPass_CrossSection"]; w5.column_dimensions["A"].width = 34
        w5.column_dimensions["K"].width = 70
        for c in w5[1]: c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="center")
        for r in w5.iter_rows(min_row=2): r[10].alignment = Alignment(wrap_text=True, vertical="top")

    # A short summary printed in the terminal so you can sanity-check the run.
    print("SECOND PASS (n=32):")
    print(sp[["Model","λ₀ (intercept)","t(λ₀)","λ₁ (β)","t(λ₁)","R²"]].head(3).round(4).to_string(index=False))
    print("avg market excess:", round(mm, 5))
    print("\nInterpretations:\n -", i1, "\n -", i2, "\n -", i3, "\n -", ov)


# This line means: only run main() when you execute this file directly.
if __name__ == "__main__":
    main()
