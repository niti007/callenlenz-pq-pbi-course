"""Clean-truth Power BI model (Builder C). Reproduces the star schema in solutions/powerbi_build.md."""
import os, sys
import pandas as pd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import master as M

HERE = os.path.dirname(os.path.abspath(__file__))
START, END = pd.Timestamp("2026-01-01"), pd.Timestamp("2026-06-30")


def tables():
    ts = M.timesheets()
    ts["Week Ending"] = pd.to_datetime(ts["Week Ending"])
    emp, rates = M.employees(), M.labour_rates()
    ts = ts.merge(emp[["Employee ID", "Grade"]], on="Employee ID", how="left").merge(rates, on="Grade", how="left")
    ts["Weekly Hours"] = ts.groupby(["Employee ID", "Week Ending"])["Hours"].transform("sum")
    ts["OT Hours"] = ((ts["Weekly Hours"] - 37.5).clip(lower=0)) * ts["Hours"] / ts["Weekly Hours"]
    r = ts["Standard Hourly Rate GBP"]
    ts["Labour Cost"] = ts["Hours"] * r + ts["OT Hours"] * r * (ts["Overtime Multiplier"] - 1)
    ts["Month"] = ts["Week Ending"].dt.to_period("M").dt.to_timestamp()
    fl = ts[["Week Ending", "Month", "Employee ID", "Project Code", "Task", "Hours", "OT Hours", "Labour Cost"]]
    po = M.purchase_orders()
    po["Order Date"] = pd.to_datetime(po["Order Date"])
    po["Line Value"] = (po["Qty"] * po["Net Unit Price"]).round(2)
    po["Month"] = po["Order Date"].dt.to_period("M").dt.to_timestamp()
    fp = po[["PO Number", "Order Date", "Month", "Project Code", "Part Number", "Qty", "Net Unit Price", "Line Value"]]
    fb = M.budget()
    fb["Month"] = pd.to_datetime(fb["Month"])
    v = pd.read_csv(os.path.join(HERE, "truth_B", "valuation.csv"))
    fv = v.rename(columns={"part": "Part Number", "source": "Price Source", "value": "Stock Value"})[["Part Number", "Price Source", "Stock Value"]]
    return dict(DimProject=M.projects(), DimEmployee=emp, DimPart=M.parts(), FactLabour=fl, FactPurchases=fp, FactBudget=fb, FactStockValuation=fv)


def period(df, col, lo=START, hi=END):
    return df[(df[col] >= lo) & (df[col] <= hi)]


def by_project(T, lo=START, hi=END):
    lab = period(T["FactLabour"], "Month", lo.replace(day=1), hi).groupby("Project Code")["Labour Cost"].sum()
    po = period(T["FactPurchases"], "Order Date", lo, hi).groupby("Project Code")["Line Value"].sum()
    bud = period(T["FactBudget"], "Month", lo.replace(day=1), hi).groupby("Project Code")["Budget GBP"].sum()
    d = pd.DataFrame({"Labour": lab, "PO": po, "Budget": bud}).fillna(0)
    d["Cost"] = d.Labour + d.PO
    d["Variance"] = d.Budget - d.Cost
    d["Variance %"] = d.Variance / d.Budget
    return d
