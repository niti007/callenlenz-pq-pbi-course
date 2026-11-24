#!/usr/bin/env python
"""Self-test for Builder C. Run from repo root: _build/.venv/bin/python _build/test_C.py"""
import json, os, sys
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import model_C as C

S = os.path.join(HERE, "..", "solutions")
c = json.load(open(os.path.join(S, "check_totals_C.json")))
a = json.load(open(os.path.join(S, "check_totals_A.json")))
b = json.load(open(os.path.join(S, "check_totals_B.json")))
Q = {q["id"]: q for q in c["questions"]}
AQ = {q["id"]: q for q in a["questions"]}; BQ = b
AI, BI = a["intermediate"], b
T = C.tables(); d = C.by_project(T)
fails = []
def check(cond, msg):
    print(("ok   " if cond else "FAIL ") + msg)
    if not cond: fails.append(msg)

lab = round(T["FactLabour"]["Labour Cost"].sum(), 2)
check(lab == 671900.00, f"labour cost Jan-Jun = {lab}")
check(AQ["Q9"]["labour_cost_ytd"] == 671900.0 and c["intermediate"]["labour_cost_jan_jun"] == 671900.0, "matches check_totals_A Q9 labour_cost_ytd")
check(round(T["FactLabour"].Hours.sum(), 2) == AQ["Q9"]["answer"], "labour hours = A Q9 19457.5")
check(len(T["FactLabour"]) == AI["L2_rows_valid_Jan_Jun"]["value"], "FactLabour rows = A L2_rows_valid_Jan_Jun")
check(round(T["FactPurchases"]["Line Value"].sum(), 2) == BI["L3_pl_total_value_ex_cancelled"]["value"], "all-dates PO value = B cleaned Purchase Log total")
check(len(T["FactPurchases"]) == BI["L3_pl_rows_after_remove_cancelled"]["value"], "FactPurchases rows = B 200")
sv = round(T["FactStockValuation"]["Stock Value"].sum(), 2)
check(sv == BQ["Q11"]["answer"] == Q["Q24"]["total"], f"stock value {sv} = B Q11")
check(abs(sum(Q["Q24"]["by_source"].values()) - BQ["Q11"]["answer"]) < 0.005, "Q24 sources sum to Q11")
bud = d.Budget["CL-P105"]
check(bud == AQ["Q16"]["answer"], "CL-P105 H1 budget = A Q16 198,551")
check(round(d.Cost.sum(), 2) == round(d.Labour.sum() + d.PO.sum(), 2), "total cost = labour + PO")
check(Q["Q21"]["project"] == d.Cost.idxmax() and Q["Q21"]["answer"] == round(d.Cost.max(), 2), "Q21")
check(Q["Q22"]["answer"] == round(d["Variance %"]["CL-P104"] * 100, 2), "Q22")
check(Q["Q23"]["project"] == d.Variance.idxmin(), "Q23")
check(abs(sum(v["total"] for v in Q["Q21"]["all_projects"].values()) - c["intermediate"]["total_project_cost_jan_jun"]) < 0.02, "Q21 per-project totals sum")
# YTD at end of May recomputed independently
lab5 = T["FactLabour"][T["FactLabour"]["Week Ending"] <= "2026-05-31"]["Labour Cost"].sum()
po5 = T["FactPurchases"][(T["FactPurchases"]["Order Date"] >= "2026-01-01") & (T["FactPurchases"]["Order Date"] <= "2026-05-31")]["Line Value"].sum()
check(Q["Q25"]["answer"] == round(lab5 + po5, 2), "Q25 YTD end of May")
check(c["intermediate"]["budget_jan_jun"] == round(T["FactBudget"][T["FactBudget"].Month <= "2026-06-01"]["Budget GBP"].sum(), 2), "budget Jan-Jun")
for i in ["Q1", "Q2", "Q21", "Q22", "Q23", "Q24", "Q25", "Q26"]:
    q = Q.get(i)
    ok = q is not None and q.get("answer") not in (None, "") and all(q.get(k) for k in ("question", "unit", "definition", "how", "lab"))
    check(ok, f"{i} has all fields and an answer")
md = open(os.path.join(S, "powerbi_build.md")).read() if os.path.exists(os.path.join(S, "powerbi_build.md")) else ""
for kw in ["DimDate", "FactLabour", "TOTALYTD", "Variance %", "Stock Value", "DIVIDE", "gateway", "Mark as date table"]:
    check(kw in md, f"powerbi_build.md mentions {kw}")
print("\nALL PASS" if not fails else f"\n{len(fails)} FAILED")
sys.exit(1 if fails else 0)
