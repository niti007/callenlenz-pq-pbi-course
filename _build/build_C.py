#!/usr/bin/env python
"""Builder C: writes solutions/check_totals_C.json from clean truth. Run from repo root."""
import json, os, sys
import pandas as pd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import model_C as C

T = C.tables()
d = C.by_project(T)
d5 = C.by_project(T, C.START, pd.Timestamp("2026-05-31"))
r2 = lambda x: round(float(x), 2)
top = d.Cost.idxmax()
worst = d.Variance.idxmin()
sv = T["FactStockValuation"].groupby("Price Source")["Stock Value"].sum().round(2)
src = {k: r2(v) for k, v in sv.items()}

PERIOD = ("Period for Q21-Q25: Jan-Jun 2026, i.e. AFTER the June timesheet has been dropped in and refreshed "
          "(labour by month of Week Ending; purchases by Order Date; budget by Month). ")
MODEL = ("Model: FactLabour (Lab 2 LabourLines, 2,320 rows), FactPurchases (cleaned Purchase Log, 200 cancelled-free lines, Line Value = Qty x Net Unit Price), "
         "FactBudget (Lab 5 unpivot, 96 rows), FactStockValuation (Lab 3, 60 parts). Total Project Cost = Labour Cost + PO Cost. ")

Q = [
 {"id": "Q1", "lab": "Power BI", "unit": "text",
  "question": "Mini-quiz: classify each as Power Query, Power Pivot (Data Model) or Power BI. (a) Removing the Total row and fixing dates in a timesheet CSV. (b) Writing the measure Total Cost := [Labour Cost] + [PO Cost]. (c) A report page with slicers that you publish so the FD can open it in a browser. (d) Unpivoting the Jan-26...Dec-26 budget columns. (e) Relating Projects to Timesheets so one slicer filters both.",
  "answer": "(a) Power Query. (b) Power Pivot / Data Model (DAX) - and the same engine inside Power BI. (c) Power BI (report + Service). (d) Power Query. (e) Power Pivot / Data Model (relationships) - Power BI has the identical Model view.",
  "definition": "Power Query = get and shape (M, one table at a time, runs at refresh). Power Pivot = the Data Model inside Excel: VertiPaq in-memory columnar tables, relationships, DAX measures. Power BI Desktop = Power Query + the same Data Model + visuals; the Power BI Service adds sharing, scheduled refresh and dashboards. Same engines, three places to use them.",
  "how": "Ask: is it changing the shape/content of ONE table before loading (PQ), relating tables or calculating on filtered data (Model/DAX), or presenting/sharing (report/Service)?"},
 {"id": "Q2", "lab": "Power BI", "unit": "text",
  "question": "JE asks: should Callen-Lenz build ONE big model with everything (labour, purchases, stock, budget, cards, assets) or many small ones? What would you do?",
  "answer": "Start with one focused model per business question, built as a star schema: a 'Project Cost' model (labour, purchases, budget, stock valuation) sharing conformed dimensions (Project, Employee, Part, Date). Do not dump everything into one giant model: it is slow to refresh, hard to test and to secure, and every report author sees all the data (cards and payroll-type data need tighter access). Do not build many isolated models either: duplicated logic gives two answers for the same number. Practical route: one certified shared semantic model (dataset) per subject area, published once in the Service; people build thin reports on top of it ('live connection') and reuse the same measures. Add a second model (e.g. Treasury: cards, assets) only when the subject and audience differ.",
  "definition": "Rule of thumb: one model per subject area / audience, star schema, shared dimensions, measures defined once. Split when refresh time, security or ownership differ; merge when two reports keep needing the same dimensions and measures.",
  "how": "Discussion question; mark on: star schema mentioned, single source of truth, security/audience, refresh cost, avoids duplicate logic."},
 {"id": "Q21", "lab": "Power BI", "unit": "GBP", "answer": r2(d.Cost[top]),
  "question": "Build the model and the measure Total Project Cost = Labour Cost + PO Cost. With the date filter at Jan-Jun 2026 (June refreshed), which project has the highest Total Project Cost and what is it? (Answer: project code and amount.)",
  "answer_text": f"{top}: GBP {d.Cost[top]:,.2f} (Labour {d.Labour[top]:,.2f} + PO {d.PO[top]:,.2f})",
  "definition": PERIOD + MODEL + "Labour Cost per line includes the Lab 2 overtime rule. PO Cost = sum of Qty x Net Unit Price for Order Date in Jan-Jun 2026 (2025 lines excluded by the date filter).",
  "how": "Bar chart: DimProject[Project Code] x [Total Project Cost], filter DimDate to Jan-Jun 2026, sort descending.",
  "project": top, "labour": r2(d.Labour[top]), "po": r2(d.PO[top]),
  "all_projects": {k: {"labour": r2(v.Labour), "po": r2(v.PO), "total": r2(v.Cost), "budget": r2(v.Budget)} for k, v in d.iterrows()}},
 {"id": "Q22", "lab": "Power BI", "unit": "percent", "answer": round(float(d["Variance %"]["CL-P104"]) * 100, 2),
  "question": "What is the Variance % for CL-P104 for Jan-Jun 2026, where Variance = Budget - Total Project Cost and Variance % = Variance / Budget? (Give as a percentage, 2dp.)",
  "definition": PERIOD + MODEL + f"CL-P104: Budget {d.Budget['CL-P104']:,.2f}, Total Project Cost {d.Cost['CL-P104']:,.2f}, Variance {d.Variance['CL-P104']:,.2f}. Positive = under budget.",
  "how": "Variance % := DIVIDE([Variance],[Budget]); put on a card with Project slicer = CL-P104 and Month slicer Jan-Jun.",
  "budget": r2(d.Budget["CL-P104"]), "total_cost": r2(d.Cost["CL-P104"]), "variance": r2(d.Variance["CL-P104"])},
 {"id": "Q23", "lab": "Power BI", "unit": "text",
  "answer": f"{worst}: overspent by GBP {-d.Variance[worst]:,.2f} (cost {d.Cost[worst]:,.2f} vs budget {d.Budget[worst]:,.2f})",
  "question": "Jan-Jun 2026: which project has the largest overspend (most negative Variance) and by how much? Which other project is also over budget?",
  "definition": PERIOD + MODEL + "Overspend = Total Project Cost - Budget where positive. Only two projects are over budget in H1: " + ", ".join(f"{k} ({-v:,.2f})" for k, v in d.Variance[d.Variance < 0].sort_values().items()) + ".",
  "how": "Table by project sorted by Variance ascending; red conditional formatting on negative variance.",
  "project": worst, "overspend": r2(-d.Variance[worst]), "over_budget_projects": {k: r2(-v) for k, v in d.Variance[d.Variance < 0].items()}},
 {"id": "Q24", "lab": "Power BI", "unit": "text",
  "answer": f"Xero {src['Xero']:,.2f}; Purchase Log {src['Purchase Log']:,.2f}; Wasp {src['Wasp']:,.2f}; total {sum(src.values()):,.2f} (= Lab 3 Q11)",
  "question": "Stock Value by Price Source (donut/bar chart on FactStockValuation). Give the value for each of Xero, Purchase Log and Wasp, and show that they add up to the Lab 3 total valuation (check_totals_B Q11).",
  "definition": "Stock Value = Total Available x Price (KR's rule: latest Xero, else latest Purchase Log, else Wasp item cost), from the Lab 3 Stock_Valuation query; 60 parts (Xero 40, Purchase Log 10, Wasp 10). Not date-filtered (stock is a snapshot).",
  "how": "Stock Value := SUM(FactStockValuation[Stock Value]); axis = FactStockValuation[Price Source].",
  "by_source": src, "total": r2(sum(src.values())), "parts_by_source": {"Xero": 40, "Purchase Log": 10, "Wasp": 10}},
 {"id": "Q25", "lab": "Power BI", "unit": "GBP", "answer": r2(d5.Cost.sum()),
  "question": "YTD Total Cost at the end of May 2026: what does the measure YTD Total Cost = TOTALYTD([Total Project Cost], DimDate[Date]) show for all projects when the visual is at May 2026? What is it for June (the next month)?",
  "answer_text": f"May YTD: GBP {d5.Cost.sum():,.2f} (Labour {d5.Labour.sum():,.2f} + PO {d5.PO.sum():,.2f}); June YTD: GBP {d.Cost.sum():,.2f}",
  "definition": MODEL + f"1 Jan - 31 May 2026, all eight projects, labour by Week Ending, PO by Order Date. June YTD (Jan-Jun) = {d.Cost.sum():,.2f}; June month alone = {d.Cost.sum()-d5.Cost.sum():,.2f}.",
  "how": "Matrix/line: DimDate[Month] x [YTD Total Cost]; read May.",
  "labour_ytd_may": r2(d5.Labour.sum()), "po_ytd_may": r2(d5.PO.sum()), "ytd_june": r2(d.Cost.sum())},
 {"id": "Q26", "lab": "Power BI", "unit": "text",
  "question": "A student adds a slicer on DimEmployee[Employee Name] and a card showing Budget (= SUM(FactBudget[Budget GBP])). Whatever employee they click, the card shows the same number. Why, and what are the choices to fix it?",
  "answer": "There is no relationship path from DimEmployee to FactBudget: budget is held by project and month, not by employee. A slicer only filters tables it is related to (DimEmployee -> FactLabour), so FactBudget never receives the filter and the SUM ignores it. Choices: (1) accept it - 'budget has no employee grain'; it is the correct answer, and better to hide/blank the card when an employee is selected, e.g. Budget (safe) = IF(ISFILTERED(DimEmployee), BLANK(), [Budget]); (2) allocate the budget to employees in the model (needs a defined allocation rule, e.g. by hours share) - do not invent one; (3) do NOT 'fix' it with a bi-directional or many-to-many relationship between employee and budget - it hides the problem and gives wrong totals. The same logic explains why PO Cost does not react to a Department slicer: FactPurchases has no employee/department column.",
  "definition": "Filters flow from the one side of a relationship to the many side. FactBudget is related to DimProject and DimDate only. A measure over a table with no relationship to the slicer's table is unaffected by it.",
  "how": "Model view: check that DimEmployee has no line to FactBudget; test with a table visual of Employee x Budget (identical value on every row)."},
]
CHK = {
 "model_rows": {"DimProject": 8, "DimEmployee": 25, "DimPart": 60, "DimDate": "Jan 2025 - Dec 2026 = 730 rows (CALENDAR(DATE(2025,1,1), DATE(2026,12,31)))",
                "FactLabour": len(T["FactLabour"]), "FactPurchases": len(T["FactPurchases"]), "FactBudget": len(T["FactBudget"]), "FactStockValuation": len(T["FactStockValuation"])},
 "labour_hours_jan_jun": r2(T["FactLabour"].Hours.sum()),
 "labour_cost_jan_jun": r2(T["FactLabour"]["Labour Cost"].sum()),
 "po_cost_jan_jun_2026": r2(d.PO.sum()),
 "po_cost_all_dates_2025_2026": r2(T["FactPurchases"]["Line Value"].sum()),
 "po_lines_jan_jun_2026": int(len(C.period(T["FactPurchases"], "Order Date"))),
 "budget_jan_jun": r2(d.Budget.sum()), "budget_full_year_2026": r2(T["FactBudget"]["Budget GBP"].sum()),
 "total_project_cost_jan_jun": r2(d.Cost.sum()), "variance_jan_jun": r2(d.Variance.sum()),
 "stock_value_total": r2(T["FactStockValuation"]["Stock Value"].sum()),
}
for q in Q:
    q.setdefault("answer_text", q["answer"] if isinstance(q["answer"], str) else None)
    if q["answer_text"] is None: del q["answer_text"]
out = {"builder": "C", "period": "Jan-Jun 2026 (after June refresh)", "intermediate": CHK, "questions": Q}
p = os.path.join(C.HERE, "..", "solutions", "check_totals_C.json")
json.dump(out, open(p, "w"), indent=2, ensure_ascii=False)
print("wrote", os.path.abspath(p))
