# Power BI build guide (15:00-17:00) - Project Cost dashboard

Trainer-grade, click-by-click. Audience: beginner finance staff who finished the Excel Power Query labs this morning. Every number quoted here is checked in `solutions/check_totals_C.json` (Q21-Q26) and `_build/test_C.py`.

**Timing (120 min)**

| Time | Block |
|---|---|
| 15:00-15:15 | Concepts (script in section 1) |
| 15:15-15:40 | Get data: bring the queries in (section 2) |
| 15:40-16:05 | Star schema, relationships, DimDate (sections 3-4) |
| 16:05-16:25 | DAX measures (section 5) |
| 16:25-16:50 | Report page (section 6) |
| 16:50-17:00 | Refresh with June, publish, wrap-up (section 7) |

Period used in all checks: **Jan-Jun 2026, i.e. AFTER the June timesheet has been dropped in and refreshed.** Before the June refresh your Labour numbers will be lower (Jan-May labour only).

Prerequisite: the Power BI Desktop (free) is installed and the student data folder is the same one used in the labs (e.g. `C:\CallenLenz_Course\data`). Power BI Desktop is Windows-only; if someone is on a Mac, pair them up.

---

## 1. Concepts (10-15 minute script)

### 1.1 Power Query vs Power Pivot vs Power BI - one engine, three places

Say it like this:

> "You already know Power Query: it gets data and shapes it - one table at a time - and it runs when you press Refresh. Power Pivot is the *next* box in the chain: the Data Model. It stores your tables in memory (an engine called VertiPaq, which is columnar and compressed), lets you draw relationships between tables, and lets you write measures in a language called DAX. Power BI Desktop is simply both of those - the same Power Query, the same Data Model, the same DAX - plus a canvas of visuals. Power BI Service is where you publish it so other people can open it in a browser and where refresh can be scheduled."

Whiteboard:

```
 SOURCES --> POWER QUERY --> DATA MODEL (Power Pivot) --> REPORT (visuals)
 Excel,       get & shape     tables + relationships       charts, slicers
 CSV, SQL     (M language)    + measures (DAX)             Power BI Desktop
                              [same engine: VertiPaq]      --> Service (share, schedule)
```

| | Power Query | Power Pivot / Data Model | Power BI Desktop |
|---|---|---|---|
| Job | Get + shape data | Relate tables + calculate | Query + model + visualise |
| Language | M (buttons write it) | DAX | M + DAX + clicks |
| Lives in | Excel and Power BI | Excel (Manage Data Model) and Power BI | Power BI |
| Runs | At refresh | When you filter/click (calculated live) | Both |
| Limit | One table at a time | 1M-row Excel limit does not apply, but file size/Excel memory does | Data model limit is memory only |

Rule of thumb for the class: **"Shape it in Power Query, relate it in the model, calculate it in DAX, show it in a visual."** Push cleaning *upstream* (PQ), not into DAX.

Quick check - ask the room: "Unpivot the budget: which tool?" (PQ.) "Total Cost = Labour + PO?" (DAX.) "One slicer filters two tables?" (relationship.) See check_totals_C Q1.

### 1.2 "What are Data Cubes in Power BI?"

Someone always asks, because they have heard "cube" from the ERP/BI world.

- An **OLAP cube** (classic SQL Server Analysis Services multidimensional) pre-computes aggregates over *dimensions* (Time x Product x Region) and is queried with MDX. IT builds it; it is slow to change.
- Power BI does not have cubes. It has a **tabular semantic model** (formerly called a *dataset*): tables + relationships + measures, held in memory (VertiPaq) and queried with DAX. It does the same job as a cube - "slice a measure by dimensions" - but it is built by analysts, aggregates on the fly, and is quick to change.
- Mental translation: cube dimension = **Dim table**; cube measure = **DAX measure**; cube fact = **Fact table**; cube hierarchy = **hierarchy** (Year > Month).
- When people say "data cube in Power BI" they mean: **the Data Model - a star schema with measures**. (Power BI can also *connect to* an existing SSAS cube, live, if the company has one.)
- Hand-out line: *"The model is the cube. Build the model once, and every report is a different view of it."*

### 1.3 One big model or many small ones? (JE's question)

- Model per **subject area and audience**, not per department and not one-for-everyone. Callen-Lenz: one *Project Cost* model (labour, purchases, budget, stock valuation) is enough for today.
- Shared **dimensions** (Project, Employee, Part, Date) are reused so every number agrees. Measures are defined **once**.
- Split when: different owners, very different refresh times, or restricted data (cards/payroll). Merge when the same dimensions and measures keep being rebuilt.
- In the Service: publish the model once as a shared **semantic model**; others build thin reports on it (Get data > Power BI semantic models). Model answer in check_totals_C Q2.

### 1.4 How Power Query feeds Power BI (three routes)

1. **Copy the M (recommended today).** In Power BI: Home > Get data > Blank query > Advanced Editor > paste the M from `solutions/m/*.m`. Same M, same results. Needs the same `FolderPath` parameter.
2. **Import Excel workbook contents / Power Query template.** If the students' lab workbook still has its queries: in Power BI Desktop use *File > Import > Excel workbook contents* (this brings across the workbook's Power Query queries, tables and Data Model measures - choose *Start* and let it migrate), or export a template from Excel's Power Query Editor and use *File > Import > Power Query templates*. All queries and parameters arrive in one go; then continue at section 2.4.
3. **Get Data from the cleaned Excel output tables.** If the students loaded their results to sheets (tables like `LabourLines`, `PurchaseLog_Clean`), Get data > Excel workbook > tick the tables. Quick but **not refreshable from source** - it reads yesterday's clean output; refresh the Excel file first.

Trainer tip: route 1 for everyone; keep route 3 as the rescue for anyone whose queries are broken ("here's a workbook with the cleaned tables").

---

## 2. Get data (15:15-15:40)

### 2.1 Open and set up
1. Open Power BI Desktop. Close the splash screen. **File > Save as** `C:\CallenLenz_Course\CallenLenz_ProjectCost.pbix` (save early).
2. **Home > Transform data** opens the Power Query Editor - it looks exactly like Excel's.

### 2.2 Create the FolderPath parameter
1. In the Power Query Editor: **Home > Manage Parameters > New Parameter**.
2. Name `FolderPath`, Type `Text`, Current Value `C:\CallenLenz_Course\data\` (keep the trailing backslash; Labs 1, 3 and 4 need the trailing backslash; the others accept either). OK.

### 2.3 Bring in the queries (paste M)
Do each of these: **Home > New Source > Blank Query**, then **Home > Advanced Editor**, select all, paste the file contents, **Done**, then rename the query in the right-hand *Properties > Name* box. Do them in this order (later ones depend on earlier ones):

| # | File to paste | Rename query to |
|---|---|---|
| 1 | `solutions/m/lab1_wasp.m` | `Wasp_Clean` |
| 2 | `solutions/m/lab3_purchase_log.m` | `PurchaseLog_Clean` |
| 3 | `solutions/m/lab3_xero.m` | `Xero_Clean` |
| 4 | `solutions/m/lab3_valuation.m` | `Stock_Valuation` |
| 5 | `solutions/m/lab2_timesheets.m` | `LabourLines` (ends at the line-level table, step "Final Labour Lines") |
| 6 | `solutions/m/lab5_unpivot_budget.m` | `Budget_Unpivot` |
| 7 (optional) | `solutions/m/lab2_project_month_summary.m` | `ProjectMonthSummary` (Group By on LabourLines; connection only, not in the model) |

Also create `Projects_Ref` / `Employees_Ref` in 2.4 below.

Notes: `LabourLines` already ends at the line level (one row per cleaned booking line, 2,320 rows with June), so no editing of the M is needed. `Wasp_Clean` has its own `Stock Value` column (at Wasp cost); `Stock_Valuation` drops it internally and recomputes Stock Value with KR's price rule, so never take Stock Value from Wasp_Clean into the model.

Check each query for the yellow/red banner. If a query shows "Formula.Firewall: ... references other queries or steps": **File > Options > Current File > Privacy > Ignore the Privacy Levels**, or set all sources to *Organizational*. (Common with `Stock_Valuation`.)

### 2.4 Shape the model tables (reference queries)
Right-click a query > **Reference** to make a new query that builds on it; then name them exactly as below. Load only the tables below; everything else stays **Connection only** (right-click query > untick *Enable load*).

| Model table | Built from | Steps | Columns kept |
|---|---|---|---|
| `DimProject` | Blank query reading `00_Reference\Projects.xlsx` (`Excel.Workbook(File.Contents(FolderPath & "00_Reference\Projects.xlsx"), true){[Item="Sheet1",Kind="Sheet"]}[Data]`) | Remove `Budget 2026 GBP` (budget lives in FactBudget; keeping it invites double counting) | Project Code, Project Name, Customer, Project Manager, Status |
| `DimEmployee` | Same pattern, `Employees.xlsx` | **Remove** Email, Card Last Four, Manager (not needed, personal data) | Employee ID, Employee Name, Department, Grade |
| `DimPart` | Reference of `Wasp_Clean` | Keep columns (do NOT keep Wasp's Stock Value) | Part Number, Description, Category |
| `FactLabour` | Reference of `LabourLines` | Remove Employee Name, Project Name, Month, Weekly Hours, Standard Rate (Standard Rate may be kept for audit) | Week Ending, Employee ID, Project Code, Task, Hours, OT Hours, Labour Cost |
| `FactPurchases` | Reference of `PurchaseLog_Clean` | Rename `Date` > `Order Date`, `Net Value` > `Line Value`; keep columns | PO Number, Order Date, Project Code, Part Number, Qty, Net Unit Price, Line Value |
| `FactBudget` | Reference of `Budget_Unpivot` | Remove `Project Name` | Project Code, Month, Budget GBP |
| `FactStockValuation` | Reference of `Stock_Valuation` | Keep columns | Part Number, Price Source, Stock Value (optionally Total Available) |

Definitions check:
- **FactPurchases** is the cleaned Purchase Log: **cancelled lines already removed** (Lab 3), part/project codes normalised, **Line Value = Qty x Net Unit Price** (the `Net Value` column of Lab 3, rounded to 2dp). 200 rows in all, of which 51 are dated Jan-Jun 2026 (the rest are 2025 - the Date table will filter them out).
- **FactBudget** is 96 rows (8 projects x 12 months, Jan-Dec 2026). Month is the 1st of the month.
- **FactLabour grain** = one row per project x employee x week x task line (Lab 2 `LabourLines`). `Labour Cost` includes the overtime rule (weekly hours above 37.5 earn a 0.5 premium, pro-rated across the week's lines).
- Set data types: Week Ending / Order Date / Month = **Date**; Hours, Labour Cost, Line Value, Budget GBP, Stock Value = **Fixed decimal number** (or Decimal); codes = **Text**.

Click **Home > Close & Apply**. Save.

> **Alternative if the lab queries fail on someone's PC (route 3):** *Get data > Excel workbook* > pick the workbook where the cleaned tables were loaded > tick the tables > Load. Then do the column trims above in Transform data.

---

## 3. Star schema

```
                         DimProject (8)          DimEmployee (25)       DimPart (60)
                          Project Code               Employee ID          Part Number
                          /     |     \                  |                 /      \
                         *      *      *                 *                *        *
              FactLabour   FactPurchases   FactBudget                FactPurchases  FactStockValuation
              (2,320)      (200)           (96)                      
                  \            |            /
                   *           *           *
                          DimDate (Date)  <-- Week Ending / Order Date / Month
```

| Table | Type | Grain / key |
|---|---|---|
| DimProject | Dimension | 1 row per project; key **Project Code** |
| DimEmployee | Dimension | 1 row per employee; key **Employee ID** (Department and Grade live here) |
| DimPart | Dimension | 1 row per part; key **Part Number** |
| DimDate | Dimension | 1 row per calendar day (Jan 2025 - Dec 2026 = 730 rows) |
| FactLabour | Fact | project x employee x week x task line |
| FactPurchases | Fact | PO line (PO Number x Part) |
| FactBudget | Fact | project x month |
| FactStockValuation | Fact (snapshot) | 1 row per part (60) |

### 3.1 DimDate - DAX version (recommended, 3 minutes)
**Modeling > New table**, paste:

```dax
DimDate =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2025, 1, 1 ), DATE ( 2026, 12, 31 ) ),
    "Year", YEAR ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month", FORMAT ( [Date], "mmm" ),
    "Month Start", DATE ( YEAR ( [Date] ), MONTH ( [Date] ), 1 ),
    "Year-Month", FORMAT ( [Date], "yyyy-mm" ),
    "Quarter", "Q" & ROUNDUP ( MONTH ( [Date] ) / 3, 0 )
)
```
Then **sort `Month` by `Month Number`**: select the Month column > *Column tools > Sort by column > Month Number*.
Better: derive the range from the data so it never needs maintenance: replace the CALENDAR line with
`CALENDAR ( DATE ( YEAR ( MIN ( FactPurchases[Order Date] ) ), 1, 1 ), DATE ( YEAR ( MAX ( FactBudget[Month] ) ), 12, 31 ) )`.

### 3.2 DimDate - Power Query (M) alternative
New blank query `DimDate`:

```m
let
    StartDate = #date(2025, 1, 1),
    EndDate   = #date(2026, 12, 31),
    Dates     = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)),
    AsTable   = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}, null, ExtraValues.Error),
    Typed     = Table.TransformColumnTypes(AsTable, {{"Date", type date}}),
    AddYear   = Table.AddColumn(Typed, "Year", each Date.Year([Date]), Int64.Type),
    AddMNum   = Table.AddColumn(AddYear, "Month Number", each Date.Month([Date]), Int64.Type),
    AddMonth  = Table.AddColumn(AddMNum, "Month", each Date.ToText([Date], "MMM", "en-GB"), type text),
    AddStart  = Table.AddColumn(AddMonth, "Month Start", each Date.StartOfMonth([Date]), type date),
    AddQtr    = Table.AddColumn(AddStart, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text)
in
    AddQtr
```
Use either DAX or M - not both. (Teaching point: M is refreshed with the data; DAX is calculated in the model. Both fine for a calendar.)

### 3.3 Mark as date table
Select DimDate in the Data pane > right-click > **Mark as date table > Date column = `Date`** > OK. This is required for the time-intelligence functions (TOTALYTD, DATESYTD) to behave.

### 3.4 Relationships (Model view)
Click **Model view** (third icon, left rail). Drag from the **dimension key** to the **fact column**. Each must show **1 ... \*** (one-to-many), **single** cross-filter direction, **active**.

| From (one side) | To (many side) |
|---|---|
| DimProject[Project Code] | FactLabour[Project Code] |
| DimProject[Project Code] | FactPurchases[Project Code] |
| DimProject[Project Code] | FactBudget[Project Code] |
| DimEmployee[Employee ID] | FactLabour[Employee ID] |
| DimPart[Part Number] | FactPurchases[Part Number] |
| DimPart[Part Number] | FactStockValuation[Part Number] |
| DimDate[Date] | FactLabour[Week Ending] |
| DimDate[Date] | FactPurchases[Order Date] |
| DimDate[Date] | FactBudget[Month] |

Rules to teach:
- Filters flow **downhill**: dimension (the 1 side) filters the fact (the * side). Keep **single direction**. Never use bi-directional "to make it work".
- Facts never relate to each other. They connect through the shared dimensions (that is what makes "Cost vs Budget by Project" work).
- If Power BI proposes a *many-to-many* line, stop: there is a duplicate key in a dimension (check DimProject / DimPart for duplicates).
- Stock value has no date: DimDate does not filter it. That is correct (a snapshot). Say so out loud.
- FactBudget[Month] is the 1st of each month; relate it to DimDate[Date] and always slice budget by **Month**, not by individual days.

### 3.5 Hide the clutter
Select these and **right-click > Hide in report view** (or use the eye icon in Model view):
- All key columns on fact tables: FactLabour[Project Code, Employee ID, Week Ending], FactPurchases[Project Code, Part Number, Order Date], FactBudget[Project Code, Month], FactStockValuation[Part Number].
- All numeric columns we replace with measures: Hours, OT Hours, Labour Cost, Line Value, Qty, Net Unit Price, Budget GBP, Stock Value. (Measures are shown, columns hidden - this stops people dragging a column and getting an un-governed *Sum of ...*.)
- Optionally mark `Month Start`, `Month Number` hidden on DimDate.
Select all dimension columns you do want to filter on and leave them visible: DimProject[Project Code, Project Name, Customer, Project Manager], DimEmployee[Employee Name, Department, Grade], DimDate[Month, Year, Quarter], DimPart[Part Number, Category].

---

## 4. Create a measures table (tidy)
**Home > Enter data**, leave it empty, name the table `_Measures`, Load. Right-click each measure's home table later and **Move** it (or just create each measure with the `_Measures` table selected). Optional but helps beginners find things.

## 5. DAX measures (16:05-16:25)

Select the `_Measures` table, then **Modeling > New measure** for each. Copy exactly.

```dax
Labour Hours = SUM ( FactLabour[Hours] )
```
```dax
Labour Cost = SUM ( FactLabour[Labour Cost] )
```
```dax
PO Cost = SUM ( FactPurchases[Line Value] )
```
(Alternative if they ask: `SUMX ( FactPurchases, FactPurchases[Qty] * FactPurchases[Net Unit Price] )` - same answer to the penny except for rounding of the 2dp Line Value; teach the SUM version because it is simple and matches check_totals_C.)
```dax
Total Project Cost = [Labour Cost] + [PO Cost]
```
```dax
Budget = SUM ( FactBudget[Budget GBP] )
```
```dax
Variance = [Budget] - [Total Project Cost]
```
```dax
Variance % = DIVIDE ( [Variance], [Budget] )
```
```dax
Stock Value = SUM ( FactStockValuation[Stock Value] )
```
```dax
YTD Total Cost = TOTALYTD ( [Total Project Cost], DimDate[Date] )
```
Equivalent (the one that is easier to extend): `YTD Total Cost = CALCULATE ( [Total Project Cost], DATESYTD ( DimDate[Date] ) )`.

Optional, teaches `ISFILTERED`/`IF` (see Q26):
```dax
Budget (safe) = IF ( ISFILTERED ( DimEmployee ), BLANK (), [Budget] )
```

**Format strings** (select the measure > *Measure tools* ribbon):

| Measure | Format | Decimals | Notes |
|---|---|---|---|
| Labour Hours | Decimal number, thousands separator | 1 | |
| Labour Cost, PO Cost, Total Project Cost, Budget, Variance, Stock Value, YTD Total Cost | Currency > English (United Kingdom) GBP, or custom `£#,##0;(£#,##0)` | 0 (cards) / 2 (checks) | Negative in brackets reads naturally to finance |
| Variance % | Percentage | 1 | Custom `0.0%;(0.0%)` |

**Why `DIVIDE`?** `/` errors when Budget is zero or blank (e.g. a month with no budget); `DIVIDE` returns blank. Say so.

**Teaching the filter context in 2 minutes:** drop `Project Code` and `Total Project Cost` into a table. Ask: "why is each row different when the formula is the same?" Answer: *filter context* - each row filters the model to that project; the relationship carries the filter to the fact tables.

### Checkpoint values (Jan-Jun 2026 with DimDate filter = 2026, months Jan-Jun, June refreshed)

| Measure | Value |
|---|---|
| Labour Hours | 19,457.5 |
| Labour Cost | £671,900.00 |
| PO Cost | £608,095.02 |
| Total Project Cost | £1,279,995.02 |
| Budget | £1,894,429.00 |
| Variance | £614,433.98 (32.4% under budget) |
| Stock Value (no date) | £293,947.12 |

Per-project (Jan-Jun): see `check_totals_C.json` Q21.

> Common gotcha: if students see PO Cost of £2,104,847.74 they have not filtered to 2026 - the Purchase Log contains 2025 orders. Add a Year slicer = 2026 or a page-level filter on DimDate[Year] = 2026 (and Month <= Jun).

---

## 6. The report page (16:25-16:50)

Go to **Report view**. Page size: View > Page view > Fit to page (16:9). Rename the page "Project Cost".

### Wireframe (1 page)

```
+-----------------------------------------------------------------------------------+
| CALLEN-LENZ  Project Cost Dashboard  (Jan-Jun 2026)                               |
+--------------+--------------+--------------+--------------+----------------+------+
| Total Cost   | Budget       | Variance     | Variance %   | Stock Value    |SLICER|
| £1,279,995   | £1,894,429   | £614,434     | 32.4%        | £293,947       |Month |
| (card)       | (card)       | (card)       | (card)       | (card)         | (Jan-|
+--------------+--------------+--------------+--------------+----------------+ Jun) |
| Cost vs Budget by Project             | Monthly Cost vs Budget (line)      |      |
| (clustered bar: Total Project Cost    |  lines: Total Project Cost, Budget |SLICER|
|  and Budget by Project Name)          |  axis: DimDate[Month]              |Project|
|  CL-P101 ||||||||                     |      /\    ___                     |      |
|  CL-P102 |||||||||||||                |  ___/  \__/   \                    |SLICER|
+---------------------------------------+------------------------------------+Dept  |
| Project table / matrix                | Stock Value by Price Source        |      |
| Project | Cost | Budget | Var | Var%  |  (donut: Xero | Log | Wasp)        |      |
| CL-P101 | 317k | 316k   | -2k | -0.6% |                                    |      |
| (conditional formatting on Var / Var%)|                                    |      |
+---------------------------------------+------------------------------------+------+
```
(Variance % card = Variance / Budget = 614,434 / 1,894,429 = **32.4%** for all projects; CL-P104 alone is **57.8%**.)

### Build steps
1. **KPI cards** (Visualizations > Card): drag `Total Project Cost`, `Budget`, `Variance`, `Variance %`, `Stock Value` one per card. Format > Callout value > display units *None*/Thousands. Title on. (Use the new *Card* or *Multi-row card*, both fine.)
2. **Clustered bar chart** - *Cost vs Budget by Project*: Y-axis = `DimProject[Project Name]`, X-axis = `Total Project Cost` and `Budget`. Sort by Total Project Cost descending (... menu > Sort axis). Colour: cost one colour, budget a neutral grey.
3. **Line chart** - *Monthly Cost vs Budget*: X-axis = `DimDate[Month]` (sorted by Month Number), Y = `Total Project Cost` and `Budget`. If it shows the whole of 2025-2026, set a visual-level filter on DimDate[Year] = 2026.
4. **Table / matrix**: Rows `DimProject[Project Code]`, `DimProject[Project Name]`; Values `Total Project Cost`, `Budget`, `Variance`, `Variance %`. **Conditional formatting**: on the field well click the arrow beside `Variance` > *Conditional formatting > Font color*; Format style *Rules*: if value < 0 then red; if >= 0 then green. For `Variance %` use *Data bars* or an icon (red triangle for < 0). Under *Cell elements* switch it on. (Two projects are red in H1: CL-P107 -£27,030 and CL-P101 -£1,767.)
5. **Donut** (or bar) - *Stock Value by Price Source*: Legend = `FactStockValuation[Price Source]`, Values = `Stock Value`. Details labels on, showing value and %. Expect: Xero £180,256.41, Purchase Log £59,526.00, Wasp £54,164.71.
6. **Slicers**: (a) `DimProject[Project Code]` - Format > Slicer settings > Options > style *Dropdown* or *List*; (b) `DimEmployee[Department]` list; (c) `DimDate[Month]` - or better a *Between* slicer on `DimDate[Date]`. For budget, month-level picks are safest because the budget row sits on the 1st of the month. For a clean default: add a **page-level filter** DimDate[Year] = 2026 and DimDate[Month Number] <= 6.
7. **Sync/clean**: Format > Edit interactions if the stock donut should *not* respond to date slicers (it is a snapshot: select the Month slicer > Format > Edit interactions > donut > None). Align and distribute (Format > Align). Add a text box title. Lock objects (View > Lock objects) before handing the page over.

> Teaching moment: select Department = FINANCE. Labour Cost moves, but PO Cost and Budget do not - because FactPurchases and FactBudget have no relationship to DimEmployee. That is correct and a good prompt for Q26 and for "what do you want to see?" (A purchase has no department in the cleaned log; the dashboard says honestly "n/a".) Don't "fix" with bi-directional relationships.

---

## 7. Refresh, publish, schedule (16:50-17:00)

### 7.1 The June refresh (live demo, 3 minutes)
Before: build and number-check with Jan-May data (the `_Later` file not in folder), noting Labour Cost and YTD at May.
1. Copy `Timesheet_2026-06.csv` from `02_Timesheets\_Later` into `02_Timesheets`.
2. Power BI Desktop > **Home > Refresh**.
3. Watch the row counts: FactLabour 1,961 > 2,320; Labour Hours 16,470.0 > 19,457.5; Labour Cost £671,900.00; June month alone adds £103,156.25 labour cost (Jan-May was £568,743.75).
Nothing else was touched: no copy/paste, no formula edits - say "That is the whole point of the morning."

(YTD Total Cost at May-end is £1,084,717.25; at June-end £1,279,995.02 - check_totals_C Q25.)

### 7.2 Publish to the Service
1. Sign in (top right) with the work account. **Home > Publish** > choose a workspace (e.g. *Finance - Project Cost*, **not** "My workspace" if others must see it).
2. Open the link "Open ... in Power BI". Report appears; the **semantic model** appears next to it in the workspace.
3. Share: **Share** button (needs Pro/Premium Per User or a Fabric capacity for others to open it). Or **Publish to web** - **never** for company data.

### 7.3 Scheduled refresh - the gateway question
- The data sits in **local files** (`C:\CallenLenz_Course\data`). The cloud Service cannot reach your C: drive. For scheduled refresh you need either:
  - an **on-premises data gateway** (installed on an always-on PC/server that can see the folder; Settings > semantic model > Gateway connection > Add data source; credentials Windows), or
  - **move the files to SharePoint/OneDrive for Business** and switch the sources to the *SharePoint folder* connector (`SharePoint.Files("https://tenant.sharepoint.com/sites/Finance", [ApiVersion = 15])`) - **recommended**: no gateway, the cloud reads them directly, and the owners keep dropping the monthly timesheet into the library.
- Then: workspace > semantic model > **Settings > Scheduled refresh** > On > Frequency Daily, add times (Pro: 8/day) > Apply. Set *Send refresh failure notifications*.
- Parameters: when you move the data, only the **FolderPath parameter** (and the `Folder.Files` call) change - one reason we parameterised.

### 7.4 File-saving and link-breaking tips (read this to the class)
- **Save As .pbix on a local drive first**, not straight on a network share - saving to a flaky share can corrupt the file. Keep a version suffix: `_v1`, `_v2`.
- **Do not rename or move** the data folders after you build: every query points at them. The fix, if you must: Home > Transform data > **Manage Parameters > FolderPath** (one edit), then Refresh. *File > Options > Data source settings* shows (and can change/clear) the stored paths and credentials.
- **Do not edit the source files while the report is refreshing**; Excel lock files (`~$...xlsx`) in the folder will break `Folder.Files` - the lab M already filters on `Timesheet_*.csv`, but an open Excel on `Purchase_Log.xlsx` causes a "file in use" error. Close it, refresh.
- The .pbix contains the data: it can be big and **contains payroll-derived numbers**. Don't email it; share via the Service.
- "Credentials" errors after moving a file: *File > Options and settings > Data source settings > Edit permissions > Clear*, then Refresh and sign in again.
- Don't hard-code paths in M. Always `FolderPath & "..."`.
- Back up the **PBIT template** (File > Export > Power BI template) before experiments: it has the model and queries without the data.

---

## 8. Trainer answer sheet (see `solutions/check_totals_C.json`)

| Q | Answer (Jan-Jun 2026, June refreshed) |
|---|---|
| Q21 | CL-P101 has the highest Total Project Cost: £317,332.65 (Labour £117,182.60 + PO £200,150.05) |
| Q22 | CL-P104 Variance % = 57.76% (Budget £284,408 - Cost £120,135.09 = £164,272.91) |
| Q23 | Largest overspend: CL-P107, £27,030.35 over (cost £190,414.35 vs budget £163,384). CL-P101 also over, by £1,766.65 |
| Q24 | Stock Value: Xero £180,256.41 + Purchase Log £59,526.00 + Wasp £54,164.71 = £293,947.12 = Lab 3 Q11 |
| Q25 | YTD Total Cost at end of May = £1,084,717.25 (June YTD £1,279,995.02) |
| Q26 | Budget card ignores the Employee slicer: no relationship path DimEmployee > FactBudget (budget is project x month). Options: accept / blank with ISFILTERED / allocate with an agreed rule. Never bi-directional |
| Q1 | (a) PQ (b) Model/DAX (c) Power BI (d) PQ (e) Model/relationship |
| Q2 | One focused star-schema model per subject area with shared dimensions; separate models only for different audience/security/refresh |

**Intermediate checkpoints:** DimProject 8, DimEmployee 25, DimPart 60, DimDate 730, FactLabour 2,320, FactPurchases 200 (51 in Jan-Jun 2026), FactBudget 96, FactStockValuation 60; Labour Cost £671,900.00 (matches check_totals_A Q9); PO Cost Jan-Jun 2026 £608,095.02 (all dates £2,104,847.74 = check_totals_B cleaned log total); Budget Jan-Jun £1,894,429.00 (full-year £3,705,000.00).

### Troubleshooting
| Symptom | Cause / fix |
|---|---|
| Same value for every project in a table | Missing/inactive relationship from DimProject to that fact; check Model view |
| Totals look 3x too high | A many-to-many/duplicate key in a dimension, or wrong direction |
| Date slicer shows 2025 dates | DimDate range wider than the data - add page filter Year = 2026 |
| Months sort alphabetically (Apr, Aug...) | Sort Month by Month Number |
| Budget shows only on the 1st | Day-level date slicer; use Month level |
| Variance % shows 0.4 not 40% | Format the measure as Percentage |
| TOTALYTD returns blank | DimDate not marked as date table, or visual not using DimDate[Date] on the axis |
| "Formula.Firewall" | Privacy levels: Options > Current file > Privacy > Ignore |
| Refresh fails after folder move | Edit FolderPath parameter; clear credentials |
