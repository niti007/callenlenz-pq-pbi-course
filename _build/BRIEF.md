# Build brief: Callen-Lenz Power Query + Power BI course data

## Audience and purpose
- **Audience:** a 1-day in-person course for the Callen-Lenz Associates Ltd finance team, a UK manufacturer using GBP, VAT 20% and Xero.
- **Their level:** beginners in Power Query.
- **Their systems:** Timesheet system (time booked to projects), Wasp (stock / part numbers), Purchase Log (POs raised against projects), Xero (accounts).
- **What the data is for:** students clean it up step by step in Excel Power Query, then build a Power BI project-cost dashboard. It must be **realistic, relevant and heavily messy**. Every mess must have **one deterministic, teachable fix**, so the check figures are unambiguous.

All names, values and emails are fictional. Emails use `@callenlenz-demo.example`. Never copy anything from ~/Downloads.

## Tech rules
- **Python:** always use `_build/.venv/bin/python` (pandas, openpyxl and numpy are installed). Use seed 42 everywhere (`numpy.random.default_rng(42)` plus sub-seeds per dataset), so regeneration is identical.
- **Scripts:** live in `_build/`. Run them from the repo root: `/Users/tarunsachdeva/dev/callenlenz-pq-pbi-course`.
- **Real Excel formatting:** write xlsx with openpyxl when you need it (title rows, multi-line headers, blank columns, text-stored numbers). Data in data-tab sheets should NOT be an Excel Table, unless stated otherwise. Students create the tables themselves.
- **Money:** round to 2 dp. Store dates as real dates, unless the mess says text or serial.

## Phase 1: master data and clean truth (`_build/master.py`)
`master.py` exposes functions that return the CLEAN pandas DataFrames. Running it as a script writes:
- (a) the reference files to `data/00_Reference/`;
- (b) every clean truth table as CSV to `_build/truth/`, for builders and validators.

### Reference files (clean, students use them as lookup tables)

**Projects.xlsx**, 8 rows
- Columns: Project Code (CL-P101..CL-P108), Project Name (manufacturing/marine/defence-flavoured, e.g. "Maritime Sensor Pod", "Autonomous Survey Vessel"), Customer, Project Manager (an employee name), Start Date, Status (Active, with one "Closed"), Budget 2026 GBP.
- Codes that do NOT exist (CL-P199, CL-P100) are reserved for injected "unknown code" rows later.

**Employees.xlsx**, 25 rows
- Columns: Employee ID ("E001".."E025"), Employee Name, Department, Grade (G1..G5), Manager, Email (first.last@callenlenz-demo.example), Card Last Four (4-digit text, filled for exactly 10 cardholders, blank for the others).
- Departments: ENGINEERING, PRODUCTION, SUPPORT, MARITIME, PROCUREMENT, FINANCE, SLT.

**Labour_Rates.xlsx**
- Columns: Grade, Standard Hourly Rate GBP (G1 22.50 .. G5 65.00), Overtime Multiplier (1.5).
- Overtime rule (used in Lab 2): weekly hours per employee above 37.5 are paid at the multiplier.

**Suppliers_Master.xlsx**, 15 rows
- Columns: Supplier ID (S01..S15), Standard Supplier Name (e.g. "Apex Fasteners Ltd", "Northgate Composites Limited", "Solent Marine Electronics Ltd"), Category.

**Customers.xlsx**, 15 rows
- Columns: Customer ID (C01..C15), Customer Name (memorable two-word names that truncate badly, e.g. "Flintwood Marine Ltd", "Rabbitson Engineering", "Duckworth Defence Systems").

### Clean truth tables (`_build/truth/*.csv`)

**parts.csv**
- 60 parts, PRT-0001..PRT-0060.
- Columns: Description, Category (Fasteners, Composites, Electronics, Machined Parts, Cables, Consumables), Default Supplier ID, Stocking Unit, Wasp Item Cost (10–900), Total Available (0–40; about 8 parts have 0), Total In House, On Order, Committed, Location, Site.

**timesheets.csv**
- Weekly bookings from week ending Fri 2 Jan 2026 to Fri 26 Jun 2026.
- Columns: Week Ending, Employee ID, Employee Name, Project Code, Task (Design/Build/Test/Install/Admin), Hours.
- Volume: about 18 booking employees, 1–3 lines per employee per week. Weekly total per employee is 30–45 hours, so some overtime exists. About 350–450 rows per calendar month. Month = month of Week Ending.

**purchase_orders.csv**
- About 200 PO lines, Jan 2025 – Jun 2026.
- Columns: Order Date, PR Number (PR-10xxx; several lines can share one PR), PO Number (PO-55xxx), Part Number, Description, Qty, Net Unit Price, Supplier ID, Project Code, Department, Payment Terms (30 days / 60 days / Pro-forma), Jira (CL-2xxx).
- Pricing: a part bought several times should have prices that drift over time, so "latest" matters.

**xero_invoices.csv**
- About 150 payable invoice lines, Jan 2025 – Jun 2026, each tied to a PO line.
- Columns: Invoice Date (after the PO date), Contact (Supplier ID), Invoice Number (INV-xxxx), Reference (PO Number), Item Code (Part Number), Description, Quantity, Unit Price ex VAT.
- Coverage: about 40 of the 60 parts have at least one invoice. About 10 more have only POs. About 10 have neither, so they fall back to the Wasp Item Cost.

**receipts.csv**
- About 120 customer receipts, 1 Jan – 31 Mar 2026.
- Columns: Date, Ref (10010+), Customer ID, Amount.

**budget.csv**
- Project Code × month (Jan–Dec 2026), with Budget GBP for each project-month.
- The 12 months must sum to the Projects.xlsx Budget 2026 for each project.

**assets.csv**
- Asset Class (Plant & Machinery, Vehicles, IT Equipment, Fixtures & Fittings) × Month (Jan–Oct 2026).
- Columns: Additions, Disposals, Write Offs (all positive numbers).
- An Opening Cost per class at 1 Jan 2026 is also needed. Design it so that the **total across classes** in Month 10 (Oct) is: BF 250,000.00, Additions 10,000.00, Disposals 25,000.00, Write Offs 0, CF 235,000.00. This matches the client's own example.

**cards.csv**
- About 80 corporate card transactions per month, Jan–Mar 2026, for the 10 cardholders.
- Columns: Date, Posted Date, Merchant, Amount GBP, Category (Travel, Accommodation, Subsistence, Fuel, Materials, Software), City, Last Four, Name on Card, MCC, Trans Type, Bank Ref (unique), Original Currency Amount, Original Currency (GBP/USD/EUR, with about 8 non-GBP rows), Merchant Country.
- Merchants: Trainline, Premier Inn, Shell, Screwfix, RS Components, Microsoft, Travelodge, Pret, hotels abroad, etc.

`master.py` must also export `SUPPLIER_VARIANTS` and `CUSTOMER_VARIANTS` dicts. Each maps an ID to 3–5 realistic messy spellings: "Ltd" / "Limited" / "Ltd.", typos, ALL CAPS, truncated, extra words like "Recd from" / "BACS". Phase 2 builders use these.

## Phase 2: messy lab files (two builders in parallel)
Both builders import `master.py`. They don't write any file except their own outputs below. Each writes its own:
- `solutions/check_totals_{A|B}.json`: question id → value + one-line definition.
- `_build/mess_log_{A|B}.json`: a list of {file, sheet, location, mess_type, description, teaching_fix, lab}.
- `solutions/m/*.m` files.

**Check totals** must be computed by a script (`_build/build_{A|B}.py`) that applies the documented cleaning steps to the MESSY file with pandas, mimicking Power Query. Where it applies, assert that the result equals the clean truth.

### Mess catalogue
Use plenty of these, spread over the labs. Each lab should have 6+ distinct mess types:
- **Text:** leading, trailing and double spaces; non-breaking space CHAR(160); embedded line breaks; mixed case; Ltd / Limited / Ltd. variants; typos; truncated names.
- **Keys:** part numbers like "prt-12", "PRT-0012 ", "PRT12", "PRT-012", "PRT 0012". Project codes like "cl-p101", "CL P101", "CLP101". Unknown or retired codes. Employee ID as text "E007" in one file and number 7 in another.
- **Dates:** mixed UK dd/mm and US mm/dd text in one file (teach Using Locale). Text dates. Excel serial numbers. Blank dates. A typo year 2062 (rule: replace with 2026). Week-ending vs transaction dates.
- **Numbers:** "£1,250.00" stored as text; trailing "CR"; brackets "(250.00)" for negatives; blank vs 0; "N/A" and "#REF!" values (teach Replace Errors / try otherwise); VAT-inclusive vs exclusive confusion; FX rows.
- **Structure:** title rows above the header; multi-line headers with trailing spaces; blank columns; merged-cell-style blanks needing Fill Down; subtotal and "Total" rows inside the data; repeated header rows mid-file; wide month columns; a renamed column in one monthly file.
- **Duplicates:** exact duplicates; near-duplicates (case or space differences); the same card transaction on two statements.

### Builder A outputs

**02_Timesheets**
- `Timesheet_2026-01.csv` .. `Timesheet_2026-05.csv`, plus `_Later/Timesheet_2026-06.csv`.
- Mess:
  - The March file uses "Proj Code" instead of "Project Code" (it breaks the combine; the fix is a rename in the Transform Sample File).
  - The April file writes Week Ending as US mm/dd/yyyy text, while the other months use UK dd/mm/yyyy (so students teach Using Locale).
  - Lowercase and spaced project codes.
  - About 6 rows with unknown codes CL-P199 / CL-P100 (found with Left Anti).
  - Employee ID is a number in one file.
  - Hours are text with "hrs" in one file.
  - About 5 exact duplicate rows.
  - A "Total" row at the bottom of each file.
  - Blank lines.
- `Control_Totals.xlsx`: system hours per month, Jan–Jun. Exactly one month (Feb) is off by a stated amount, such as 7.5 hours, because of a duplicate or unknown-code row that students must explain.

**05_Budget_Assets**
- `Project_Budget_2026.xlsx`: a wide layout (Project Code, Project Name, Jan..Dec, Total).
  - Mess: a title row, a Total column and Total row to remove before Unpivot, month headers as "Jan-26"-style text, one project code in lowercase.
- `Asset_Movements_2026.xlsx`: monthly movements by class, in a long layout with mess.
  - Mess: Fill Down needed for Asset Class (merged-cell style), disposals as "(25,000.00)" text, one blank month.
  - Also sheet `Roll_Forward_Template`: the client's sheet-2 layout (Cost Bfwd, Additions, Disposals, Write offs, Cost Cfwd for Month 10, with Month 11 blank).
  - The lab uses a Period parameter so Month 11 fills on refresh. Include the November movements in a separate `_Later/Asset_Movements_2026-11.xlsx` drop-in.

**06_Cards**
- `Card_Statement_2026-01.csv` .. `-03.csv`, in the client's 15-column layout: Date, Posted date, Status, Merchant, Amount, Category, City, Last four, Name on card, MCC, Trans Type, Bank Ref, Original Currency Amount, Merchant Currency Description, Merchant Country.
- Mess:
  - Blank and 0 amounts (rule: a blank Amount on a non-GBP row = Original Currency Amount × a stated rate from `FX_Rates.xlsx`, which you create; a blank on a GBP row = Original Currency Amount).
  - "NEW  YORK" double spaces and mixed-case cities.
  - Last Four stored as number (loses a leading 0; at least one card starts with 0).
  - The same transactions (same Bank Ref) appear at the end of Jan and the start of Feb (dedupe on Bank Ref).
  - "Pending" status rows to exclude.
  - Refunds with "CR".
- `Employee_Lookup.xlsx`: the client's 3C layout (Employee, Department, Manager, Email), with a name-spelling mess vs the card file. The join is on Last Four via Employees, so the name mess is a distraction to discuss.
- `Expense_Template.xlsx`: the client's 3B header layout.

**M files**
- `lab2_timesheets.m` (the full combine, clean, merge rates, overtime, group by project-month).
- `lab2_control_checks.m`.
- `lab5_unpivot_budget.m`, `lab5_pivot_hours.m`, `lab5_rollforward.m`.
- `lab6_cards.m`, `lab6_fn_cardholder.m` (a function that takes a cardholder name and returns their rows).
- `lab6_power_automate.md`: step-by-step build of a cloud flow.
  - Trigger: manual or scheduled.
  - Steps: list rows in an Excel table `CardholderSummary` → Apply to each → Filter array on the transactions table → Create HTML table → Send an email (V2) to the Email column.
  - Also an alternative "one tab per person" via a Pivot Table "Show Report Filter Pages" trick, plus the Word mail-merge option.

**Questions owned by A:** Q6–Q9 (Lab 2), Q16–Q18 (Lab 5), Q19–Q20 (Lab 6). Design each question as a concrete student task with a numeric or exact answer. Examples:
- "Total labour cost for CL-P103 in Q1 including overtime".
- "Which month fails the control check and by how many hours, and why?"
- "How many unknown-code rows, and their hours?"
- "After adding June and refreshing, total hours YTD".
- "Unpivoted budget: total for CL-P105 in H1".
- "Cost Cfwd for Month 11".
- "Total GBP spend for cardholder X in Feb after dedupe and FX".

### Builder B outputs

**01_Wasp_Stock**
- `Wasp_Stock_Export.xlsx`: the client's 4A 24-column layout: Item Number, Type, Description, Total Available, Total Checked Out, Total In House, On Order, Committed, Tax Code, Item Cost, Cost Method (FIFO), Item List Price, Item Sale Price, Lot, Date Code, Manufacturer, Category, Stocking Unit, Alt Item Number, Serial Number, Attachment, Location, Site, Drawing Issue.
- Mess:
  - 2 export title rows.
  - Messy part numbers (all the variants listed above).
  - Item Cost as "£" text in some rows and "N/A" / "#REF!" in a few.
  - Blank rows.
  - A repeated header row mid-sheet.
  - Duplicate parts (same part, different case or spaces; rule: keep the first).
  - Description with embedded line breaks and CHAR(160).
  - Category blank where Fill Down applies (sorted by category, merged style).
- The Extract lesson: Description like "M6 Hex Bolt - SS316 [Drawing D-1043]" → extract the drawing number between delimiters.
- Conditional column: Stock Status (Out of stock / Reorder (≤5) / OK).

**03_Valuation**
- `Xero_Payable_Invoice_Detail.xlsx`: the client's 4B layout with 4 title rows ("Payable Invoice Detail", "Callen-Lenz Associates Limited", "For the period 1 January 2025 to 30 June 2026", blank). Columns: Invoice Date, Contact, Source, Reference, Item Code, Description, Quantity, Unit Price (ex) (GBP).
  - Mess: Contact uses SUPPLIER_VARIANTS; Item Code mess; a near-duplicate invoice line; a credit note line with negative qty; dates as text in one block.
- `Purchase_Log.xlsx`: the client's 4C wide layout (about 48 columns), with headers that contain line breaks ("Ord'd\nBy", "Ath'd\nBy"), trailing spaces ("Manufactured Part Numbers "), and about 5 completely blank columns with blank headers.
  - Mess: Net Unit Price text with "£"; repeated PR numbers; Notes like "PO-55123 / Jira CL-2231 / urgent" for Extract; Supplier variants; mixed date formats; a cancelled line ("CANCELLED" in Notes; rule: exclude).
- **Valuation rule (KR's request):** for each Wasp part:
  - Price = latest Xero invoice Unit Price (by Invoice Date, excluding credit notes).
  - Else the latest Purchase Log Net Unit Price (by Date, excluding cancelled lines).
  - Else the Wasp Item Cost.
  - Price Source column = "Xero" / "Purchase Log" / "Wasp".
  - Stock Value = Total Available × Price.

**04_Receipts**
- `Bank_Receipts_Q1.xlsx`: the client's 1A layout (Date, Ref, Description, Amount), about 120 rows.
  - Descriptions use CUSTOMER_VARIANTS ("Recd from FLINTWD MARINE", "BACS RABBITSON ENG", "Duckworth Def Sys - INV 2231").
  - Amounts include text "£", 3 duplicates and 1 refund "(150.00)".
- `Customer_List.xlsx`: the client's 1B layout (Customer, "£ paid?", "# payments" blank).
- `Fuzzy_Transform_Table.xlsx`: From / To columns.
- The teaching design should make a threshold of 0.8 alone miss some names, so the transform table plus a lower threshold fixes them. Document the exact expected matches. Because the fuzzy algorithm cannot be run here, design the variants so the intended answer comes from the transform table plus Extract (text after "Recd from " / "BACS ", before " - "). Note the fuzzy-only behaviour as "approximate; confirm on the day".

**M files**
- `lab1_wasp.m`.
- `lab3_xero.m`, `lab3_purchase_log.m`, `lab3_valuation.m`.
- `lab4_receipts_fuzzy.m` (use `Table.FuzzyNestedJoin` with `[Threshold=0.8, IgnoreCase=true, IgnoreSpace=true, TransformationTable=...]`).

**Questions owned by B:** Q3–Q5 (Lab 1), Q10–Q12 (Lab 3), Q13–Q15 (Lab 4). Concrete numeric or exact answers. Examples:
- "How many distinct valid parts after cleaning?"
- "Total stock units out of stock".
- "Total stock valuation £".
- "How many parts are priced from each source?"
- "Price used for PRT-0017 and why".
- "Total paid by Flintwood Marine in Q1 and number of payments".
- "Which receipts remain unmatched?"

## Phase 2.5: Power BI (Builder C, after Phase 2)
- Write `solutions/powerbi_build.md`:
  - The star schema: DimProject, DimEmployee, DimPart, DimDate (M or DAX CALENDAR); FactLabour (Lab 2 output), FactPurchases (Lab 3 PO clean), FactBudget (Lab 5 unpivot), FactStockValuation (Lab 3).
  - The relationships.
  - DAX measures: Labour Cost, PO Cost, Total Project Cost, Budget, Variance, Variance %, Stock Value, Labour Hours, plus one time-intelligence measure (YTD Cost).
  - A 1-page report layout.
  - Refresh and publish steps.
  - Concept notes: PQ vs Power Pivot vs PBI, and why "data cube" = semantic model (tabular).
- Write `check_totals_C.json` for Q21–Q26, computed from the clean outputs. Examples: "Total project cost CL-P101 H1", "Variance % for CL-P104 Jan–Jun", "Project with largest overspend".
- Questions Q1–Q2 are discussion questions with model answers (no figure).
