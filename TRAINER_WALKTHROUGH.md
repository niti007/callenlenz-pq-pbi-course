# Callen-Lenz — Power Query for Finance + Power BI: Trainer Walkthrough

**How to use this document**
1. Read it top to bottom: each block says what to teach first, then the student question, then the solution directly below it.
2. Each "Teach" step names the specific mess in the data it fixes and the checkpoint to hit before you move on.
3. Only the `data/` folder goes to students. Solutions, M code and checkpoints in this document are for the trainer.

---

## Day run sheet

| Time | Block | Questions |
|---|---|---|
| 9:00–9:20 | 0. Welcome & landscape: PQ vs Power Pivot vs PBI, what M is, Input / Query / Output workbook structure, one big model vs many small, how Applied Steps work | Q1–Q2 |
| 9:20–10:15 | Lab 1: Wasp stock clean-up | Q3–Q5 |
| 10:15–10:30 | Break | |
| 10:30–11:30 | Lab 2: Timesheet consolidation (From Folder) + control checks + refresh with June | Q6–Q9 |
| 11:30–12:30 | Lab 3: Stock valuation template (Wasp + Xero + Purchase Log) | Q10–Q12 |
| 12:30–13:15 | Lunch | |
| 13:15–14:00 | Lab 4: Fuzzy matching receipts to customers | Q13–Q15 |
| 14:00–14:35 | Lab 5: Unpivot budget, Pivot hours, monthly roll-forward, version control & saving files | Q16–Q18 |
| 14:35–15:00 | Lab 6: Card statements + split per person + Power Automate demo (trainer-led) | Q19–Q20 |
| 15:00–17:00 | Power BI: project-cost dashboard (with a 10-minute break at about 16:00) | Q21–Q26 |

---

## Client requests → block mapping

| # | Request (verbatim) | Who | Block |
|---|---|---|---|
| R1 | "Consolidating data from multiple sources." | JE | Lab 2, Lab 3 |
| R2 | "Best practice in structuring data" | JE | Block 0, Lab 2 |
| R3 | "Is there a way to build in control checks? Ie. Ensuring data is all valid and pulling through." | JE | Lab 2 |
| R4 | "Is it better to have multiple, isolated powerquery documents – or is it better to have less, larger models?" | JE | Block 0 (Q2), Power BI |
| R5 | "How does Powerquery feed into PowerBI?" | JE | Power BI |
| R6 | "How can you manage version control without losing the Power Query rules - if you want to keep each month's query separate and not overwrite it" | CP | Lab 5 |
| R7 | "Doing lookups/sumifs between data where you have say a list of Supplier Names that's defined but those Names have been included in the description field in a different report sometimes not with an exact match - fuzzy matching" | PG | Lab 4 |
| R8 | "Can it be used to roll forward spreadsheets each month" | CP | Lab 5 |
| R9 | "Can PowerQuery be used to split data into tabs with each change in name, inputting data into specific columns and create a mail merge to send to individuals?" | CP | Lab 6 |
| R10 | "Create a valuation template consolidated from 4 different reports" (Wasp part numbers, Xero latest invoice, Purchase Log latest order if no invoice) | KR | Lab 3 |
| R11 | "How do PowerQuery, PowerPivot and PowerBI work together?" | CP | Block 0 (Q1), Power BI |
| R12 | "What are Data Cubes in PowerBI" | CP | Power BI |
| R13 | "What to look out for in terms of how to save our files - what would cause a power query link to be lost, can it be re-linked?" | CP | Lab 5, Appendix |
| R14 | "What is Power Query M Language" | CP | Block 0, Appendix |

Functions the brief asks us to teach by hand, and where: **Extract** (Lab 1 drawing number, Lab 3 Notes, Lab 4 payer). **Conditional Columns** (Lab 1 Stock Status, Lab 3 Price Source). **Merge Queries** (Labs 2, 3, 4, 6). **Data from Folder** (Lab 2). **Unpivot** (Lab 5).

---

## Before the day

- [ ] Copy the course `data/` folder to `C:\CallenLenz_Course\data` on every training PC. Every query reads `FolderPath`, which points at this folder with a trailing backslash (`C:\CallenLenz_Course\data\`).
- [ ] Excel 365 for Windows on every PC. Fuzzy merge and some connectors are limited on Mac, so do not use Mac for the labs.
- [ ] Power BI Desktop installed (needed for the 15:00 block).
- [ ] Dry-run every lab once before the day. The M code was checked by review and the figures by independent Python recomputation. **The M was NOT run in Excel.** Expect to fix small things (a column name, a step order) on the first run. The checkpoints in this document are the figures to check against.
- [ ] Test the Microsoft 365 mailbox used for the Power Automate demo (Lab 6, trainer-led).
- [ ] Projector zoom set so the Applied Steps pane and the formula bar are both readable.
- [ ] Students receive only the `data/` folder. There is no student handout; this document is the trainer's script.

---

## 0. Welcome & landscape ⏱ 9:00–9:20

> **Who asked:** "Best practice in structuring data" (JE); "How do PowerQuery, PowerPivot and PowerBI work together?" (CP); "Is it better to have multiple, isolated powerquery documents – or is it better to have less, larger models?" (JE); "What is Power Query M Language" (CP)
> **Why it matters to them:** Timesheets, Wasp stock, the Purchase Log and Xero all need the same three-stage pipeline. Once they can see which stage a job belongs to, they stop fixing the same mess in three places.

**Files:** none for this block. Use a blank Excel 365 workbook.

### Teach — step by step

**Step 1 – Draw the pipeline**
- Click path: none yet. Open a blank workbook and point at Data ▸ Get Data, the entry point for every source.
- 🗣 Say: "These three are one pipeline, not three competing products. Power Query cleans one table at a time. The Data Model (Power Pivot) relates the tables and holds the sums. Power BI, or an Excel PivotTable, shows the answer."
- Draw this on the board:

```text
 Sources                Power Query               Data Model / Power Pivot          Output
 (Excel, CSV,     ──▶   (get & clean ONE table,  ──▶ (relationships + DAX          ──▶ Excel PivotTables
  folders, Xero,        runs at refresh,            measures, VertiPaq in memory)      or Power BI visuals
  Purchase Log)         M language)                                                    (report + Service)
```

- 🧹 Mess this fixes: none yet. Students meet the real mess in Lab 1.
- ✅ Checkpoint: students can say which stage does which job. Q1 checks this at the end of the block.

**Step 2 – Show that the clicks are M**
- Click path: Data ▸ Get Data ▸ Launch Power Query Editor. Then Home ▸ Advanced Editor.
- Show the query on screen. Each Applied Step on the right is one line of M in the editor.
- The first lines of the Wasp query (from Lab 1) are a good example to show:

```m
let
    Source = Excel.Workbook(File.Contents(FolderPath & "01_Wasp_Stock\Wasp_Stock_Export.xlsx"), null, true),
    StockSheet = Source{[Item = "Stock", Kind = "Sheet"]}[Data],
    RemovedTopRows = Table.Skip(StockSheet, 2)
in
    RemovedTopRows
```

- 🗣 Say: "M is the language behind every click. `Table.Skip(StockSheet, 2)` means 'skip the first two rows'. `Excel.Workbook(...)` means 'open this workbook'. You rarely type M from scratch. You edit it when a fix is beyond the buttons."
- 🧹 Mess this fixes: none yet.
- ✅ Checkpoint: each Applied Step has a matching line of M. Rename a step in the pane and show it changes in the editor.

**Step 3 – Input / Query / Output**
- Click path: Data ▸ Queries & Connections to show the queries pane.
- 🗣 Say: "Input is the untouched files in the data folder. Query is the list of Power Query steps that clean them. Output is what you load back to the sheet or the model: a table, a PivotTable, a chart."
- 🧹 Mess this fixes: none yet.
- ✅ Checkpoint: each student can point to the input, query and output layers in their own workbook.

**Step 4 – The three golden rules**
- Never edit the raw file. Clean in a query, so the mess stays in the input and the fix is repeatable. (Lab 1: the Wasp export has two title rows; delete them in the query, never in Excel.)
- Name every step and every query. "Changed Type1" tells nobody anything. Rename steps to what they do, for example "Dates converted", and name queries `Wasp_Clean`, `LabourLines`.
- One query per job. One query loads and cleans one source. Merge queries combine them; do not build one giant query that does everything.
- ✅ Checkpoint: a student renames "Changed Type1" to "Dates converted" in one of their queries.

### 🙋 Student question Q1 – Which tool does which job?
Mini-quiz: classify each as Power Query, Power Pivot (Data Model) or Power BI. (a) Removing the Total row and fixing dates in a timesheet CSV. (b) Writing the measure Total Cost := [Labour Cost] + [PO Cost]. (c) A report page with slicers that you publish so the FD can open it in a browser. (d) Unpivoting the Jan-26...Dec-26 budget columns. (e) Relating Projects to Timesheets so one slicer filters both.

<details open><summary><b>Solution Q1</b></summary>

- Steps: ask one question per item. Is it changing the shape or content of ONE table before it loads? That is Power Query. Is it relating tables or calculating on filtered data? That is the Data Model (DAX). Is it presenting or sharing? That is Power BI.
- Definition to say out loud: Power Query = get and shape (M, one table at a time, runs at refresh). Power Pivot = the Data Model inside Excel: VertiPaq in-memory columnar tables, relationships, DAX measures. Power BI Desktop = Power Query + the same Data Model + visuals; the Power BI Service adds sharing, scheduled refresh and dashboards. Same engines, three places to use them.
- **Answer: (a) Power Query. (b) Power Pivot / Data Model (DAX) - and the same engine inside Power BI. (c) Power BI (report + Service). (d) Power Query. (e) Power Pivot / Data Model (relationships) - Power BI has the identical Model view.**
</details>

### 🙋 Student question Q2 – One big model or many small ones?
JE asks: should Callen-Lenz build ONE big model with everything (labour, purchases, stock, budget, cards, assets) or many small ones? What would you do?

<details open><summary><b>Solution Q2</b></summary>

- Steps: this is a discussion. Mark the answer on five points: star schema mentioned; single source of truth; security and audience; refresh cost; avoids duplicate logic.
- Rule of thumb to give: one model per subject area or audience, star schema, shared dimensions, measures defined once. Split when refresh time, security or ownership differ. Merge when two reports keep needing the same dimensions and measures.
- **Answer: Start with one focused model per business question, built as a star schema: a 'Project Cost' model (labour, purchases, budget, stock valuation) sharing conformed dimensions (Project, Employee, Part, Date). Do not dump everything into one giant model: it is slow to refresh, hard to test and to secure, and every report author sees all the data (cards and payroll-type data need tighter access). Do not build many isolated models either: duplicated logic gives two answers for the same number. Practical route: one certified shared semantic model (dataset) per subject area, published once in the Service; people build thin reports on top of it ('live connection') and reuse the same measures. Add a second model (e.g. Treasury: cards, assets) only when the subject and audience differ.**
</details>

### Common errors and fixes

| What they see | Why | Fix |
|---|---|---|
| "Where is the M? I only see the steps." | Advanced Editor was never opened | Home ▸ Advanced Editor. Each Applied Step is one line of M. |
| Steps named "Changed Type1", "Filtered Rows2" | Default names were left in place | Right-click the step ▸ Rename. Name by what it does. |
| Someone edits the CSV in Excel and the query breaks next month | Raw file was changed by hand | Restore the raw file. Put the fix in a query step. |
| Two students get different answers to the same question | One built a second copy of the logic | One query per job; reference or merge the first query instead of copying it. |

---

## Lab 1 — Wasp stock clean-up ⏱ 9:20–10:15

> **Who asked:** No single client request maps to this lab on its own. It is the first half of R10 (KR), the stock valuation in Lab 3.
> **Why it matters to them:** Stock is the number nobody counts in one place. The Wasp export has title rows, part numbers typed six different ways, text prices, a repeated header and duplicate lines. Once this list is clean, the stock value and the reorder list are one refresh away.

**Files:** `data/01_Wasp_Stock/Wasp_Stock_Export.xlsx` (sheet `Stock`, 75 rows including two title rows and the header on row 3).

### Teach — step by step

**Step 1 – Create the FolderPath parameter**
- Click path: Data ▸ Get Data ▸ Launch Power Query Editor, then in the Power Query Editor, Home ▸ Manage Parameters ▸ New Parameter.
- Name: `FolderPath`. Type: Text. Current Value: `C:\CallenLenz_Course\data\` (keep the trailing backslash, the M adds the file names to it).
- 🗣 Say: "Every query starts from this one box. When the course moves to another PC, you change one value, not forty file paths."
- 🧹 Mess this fixes: none. This is setup.
- ✅ Checkpoint: `FolderPath` shows in the Queries pane. Later steps depend on it.

**Step 2 – Connect and remove the two title rows, then promote the header**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ choose the Wasp file ▸ tick `Stock` in the Navigator ▸ Transform Data.
- Then in Advanced Editor, replace the file path with `FolderPath & "01_Wasp_Stock\Wasp_Stock_Export.xlsx"`.
- Home ▸ Remove Rows ▸ Remove Top Rows ▸ 2. Then Home ▸ Use First Row as Headers.
- 🗣 Say: "The export puts two title lines above the real table. We skip them. The real header is now row 3 of the sheet, and PQ takes it as the column names."
- 🧹 Mess this fixes: sheet rows 1 and 2 are export titles (`A1:X1`, `A2:X2`). The real header is row 3.
- ✅ Checkpoint: **72 data rows** after promoting headers. This includes the blank, repeated-header, test and Total rows, which the next steps remove.

**Step 3 – Remove blank rows and the repeated header**
- Click path: Home ▸ Remove Rows ▸ Remove Blank Rows. Then click the Item Number filter arrow ▸ untick `Item Number` (this removes the repeated header line).
- 🗣 Say: "Blank lines are spacers from the export. The repeated header is a page break in the export. Neither one is a part."
- 🧹 Mess this fixes: blank rows at sheet rows 14, 49, 57 and 74. Repeated header row at sheet row 31 (`A31:X31`).
- ✅ Checkpoint: **68 rows** after Remove Blank Rows. **67 rows** after removing the repeated header.

**Before Step 4 – add the two helper functions**
- Click path: Home ▸ Advanced Editor. Copy `CleanText` and `NormPart` from the full M under Q3 and paste them directly under `let`. Click Done. Then use Add Column ▸ Custom Column with `NormPart([Item Number])` and `CleanText([Description])`, remove the old columns and rename the new ones.

**Step 4 – Clean the Description text**
- Click path: select Description ▸ Transform ▸ Format ▸ Trim.
- The export also has line breaks inside cells (Alt+Enter) and non-breaking spaces (CHAR 160). Trim does not remove those. Use the `CleanText` function (in the full M below) which replaces both, collapses double spaces and trims. In the GUI this means Replace Values for each character, so the M route is faster in class.
- 🗣 Say: "A non-breaking space looks like a space but it is a different character, so a match on the text fails. CleanText turns it into a normal space first."
- 🧹 Mess this fixes: line breaks inside Description at C8, C15, C19 (12 in the file). Non-breaking spaces at C18, C22, C30 (7 in the file). Leading or trailing spaces at C9, C17, C25 (6 in the file).
- ✅ Checkpoint: no count in the check totals for this step. Check that C8 shows on one line after the step.

**Step 5 – Standardise part numbers and drop the non-parts**
- The Item Number column is typed 41 different ways in this file: `prt-12`, `PRT12`, `PRT-012`, `PRT 0012`, a trailing space, a non-breaking space. The first ones are at A7, A8, A11, A12 and A13.
- Use the `NormPart` function. It removes spaces and hyphens, makes the text upper case, strips the `PRT` prefix, pads the digits to 4 with leading zeros, and adds `PRT-` back.
- 🗣 Say: "`Text.PadStart` adds leading zeros up to a set length. So `12` becomes `0012`, and every spelling of part 12 becomes `PRT-0012`."
- Anything that is not a real part becomes blank. Those are the test items at A20 and A53, and the `Total` footer at A75. Then click the Item Number filter arrow ▸ untick (null).
- 🧹 Mess this fixes: part number variants (41 in the file). Test/sample items at A20 and A53. Total footer at A75.
- ✅ Checkpoint: **64 rows** after the valid-parts filter. 4 duplicates are still in the table; the next step removes them.

**Step 6 – Fill Down Category**
- Click path: Transform ▸ Fill ▸ Down, on the Category column.
- 🗣 Say: "The export types the category once per group, like a merged cell. Fill Down copies it into the rows underneath."
- Order note: the M file runs this before Step 5. Either order gives the same 64 rows, because fill-down does not change the row count.
- 🧹 Mess this fixes: Category only typed on the first row of each group. 58 blank Category cells in the file (Q5 to Q8 are the first four logged).
- ✅ Checkpoint: no count in the check totals. Check that no Category cell is blank in the rows you kept.

**Step 7 – Fix Item Cost (text pounds, N/A and #REF!)**
- Three kinds of mess in one column:
  - Text with a pound sign: J11, J22, J26 (10 in the file). Fix: replace `£` and `,` with nothing, then change the type to Decimal Number.
  - The text `N/A`: J10, J33, J40, J46. Fix: `Number.From` inside a `try`, otherwise null.
  - Excel error `#REF!`: J17, J52. Fix: right-click the Item Cost column ▸ Replace Errors ▸ leave blank (null).
- Replace Errors must run before any row is filtered. In the M it runs straight after the header step. Do it early, as the M does.
- 🗣 Say: "A cell with an error value breaks any sum over the column. We turn it into a blank, and a blank adds nothing."
- 🧹 Mess this fixes: Item Cost text `£` (J11, J22, J26). `N/A` (J10, J33, J40, J46). `#REF!` (J17, J52).
- ✅ Checkpoint: four parts have no usable cost: **PRT-0003, PRT-0013, PRT-0023, PRT-0047**. They count as £0 in the Q5 total. Lab 3 shows each of them has a Xero or Purchase Log price, so the valuation is not affected.

**Step 8 – Remove duplicates, keep the first**
- Click path: Home ▸ Remove Rows ▸ Remove Duplicates, on the Item Number column.
- Before this, add `Table.Buffer`. The M does this on purpose. Buffering locks the row order in memory, so "first" means the first row in the file. Without it, Power Query may keep a different duplicate.
- Click path: after Remove Duplicates writes its step, Home ▸ Advanced Editor. Insert `Buffered = Table.Buffer(<previous step name>),` above it and make the Table.Distinct line read from `Buffered`. There is no button for this.
- 🗣 Say: "The same part appears again with a different spelling and an old stock figure. We keep the first line, the one that comes first in the file."
- 🧹 Mess this fixes: duplicate lines at A33, A46, A59 and A67. Same part, different case or spacing, and a stale Total Available.
- ✅ Checkpoint: **60 rows**. That is the number of distinct valid parts (Q3).

**Step 9 – Extract the drawing number**
- Click path: Add Column ▸ Extract ▸ Text Between Delimiters. Start `[Drawing `, end `]`, on the Description column. Name the new column `Drawing No`.
- 🗣 Say: "Extract pulls out the piece of text between two markers. Here it is the drawing number in square brackets."
- 🧹 Mess this fixes: none. The drawing number is typed inside Description, so this is an extract, not a fix.
- ✅ Checkpoint: no count in the check totals. Check that a part with a drawing number shows the number on its own.

**Step 10 – Stock Status with a Conditional Column**
- Click path: Add Column ▸ Conditional Column. New column `Stock Status`. Rules, in order: if Total Available equals 0, then `Out of stock`. Else if Total Available is less than or equal to 5, then `Reorder`. Otherwise `OK`.
- 🗣 Say: "Conditional Column is IF / ELSE IF in plain English. The rules are checked top to bottom, so the first match wins."
- ✅ Checkpoint: **Out of stock 8, Reorder 8, OK 44** (60 parts in total). See Q4.

**Step 11 – Stock Value at Wasp cost**
- Click path: Add Column ▸ Custom Column. Name `Stock Value`. Formula: `[Total Available] * (if [Item Cost] = null then 0 else [Item Cost])`.
- 🗣 Say: "A part with no usable cost adds nothing. It does not break the sum."
- ✅ Checkpoint: **1,135 units** of Total Available across the 60 parts (sum before the stock value). The total stock value is Q5.

**Step 12 – Close & Load**
- Click path: Home ▸ Close & Load. Rename the query `Wasp_Clean`. Keep it loaded to a table on a sheet (you need it again in Lab 3).
- Save the workbook.
- ✅ Checkpoint: `Wasp_Clean` has **60 rows**.

### 🙋 Student question Q3 – How many distinct valid parts?
Clean the Wasp stock export (remove title rows, blank rows, repeated header, test/total rows; normalise part numbers; remove duplicates keeping the first). How many distinct valid parts are left?

<details open><summary><b>Solution Q3</b></summary>

- Steps: Power Query ▸ Home ▸ Use First Row as Headers after Remove Top Rows = 2. Remove Blank Rows. Filter out the repeated header. Normalise Item Number with `NormPart`, then filter out blanks (removes A20, A53, A75). Table.Buffer, then Remove Duplicates on Item Number, keep first. Count the rows.
- Checkpoints on the way: 72 after promoting, 68 after blanks, 67 after repeated header, 64 after the valid-parts filter, **60 after de-duplication**.
- M (full, from `solutions/m/lab1_wasp.m`):

<details><summary>Full M code</summary>

```m
// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 1 - Wasp stock clean-up.  Query name: Wasp_Clean
// Rule order matters: errors -> blank rows -> repeated header -> fill down -> normalise -> filter -> dedupe.
let
    // ---- reusable helper functions (each one is a custom function) ----
    // CleanText: line breaks and CHAR(160) -> space, collapse repeated spaces, trim
    CleanText = (t as nullable text) as nullable text =>
        if t = null then null
        else Text.Trim(Text.Combine(List.Select(Text.Split(Text.Replace(Text.Replace(Text.Replace(t, "#(00A0)", " "), "#(lf)", " "), "#(cr)", " "), " "), each _ <> ""), " ")),
    // NormPart: any spelling of a part number -> PRT-0000 ; anything that is not a PRT part -> null
    NormPart = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                Squashed = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"})),
                Digits = Text.Middle(Squashed, 3)
            in
                if Text.StartsWith(Squashed, "PRT") and Digits <> "" then "PRT-" & Text.PadStart(Digits, 4, "0") else null,

    // ---- 1. load and remove the 2 export title rows, promote the real header (row 3) ----
    Source = Excel.Workbook(File.Contents(FolderPath & "01_Wasp_Stock\Wasp_Stock_Export.xlsx"), null, true),
    StockSheet = Source{[Item = "Stock", Kind = "Sheet"]}[Data],
    RemovedTopRows = Table.Skip(StockSheet, 2),
    PromotedHeaders = Table.PromoteHeaders(RemovedTopRows, [PromoteAllScalars = true]),

    // ---- 2. Replace Errors: the #REF! cells in Item Cost become null (must happen BEFORE row filters) ----
    ReplacedErrors = Table.ReplaceErrorValues(PromotedHeaders, {{"Item Cost", null}}),

    // ---- 3. structure clean-up ----
    RemovedBlankRows = Table.SelectRows(ReplacedErrors, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    RemovedRepeatedHeaders = Table.SelectRows(RemovedBlankRows, each [Item Number] <> "Item Number"),
    FilledDownCategory = Table.FillDown(RemovedRepeatedHeaders, {"Category"}),

    // ---- 4. normalise part numbers; non-PRT rows (TEST-ITEM, ZZ SAMPLE, Total) become null and are dropped ----
    NormalisedItemNumber = Table.TransformColumns(FilledDownCategory, {{"Item Number", NormPart, type text}}),
    ValidPartsOnly = Table.SelectRows(NormalisedItemNumber, each [Item Number] <> null),

    // ---- 5. remove duplicates, keeping the FIRST row ----
    // NOTE: the UI's Remove Duplicates only reliably keeps the first row if the table is buffered first;
    // without Table.Buffer Power Query may evaluate the rows in another order and keep a different duplicate.
    BufferedRows = Table.Buffer(ValidPartsOnly),
    RemovedDuplicates = Table.Distinct(BufferedRows, {"Item Number"}),

    // ---- 6. tidy text, numbers and types ----
    CleanedDescription = Table.TransformColumns(RemovedDuplicates, {{"Description", CleanText, type text}}),
    // Locale-safe cost parse: numbers pass through; text has £ and , removed then is read with en-GB; N/A etc -> null
    CleanedItemCost = Table.AddColumn(CleanedDescription, "Item Cost Clean",
        each if Value.Is([Item Cost], type number) then [Item Cost]
             else try Number.From(Text.Replace(Text.Replace(Text.Trim([Item Cost]), "£", ""), ",", ""), "en-GB") otherwise null, type nullable number),
    ChangedTypes = Table.TransformColumnTypes(CleanedItemCost, {{"Total Available", Int64.Type}, {"Total In House", Int64.Type}, {"On Order", Int64.Type}, {"Committed", Int64.Type}}),

    // ---- 7. Extract: drawing number between "[Drawing " and "]" ----
    AddedDrawingNo = Table.AddColumn(ChangedTypes, "Drawing No", each Text.BetweenDelimiters([Description], "[Drawing ", "]"), type text),

    // ---- 8. Conditional column: Stock Status ----
    AddedStockStatus = Table.AddColumn(AddedDrawingNo, "Stock Status",
        each if [Total Available] = 0 then "Out of stock" else if [Total Available] <= 5 then "Reorder" else "OK", type text),

    // ---- 9. Stock Value at Wasp cost (unusable cost = blank = adds 0), then keep the useful columns ----
    AddedStockValue = Table.AddColumn(AddedStockStatus, "Stock Value", each [Total Available] * (if [Item Cost Clean] = null then 0 else [Item Cost Clean]), type number),
    SelectedColumns = Table.SelectColumns(AddedStockValue, {"Item Number", "Description", "Drawing No", "Category", "Total Available", "Total In House", "On Order", "Committed", "Item Cost Clean", "Stocking Unit", "Location", "Site", "Stock Status", "Stock Value"}),
    Renamed = Table.RenameColumns(SelectedColumns, {{"Item Number", "Part Number"}, {"Item Cost Clean", "Item Cost"}})
in
    Renamed
```

</details>

- **Answer: 60 parts**
</details>

### 🙋 Student question Q4 – How many parts are Out of stock?
After adding the Stock Status conditional column (Total Available = 0 -> Out of stock; <= 5 -> Reorder; else OK), how many parts are 'Out of stock'?

<details open><summary><b>Solution Q4</b></summary>

- Steps: Add Column ▸ Conditional Column on Total Available. Rules as in Step 10. Then filter Stock Status to `Out of stock` (or Group By Stock Status and count).
- Keep the first duplicate, as Step 8 does.
- **Answer: 8 parts (Out of stock). Reorder 8, OK 44.**
- ⚠️ If a student gets **7** → they kept the LAST duplicate instead of the first (the Table.Buffer / keep-first step was skipped or the sort order was reversed).
</details>

### 🙋 Student question Q5 – Total stock value at Wasp cost
What is the total stock value at Wasp Item Cost: the sum of Total Available × Item Cost over the 60 de-duplicated parts (keep the first row of each duplicate)? Where Item Cost is N/A or an error (#REF!), treat the cost as blank, so that part adds £0. Round to 2 dp.

<details open><summary><b>Solution Q5</b></summary>

- Steps: Use the Custom Column in Step 11 (Total Available × Item Cost, blank cost counts as 0). Then Home ▸ Group By with no grouping, or select the Stock Value column and read the sum in the status bar.
- Unusable costs are PRT-0003, PRT-0013, PRT-0023 and PRT-0047. They add £0.
- **Answer: £258,431.64 (258431.64)**
- ⚠️ If a student gets **243,918.90** → they kept the LAST duplicate instead of the first. The same-named part then carries a stale Total Available.
</details>

### Common errors and fixes

| What they see | Why | Fix |
|---|---|---|
| "The column 'Item Number' of the table was not found" | Remove Top Rows was not set to 2, or the header was not promoted | Remove Top Rows = 2, then Use First Row as Headers. The real header is row 3. |
| Out of stock shows 7, not 8 | Duplicates kept the last line | Put Table.Buffer before Remove Duplicates, then keep first. |
| Stock value shows 243,918.90 | Same cause: last duplicate kept | As above. The right answer is £258,431.64. |
| Item Cost will not sum; a cell shows an error | `#REF!` not replaced | Right-click Item Cost ▸ Replace Errors, before the row filters. |
| Item Cost text like `£1,234.50` will not change to a number | `£` or the comma left in | Replace `£` and `,` with nothing, then change type to Decimal Number. |
| Duplicates are not removed, or a non-part row stays in | Part numbers not normalised before the filter and dedupe | Run NormPart first, then filter blanks, then dedupe. |
| "We couldn't find a part of the path" on the first query | FolderPath typed without the trailing backslash (Labs 1, 3 and 4 add the folder name straight on) | Edit the parameter to C:\CallenLenz_Course\data\ with the trailing backslash |

---

## Lab 2 — Timesheet consolidation & control checks ⏱ 10:30–11:30

> **Who asked:** "Consolidating data from multiple sources." (JE); "Best practice in structuring data" (JE); "Is there a way to build in control checks? Ie. Ensuring data is all valid and pulling through." (JE)
> **Why it matters to them:** Six months of hand-typed timesheets feed labour cost for every project. Before anyone costs a project, the hours have to match the timesheet system's own control totals, and every booking has to be on a valid project code.

**Files:** `data/02_Timesheets/Timesheet_2026-01.csv` to `Timesheet_2026-05.csv`; `data/02_Timesheets/Control_Totals.xlsx` (sheet `Control`); `data/00_Reference/Employees.xlsx`, `Labour_Rates.xlsx`, `Projects.xlsx` (sheet `Sheet1` in each). The June file, `data/02_Timesheets/_Later/Timesheet_2026-06.csv`, stays in `_Later` until Step 12.

Lab 2 starts after the 10:15–10:30 break.

### Teach — step by step

**Step 1 – Point at the folder, not the file**
- Click path: Data ▸ Get Data ▸ From File ▸ From Folder ▸ Browse to C:\CallenLenz_Course\data\02_Timesheets ▸ OK ▸ **Transform Data** (not Combine yet). Click the filter arrow on the Folder Path column and tick only the value that ends 02_Timesheets\ (untick the _Later one). Then click the Combine Files icon (double down arrow) at the top of the Content column ▸ OK.
- 🗣 Say: "We point at the timesheets folder, not the whole data folder, or the reference files would be combined too. From Folder also reads sub-folders, so we filter Folder Path to keep `_Later` out. The finished M does the same filter."
- 🧹 Mess this fixes: none yet.
- ✅ Checkpoint: **5 monthly files** in the folder (Jan–May). June is in `_Later`.

**Step 2 – The March file breaks the combine**
- Mess: `Timesheet_2026-03.csv` names its column `Proj Code`, not `Project Code`. Once combined, the March rows have no `Project Code` column and the combine shows an error on that column.
- Fix: open the Transform Sample File query (the function Combine builds from the Jan sample) and add a rename step: `Proj Code` to `Project Code`, with `MissingField.Ignore`, so the other files are not broken by a missing column. The line in the M is:

```m
#"Renamed Proj Code" = Table.RenameColumns(#"Trimmed Headers", {{"Proj Code", "Project Code"}}, MissingField.Ignore),
```

- 🗣 Say: "One file's header differs by two letters. Rename it inside the function, so every file lines up before they are stacked."
- 🧹 Mess this fixes: `Timesheet_2026-03.csv` header `Proj Code` (line 1).
- ✅ Checkpoint: the combine runs with no error on the Project Code column.

**Step 3 – Remove blank lines and the Total footer**
- Click path: in the sample function, Home ▸ Remove Rows ▸ Remove Blank Rows. Then filter Week Ending to exclude `Total`.
- 🗣 Say: "The files have blank lines, lines with only commas, and a Total footer at the bottom. None of these is a booking. The Total would double-count the hours."
- 🧹 Mess this fixes: blank lines (e.g. Jan lines 250, 294, 455). Comma-only lines (e.g. Jan line 302, Feb line 172). Total footer rows (Jan line 456 shows 3743.0; Feb line 360 shows 3017.5; each month has one).
- ✅ Checkpoint: about **1,983** data lines after combining (check-totals figure, marked approximate, confirm on the day). **1,973** rows after blank and Total rows are removed.

**Step 4 – April has US dates**
- Mess: `Timesheet_2026-04.csv` stores Week Ending as month/day/year. `04/03/2026` in April is **3 April**, not 4 March. The other files are day/month/year.
- Fix: Transform ▸ Data Type ▸ Using Locale ▸ Date, English (United States), for April only. The M picks the culture from the file name. That is a teaching shortcut. In real life you would fix the export.
- GUI limit: the Transform Sample File query does not know which file it is reading, so it cannot switch to US dates for April only. Teach Steps 1–3 in the GUI to show the idea. Then Data ▸ Get Data ▸ From Other Sources ▸ Blank Query ▸ Home ▸ Advanced Editor ▸ paste solutions/m/lab2_timesheets.m. In that M, fnLoadTimesheet(fileContent, fileName) uses en-US when the file name contains 2026-04.
- 🗣 Say: "The same text means two different days depending on the country. Tell Power Query which country's date order each file uses."
- 🧹 Mess this fixes: the April file (e.g. `04/10/2026` is 10 April, not 4 October).
- ✅ Checkpoint: April dates fall in April, not in March or October.

**Step 5 – Fix the 2062 typo**
- Mess: `Timesheet_2026-05.csv` line 148 has Week Ending `08/05/2062`. It should be 2026. Q9 below mentions the same typo: it is in the May file, not June.
- Fix: a Custom Column or Transform step that replaces the year when it is 2062:

```m
#"Fixed Year Typo" = Table.TransformColumns(Combined, {{"Week Ending", each if Date.Year(_) = 2062 then #date(2026, Date.Month(_), Date.Day(_)) else _, type date}}),
```

- 🗣 Say: "A year typed as 2062 is an easy typo. Check the year is plausible, then correct it, and say so out loud so it is not a silent fix."
- 🧹 Mess this fixes: `Timesheet_2026-05.csv` line 148, Week Ending `08/05/2062`.
- ✅ Checkpoint: no Week Ending in 2062 remains. The May weeks group under May.

**Step 6 – Clean the codes, IDs, names, hours and tasks**
- Project codes come in at least six shapes. Examples: `cl-p101` (lower case), `Cl-P101` (mixed), `CL P101` (space), `CLP101` (no hyphen), `CL-P104 ` (trailing space), ` CL-P102` (leading space). Fix: Text.Upper, strip spaces and hyphens, rebuild as `CL-P` + the last three digits (the `fnProjectCode` function in the M).
- Employee IDs: `e009` (lower case), `E012 ` (trailing space), and bare numbers in May (`7` for `E007`). Fix: trim, upper case, then pad to three digits after an E (`fnEmployeeId`).
- Employee names with double spaces (`Oliver  Hartley`) or a trailing space (`Callum Reid `). Fix: collapse spaces (`fnCollapse`). In Step 9 we take the name from Employees anyway.
- Task in lower or upper case (`design`, `DESIGN`). Fix: Transform ▸ Format ▸ Capitalize Each Word.
- Hours typed as text with a suffix, e.g. `12.5 hrs` on the January file (line 2). Fix: keep only digits and the point (`Text.Select`), then change the type to Decimal Number.
- 🗣 Say: "Text.Select keeps only the characters we list. Here, digits and the decimal point. So `12.5 hrs` keeps `12.5`."
- 🧹 Mess this fixes: the project code, employee ID, name, task and hours variants above, which appear in every monthly file (`_build/mess_log_A.json` lists the first examples per file). Employee ID variants: Feb `e009`, May numeric `6`, `E012 ` in Jan.
- ✅ Checkpoint: row count is unchanged at **1,973**. Every Project Code now reads `CL-Pnnn`.

**Step 7 – Clean first, then remove duplicates**
- Click path: Home ▸ Remove Rows ▸ Remove Duplicates, after the clean steps.
- 🗣 Say: "The order matters. Clean first, so a near-duplicate becomes an exact duplicate, and only then remove duplicates."
- 🧹 Mess this fixes: exact duplicate lines (Jan line 178, Feb line 249, Mar line 126, Apr line 53, May line 144). Near-duplicates: Jan line 289 (`cl-p103` with a trailing space in the name) becomes an exact copy once cleaned.
- ✅ Checkpoint: **6 rows removed**, leaving **1,967 rows** for Jan–May.

**Step 7b – Load the three reference tables**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ data\00_Reference\Projects.xlsx ▸ Sheet1 ▸ Transform Data ▸ Home ▸ Close & Load ▸ Close & Load To… ▸ Only Create Connection. Repeat for Employees.xlsx and Labour_Rates.xlsx (sheet `Sheet1` in each). Name the queries Projects, Employees and Labour_Rates.

**Step 8 – Find the unknown project codes (Left Anti join)**
- Click path: Home ▸ Merge Queries ▸ Merge Queries as New. Left table: the timesheets. Right table: `Projects`, matched on Project Code. Join Kind: **Left Anti**.
- 🗣 Say: "A Left Anti join keeps the rows on the left that have no match on the right. These are bookings to codes that are not on the project list."
- 🧹 Mess this fixes: bookings to codes not in Projects.xlsx: `CL-P199` (4 rows) and `CL-P100` (2 rows). All are in the February file (lines 76, 80, 106, 111, 173, 316).
- ✅ Checkpoint: **6 rows, 7.5 hours**. This is Q8. Then merge again with Join Kind **Inner** to keep only valid rows: **1,961 valid rows** for Jan–May.

**Step 9 – Merge Employees and Labour_Rates**
- Click path: Merge Queries, Employees on Employee ID (Left Outer), then expand `Employee Name` and `Grade`. Then merge Labour_Rates on Grade (Left Outer), then expand `Standard Hourly Rate GBP` and `Overtime Multiplier`.
- 🗣 Say: "Merge Queries is a lookup. Employees tells us the grade, and the grade tells us the rate. We do not type rates into the timesheet."
- 🧹 Mess this fixes: typed names in the timesheets (double spaces, trailing spaces) are replaced by the name in Employees.
- ✅ Checkpoint: every valid row has a Standard Rate and an OT Multiplier.

**Step 10 – Overtime, with a worked example**
- The rule: for each employee and Week Ending, hours above **37.5** are overtime. Count them after unknown codes and duplicates are removed. Each line gets a share of that overtime in proportion to its hours. Labour cost per line = Hours × Standard Rate + that line's OT Hours × Standard Rate × (Multiplier − 1).
- Worked example to put on the board. An employee at a rate of £20 an hour, with an overtime multiplier of 1.5, books 40 hours in one week: 30 hours on CL-P101 and 10 hours on CL-P102.
  - Weekly hours = 40. Overtime = 40 − 37.5 = **2.5 hours**.
  - CL-P101 line: OT share = 2.5 × 30 / 40 = **1.875**. Cost = 30 × 20 + 1.875 × 20 × 0.5 = 600 + 18.75 = **£618.75**.
  - CL-P102 line: OT share = 2.5 × 10 / 40 = **0.625**. Cost = 10 × 20 + 0.625 × 20 × 0.5 = 200 + 6.25 = **£206.25**.
  - The two lines add to **£825.00**. Check: 40 × 20 + 2.5 × 20 × 0.5 = 800 + 25 = £825.00.
- 🗣 Say: "Every hour is paid at the standard rate. The overtime hours get the extra half, and we split that extra over the lines the hours came from."
- Click path: Transform ▸ Group By, Employee ID and Week Ending, new column Weekly Hours = Sum of Hours. Merge the result back on Employee ID and Week Ending. Add Custom Column `OT Hours` = if Weekly Hours > 37.5 then (Weekly Hours − 37.5) × Hours / Weekly Hours else 0. Add Custom Column `Labour Cost` = Hours × Standard Rate + OT Hours × Standard Rate × (OT Multiplier − 1).
- ✅ Checkpoint: **CL-P103, Jan–Mar: 801.0 hours and 13.37 overtime hours**, giving labour cost **£28,779.27** (Q6). Across Jan–May, labour cost is **£568,743.75**. Across Jan–Jun it is **£671,900.00**, after the June refresh in Step 12. Overtime hours Jan–Jun total **475.5**.
- Rename this final query **LabourLines** (right-click ▸ Rename). Lab 5, the control check and Power BI all refer to it by that name.

**Step 11 – Group by project and month, then the control check**
- Click path: Transform ▸ Group By, with Project Code and Month as the groups. Sum Hours, OT Hours and Labour Cost. The result has one row per project and month.
- ✅ Checkpoint: **45 project-month rows** in the Group By output (check totals).
- Control check (answers R3): in Queries & Connections right-click LabourLines ▸ Duplicate, and rename the copy Timesheets_Deduped. In Applied Steps right-click the step after "Removed Duplicates" ▸ Delete Until End, so the unknown-code rows stay in. Then Data ▸ Get Data ▸ From Other Sources ▸ Blank Query ▸ Advanced Editor ▸ paste solutions/m/lab2_control_checks.m and name it Control_Check. It groups by month and merges Control_Totals (Remove Top Rows = 3, then headers) for you.
- 🗣 Say: "A control check is a second number that has to match. If it does not, the difference tells us where to look."
- The result compares **loaded hours** with **system hours**, month by month:
  - Jan: 3,727.0 loaded vs 3,727.0 system. Difference 0.0.
  - **Feb: 3,006.5 loaded vs 2,999.0 system. Difference +7.5.**
  - Mar: 3,037.0 vs 3,037.0. Apr: 3,024.5 vs 3,024.5. May: 3,682.5 vs 3,682.5. All 0.0.
- The extra column "Hours on Unknown Codes" shows 7.5 for February, so "Unexplained Hours" is 0. The unknown-code rows explain the whole difference.
- 🧹 Mess this fixes: the unknown codes (Step 8).

**Step 12 – Drop in June and Refresh All**
- File Explorer: copy `data\02_Timesheets\_Later\Timesheet_2026-06.csv` into `data\02_Timesheets`.
- Data ▸ Refresh All.
- 🗣 Say: "We did not change the query. We added a file to the folder and refreshed. That is the whole point of From Folder."
- 🧹 Mess this fixes: June has the same kinds of mess as the other files (codes, spaces, blank lines, the Total footer), plus a near-duplicate at line 56 (`cl-p106`, trailing space in the name). The June footer shows 2,993.5; the gap to the cleaned June total of 2,987.5 is exactly 6.0, the hours on that near-duplicate line.
- ✅ Checkpoint: **2,326 rows** after removing duplicates with June included (7 removed in total). **2,320 valid rows** for Jan–Jun. June alone is **2,987.5 hours**. Year-to-date **19,457.5 hours** (Q9).

### 🙋 Student question Q6 – CL-P103 labour cost, Jan–Mar
Combine the January-May timesheets and clean them (headers, dates, codes, hours, Total/blank rows, duplicates, unknown codes), then merge Employees and Labour_Rates. One Week Ending has a typo year 2062; correct it to 2026. Overtime rule: for each employee and Week Ending, hours above 37.5 are overtime, counted AFTER removing duplicates and unknown-code rows. Each line gets its share of that overtime in proportion to its hours. Labour cost per line = Hours x Standard Rate + that line's OT hours x Standard Rate x (OT Multiplier - 1). What is the total labour cost for CL-P103 for Jan-Mar (month of Week Ending)?

<details open><summary><b>Solution Q6</b></summary>

- Steps: Folder combine (Lab 2 Steps 1–3). Fix the March header (Step 2). Clean dates, codes, IDs, names and hours (Steps 4–6). Remove duplicates after cleaning (Step 7). Left Anti then Inner join on Projects (Step 8). Merge Employees and Labour_Rates (Step 9). Group By Employee + Week Ending for Weekly Hours, merge it back, then add OT Hours and Labour Cost (Step 10). Filter Project Code = CL-P103 and Month = Jan–Mar. Sum Labour Cost.
- Checkpoints: CL-P103 Jan–Mar = **801.0 hours**, **13.37 OT hours**.
- M (full, from `solutions/m/lab2_timesheets.m`):

<details><summary>Full M code</summary>

```m
// =====================================================================================
// LAB 2 - TIMESHEETS: combine, clean, merge rates, overtime, group by project-month
// -------------------------------------------------------------------------------------
// BEFORE YOU PASTE: create a parameter  (Home > Manage Parameters > New Parameter)
//     FolderPath   Type = Text   Current value = the folder that holds 00_Reference, 02_Timesheets ...
//                  e.g.  C:\CallenLenz_Course\data      (no trailing backslash needed)
// Then: Home > New Source > Blank Query > Advanced Editor > paste this whole file.
//
// OVERTIME RULE (same in the check totals):
//   Weekly Hours   = sum of an employee's hours for one Week Ending (after removing unknown codes)
//   Overtime Hours = max(0, Weekly Hours - 37.5)
//   Each line gets its share of the overtime pro-rata by hours:
//        OT Hours (line) = Overtime Hours x Hours / Weekly Hours
//   Labour Cost (line)  = Hours x Standard Rate  +  OT Hours (line) x Standard Rate x (Multiplier - 1)
//   i.e. every hour is paid at standard rate, and overtime hours get the extra 0.5 premium.
//
// STAGES (useful checkpoints):  Combined -> Removed Duplicates -> Valid rows -> Final Labour Lines
// NAME THIS QUERY  LabourLines  (one row per cleaned line). The project-month Group By is a separate query:
// see lab2_project_month_summary.m, and lab5_pivot_hours.m / lab2_control_checks.m which also use it.
// =====================================================================================
let
    // ---------- 1. Where are the files? (the _Later sub-folder is deliberately ignored) ----------
    TimesheetFolder = Text.TrimEnd(FolderPath, "\") & "\02_Timesheets\",
    ReferenceFolder = Text.TrimEnd(FolderPath, "\") & "\00_Reference\",
    Source = Folder.Files(TimesheetFolder),
    #"Filtered Timesheet Files" = Table.SelectRows(
        Source,
        each [Folder Path] = TimesheetFolder
            and [Extension] = ".csv"
            and Text.StartsWith([Name], "Timesheet_")
    ),

    // ---------- 2. Small helper functions (each one is a single, teachable fix) ----------
    // Trim spaces, turn non-breaking spaces (CHAR 160) into normal spaces, null -> ""
    fnTrim = (x as any) as text =>
        if x = null then "" else Text.Trim(Text.Replace(Text.From(x), Character.FromNumber(160), " ")),
    // "Oliver   Hartley" -> "Oliver Hartley"
    fnCollapse = (x as text) as text =>
        Text.Combine(List.Select(Text.Split(x, " "), each _ <> ""), " "),
    // "cl p101" / "CLP101" / "CL-P101 " -> "CL-P101"
    fnProjectCode = (x as text) as text =>
        let
            Letters = Text.Select(Text.Upper(x), {"A".."Z", "0".."9"})
        in
            if Text.Length(Letters) >= 4 then "CL-P" & Text.End(Letters, 3) else Letters,
    // 7 -> "E007" ; "E007" stays "E007"
    fnEmployeeId = (x as text) as text =>
        let
            Upper = Text.Upper(x)
        in
            if Upper = "" then "" else if Text.StartsWith(Upper, "E") then Upper else "E" & Text.PadStart(Upper, 3, "0"),

    // ---------- 3. The 'Transform Sample File' function: loads and cleans ONE monthly file ----------
    fnLoadTimesheet = (fileContent as binary, fileName as text) as table =>
        let
            Raw = Csv.Document(fileContent, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            #"Promoted Headers" = Table.PromoteHeaders(Raw, [PromoteAllScalars = true]),
            #"Trimmed Headers" = Table.TransformColumnNames(#"Promoted Headers", Text.Trim),
            // March calls the column "Proj Code" - rename it here so every file lines up
            #"Renamed Proj Code" = Table.RenameColumns(#"Trimmed Headers", {{"Proj Code", "Project Code"}}, MissingField.Ignore),
            #"Kept Columns" = Table.SelectColumns(
                #"Renamed Proj Code",
                {"Week Ending", "Employee ID", "Employee Name", "Project Code", "Task", "Hours"}
            ),
            // Trim every column as text first (also turns nulls into "")
            #"Trimmed Text" = Table.TransformColumns(
                #"Kept Columns",
                List.Transform(Table.ColumnNames(#"Kept Columns"), each {_, fnTrim, type text})
            ),
            // Blank lines have an empty Week Ending; the footer line says "Total"
            #"Removed Blank and Total Rows" = Table.SelectRows(
                #"Trimmed Text",
                each [Week Ending] <> "" and Text.Lower([Week Ending]) <> "total"
            ),
            // April is US month/day/year text; every other file is UK day/month/year.
            // NOTE: choosing the culture from the FILE NAME is a teaching shortcut. In real life you would
            // fix the source export or detect the format from the data (e.g. a day > 12).
            Culture = if Text.Contains(fileName, "2026-04") then "en-US" else "en-GB",
            #"Changed Week Ending Type" = Table.TransformColumnTypes(
                #"Removed Blank and Total Rows", {{"Week Ending", type date}}, Culture
            ),
            #"Cleaned Fields" = Table.TransformColumns(
                #"Changed Week Ending Type",
                {
                    {"Employee ID", fnEmployeeId, type text},
                    {"Employee Name", fnCollapse, type text},
                    {"Project Code", fnProjectCode, type text},
                    {"Task", each Text.Proper(_), type text},
                    // "7.5 hrs" -> 7.5
                    {"Hours", each Number.FromText(Text.Select(_, {"0".."9", "."}), "en-GB"), type number}
                }
            )
        in
            #"Cleaned Fields",

    // ---------- 4. Combine all monthly files ----------
    #"Added Cleaned Data" = Table.AddColumn(
        #"Filtered Timesheet Files", "Data", each fnLoadTimesheet([Content], [Name])
    ),
    Combined = Table.Combine(#"Added Cleaned Data"[Data]),

    // ---------- 5. Typo year: 08/05/2062 should be 2026 ----------
    #"Fixed Year Typo" = Table.TransformColumns(
        Combined,
        {{"Week Ending", each if Date.Year(_) = 2062 then #date(2026, Date.Month(_), Date.Day(_)) else _, type date}}
    ),

    // ---------- 6. Remove exact duplicates (AFTER cleaning, so near-duplicates now match) ----------
    #"Removed Duplicates" = Table.Distinct(#"Fixed Year Typo"),

    // ---------- 7. Reference data ----------
    Projects = Excel.Workbook(File.Contents(ReferenceFolder & "Projects.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    Employees = Excel.Workbook(File.Contents(ReferenceFolder & "Employees.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    Rates = Excel.Workbook(File.Contents(ReferenceFolder & "Labour_Rates.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],

    // ---------- 8. Unknown project codes (Left Anti) - inspect this step: it is Q8 ----------
    #"Unknown Code Rows" = Table.NestedJoin(
        #"Removed Duplicates", {"Project Code"}, Projects, {"Project Code"}, "Match", JoinKind.LeftAnti
    ),
    // Valid rows only (Inner join) and bring in the project name
    #"Merged Projects" = Table.NestedJoin(
        #"Removed Duplicates", {"Project Code"}, Projects, {"Project Code"}, "Project", JoinKind.Inner
    ),
    #"Expanded Project" = Table.ExpandTableColumn(#"Merged Projects", "Project", {"Project Name"}, {"Project Name"}),

    // ---------- 9. Employee name + grade, then rate ----------
    #"Removed Typed Name" = Table.RemoveColumns(#"Expanded Project", {"Employee Name"}),
    #"Merged Employees" = Table.NestedJoin(
        #"Removed Typed Name", {"Employee ID"}, Employees, {"Employee ID"}, "Emp", JoinKind.LeftOuter
    ),
    #"Expanded Employees" = Table.ExpandTableColumn(
        #"Merged Employees", "Emp", {"Employee Name", "Grade"}, {"Employee Name", "Grade"}
    ),
    #"Merged Rates" = Table.NestedJoin(#"Expanded Employees", {"Grade"}, Rates, {"Grade"}, "Rate", JoinKind.LeftOuter),
    #"Expanded Rates" = Table.ExpandTableColumn(
        #"Merged Rates",
        "Rate",
        {"Standard Hourly Rate GBP", "Overtime Multiplier"},
        {"Standard Rate", "OT Multiplier"}
    ),

    // ---------- 10. Weekly hours per employee, merged back onto every line ----------
    #"Weekly Hours Table" = Table.Group(
        #"Expanded Rates",
        {"Employee ID", "Week Ending"},
        {{"Weekly Hours", each List.Sum([Hours]), type number}}
    ),
    #"Merged Weekly Hours" = Table.NestedJoin(
        #"Expanded Rates", {"Employee ID", "Week Ending"}, #"Weekly Hours Table", {"Employee ID", "Week Ending"}, "Weekly", JoinKind.LeftOuter
    ),
    #"Expanded Weekly Hours" = Table.ExpandTableColumn(#"Merged Weekly Hours", "Weekly", {"Weekly Hours"}, {"Weekly Hours"}),

    // ---------- 11. Overtime and labour cost per line ----------
    #"Added OT Hours" = Table.AddColumn(
        #"Expanded Weekly Hours",
        "OT Hours",
        each if [Weekly Hours] > 37.5 then ([Weekly Hours] - 37.5) * [Hours] / [Weekly Hours] else 0,
        type number
    ),
    #"Added Labour Cost" = Table.AddColumn(
        #"Added OT Hours",
        "Labour Cost",
        each [Hours] * [Standard Rate] + [OT Hours] * [Standard Rate] * ([OT Multiplier] - 1),
        type number
    ),
    #"Added Month" = Table.AddColumn(#"Added Labour Cost", "Month", each Date.StartOfMonth([Week Ending]), type date),
    // Final line-level table: one row per cleaned booking line (this is the query called LabourLines)
    #"Final Labour Lines" = Table.SelectColumns(
        #"Added Month",
        {"Week Ending", "Month", "Employee ID", "Employee Name", "Project Code", "Project Name", "Task", "Hours", "Standard Rate", "Weekly Hours", "OT Hours", "Labour Cost"}
    )
in
    #"Final Labour Lines"
```

</details>

- **Answer: £28,779.27 (28779.27)**
- ⚠️ If a student gets **28,797.30** → they worked out overtime BEFORE removing the unknown-code rows.
- ⚠️ If a student gets **29,039.20** → they kept the duplicate rows (no Remove Duplicates step).
- ⚠️ If a student gets **28,537.50** → they used hours × rate only and forgot the overtime premium (the OT Multiplier − 1 part).
- Note: the 2062 typo does not change the Jan–Mar answer. It only affects May weeks.
</details>

### 🙋 Student question Q7 – Which month fails the control check?
Compare loaded hours per month (after cleaning and removing duplicates, BEFORE removing unknown project codes) with Control_Totals.xlsx. Which month fails and by how many hours, and why?

<details open><summary><b>Solution Q7</b></summary>

- Steps: Duplicate LabourLines as Timesheets_Deduped and delete every step after Removed Duplicates. Group by Month. Merge with Control_Totals (Remove Top Rows = 3, then Use First Row as Headers). Difference = Loaded − System. Then Left Anti merge with Projects to see the rows that explain the difference.
- M (control check, from `solutions/m/lab2_control_checks.m`):

<details><summary>Full M code</summary>

```m
// =====================================================================================
// LAB 2 - CONTROL CHECKS: loaded hours per month vs the timesheet system's control totals
// -------------------------------------------------------------------------------------
// Needs:  parameter FolderPath (Text), and a query called Timesheets_Deduped.
// Make Timesheets_Deduped like this: right-click the Lab 2 query (LabourLines) > Duplicate, rename it
// Timesheets_Deduped, then delete every step AFTER  #"Removed Duplicates"  (so it still
// contains the unknown-code rows - that is the point of the check).
// Expected result: every month OK except Feb, which is 7.5 hours OVER the control total;
// the 'Hours on Unknown Codes' column explains the whole difference (6 rows).
// =====================================================================================
let
    Source = Timesheets_Deduped,
    ReferenceFolder = Text.TrimEnd(FolderPath, "\") & "\00_Reference\",
    ControlFile = Text.TrimEnd(FolderPath, "\") & "\02_Timesheets\Control_Totals.xlsx",
    Projects = Excel.Workbook(File.Contents(ReferenceFolder & "Projects.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    ValidCodes = Projects[Project Code],

    #"Added Month" = Table.AddColumn(Source, "Month", each Date.StartOfMonth([Week Ending]), type date),
    #"Added Valid Flag" = Table.AddColumn(#"Added Month", "Is Valid Code", each List.Contains(ValidCodes, [Project Code]), type logical),

    // Loaded hours, lines, and what part of them sits on unknown codes
    #"Loaded Totals" = Table.Group(
        #"Added Valid Flag",
        {"Month"},
        {
            {"Loaded Hours", each List.Sum([Hours]), type number},
            {"Loaded Lines", each Table.RowCount(_), Int64.Type},
            {"Hours on Unknown Codes", each List.Sum(List.Combine({{0}, Table.SelectRows(_, each not [Is Valid Code])[Hours]})), type number},
            {"Unknown Code Lines", each Table.RowCount(Table.SelectRows(_, each not [Is Valid Code])), Int64.Type}
        }
    ),

    // Control totals sheet: 3 rows above the header (title, source note, blank line)
    ControlRaw = Excel.Workbook(File.Contents(ControlFile), false, true){[Item = "Control", Kind = "Sheet"]}[Data],
    #"Skipped Title Rows" = Table.Skip(ControlRaw, 3),
    #"Promoted Headers" = Table.PromoteHeaders(#"Skipped Title Rows", [PromoteAllScalars = true]),
    #"Typed Control" = Table.TransformColumnTypes(
        #"Promoted Headers", {{"Month", type date}, {"System Hours", type number}, {"System Lines", Int64.Type}}
    ),
    #"Control Month Start" = Table.TransformColumns(#"Typed Control", {{"Month", each Date.StartOfMonth(_), type date}}),

    Merged = Table.NestedJoin(#"Loaded Totals", {"Month"}, #"Control Month Start", {"Month"}, "Control", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "Control", {"System Hours", "System Lines"}, {"System Hours", "System Lines"}),

    #"Added Difference" = Table.AddColumn(Expanded, "Hours Difference", each [Loaded Hours] - [System Hours], type number),
    #"Added Line Difference" = Table.AddColumn(#"Added Difference", "Line Difference", each [Loaded Lines] - [System Lines], Int64.Type),
    #"Added Status" = Table.AddColumn(
        #"Added Line Difference", "Status", each if Number.Abs([Hours Difference]) < 0.001 then "OK" else "CHECK", type text
    ),
    // 0 means the unknown-code rows fully explain the difference
    #"Added Unexplained" = Table.AddColumn(
        #"Added Status", "Unexplained Hours", each Number.Round([Hours Difference] - [Hours on Unknown Codes], 3), type number
    ),
    #"Sorted Rows" = Table.Sort(#"Added Unexplained", {{"Month", Order.Ascending}})
in
    #"Sorted Rows"
```

</details>

- **Answer: Feb: loaded 7.5 hours MORE than control (3006.5 vs 2999.0); explained by 6 unknown-code rows.**
- ⚠️ If a student gets **"every month shows difference 0.0"** → they removed the unknown codes BEFORE the control check. February then looks fine and the discrepancy is hidden.
</details>

### 🙋 Student question Q8 – How many rows have an unknown project code?
How many timesheet rows carry an unknown project code (not in Projects.xlsx), and how many hours are on them?

<details open><summary><b>Solution Q8</b></summary>

- Steps: Merge Queries ▸ Merge as New, timesheets and Projects on Project Code, Join Kind **Left Anti**. Read the row count and sum the Hours column.
- Rows: CL-P199 (4 rows) and CL-P100 (2 rows), all in the February file.
- **Answer: 6 rows; 7.5 hours**
</details>

### 🙋 Student question Q9 – Refresh with June: total valid hours Jan–Jun
Copy Timesheet_2026-06.csv from the _Later folder into 02_Timesheets and refresh. Remember to remove duplicates and unknown-code rows, and that one Week Ending has a typo year 2062; correct it to 2026. What are the total valid hours year to date (Jan-Jun)?

<details open><summary><b>Solution Q9</b></summary>

- Steps: copy the June file into `02_Timesheets` (Step 12). Data ▸ Refresh All. Check the total of Hours on the result table, after duplicates and unknown codes are removed.
- The June file adds **2,987.5 hours**. The year-typo fix is already in the query (it is the May file that has the 2062 date).
- Cross-check for the Power BI block: labour cost Jan–Jun is **£671,900.00**.
- **Answer: 19,457.5 hours**
</details>

### Common errors and fixes

| What they see | Why | Fix |
|---|---|---|
| Combine stops with a column-not-found error on March | `Proj Code` in the March file | Rename `Proj Code` to `Project Code` in the sample function, with MissingField.Ignore. |
| April weeks fall in the wrong month, or a day is 3 or 4 from what it should be | UK date order used on the US-format April file | Change Type ▸ Using Locale ▸ Date, English (United States), for April only. |
| A Week Ending in 2062 | Year typo in the May file | Replace year 2062 with 2026 (Step 5). |
| Feb control check shows 0.0 difference, not +7.5 | Unknown codes removed before the control check | Duplicate LabourLines as Timesheets_Deduped, delete every step after Removed Duplicates, and check there. |
| Labour total is £29,039.20 for CL-P103 | Duplicates kept | Remove Duplicates after cleaning, not before. |
| Overtime shows nothing, or the overtime line is missing its premium | Multiplier not joined, or cost uses hours × rate only | Merge Labour_Rates on Grade. Labour Cost = Hours × Rate + OT Hours × Rate × (Multiplier − 1). |
| June rows do not appear after Refresh All | The file is still in `_Later` | Copy `Timesheet_2026-06.csv` into `02_Timesheets`, then Refresh All. |

---

## Lab 3 — Stock valuation template ⏱ 11:30–12:30

> **Who asked:** "Create a valuation template consolidated from 4 different reports" (KR), with the rule: Wasp part numbers, Xero latest invoice, Purchase Log latest order if no invoice. Also "Consolidating data from multiple sources." (JE)
> **Why it matters to them:** Stock value is the Wasp count multiplied by a price. The price should come from the latest invoice, else the latest purchase order, else Wasp's own cost. Get the sources wrong and the stock total moves. The check totals show the two traps move it by £2,165.64 and £9,395.18.

**Files:** `data/03_Valuation/Xero_Payable_Invoice_Detail.xlsx` (sheet `Xero`); `data/03_Valuation/Purchase_Log.xlsx` (sheet `Purchase Log`, 48 columns); `data/03_Valuation/Supplier_Name_Map.xlsx` (sheet `Map`, columns From and To). Also uses `Wasp_Clean` from Lab 1, which should still be in this workbook.

If a student did not finish Lab 1, the trainer copy is `solutions/m/lab1_wasp.m`. Load it as `Wasp_Clean` and carry on. If they fall behind in this lab, paste `solutions/m/lab3_valuation.m` as `Stock_Valuation` once the Xero and Purchase Log queries exist (they are named `Xero_Clean` in Step 10 and `PurchaseLog_Clean` in Step 16).

### Teach — step by step

**Step 1 – Xero: connect, remove the four title rows, promote the header**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ Xero_Payable_Invoice_Detail ▸ tick `Xero` ▸ Transform Data. Home ▸ Remove Rows ▸ Remove Top Rows ▸ 4. Home ▸ Use First Row as Headers.
- 🗣 Say: "This report has three title lines and a blank line above the header. Skip four rows and the real header is the first row."
- 🧹 Mess this fixes: title lines `A1:H1`, `A2:H2`, `A3:H3`, and a blank row 4. The header is on row 5.
- ✅ Checkpoint: **158 rows** after promoting the header. That includes a blank row and the Total footer.

**Step 2 – Remove the blank row and the Total footer**
- Click path: Home ▸ Remove Rows ▸ Remove Blank Rows. Then filter Item Code to exclude blanks.
- 🗣 Say: "The footer has no Item Code. A filter on Item Code is enough to remove it, and it stays removed if the footer moves."
- 🧹 Mess this fixes: blank row at sheet row 162, Total footer at row 163.
- ✅ Checkpoint: **156 rows**.

**Step 3 – Fix the text dates**
- Mess: a block of invoice dates is stored as UK text `dd/mm/yyyy`, while the rest are real dates (37 in the file; first ones at A65, A66, A67, A68).
- Fix: the `ToDate` function. It reads real dates as they are and reads text with `Date.FromText(..., [Culture = "en-GB"])`, so `02/01/2026` is 2 January.
- 🗣 Say: "Text that looks like a date is not a date until we tell Power Query the order: day, then month. We tell it en-GB."
- ✅ Checkpoint: no count in the check totals. Sort the Invoice Date column and check the earliest and latest dates look right.

**Step 4 – Unit Price text with a pound sign**
- Mess: Unit Price stored as `£` text at H14, H35, H38 and H41 (8 in the file).
- Fix: `ParseMoney` removes `£` and commas, then reads the number. (It also handles brackets for negatives, which matter in Lab 4.)
- ✅ Checkpoint: Unit Price is a number on every row.

**Step 5 – Clean the text and the references**
- Contact and Description: trailing spaces (B9, B47, B53), double spaces (B78, B112), non-breaking spaces (F6, F21, F43). Fix: `CleanText`.
- Reference: lower case or trailing spaces (D13, D33, D34, D55, D71, D81). Fix: Text.Upper on the cleaned Reference.
- ✅ Checkpoint: no count in the check totals. Reference values all look like upper-case invoice numbers.

**Step 6 – Clean the Item Code**
- Mess: Item Code typed several ways at E6, E8, E9, E10 (106 in the file). Examples: `prt-12`, `PRT 0012`, trailing space.
- Fix: `NormPart`, the same function as Lab 1.
- ✅ Checkpoint: every Item Code reads `PRT-0000`.

**Step 7 – Remove the three credit notes BEFORE choosing any price**
- Click path: filter Source to exclude `Payable Credit Note`.
- Mess: credit-note lines at sheet rows 82, 127 and 149. Each has a negative Quantity and is dated AFTER the latest invoice for that part.
- 🗣 Say: "A credit note is money coming back, not a new price. If we leave it in, the 'latest' line for the part is the credit note, and its price is the wrong one. Remove credit notes first, before anything looks for the latest price."
- ✅ Checkpoint: **153 rows** after removing the three credit notes. Three near-duplicate pairs are still there; the next step removes them.

**Step 8 – Remove the near-duplicate lines**
- Click path: Home ▸ Remove Rows ▸ Remove Duplicates, on Reference and Item Code, keep the first. Use Table.Buffer before it, as in Lab 1.
- Mess: an invoice line imported twice with different case or spacing in Item Code, Contact or Reference, at sheet rows 58, 74 and 113. Steps 5 and 6 make them exact copies first.
- ✅ Checkpoint: **150 invoice lines**.

**Step 9 – Standard supplier name from the map**
- Click path: Home ▸ Merge Queries ▸ Merge Queries as New. Contact on `Supplier_Name_Map` From (Left Outer). Expand To, rename it `Supplier Standard`.
- Mess: Contact name variants (Ltd, Limited, Ltd., typos, caps, truncated). Examples at B6, B8, B10, B11 (99 in the file).
- 🗣 Say: "Merge Queries with the map gives every variant one standard name, so the supplier totals add up."
- ✅ Checkpoint: no count in the check totals.

**Step 10 – Value the Xero lines (for Q10)**
- Add Custom Column `Line Value` = Quantity × Unit Price. Sum it.
- ✅ Checkpoint: the Q10 total is **£1,318,886.68 ex VAT**. See Q10.
- Close & Load To ▸ Only Create Connection. Rename the query **Xero_Clean**.

**Step 11 – Load the Purchase Log and promote its header**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ Purchase_Log ▸ tick `Purchase Log` ▸ Transform Data. Home ▸ Use First Row as Headers.
- Mess: the 48 columns have headers with line breaks (`Ord'd/By`, `Required/By`, e.g. at D1, E1, J1, W1) and trailing spaces (`Part Number `, `Net Unit Price `, `Manufactured Part Numbers ` at K1, L1, Q1).
- Fix: Table.TransformColumnNames with a line-break replace and Text.Trim. The M is `Table.TransformColumnNames(PromotedHeaders, each Text.Trim(Text.Replace(...)))`, see the full M.
- ✅ Checkpoint: no count in the check totals. The header names now have no trailing spaces, so `Net Unit Price` can be used.

**Step 12 – Remove the five blank columns and keep the eight you need**
- Click path: Home ▸ Choose Columns. Tick: PR Number, Date, Project Code, Supplier, Part Number, Qty, Net Unit Price, Notes.
- 🗣 Say: "Five columns have blank headers. Power Query names them Column8, Column16 and so on. Choose Columns keeps the eight we use and drops the rest."
- 🧹 Mess this fixes: five blank columns with blank headers at H1, P1, Z1, AI1, AR1.
- ✅ Checkpoint: **204 data rows** still, with eight columns.

**Step 13 – Dates in two text styles**
- Mess: Date typed as `dd/mm/yyyy` text (C2, C3, C5; 59 in the file) and as `dd-Mon-yyyy` text such as `06-Jan-2025` (C4, C6, C10; 42 in the file).
- Fix: the same `ToDate` function from Step 3. `en-GB` handles both styles.
- ✅ Checkpoint: no count. Every Date is a real date.

**Step 14 – Net Unit Price is text with a pound sign**
- Mess: `£1,234.50`-style text in about 40% of rows (Q2, Q8, Q10, Q11 are examples).
- Fix: `ParseMoney`, the same as Step 4.
- ✅ Checkpoint: Net Unit Price is a number on every row.

**Step 15 – Extract the PO and Jira from Notes**
- Click path: Add Column ▸ Custom Column `PO Number` = Text.BeforeDelimiter([Notes], " / "). Then Custom Column `Jira` = Text.BetweenDelimiters([Notes], "Jira ", " / ").
- Mess: Notes holds `PO-55123 / Jira CL-2231 / urgent` in one cell (AQ2, AQ3, AQ4 and 204 in total). The PO Number column is blank on about 20% of rows (B8, B12, B19 in the file). Ignore that column and derive the PO from Notes for every row.
- 🗣 Say: "Extract pulls the PO and the Jira number out of the one text cell. We stop trusting the PO Number column, because it is blank on one row in five."
- ✅ Checkpoint: every row has a PO Number.

**Step 16 – Remove the cancelled lines**
- Click path: filter Notes to exclude text containing `CANCELLED`. In the M this is `Text.Contains([Notes], "CANCELLED", Comparer.OrdinalIgnoreCase)`, which ignores capitals.
- Mess: cancelled lines at AQ160, AQ165, AQ191 and AQ202.
- ✅ Checkpoint: **204 rows → 4 cancelled → 200 PO lines**. The sum of Qty × Net Unit Price over the 200 lines is **£2,104,847.74**. That equals the clean PO truth used in Power BI.
- Close & Load To ▸ Only Create Connection. Rename the query **PurchaseLog_Clean**.

**Step 17 – Supplier, part number and project code**
- Supplier: variants (I2, I6, I7, I8; 143 in the file). Fix: CleanText, then merge to `Supplier_Name_Map` as in Lab 3 Step 9.
- Part Number: variants (K2, K3, K4, K6; 149 in the file). Fix: `NormPart`.
- Project Code: `cl-p101`, `CL P101`, `CLP101` (G2, G4, G5, G7; 78 in the file). Fix: `NormProject` removes spaces and hyphens, makes upper case, rebuilds as `CL-Pnnn`.
- ✅ Checkpoint: every Project Code reads `CL-Pnnn`.

**Step 18 – Latest price per part (sort, buffer, keep the first)**
- Explain it in one sentence: "Sort so the newest line is at the top for each part, then Remove Duplicates keeps the first line, which is the newest."
- Click path: Home ▸ Sort Ascending on Part Number, then Sort Descending on Invoice Date (or Date), then Table.Buffer, then Remove Duplicates on the part.
- Ties on the same date: break by the higher Reference (Xero) or PO Number (Purchase Log), as the M does.
- Alternative for the advanced group: Group By the part with Max of Date, then merge back on Part and Date. Both give the same answer.
- Do this twice, as Reference queries: right-click Xero_Clean ▸ Reference, rename it Xero_Latest and keep Item Code and Unit Price (rename Unit Price to Xero Price). Right-click PurchaseLog_Clean ▸ Reference, rename it Log_Latest and keep Part Number and Net Unit Price (rename Net Unit Price to Log Price). Credit notes and cancelled lines are already gone, so the first line is a real price.

**Step 19 – Merge the three sources into one row per part**
- Click path: Home ▸ Merge Queries as New on `Wasp_Clean` with `Xero_Latest`, Part Number to Item Code, Left Outer. Expand `Xero Price`. Merge again with `Log_Latest`, Part Number to Part Number, Left Outer. Expand `Log Price`. Use Merge Queries as New so the result is a new query. Rename the final query **Stock_Valuation** in Step 21.
- Remove the old Stock Value from `Wasp_Clean` (it is valued at Wasp cost, and we recalculate).
- 🗣 Say: "Left Outer keeps every Wasp part. If a part has no Xero line, its Xero price is blank, and the next source takes over."
- ✅ Checkpoint: **60 parts**.

**Step 20 – The price waterfall and Price Source**
- Click path: Add Column ▸ Conditional Column ▸ `Price`: if Xero Price is not blank, then Xero Price; else if Log Price is not blank, then Log Price; else Item Cost. Then Add Column ▸ Conditional Column ▸ `Price Source`: if Xero Price is not blank, then `Xero`; else if Log Price is not blank, then `Purchase Log`; else `Wasp`.
- 🗣 Say: "Three rules, checked in order. Xero first, then the Purchase Log, then the Wasp cost. The Price Source column tells us which one won."
- ✅ Checkpoint: **Xero 40, Purchase Log 10, Wasp 10** (60 parts).

**Step 21 – Stock Value and load**
- Add Custom Column `Stock Value` = Round(Total Available × Price, 2).
- Home ▸ Close & Load, query name `Stock_Valuation`.
- ✅ Checkpoint: total stock value **£293,947.12**. By source: Xero **£180,256.41**, Purchase Log **£59,526.00**, Wasp **£54,164.71**. This is Q11.

**Show the traps (do this live, before the questions)**
- Remove the credit-note filter from Step 7 for a moment. Total moves to **£291,781.48**. Three parts change price: PRT-0015, PRT-0016, PRT-0052.
- Remove the cancelled filter from Step 16. Total moves to **£303,342.30**. Four parts change price or source: PRT-0008, PRT-0013, PRT-0022, PRT-0051.
- 🗣 Say: "Both mistakes give a total that looks reasonable. That is why the checks in this lab matter."
- Put the filters back.

### 🙋 Student question Q10 – Total Xero invoice value ex VAT
After cleaning the Xero payable invoice detail (fix dates, prices and item codes; remove - do not net off - the 3 credit-note lines and remove the near-duplicate lines, keeping the first), what is the total invoice value ex VAT, i.e. the sum of Quantity × Unit Price (ex), to 2 dp?

<details open><summary><b>Solution Q10</b></summary>

- Steps: Xero Steps 1–9 above. Filter Source to exclude `Payable Credit Note`. Normalise codes, dates and prices. Remove Duplicates on Reference and Item Code, keep first. Add Custom Column = Quantity × Unit Price (ex). Sum.
- Checkpoints: 158 after promoting, 156 after the blank and Total rows, **153 after removing the 3 credit notes**, **150 after de-duplication**.
- M (full, from `solutions/m/lab3_xero.m`):

<details><summary>Full M code</summary>

```m
// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 3 - Xero payable invoice lines.  Query name: Xero_Clean
// Also needs the Supplier_Name_Map.xlsx file (From/To) in 03_Valuation.
let
    // ---- reusable helper functions (each one is a custom function) ----
    // CleanText: line breaks and CHAR(160) -> space, collapse repeated spaces, trim
    CleanText = (t as nullable text) as nullable text =>
        if t = null then null
        else Text.Trim(Text.Combine(List.Select(Text.Split(Text.Replace(Text.Replace(Text.Replace(t, "#(00A0)", " "), "#(lf)", " "), "#(cr)", " "), " "), each _ <> ""), " ")),
    // NormPart: any spelling of a part number -> PRT-0000 ; anything that is not a PRT part -> null
    NormPart = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                Squashed = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"})),
                Digits = Text.Middle(Squashed, 3)
            in
                if Text.StartsWith(Squashed, "PRT") and Digits <> "" then "PRT-" & Text.PadStart(Digits, 4, "0") else null,
    // ParseMoney: 1234.5 / "£1,234.50" / "(150.00)" -> number ; anything unusable -> null
    ParseMoney = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else
            let
                S = Text.Trim(Text.Replace(Text.Replace(Text.From(v), "£", ""), ",", "")),
                IsNeg = Text.StartsWith(S, "(") and Text.EndsWith(S, ")"),
                N = try Number.FromText(Text.Remove(S, {"(", ")"}), "en-GB") otherwise null
            in
                if N = null then null else if IsNeg then -N else N,
    // ToDate: real Excel dates pass through, UK text dates (dd/mm/yyyy or dd-Mon-yyyy) are parsed with en-GB
    ToDate = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type text) then Date.FromText(Text.Trim(v), [Culture = "en-GB"])
        else Date.From(v),

    // ---- 1. load; the real header is on row 5, so skip the 4 title rows ----
    Source = Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Xero_Payable_Invoice_Detail.xlsx"), null, true),
    XeroSheet = Source{[Item = "Xero", Kind = "Sheet"]}[Data],
    RemovedTopRows = Table.Skip(XeroSheet, 4),
    PromotedHeaders = Table.PromoteHeaders(RemovedTopRows, [PromoteAllScalars = true]),

    // ---- 2. remove blank rows and the 'Total' footer (it has no Item Code) ----
    RemovedBlankRows = Table.SelectRows(PromotedHeaders, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    RemovedFooter = Table.SelectRows(RemovedBlankRows, each [Item Code] <> null),

    // ---- 3. clean text columns and the keys ----
    CleanedText = Table.TransformColumns(RemovedFooter, {{"Contact", CleanText, type text}, {"Source", CleanText, type text}, {"Description", CleanText, type text}}),
    CleanedReference = Table.TransformColumns(CleanedText, {{"Reference", each Text.Upper(CleanText(_)), type text}}),
    NormalisedItemCode = Table.TransformColumns(CleanedReference, {{"Item Code", NormPart, type text}}),

    // ---- 4. credit notes must be removed BEFORE we ever look for the latest price ----
    RemovedCreditNotes = Table.SelectRows(NormalisedItemCode, each [Source] <> "Payable Credit Note"),

    // ---- 5. types: dates (mixed real/text) and prices (some are '£' text) ----
    ConvertedDates = Table.TransformColumns(RemovedCreditNotes, {{"Invoice Date", ToDate, type date}}),
    ConvertedPrices = Table.TransformColumns(ConvertedDates, {{"Unit Price (ex) (GBP)", ParseMoney, type number}}),
    ChangedTypes = Table.TransformColumnTypes(ConvertedPrices, {{"Quantity", Int64.Type}}),
    RenamedPrice = Table.RenameColumns(ChangedTypes, {{"Unit Price (ex) (GBP)", "Unit Price"}}),

    // ---- 6. near-duplicate lines: same Reference + Item Code (after cleaning) -> keep first ----
    BufferedRows = Table.Buffer(RenamedPrice),
    RemovedDuplicates = Table.Distinct(BufferedRows, {"Reference", "Item Code"}),

    // ---- 7. standard supplier name via the From/To map ----
    SupplierMap = Table.PromoteHeaders(Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Supplier_Name_Map.xlsx"), null, true){[Item = "Map", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    MergedSupplier = Table.NestedJoin(RemovedDuplicates, {"Contact"}, SupplierMap, {"From"}, "SupplierMatch", JoinKind.LeftOuter),
    ExpandedSupplier = Table.ExpandTableColumn(MergedSupplier, "SupplierMatch", {"To"}, {"Supplier Standard"}),
    Result = Table.SelectColumns(ExpandedSupplier, {"Invoice Date", "Supplier Standard", "Source", "Reference", "Item Code", "Description", "Quantity", "Unit Price"}),
    RenamedSupplier = Table.RenameColumns(Result, {{"Supplier Standard", "Contact"}})
in
    RenamedSupplier
```

</details>

- **Answer: £1,318,886.68 (1318886.68)**
- ⚠️ If a student gets **1,317,095.21** → they netted the credit notes off instead of removing them.
- ⚠️ If a student gets **1,358,871.85** → they kept the near-duplicate lines.
</details>

### 🙋 Student question Q11 – Total stock valuation with KR's rule
Using KR's valuation rule, what is the total stock valuation across the 60 parts, to 2 dp? Price = the latest Xero invoice Unit Price (by Invoice Date, credit notes excluded); else the latest Purchase Log Net Unit Price (by Date, cancelled lines excluded); else the Wasp Item Cost. Stock Value = Total Available × Price.

<details open><summary><b>Solution Q11</b></summary>

- Steps: Purchase Log Steps 11–17. Latest price per part for Xero and for the Purchase Log (Step 18). Merge both into Wasp_Clean (Step 19). Price and Price Source waterfall (Step 20). Stock Value = Total Available × Price, sum over 60 parts (Step 21).
- Checkpoints: **60 parts**. Price Source: Xero 40, Purchase Log 10, Wasp 10. Purchase Log total ex cancelled lines = £2,104,847.74.
- M (full, from `solutions/m/lab3_purchase_log.m` and `solutions/m/lab3_valuation.m`):

<details><summary>Full M code: Purchase Log</summary>

```m
// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 3 - Purchase Log (48 columns).  Query name: PurchaseLog_Clean
// Also needs Supplier_Name_Map.xlsx in 03_Valuation.
let
    // ---- reusable helper functions (each one is a custom function) ----
    // CleanText: line breaks and CHAR(160) -> space, collapse repeated spaces, trim
    CleanText = (t as nullable text) as nullable text =>
        if t = null then null
        else Text.Trim(Text.Combine(List.Select(Text.Split(Text.Replace(Text.Replace(Text.Replace(t, "#(00A0)", " "), "#(lf)", " "), "#(cr)", " "), " "), each _ <> ""), " ")),
    // NormPart: any spelling of a part number -> PRT-0000 ; anything that is not a PRT part -> null
    NormPart = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                Squashed = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"})),
                Digits = Text.Middle(Squashed, 3)
            in
                if Text.StartsWith(Squashed, "PRT") and Digits <> "" then "PRT-" & Text.PadStart(Digits, 4, "0") else null,
    // NormProject: cl-p101 / CL P101 / CLP101 -> CL-P101
    NormProject = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                S = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"}))
            in
                if Text.StartsWith(S, "CLP") then "CL-" & Text.Middle(S, 2) else S,
    // ParseMoney: 1234.5 / "£1,234.50" / "(150.00)" -> number ; anything unusable -> null
    ParseMoney = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else
            let
                S = Text.Trim(Text.Replace(Text.Replace(Text.From(v), "£", ""), ",", "")),
                IsNeg = Text.StartsWith(S, "(") and Text.EndsWith(S, ")"),
                N = try Number.FromText(Text.Remove(S, {"(", ")"}), "en-GB") otherwise null
            in
                if N = null then null else if IsNeg then -N else N,
    // ToDate: real Excel dates pass through, UK text dates (dd/mm/yyyy or dd-Mon-yyyy) are parsed with en-GB
    ToDate = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type text) then Date.FromText(Text.Trim(v), [Culture = "en-GB"])
        else Date.From(v),

    // ---- 1. load; header is on row 1 ----
    Source = Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Purchase_Log.xlsx"), null, true),
    LogSheet = Source{[Item = "Purchase Log", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(LogSheet, [PromoteAllScalars = true]),

    // ---- 2. headers: line breaks -> space, trim trailing spaces ----
    CleanedHeaders = Table.TransformColumnNames(PromotedHeaders, each Text.Trim(Text.Replace(Text.Replace(_, "#(lf)", " "), "#(cr)", " "))),

    // ---- 3. the 5 empty columns have blank headers, so PQ called them Column8, Column16 ... remove them ----
    BlankColumnNames = List.Select(Table.ColumnNames(CleanedHeaders), each Text.StartsWith(_, "Column")),
    RemovedBlankColumns = Table.RemoveColumns(CleanedHeaders, BlankColumnNames),
    KeptColumns = Table.SelectColumns(RemovedBlankColumns, {"PR Number", "Date", "Project Code", "Supplier", "Part Number", "Qty", "Net Unit Price", "Notes"}),

    // ---- 4. extract PO number and Jira from Notes ("PO-55123 / Jira CL-2231 / urgent") ----
    CleanedNotes = Table.TransformColumns(KeptColumns, {{"Notes", CleanText, type text}}),
    AddedPO = Table.AddColumn(CleanedNotes, "PO Number", each Text.BeforeDelimiter([Notes], " / "), type text),
    AddedJira = Table.AddColumn(AddedPO, "Jira", each Text.BetweenDelimiters([Notes], "Jira ", " / "), type text),

    // ---- 5. cancelled lines out (any capitalisation of CANCELLED) ----
    RemovedCancelled = Table.SelectRows(AddedJira, each not Text.Contains([Notes], "CANCELLED", Comparer.OrdinalIgnoreCase)),

    // ---- 6. keys, dates, prices ----
    CleanedKeys = Table.TransformColumns(RemovedCancelled, {{"Part Number", NormPart, type text}, {"Project Code", each NormProject(_), type text}, {"Supplier", CleanText, type text}, {"PR Number", CleanText, type text}}),
    ConvertedDates = Table.TransformColumns(CleanedKeys, {{"Date", ToDate, type date}}),
    ConvertedPrices = Table.TransformColumns(ConvertedDates, {{"Net Unit Price", ParseMoney, type number}}),
    ChangedTypes = Table.TransformColumnTypes(ConvertedPrices, {{"Qty", Int64.Type}}),
    AddedNetValue = Table.AddColumn(ChangedTypes, "Net Value", each Number.Round([Qty] * [Net Unit Price], 2), type number),

    // ---- 7. standard supplier name via the From/To map ----
    SupplierMap = Table.PromoteHeaders(Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Supplier_Name_Map.xlsx"), null, true){[Item = "Map", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    MergedSupplier = Table.NestedJoin(AddedNetValue, {"Supplier"}, SupplierMap, {"From"}, "SupplierMatch", JoinKind.LeftOuter),
    ExpandedSupplier = Table.ExpandTableColumn(MergedSupplier, "SupplierMatch", {"To"}, {"Supplier Standard"}),
    Result = Table.SelectColumns(ExpandedSupplier, {"PR Number", "PO Number", "Date", "Project Code", "Supplier Standard", "Part Number", "Qty", "Net Unit Price", "Net Value", "Jira"})
in
    Result
```

</details>

<details><summary>Full M code: Stock_Valuation</summary>

```m
// Lab 3 - Stock valuation (KR's rule).  Query name: Stock_Valuation
// Needs three queries already loaded (Connection only is fine): Wasp_Clean, Xero_Clean, PurchaseLog_Clean.
// Price = latest Xero invoice Unit Price (by Invoice Date; credit notes already removed)
//         else latest Purchase Log Net Unit Price (by Date; cancelled lines already removed)
//         else Wasp Item Cost.   Stock Value = Total Available x Price.
let
    // ---- latest Xero price per part: sort newest first (ties: higher Reference first), buffer, keep first row per Item Code ----
    XeroSorted = Table.Buffer(Table.Sort(Xero_Clean, {{"Item Code", Order.Ascending}, {"Invoice Date", Order.Descending}, {"Reference", Order.Descending}})),
    XeroLatest = Table.SelectColumns(Table.Distinct(XeroSorted, {"Item Code"}), {"Item Code", "Unit Price"}),
    XeroLatestRenamed = Table.RenameColumns(XeroLatest, {{"Unit Price", "Xero Price"}}),

    // ---- latest Purchase Log price per part ----
    LogSorted = Table.Buffer(Table.Sort(PurchaseLog_Clean, {{"Part Number", Order.Ascending}, {"Date", Order.Descending}, {"PO Number", Order.Descending}})),
    LogLatest = Table.SelectColumns(Table.Distinct(LogSorted, {"Part Number"}), {"Part Number", "Net Unit Price"}),
    LogLatestRenamed = Table.RenameColumns(LogLatest, {{"Part Number", "Log Part Number"}, {"Net Unit Price", "Log Price"}}),

    // ---- left joins onto the cleaned Wasp parts (drop Wasp_Clean's own cost-based Stock Value; it is recalculated below) ----
    WaspParts = Table.RemoveColumns(Wasp_Clean, {"Stock Value"}),
    JoinedXero = Table.NestedJoin(WaspParts, {"Part Number"}, XeroLatestRenamed, {"Item Code"}, "XeroMatch", JoinKind.LeftOuter),
    ExpandedXero = Table.ExpandTableColumn(JoinedXero, "XeroMatch", {"Xero Price"}),
    JoinedLog = Table.NestedJoin(ExpandedXero, {"Part Number"}, LogLatestRenamed, {"Log Part Number"}, "LogMatch", JoinKind.LeftOuter),
    ExpandedLog = Table.ExpandTableColumn(JoinedLog, "LogMatch", {"Log Price"}),

    // ---- the pricing waterfall ----
    AddedPrice = Table.AddColumn(ExpandedLog, "Price",
        each if [Xero Price] <> null then [Xero Price] else if [Log Price] <> null then [Log Price] else [Item Cost], type number),
    AddedSource = Table.AddColumn(AddedPrice, "Price Source",
        each if [Xero Price] <> null then "Xero" else if [Log Price] <> null then "Purchase Log" else "Wasp", type text),
    AddedStockValue = Table.AddColumn(AddedSource, "Stock Value", each Number.Round([Total Available] * [Price], 2), type number),
    Result = Table.SelectColumns(AddedStockValue, {"Part Number", "Description", "Category", "Total Available", "Item Cost", "Price", "Price Source", "Stock Value", "Stock Status"})
in
    Result
```

</details>

- **Answer: £293,947.12 (293947.12)**
- ⚠️ If a student gets **291,781.48** → they kept the credit notes in Xero (a credit note became the latest price for three parts).
- ⚠️ If a student gets **303,342.30** → they kept the cancelled Purchase Log lines (four parts take a cancelled price or source).
</details>

### 🙋 Student question Q12 – Which price does PRT-0015 use, and why?
(a) What unit price (2 dp) does the valuation use for PRT-0015? (b) Which source does it come from, and why is it not the price on the most recent Xero line for that part?

<details open><summary><b>Solution Q12</b></summary>

- Steps: filter the Xero lines for PRT-0015. There is a credit note dated after the latest invoice. Remove credit notes (Step 7), then the latest invoice line dated 06/02/2026 sets the price.
- Format of the answer: price | source.
- **Answer: 185.66 | Xero**
- ⚠️ If a student gets **148.53** → they kept the credit note, which is dated later than the invoice and has unit price 148.53.
</details>

### Common errors and fixes

| What they see | Why | Fix |
|---|---|---|
| Xero dates show as errors or wrong months | Text dates read in US order | Use ToDate with en-GB, for both `dd/mm/yyyy` and `dd-Mon-yyyy`. |
| Xero total does not match £1,318,886.68 | Credit notes netted off, or near-duplicates kept | Remove credit notes first; remove duplicates on Reference and Item Code, keep first. |
| A part's price is a credit note price (e.g. PRT-0015 shows 148.53) | Credit notes kept when choosing the latest price | Filter Source before any latest-price step. Answer is 185.66, Xero. |
| Stock total is £291,781.48 | Credit notes kept | Same fix as above. |
| Stock total is £303,342.30 | Cancelled Purchase Log lines kept | Filter Notes for CANCELLED (any case) before the latest-price step. |
| Purchase Log PO Number is blank on some rows | Relying on the PO Number column | Derive PO Number from Notes for every row. |
| Purchase Log columns named `Column8`, `Column16` | Five blank columns left in | Choose Columns, or remove every column that starts with `Column`. |
| Some parts have no price at all | Left join not used, or the Price column is built before the merge | Left Outer merge into Wasp_Clean, then the Price waterfall. |

Lunch 12:30–13:15


---

## Lab 4 — Fuzzy matching receipts to customers ⏱ 13:15–14:00

> **Who asked:** "Doing lookups/sumifs between data where you have say a list of Supplier Names that's defined but those Names have been included in the description field in a different report sometimes not with an exact match - fuzzy matching" (PG)
> **Why it matters to them:** Month-end cash allocation. The bank describes each receipt in its own words ("Recd from FLINTWD MARINE - INV 2231"), but the customer ledger says "Flintwood Marine Ltd". Someone has to match them before AR can say who has paid what.

**Files:** `data/04_Receipts/Bank_Receipts_Q1.xlsx` (sheet `Receipts`), `data/04_Receipts/Customer_List.xlsx` (sheet `Customers`), `data/04_Receipts/Fuzzy_Transform_Table.xlsx` (sheet `Transform`, headers From / To, 35 rows)

**Trainer note:** The M in `solutions/m/lab4_receipts_fuzzy.m` and `lab4_customer_summary.m` was checked by review and the figures by independent Python recomputation. It was NOT run in Excel. Dry-run this lab once before the day. Fuzzy merge needs Excel 365 for Windows.

### Teach — step by step
**Step 1 – Show why SUMIFS fails**
- Click path: Open `Bank_Receipts_Q1.xlsx` and `Customer_List.xlsx` side by side in Excel. No query yet.
- 🗣 Say: "Find Ref 10053. The bank says 'Recd from FLINTWD MARINE'. Our customer list says 'Flintwood Marine Ltd'. A SUMIFS looks for an exact match, so it finds nothing for this receipt."
- 🧹 Mess this fixes: in Bank_Receipts_Q1 the Description column has a 'Recd from ' or 'BACS ' prefix (48 instances, logged at C2, C12, C14, C17), a ' - INV 2231' style suffix (34 instances, logged at C2, C3, C4, C11), lower case (C35, C36, C53) and abbreviations or typos such as 'FLINTWD', 'Gosoeberry', 'Flnitwood'.
- ✅ Checkpoint: none yet. The point is that the text does not match.

**Step 2 – Load the receipts and set FolderPath**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ `04_Receipts\Bank_Receipts_Q1.xlsx` ▸ choose sheet `Receipts` ▸ Transform Data.
- 🗣 Say: "We already have the FolderPath parameter from Lab 1. Every file path is built from it, so the whole folder can move without editing queries."
- 🧹 Mess this fixes: Row 1 is the header, so use Home ▸ Use First Row as Headers if it is not promoted already.
- ✅ Checkpoint: 126 receipt rows after promoting headers (includes 3 duplicates, 1 refund, 2 unmatched payers).
- Also load Customer_List.xlsx (sheet Customers) and Fuzzy_Transform_Table.xlsx (sheet Transform): Get Data ▸ From Workbook ▸ Transform Data ▸ Use First Row as Headers ▸ on Customers also Transform ▸ Format ▸ Trim on the Customer column ▸ Close & Load To ▸ Only Create Connection. Name them Customers and Fuzzy_Transform_Table. (The full M below also loads these itself.)

**Step 3 – Fix the types (dates, £ text, brackets)**
- Click path: easiest for a beginner class is Data ▸ Get Data ▸ From Other Sources ▸ Blank Query ▸ Home ▸ Advanced Editor ▸ paste solutions/m/lab4_receipts_fuzzy.m and name the query Receipts_Matched. Walk the Applied Steps. The helper functions ToDate, ParseMoney and ExtractPayer sit at the top of that query.
- 🗣 Say: "`ToDate` reads UK dates. `ParseMoney` strips the £ and the commas, and turns (150.00) into minus 150."
- 🧹 Mess this fixes:
  - Dates stored as text `06/01/2026` (A6, A23, A43) — parse with `Date.FromText(v,[Culture="en-GB"])`, or 06/01 becomes 1 June.
  - Amounts as text `£6,512.59` (D2, D6, D10, D11) — remove £ and commas.
  - The refund shown as text `(150.00)` (D62, Ref 10130, "BACS FLINTWD MARINE - REFUND") — brackets mean negative.
- ✅ Checkpoint: the refund is -150.00 in the Amount column. Do not move on until you see it as a negative number.

**Step 4 – Remove duplicate receipts**
- Click path: Home ▸ Remove Rows ▸ Remove Duplicates (select the Ref column first). Keep the first row.
- 🗣 Say: "The same Ref was imported twice. That is one receipt, not two."
- 🧹 Mess this fixes: duplicate Ref rows at A21:D21, A30:D30, A34:D34.
- ✅ Checkpoint: 123 receipts after Remove Duplicates on Ref.

**Step 5 – Extract the payer text**
- Click path: Add Column ▸ Custom Column ▸ name `Payer Text` ▸ formula `ExtractPayer([Description])`. ExtractPayer is one of the helper functions at the top of Receipts_Matched.
- 🗣 Say: "`ExtractPayer` removes 'Recd from ' or 'BACS ' in any case, then keeps the text before ' - '. What is left is the payer's name."
- Plain-English version of the rule: `Text.Middle` cuts off the prefix; `Text.BeforeDelimiter` style splitting on ' - ' removes the invoice tail. Transform ▸ Extract ▸ Text Before Delimiter (delimiter ' - ') does the tail on its own if you prefer clicks.
- 🧹 Mess this fixes: double spaces (C7, C8, C24), trailing spaces (C9, C15, C16), and non-breaking spaces handled by `CleanText`.
- ✅ Checkpoint: examples now read like `FLINTWD MARINE`, `Gosoeberry Subsea Ltd`, `CASH DEPOSIT`.

**Step 6 – Try the exact Merge first (so students see how few match)**
- Click path: Home ▸ Merge Queries ▸ Merge Queries ▸ choose `Customers` (from Customer_List) ▸ click `Payer Text` and `Customer` ▸ Join Kind: Left Outer ▸ OK. Leave "Use fuzzy matching" UNTICKED.
- 🗣 Say: "Now we try a normal merge. Look at how few rows find a customer."
- 🧹 Mess this fixes: Customer_List A5 and A11 have trailing spaces. Trim them first, or the exact merge loses them too.
- ✅ Checkpoint: only about 3 of 123 receipts match exactly (case-sensitive). About 30 match if you tick ignore case. These are my own counts from the data, not figures in the check files. Confirm on the day. Delete the merge step afterwards; it is only a teaching point.

**Step 7 – Fuzzy merge with the transformation table**
- Click path: Home ▸ Merge Queries ▸ Merge Queries ▸ `Customers` ▸ click `Payer Text` and `Customer` ▸ tick **Use fuzzy matching to perform the merge** ▸ expand **Fuzzy Matching Options**: Similarity threshold **0.80**, tick **Ignore case**, tick **Ignore spaces** if your build shows it (it is on in the M), Maximum number of matches **1**, Transformation table = `Fuzzy_Transform_Table` (From → To). Join Kind: Left Outer ▸ OK.
- 🗣 Say: "Fuzzy matching scores how alike two texts are. The threshold is how alike they must be. The transformation table is a hand-built list: 'FLINTWD MARINE' means 'Flintwood Marine Ltd'. With the table, the match is exact, not a guess."
- 🧹 Mess this fixes: the 35 rows in Fuzzy_Transform_Table (From/To) cover the abbreviations and typos (for example 'FLINTWD MARINE', 'Flnitwood Marine Ltd', 'Dukcworth Defence Systems', 'Gosoeberry Subsea Ltd'). Load `Fuzzy_Transform_Table` as a query first, so it appears in the merge dialog.
- ✅ Checkpoint: 121 receipts matched to a customer; 2 unmatched.
- Honesty note for the trainer: pure fuzzy matching at 0.8 WITHOUT the table is approximate. Some abbreviations and typos score below 0.8, and some score above 0.8 but point at the wrong customer. The transform table is what makes the Lab 4 answers exact. Say this out loud.

**Step 8 – Expand the match and label the unmatched receipts**
- Click path: click the expand icon on the `Customer` column ▸ tick **Customer** only ▸ rename to `Customer Matched`. Then Transform ▸ Replace Values ▸ null ▸ `UNMATCHED`.
- 🗣 Say: "A receipt that matches nothing is not an error we hide. We label it UNMATCHED and send it to AR."
- 🧹 Mess this fixes: C71 "CASH DEPOSIT" (Ref 10131) and C95 "J SMITH BUILDERS" (Ref 10132), the two non-customer payers.
- ✅ Checkpoint: 2 UNMATCHED rows worth 3,375.50 in total. Matched total is 387,392.92. Together they give the 390,768.42 of Q13.

**Step 9 – Group by customer (net £ and count of payments)**
- Click path: Home ▸ Group By ▸ group by `Customer Matched` ▸ **Advanced** ▸ add two aggregates: (1) Sum of `Amount`, name `£ paid?`; (2) a custom aggregate `# payments` = `List.Count(List.Select([Amount], each _ > 0))`. Filter out UNMATCHED first.
- 🗣 Say: "The refund reduces what the customer paid. It is not a payment, so it does not add to the count."
- ✅ Checkpoint: 15 customers. Flintwood Marine Ltd: 42,759.67 and 13 payments. Pennywhistle Offshore Ltd: 688.08 and 1 payment.
- Fill the Customer_List: load this query to a sheet, then type or look up the two columns into `Customer_List.xlsx` (`£ paid?`, `# payments`). The M for this step is `solutions/m/lab4_customer_summary.m`.

### 🙋 Student question Q13 – Total of all receipts after de-duplication and the refund
After removing duplicate receipts (same Ref, keep the first) and treating the (150.00) refund as a negative amount, what is the total of ALL Q1 receipts, including the two receipts from non-customers, to 2 dp?

<details open><summary><b>Solution Q13</b></summary>

- Steps: 1) Load `Bank_Receipts_Q1.xlsx`. 2) Parse dates (en-GB) and amounts (£ text, brackets as negative). 3) Remove Duplicates on Ref (keep first) → 123 rows. 4) Home ▸ Group By with no grouping ▸ Sum of Amount, OR add Amount column total in Excel. Include the UNMATCHED rows.
- M (if useful): `Total = List.Sum(RemovedDuplicates[Amount])` — the M uses `ParseMoney` for `Amount` and `Table.Distinct(BufferedRows, {"Ref"})`.
- **Answer: 390768.42**
- ⚠️ If a student gets **409071.03** → they kept the duplicate receipts (did not remove duplicates on Ref).
- ⚠️ If a student gets **390918.42** → they left the refund out instead of netting it off (the total is 150.00 too high).
- Check: 387,392.92 matched + 3,375.50 unmatched = 390,768.42.

<details><summary>Full M code</summary>

```m
// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 4 - Bank receipts matched to customers.  Query name: Receipts_Matched
// Deterministic answer = ExtractPayer + Fuzzy_Transform_Table. Pure fuzzy at 0.8 WITHOUT the table is approximate
// (abbreviations such as FLINTWD / DEF SYS and the typos score below 0.8) - confirm on the day.
let
    // ---- reusable helper functions (each one is a custom function) ----
    // CleanText: line breaks and CHAR(160) -> space, collapse repeated spaces, trim
    CleanText = (t as nullable text) as nullable text =>
        if t = null then null
        else Text.Trim(Text.Combine(List.Select(Text.Split(Text.Replace(Text.Replace(Text.Replace(t, "#(00A0)", " "), "#(lf)", " "), "#(cr)", " "), " "), each _ <> ""), " ")),
    // NormPart: any spelling of a part number -> PRT-0000 ; anything that is not a PRT part -> null
    NormPart = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                Squashed = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"})),
                Digits = Text.Middle(Squashed, 3)
            in
                if Text.StartsWith(Squashed, "PRT") and Digits <> "" then "PRT-" & Text.PadStart(Digits, 4, "0") else null,
    // ParseMoney: 1234.5 / "£1,234.50" / "(150.00)" -> number ; anything unusable -> null
    ParseMoney = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else
            let
                S = Text.Trim(Text.Replace(Text.Replace(Text.From(v), "£", ""), ",", "")),
                IsNeg = Text.StartsWith(S, "(") and Text.EndsWith(S, ")"),
                N = try Number.FromText(Text.Remove(S, {"(", ")"}), "en-GB") otherwise null
            in
                if N = null then null else if IsNeg then -N else N,
    // ToDate: real Excel dates pass through, UK text dates (dd/mm/yyyy or dd-Mon-yyyy) are parsed with en-GB
    ToDate = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type text) then Date.FromText(Text.Trim(v), [Culture = "en-GB"])
        else Date.From(v),

    // ExtractPayer: strip 'Recd from ' / 'BACS ' (any case), then keep the text before ' - ' (removes ' - INV 2231')
    ExtractPayer = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                C = CleanText(t),
                NoPrefix = if Text.StartsWith(C, "Recd from ", Comparer.OrdinalIgnoreCase) then Text.Middle(C, 10)
                           else if Text.StartsWith(C, "BACS ", Comparer.OrdinalIgnoreCase) then Text.Middle(C, 5)
                           else C
            in
                Text.Trim(List.First(Text.Split(NoPrefix, " - "))),

    // ---- 1. load receipts (header on row 1) ----
    Source = Excel.Workbook(File.Contents(FolderPath & "04_Receipts\Bank_Receipts_Q1.xlsx"), null, true),
    ReceiptSheet = Source{[Item = "Receipts", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(ReceiptSheet, [PromoteAllScalars = true]),

    // ---- 2. types: mixed real/text dates, '£' text and '(150.00)' amounts ----
    ConvertedDates = Table.TransformColumns(PromotedHeaders, {{"Date", ToDate, type date}}),
    ConvertedAmounts = Table.TransformColumns(ConvertedDates, {{"Amount", ParseMoney, type number}}),
    ChangedRef = Table.TransformColumnTypes(ConvertedAmounts, {{"Ref", Int64.Type}}),

    // ---- 3. same Ref twice = duplicate receipt, keep the first ----
    BufferedRows = Table.Buffer(ChangedRef),
    RemovedDuplicates = Table.Distinct(BufferedRows, {"Ref"}),

    // ---- 4. extract the payer text ----
    AddedPayer = Table.AddColumn(RemovedDuplicates, "Payer Text", each ExtractPayer([Description]), type text),

    // ---- 5. customers and transformation table ----
    CustomerSource = Excel.Workbook(File.Contents(FolderPath & "04_Receipts\Customer_List.xlsx"), null, true){[Item = "Customers", Kind = "Sheet"]}[Data],
    Customers = Table.TransformColumns(Table.PromoteHeaders(CustomerSource, [PromoteAllScalars = true]), {{"Customer", CleanText, type text}}),
    TransformSource = Excel.Workbook(File.Contents(FolderPath & "04_Receipts\Fuzzy_Transform_Table.xlsx"), null, true){[Item = "Transform", Kind = "Sheet"]}[Data],
    TransformTable = Table.PromoteHeaders(TransformSource, [PromoteAllScalars = true]),

    // ---- 6. fuzzy match Payer Text to Customer ----
    FuzzyJoined = Table.FuzzyNestedJoin(AddedPayer, {"Payer Text"}, Customers, {"Customer"}, "CustomerMatch", JoinKind.LeftOuter,
        [Threshold = 0.8, IgnoreCase = true, IgnoreSpace = true, NumberOfMatches = 1, TransformationTable = TransformTable]),
    ExpandedCustomer = Table.ExpandTableColumn(FuzzyJoined, "CustomerMatch", {"Customer"}, {"Customer Matched"}),
    LabelledUnmatched = Table.ReplaceValue(ExpandedCustomer, null, "UNMATCHED", Replacer.ReplaceValue, {"Customer Matched"}),
    Result = Table.SelectColumns(LabelledUnmatched, {"Date", "Ref", "Description", "Payer Text", "Amount", "Customer Matched"})
in
    Result
```

</details>

</details>

### 🙋 Student question Q14 – Flintwood Marine Ltd: net £ paid and number of payments
How much did Flintwood Marine Ltd pay in Q1 (net of its refund), and how many payments is that? Count only positive receipts as payments; the refund reduces the total but is not a payment.

<details open><summary><b>Solution Q14</b></summary>

- Steps: 1) Filter `Customer Matched` = `Flintwood Marine Ltd`. 2) Sum `Amount` (the refund is included as minus 150). 3) Count only rows where `Amount` is greater than 0.
- Or use the Group By output from Step 9 and find Flintwood Marine Ltd.
- **Answer: 42759.67 / 13 payments**
- Note: the check file has no separate trap value for this question. A student who counts every row will get one more payment than the answer, because the refund is also a row. The answer counts positive receipts only.
- Note: Flintwood's receipts include "FLINTWD MARINE", "Flintwood Marine", "Flnitwood Marine Ltd" and "Recd from FLINTWD MARINE". They all match only because of the transformation table.
</details>

### 🙋 Student question Q15 – Receipts still unmatched after the transformation table
After matching with the supplied Fuzzy_Transform_Table, which receipts (Ref numbers) remain unmatched to any customer?

<details open><summary><b>Solution Q15</b></summary>

- Steps: after the fuzzy merge, filter `Customer Matched` = `UNMATCHED` (or null). Read the Ref column.
- **Answer: 10131, 10132** (10131 = CASH DEPOSIT; 10132 = J SMITH BUILDERS; combined value 3,375.50)
- The check file has no trap answer for this question. The mistake to watch for is trusting a fuzzy match at 0.8 without the table. That can give a different list, so do not accept it as the final answer.
</details>

### Common errors and fixes
| What they see | Why | Fix |
|---|---|---|
| Only about 3 receipts match on the exact merge | Expected. Text differs (prefix, suffix, case, typos) | That is the teaching point. Go on to the fuzzy merge |
| "Recd from" still in Payer Text | Prefix is case-sensitive in the custom function | Use `Comparer.OrdinalIgnoreCase` in the StartsWith test (as in the M) |
| Total 150.00 too high (390,918.42 for the whole list) | Refund left out instead of netted off | Make sure `ParseMoney` handles brackets so `(150.00)` = -150 and is kept |
| Total 300.00 too high | Refund counted as +150 | Check the bracket handling in `ParseMoney` |
| Dates become wrong months (06/01 = 1 June) | Date parsed as US culture | `Date.FromText(v,[Culture="en-GB"])` |
| "We couldn't convert to Number" on Amount | `£` or thousands comma still in the text | Use `ParseMoney` or remove `£` and `,` first |
| Fuzzy merge matches a customer that is wrong | Threshold too low or no transformation table | Use threshold 0.80 and the transformation table. Check the Customer Matched column |
| Total 409,071.03 instead of 390,768.42 | Duplicates kept | Remove Duplicates on Ref, keep first |
| Fuzzy merge option greyed out | Fuzzy matching only in Excel 365 for Windows (limited on Mac) | Use a Windows machine |

---

## Lab 5 — Reshape, roll-forward & version control ⏱ 14:00–14:35

> **Who asked:** "How can you manage version control without losing the Power Query rules - if you want to keep each month's query separate and not overwrite it" (CP)
> "Can it be used to roll forward spreadsheets each month" (CP)
> "What to look out for in terms of how to save our files - what would cause a power query link to be lost, can it be re-linked?" (CP)
> **Why it matters to them:** Budget vs actual and the fixed asset register are rebuilt every month-end. Rolling forward by hand copies errors forward each time.

**Files:** `data/05_Budget_Assets/Project_Budget_2026.xlsx` (sheet `Budget`), `data/05_Budget_Assets/Asset_Movements_2026.xlsx` (sheets `Movements`, `Opening_Cost`, `Roll_Forward_Template`), `data/05_Budget_Assets/_Later/Asset_Movements_2026-11.xlsx` (copy in for the November step), and the `LabourLines` query from Lab 2 (`data/02_Timesheets`).

**Trainer note:** The M files are `solutions/m/lab5_unpivot_budget.m`, `lab5_pivot_hours.m` and `lab5_rollforward.m`. They were reviewed, not run in Excel. Dry-run before the day.

### Part A — Unpivot the budget

### Teach — step by step
**Step 1 – Load the budget and remove the title rows**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ `05_Budget_Assets\Project_Budget_2026.xlsx` ▸ sheet `Budget` ▸ Transform Data. Then Home ▸ Remove Rows ▸ Remove Top Rows ▸ **3** ▸ Home ▸ Use First Row as Headers.
- 🗣 Say: "Three rows of title above the table. Remove them, then the real header becomes row one."
- 🧹 Mess this fixes: title and subtitle rows (A1, A2) and a blank row. The header is on sheet row 4.
- ✅ Checkpoint: no check-file figure at this step. Check that the headers read `Project Code`, `Project Name`, `Column3`, `Jan-26` … `Dec-26`, `Total`.

**Step 2 – Remove the blank spacer column and the Total column BEFORE unpivoting**
- Click path: select `Column3` and `Total` (both headers) ▸ Home ▸ Remove Columns ▸ Remove Columns.
- 🗣 Say: "If we unpivot with the Total column still there, Total becomes a 'month'. Take it out first."
- 🧹 Mess this fixes: blank spacer column C4 (comes in as `Column3`), and the Total column at P4.

**Step 3 – Clean the project codes**
- Click path: select the `Project Code` column ▸ Transform ▸ Format ▸ UPPERCASE, then Transform ▸ Format ▸ Trim.
- 🗣 Say: "Codes must match the Projects list exactly, so 'cl-p104' and 'CL-P106 ' with a trailing space must be fixed."
- 🧹 Mess this fixes: `cl-p104` in lower case (A8); `CL-P106 ` with a trailing space (A10).

**Step 4 – Keep only real project rows**
- Click path: Click the Project Code filter arrow ▸ Text Filters ▸ Begins With ▸ `CL-P`.
- 🗣 Say: "The Total row and the footnote do not start with CL-P, so they drop out here."
- 🧹 Mess this fixes: Total row at A13; footnote row at A15.

**Step 5 – Unpivot the month columns**
- Click path: select `Project Code` and `Project Name` ▸ Transform ▸ Unpivot Columns ▸ **Unpivot Other Columns**. Rename the two new columns `Month Label` and `Budget Text`.
- 🗣 Say: "Twelve month columns become twelve rows per project. That is 96 rows for 8 projects."
- ✅ Checkpoint: 96 rows (8 projects × 12 months).

**Step 6 – Turn 'Jan-26' text into a date, and fix the text number**
- Click path: Add Column ▸ Custom Column ▸ `Month` = `Date.FromText("1 " & Text.Start([Month Label],3) & " 20" & Text.End(Text.Trim([Month Label]),2), "en-GB")` ▸ Type Date. For the amount: Transform ▸ Replace Values ▸ `,` with nothing ▸ Transform ▸ Data Type ▸ Decimal Number.
- 🗣 Say: "The header 'Jan-26' is text. We make it a real date so we can filter for H1."
- 🧹 Mess this fixes: "Jan-26" text headers (D4 onward); number stored as text in F9 for CL-P105 March ('45,000.00' style).
- ✅ Checkpoint: total of all projects is 3,705,000.00, which equals the Budget 2026 column in Projects.xlsx.

**Step 7 – Filter H1 and sum**
- Click path: filter `Month` between 01/01/2026 and 30/06/2026 ▸ filter `Project Code` = CL-P105 ▸ Home ▸ Group By (Sum of Budget GBP) or read the sum from the Excel table.
- ✅ Checkpoint: CL-P105 H1 budget = 198,551.00 (Q16).

### 🙋 Student question Q16 – CL-P105 H1 budget after unpivoting
Unpivot the wide budget (remove title rows, spacer/Total columns and Total/footnote rows first; fix lower-case code and the text number). What is the total budget for CL-P105 for H1 (Jan-Jun 2026)?

<details open><summary><b>Solution Q16</b></summary>

- Steps: Skip 3 title rows → promote headers → remove `Column3` and `Total` → UPPERCASE and Trim the code → keep codes starting `CL-P` → Unpivot Other Columns → date from `Jan-26` text → fix the text number in March → filter `Project Code` = CL-P105 and months Jan to Jun → sum `Budget GBP`.
- M: `solutions/m/lab5_unpivot_budget.m` (the whole query). The text-number fix is the `Budget GBP` step.
- **Answer: 198551.0** (GBP 198,551.00)
- The check file has no trap value for this question. The likely mistake is leaving CL-P105 March as text, which drops that month from the sum. Check that March is in the H1 total.
</details>

### Part B — Pivot hours by project and month

### Teach — step by step
**Step 8 – Start from the Lab 2 query**
- Click path: in the Queries pane, right-click `LabourLines` ▸ **Reference**. Rename the new query `Hours_Pivot`.
- 🗣 Say: "We do not reload the timesheets. We reference the clean line-level query from Lab 2. That way Lab 2 and Lab 5 always agree."
- ✅ Checkpoint: Jan to Jun hours, valid rows only (June comes in from the `_Later` copy of the timesheets, as in Q9).

**Step 9 – Keep three columns and name the months**
- Click path: Home ▸ Choose Columns ▸ `Project Code`, `Month`, `Hours`. Then Add Column ▸ Custom Column ▸ `Month Name` = `Date.ToText([Month],[Format="MMM",Culture="en-GB"])`.
- 🗣 Say: "The pivot needs one word per column heading. 'Jan', 'Feb' and so on."
- ✅ Checkpoint: one row per timesheet line with a month name.

**Step 10 – Pivot**
- Click path: select `Month Name` ▸ Transform ▸ Pivot Column ▸ Values Column = `Hours` ▸ Advanced Options ▸ Aggregate Value Function = **Sum** ▸ OK. Then Transform ▸ Replace Values (null → 0) on the month columns.
- 🗣 Say: "Rows are projects, columns are months, cells are hours."
- ✅ Checkpoints: 8 project rows. Grand total of hours Jan to Jun = **19,457.5**. Top project is CL-P101 with 3,327.0 hours. CL-P102 in March is 501.5 hours (Q17).

### 🙋 Student question Q17 – CL-P102 hours in March
Pivot the cleaned hours (from Lab 2 incl. June) so Project Code is the rows and months Jan-Jun are columns. How many hours were booked to CL-P102 in Mar-26?

<details open><summary><b>Solution Q17</b></summary>

- Steps: reference `LabourLines` → keep Project Code, Month, Hours → month name → Pivot Column (Hours, Sum) → read the row CL-P102, column Mar.
- M: `solutions/m/lab5_pivot_hours.m`.
- **Answer: 501.5** (hours)
- The check file has no trap value for this question. The likely mistake is pivoting before removing duplicates or unknown codes, which would overstate hours. Use the Lab 2 query, so the rules carry over.
</details>

### Part C — Monthly roll-forward with a Period parameter

### Teach — step by step
**Step 11 – Create the Period parameter**
- Click path: Home ▸ Manage Parameters ▸ New Parameter ▸ Name `Period` ▸ Type Decimal Number (or Whole Number) ▸ Current Value **10** ▸ OK.
- 🗣 Say: "This one cell decides which month we are looking at. Change it, refresh, and the whole roll-forward moves to that month."
- ✅ Checkpoint: `Period` appears in the Queries pane under Parameters. `FolderPath` already exists from Lab 1.

**Step 12 – Paste the roll-forward query**
- Click path: Data ▸ Get Data ▸ From Other Sources ▸ Blank Query ▸ Home ▸ Advanced Editor ▸ paste `solutions/m/lab5_rollforward.m` ▸ Done ▸ rename the query `Cost_Rollforward`.
- 🗣 Say: "The query reads every `Asset_Movements_2026*` file in the 05 folder. It does not read one fixed file, so a new month just needs a new file dropped in."
- ✅ Checkpoint: the query shows a small table with five lines: Cost Bfwd, Additions, Disposals, Write offs, Cost Cfwd.

**Step 13 – Fix the mess inside the movement file (this is what the M does for you)**
- 🧹 Mess this fixes, in `Asset_Movements_2026.xlsx`:
  - Title rows above the header (A1 to A3). The header is on row 4. `Table.Skip` 3.
  - Asset Class blank on every row except the first of each class (merged-cell style). Fill Down on `Asset Class`.
  - Subtotal rows inside the data: 'Total Plant & Machinery' (A15), 'Total Vehicles' (A26), 'Total IT Equipment' (A37), 'Total Fixtures & Fittings' (A48), and Grand Total (A49). They have no Period, so filter out rows where Period is blank.
  - Spacer column F is empty. Keep only the five needed columns.
  - Disposals as text in brackets, e.g. `(25,000.00)` (D25). Use `fnNum`.
  - Write Offs as text `£1,500.00` (E10).
  - Zero disposals shown as `-` (D6). `fnNum` returns 0.
  - Blank vs zero additions (C5, mixed). Null counts as 0.
  - One month with a period but blank amounts (C29, IT Equipment, period 3). Null counts as 0.
  - Opening_Cost sheet: IT Equipment is `£38,000.00` as text (B6). Opening total is 238,000.00.
- ✅ Checkpoint: 40 class-period rows before November (4 classes × 10 periods) after Fill Down and removing subtotals.

**Step 14 – Check Month 10 (the client's example)**
- Click path: keep `Period` = 10 ▸ Data ▸ Refresh All.
- 🗣 Say: "Month 10 must match the client's example on the Roll_Forward_Template sheet. If it does not, stop and find out why before we go on."
- ✅ Checkpoint: Cost Bfwd 250,000.00 ▸ Additions 10,000.00 ▸ Disposals -25,000.00 ▸ Write offs 0 ▸ Cost Cfwd **235,000.00**.

**Step 15 – Drop in the November file and roll forward**
- Click path: in File Explorer, copy `data\05_Budget_Assets\_Later\Asset_Movements_2026-11.xlsx` into `data\05_Budget_Assets\` (the copy, not a move). In Excel: Data ▸ Get Data ▸ Launch Power Query Editor ▸ Home ▸ Manage Parameters ▸ select Period ▸ Current Value 11 ▸ Home ▸ Close & Load, then Data ▸ Refresh All.
- 🗣 Say: "Nothing else changes. No new query, no copy-paste. The file comes in, the parameter moves to November, and the same rules run again."
- 🧹 Mess this fixes: the November file has the same title rows, Total rows and blank column as the main file. Same fixes apply.
- ✅ Checkpoint: Month 11 — Cost Bfwd 235,000.00 ▸ Additions 28,700.00 ▸ Disposals -900.00 ▸ Write offs -600.00 ▸ Cost Cfwd **262,200.00**.
- Note: the `Roll_Forward_Template` sheet has Month 10 filled and Month 11 blank. The lab fills it by refreshing and copying the output into column C.

### 🙋 Student question Q18 – Cost Cfwd for Month 11
Build the roll-forward with a Period parameter. Check Month 10 (Bfwd 250,000.00 / Additions 10,000.00 / Disposals (25,000.00) / Write offs 0 / Cfwd 235,000.00). Then copy Asset_Movements_2026-11.xlsx from _Later into 05_Budget_Assets, set Period = 11 and refresh. (The 1 January opening cost per asset class is on the Opening_Cost sheet of Asset_Movements_2026.xlsx.) What is Cost Cfwd for Month 11?

<details open><summary><b>Solution Q18</b></summary>

- Steps: Period = 10 and check the Month 10 example → copy the November file into `05_Budget_Assets` → Period = 11 → Refresh.
- Formula: Cost Bfwd(period) = opening 238,000.00 + all Additions − Disposals − Write offs in earlier periods. Cost Cfwd = Bfwd + Additions − Disposals − Write offs for the period. Disposals and write offs are shown as negatives in the M output.
- Month 11 by line: Bfwd 235,000.00; Additions 28,700.00; Disposals −900.00; Write offs −600.00.
- M: `solutions/m/lab5_rollforward.m`.
- **Answer: 262200.0** (GBP 262,200.00)
- The check file has no trap value for this question. Common mistakes are leaving the November file in `_Later` (the query never sees it, so Cost Cfwd stays at 235,000.00), or forgetting to change Period to 11 (the figure shown is still Month 10).
</details>

### Teach — version control and saving files (answers R6, R8 and R13)
Answer these three questions from the client before the break. Do not number them. They are not part of Q16 to Q18.

**Point 1 – One master workbook, not one copy per month**
- 🗣 Say: "Do not copy the query every month. Keep ONE master workbook with ONE set of rules. The month is a parameter. Every month you change `Period` and refresh. The rules never get overwritten, because nobody edits them."
- Why it matters: the rules are in the query, not in a pasted value. A copied query drifts. Someone fixes Lab 5 in one copy and forgets the others.

**Point 2 – Keep each month's answer as a values-only snapshot**
- Click path for a monthly archive: after the refresh, Copy the output table ▸ Paste Special ▸ Values into a new workbook, then File ▸ Save As `Cost_Rollforward_2026-11.xlsx` in an archive folder (for example `Archive\2026\`). Alternatively, Save As the whole master file with the month in the name, then re-open the master.
- 🗣 Say: "The archive is what was reported, as values. The master keeps the rules for the next month. Do not archive a file that still has a live link to a folder that may move."

**Point 3 – What breaks a Power Query link, and how to re-link**
- A moved, renamed or deleted source file or folder. The `FolderPath` parameter now points to nothing. Error: "We couldn't find a part of the path".
- A renamed sheet. `Source{[Item = "Budget", Kind = "Sheet"]}` fails.
- A renamed column. `Table.SelectColumns` fails with "column not found".
- A renamed table. Table names are used by the Power Automate flow in Lab 6 (`Transactions`, `CardholderSummary`).
- Privacy levels. Power Query blocks a combine of two sources with different privacy levels. Set them under Data ▸ Get Data ▸ Query Options ▸ Privacy.
- Credentials. Stored logins expire or change. Data ▸ Get Data ▸ Data Source Settings ▸ select source ▸ Edit Permissions.
- OneDrive path vs local path. A file in OneDrive has a URL-style path on one PC and a `C:\Users\...\OneDrive\` path on another. The same query can work on one machine and fail on another.
- Re-linking. Yes, it can be re-linked:
  - Data ▸ Get Data ▸ Data Source Settings ▸ select the source ▸ **Change Source** ▸ pick the file. Or, better for the team, change the `FolderPath` parameter (Data ▸ Get Data ▸ Launch Power Query Editor ▸ Home ▸ Manage Parameters ▸ select FolderPath ▸ change the Current Value ▸ Home ▸ Close & Load) once, and every query follows.
  - After a rename, check the sheet and column names still match. A rename is the one case the parameter does not fix.
- Setup path for students: `C:\CallenLenz_Course\data\` (with the trailing backslash).

**Discussion prompt (not numbered, not part of Q16–Q18)**
- 🗣 Say: "Asha saves the master to her own OneDrive and opens it on her laptop. Her `FolderPath` says `C:\CallenLenz_Course\data\`. What breaks, and what is the fix?"

<details open><summary><b>Solution — discussion prompt</b></summary>

- What breaks: the data folder is not at `C:\CallenLenz_Course\data\` on her laptop, so every `File.Contents` and `Folder.Files` call fails with "We couldn't find a part of the path".
- Fix: change the `FolderPath` parameter to the real local path of her copy of the data folder. Refresh. Every query follows, because they all use the parameter. Do not edit the queries.
- Lesson: keep the data in one agreed folder and keep the parameter as the only place that holds the path.
</details>

### Common errors and fixes
| What they see | Why | Fix |
|---|---|---|
| "Total" appears as a month in the unpivoted table | Total column not removed before unpivoting | Remove `Total` first, then Unpivot |
| CL-P105 H1 total is too low by one month | March stored as text '45,000.00' | Replace the comma and change type to Decimal Number |
| Totals include 'Total Plant & Machinery' | Subtotal rows have no Period | Filter out rows where Period is blank, after Fill Down |
| Disposals sum to 0 or error | `(25,000.00)` or `-` as text | Use the `fnNum` function (brackets = negative, dash = 0) |
| Month 10 Cost Cfwd is not 235,000.00 | An IT Equipment opening, a blank, or a dash was not handled | Check Opening_Cost (B6 `£38,000.00`), and null → 0 on the three movement columns |
| Month 11 is unchanged after refresh | November file still in `_Later`, or `Period` still 10 | Copy the file into `05_Budget_Assets`, set `Period` = 11, Refresh All |
| "column 'X' of the table wasn't found" | A column was renamed in the source | Rename it back, or update the step that names it |
| "We couldn't find a part of the path" | FolderPath points to the wrong folder | Change the `FolderPath` parameter |
| "Formula.Firewall" or privacy error | Two sources with different privacy levels | Set privacy under Query Options ▸ Privacy |

---

## Lab 6 — Card statements, split per person & Power Automate ⏱ 14:35–15:00

> **Who asked:** "Can PowerQuery be used to split data into tabs with each change in name, inputting data into specific columns and create a mail merge to send to individuals?" (CP)
> **Why it matters to them:** Every month each cardholder gets asked to justify their own card spend. Today someone splits the statement by hand and copies rows into emails.

**Files:** `data/06_Cards/Card_Statement_2026-01.csv`, `Card_Statement_2026-02.csv`, `Card_Statement_2026-03.csv`, `data/06_Cards/FX_Rates.xlsx` (sheet `FX_Rates`, headers in row 4), `data/00_Reference/Employees.xlsx` (sheet `Sheet1`, the card last four is the key).

**Do NOT use** `data/06_Cards/Employee_Lookup.xlsx`. It has spelling mismatches on names (for example `Rajesh Patell`, `Surname, First` order) and it is a distractor. Join on Last Four, never on name.

**Trainer note:** M in `solutions/m/lab6_cards.m` and `lab6_fn_cardholder.m`. Not run in Excel. The Power Automate flow in `solutions/m/lab6_power_automate.md` is the trainer demo. Do not send real emails in the demo.

### Teach — step by step
**Step 1 – Combine the three statements from the folder**
- Click path: Data ▸ Get Data ▸ From File ▸ From Folder ▸ `06_Cards` (from FolderPath) ▸ Transform Data ▸ filter `Name` begins with `Card_Statement_` ▸ Combine Files (or paste `lab6_cards.m` as the query named `Transactions`).
- 🗣 Say: "Three months of statements, three different layouts. We read each file by its columns, not by its header names."
- 🧹 Mess this fixes:
  - March has three title lines above the header (line 4 is the header). Skip rows until the cell reads `Date`. Mess log: Card_Statement_2026-03.csv line 1.
  - February headers have trailing spaces and different case: `Posted date `, `Last four `. Rename by position, because all three files have the same 15 columns. Mess log: Card_Statement_2026-02.csv line 1.
- ✅ Checkpoint: 270 data rows after combining (includes pending rows, month-end overlaps and blank-field lines).

**Step 2 – Remove the blank lines**
- Click path: Home ▸ Remove Rows ▸ Remove Blank Rows, or filter `Bank Ref` does not equal blank.
- 🗣 Say: "Blank lines are on every statement. A row with no Bank Ref is not a transaction."
- 🧹 Mess this fixes: Card_Statement_2026-01.csv line 40, 2026-02.csv line 69, 2026-03.csv line 62.
- ✅ Checkpoint: 267 rows.

**Step 3 – Remove Pending rows (keep Posted only)**
- Click path: Home ▸ filter `Status` ▸ tick **Posted**. Filter must ignore case: `POSTED` in upper case also counts. Use `Text.Upper([Status]) = "POSTED"` in a custom filter if needed.
- 🗣 Say: "A pending row has not been posted yet and has no posted date. We do not count it until it posts, or we count it twice."
- 🧹 Mess this fixes: Pending rows at lines 83 to 87 (Jan), 89 to 93 (Feb), 92 to 96 (Mar). Five per statement. Status in upper case (`POSTED`) at several lines in each file.
- ✅ Checkpoint: 15 pending rows removed.

**Step 4 – Clean the Bank Ref and remove duplicates**
- Click path: Transform ▸ Format ▸ Trim, then Transform ▸ Format ▸ UPPERCASE on `Bank Ref` ▸ Home ▸ Remove Duplicates on `Bank Ref` (keep first).
- 🗣 Say: "The last few days of one statement are repeated at the top of the next. Same Bank Ref, same transaction. Keep one. Clean the Ref first, or 'bk91391037 ' and 'BK91391037' look different."
- 🧹 Mess this fixes: month-end overlap rows (Feb statement lines 2, 4, 5, 7; Mar statement lines 5, 7, 8, 10). Near-duplicates whose Bank Ref differs only by case or a trailing space (Feb lines 3, 6; Mar lines 6, 9).
- ✅ Checkpoint: 12 duplicates removed. 240 rows left, which is 80 per month.

**Step 5 – Dates**
- Click path: Transform ▸ Data Type ▸ Date, or `Date.FromText(_, "en-GB")` in a custom column.
- 🗣 Say: "The February statement writes dates as 05-Feb-2026 text. The other two use 27/02/2026. The en-GB culture handles both."
- 🧹 Mess this fixes: text dates `29-Jan-2026` style across the February statement (86 rows). Dates on the first rows of February are January transactions. Use the **Date** column (transaction date), not the Posted date. This matters for Q19.

**Step 6 – Fix the amounts (text, CR refunds, blanks, zeros)**
- Click path: Transform ▸ Replace Values (remove £ and commas), Transform ▸ Data Type ▸ Decimal Number. For the CR refunds use the custom function `fnAmount` in `lab6_cards.m` (a value ending `CR` is negative).
- 🗣 Say: "A refund shows as '45.00 CR'. In the sum it must be minus 45."
- 🧹 Mess this fixes:
  - Amount as text `£1,250.00` (Jan lines 34, 54, 60, 68; Feb 55, 67; Mar 52, 91).
  - Refunds as text `45.00 CR` (Jan line 52, Feb line 28, Mar line 51).
  - Amount blank on a GBP row (Jan 37, 62, 76; Feb 12, 29, 83; Mar 13, 23, 46). Use the Original Currency Amount with rate 1.00.
  - Amount 0.00 on a foreign-currency row (Jan lines 6, 11, 81; Feb 37, 42, 75; Mar 34, 66). Use the Original Currency Amount times the rate.
  - Amount blank on a foreign-currency row (Jan line 4, Feb line 32, Mar line 20) and on one refund row (Feb line 14, where Original Currency Amount is already negative). Same rule: Original Currency Amount × rate.
  - Merchant with double spaces, city in upper or lower case (Transform ▸ Format ▸ Capitalize Each Word, Transform ▸ Format ▸ Trim, Clean).
- ✅ Checkpoints: 21 rows where the amount is blank or 0 (after dedupe). 4 refund rows. 8 foreign-currency rows.

**Step 7 – Work out Amount GBP with FX_Rates**
- Click path: Home ▸ Merge Queries ▸ `Transactions` with `FX_Rates` ▸ match `Merchant Currency Description` to `Currency Description` (both UPPER and TRIM first, as the M does) ▸ Left Outer ▸ expand `GBP per 1 unit`. Then Add Column ▸ Custom Column `Amount GBP` = `if [Statement Amount] <> null and [Statement Amount] <> 0 then [Statement Amount] else Number.Round([Original Currency Amount] * [FX Rate], 2)`.
- 🗣 Say: "Use the card statement amount if it is there. If it is blank or 0, work it out from the original amount and the rate. GBP is 1.00, USD 0.79, EUR 0.86."
- 🧹 Mess this fixes: FX_Rates.xlsx has title rows above the header (header on row 4). Remove top rows 3.
- ✅ Checkpoint: Q1 net spend, all cardholders, = **19,615.46**.

**Step 8 – Restore the leading zeros in Last Four**
- Click path: Transform ▸ Data Type ▸ Text, then Transform ▸ Format ▸ or custom column `Text.PadStart([Last Four], 4, "0")`.
- 🗣 Say: "`Text.PadStart` adds leading zeros up to a set length. So 937 becomes 0937. Without it, Priya Nair's card will not match her employee record."
- 🧹 Mess this fixes: the card's last four stored as a number, so the leading zero is lost. Priya Nair's card is 0937; Aisha Rahman's is 0184. The mess log flags 21 rows in the January statement (for example line 4) and 14 rows in the February statement (for example line 3).
- ✅ Checkpoint: Priya Nair and Aisha Rahman both have transactions. Without this step their rows are missing (63 rows would fail to match, per the check file).

**Step 9 – Match each card to a person (join on Last Four, not name)**
- Click path: Home ▸ Merge Queries ▸ `Transactions` with `Employees` (from `00_Reference\Employees.xlsx`, sheet `Sheet1`) ▸ match `Last Four` to `Card Last Four` ▸ Left Outer ▸ expand `Employee Name`, `Department`, `Email`.
- 🗣 Say: "We join on the card number, not the name. Names are spelled differently in different systems. The card number is the key that does not change."
- 🧹 Mess this fixes: the name spelling mismatches in `Employee_Lookup.xlsx` (not used here). Discussion point: why keys beat names. Email in upper case (Employee_Lookup D13), fix with Lower. Name on Card is lower or mixed case (for example p danvers, r patel) on 21 rows across the three statements. It is not used for the join: a discussion point.
- ✅ Checkpoint: 240 rows, each with an Employee Name.

**Step 10 – Per-person query: fnCardholder**
- Click path: Data ▸ Get Data ▸ From Other Sources ▸ Blank Query ▸ Home ▸ Advanced Editor ▸ paste `solutions/m/lab6_fn_cardholder.m` ▸ rename to `fnCardholder`.
- 🗣 Say: "A function is a recipe you can run for any name. Give it a name, get back that person's rows."
- Invoke: type a name in the function's input, for example `fnCardholder("Steve Mallory")`. Use Table.RowCount and the Sum of `Amount GBP` to check the figures.
- ✅ Checkpoint: each cardholder's Q1 net total is in the JSON (Aisha Rahman 2,352.92 is the highest; Marcus Lowe 1,334.13 is the lowest).

**Step 11 – One tab per person (PivotTable, Show Report Filter Pages)**
- Click path: Insert ▸ PivotTable ▸ from the `Transactions` table ▸ new sheet. Drag `Date`, `Merchant`, `Category`, `City` to **Rows**. Drag `Amount GBP` to **Values** (Sum). Drag `Employee Name` to **Filters**.
- Then PivotTable Analyze ▸ Options (the arrow under Options) ▸ **Show Report Filter Pages** ▸ pick `Employee Name` ▸ OK.
- 🗣 Say: "Excel makes one sheet per name, filtered to that person. Ten names, ten tabs. No copy-paste."
- ✅ Checkpoint: 10 new tabs, one per cardholder. Design ▸ Report Layout ▸ Show in Tabular Form, so each tab reads like a list.
- Note: the Show Report Filter Pages option is greyed out unless `Employee Name` is in the Filters area.

### 🙋 Student question Q19 – Callum Reid's February spend (net GBP)
Combine the three card statements, keep Posted rows only, dedupe on Bank Ref, and match to Employees on Last Four (restore leading zeros: 937 becomes 0937). Amount rule: if Amount is blank or 0, use Original Currency Amount x the rate in FX_Rates.xlsx (GBP = 1.00); an amount ending 'CR' is a refund and must be negative. What is the total GBP spend (net of refunds) for Callum Reid using the transaction Date (not the Posted date) in February 2026?

<details open><summary><b>Solution Q19</b></summary>

- Steps: after the Step 1 to Step 9 query, filter `Employee Name` = `Callum Reid` ▸ filter `Date` in February 2026 (use the Date column, not Posted Date) ▸ sum `Amount GBP`.
- M: `lab6_cards.m` (query `Transactions`). Then in a new query or the Excel table: `Table.SelectRows(Transactions, each [Employee Name] = "Callum Reid" and Date.Month([Date]) = 2 and Date.Year([Date]) = 2026)`, then `List.Sum`.
- **Answer: 1043.54**
- ⚠️ If a student gets **1040.41** → they used the Posted date instead of the transaction Date. Some February transactions post in March.
- ⚠️ The check file also lists **1043.54** under "leading zeros not restored". That is the correct answer, because Callum Reid's card has no leading zero, so the trap does not change Q19. It does change the totals for Priya Nair and Aisha Rahman (see Q20).
</details>

### 🙋 Student question Q20 – Highest net Q1 spend
Create the function fnCardholder (lab6_fn_cardholder.m) and invoke it for the cardholder with the highest net Q1 spend. Who is it, how much did they spend (net GBP) and how many transactions? Hint: two cards start with 0 (0937, 0184); if Last Four lost its leading zero those cardholders will show missing rows.

<details open><summary><b>Solution Q20</b></summary>

- Steps: create `fnCardholder` (Step 10) ▸ invoke it for each of the ten names (or Group By `Employee Name` on `Transactions` to see the ranking) ▸ for the top name, read the sum of `Amount GBP` and the row count.
- Q1 net spend by cardholder, from the check file, for comparison: Aisha Rahman 2,352.92 (33 transactions); Priya Nair 2,332.35 (30); Steve Mallory 2,202.15 (24); Oliver Hartley 2,169.61 (27); Callum Reid 2,168.44 (21); Fiona Sutherland 2,067.19 (28); Rajesh Patel 1,810.16 (18); Paul Danvers 1,664.48 (23); Angus McBride 1,514.03 (21); Marcus Lowe 1,334.13 (15).
- **Answer: Aisha Rahman; 2352.92; 33 transactions**
- The check file has no trap value for this question. The mistake is skipping the leading-zero fix. Aisha Rahman's card is 0184. Without the fix her rows do not match, her total is understated, and the ranking changes. Priya Nair (0937) is affected too.
</details>

### Step 12 – TRAINER DEMO (students watch): Power Automate, email each cardholder their own transactions

Do this live from `solutions/m/lab6_power_automate.md` (Option A). Students watch; no student builds it. Do not send real emails. Swap every Email value in `CardholderSummary` for your own test address first.

**Before the demo (prep, already in the workbook)**
1. Save `Card_Flow.xlsx` in OneDrive for Business or SharePoint (not personal OneDrive, not a local drive).
2. Sheet `Transactions`: load the `Transactions` query to a table named exactly `Transactions`. Keep these columns: Date, Merchant, Category, City, Amount GBP, Last Four, Employee Name, Email.
3. Do this in a copy: right-click Transactions ▸ Reference, add the `Date.ToText([Date], "dd/MM/yyyy")` step there (capital MM is months; lower mm is minutes), make Last Four text, and load that copy to the table the flow reads. Keep the original `Transactions` unchanged for Q19.
4. Sheet `CardholderSummary`: Reference `Transactions` ▸ Group By `Employee Name`, `Email`, `Last Four` ▸ Count Rows, and Sum of `Amount GBP` named `Total GBP` ▸ load to a table named `CardholderSummary` (10 rows).
5. Close the workbook in Excel before the flow runs.

**The flow, step by step**
1. Go to make.powerautomate.com ▸ Create ▸ **Instant cloud flow** ▸ name `Email card transactions` ▸ trigger **Manually trigger a flow**.
2. Action 1: **+ New step** ▸ Excel Online (Business) ▸ **List rows present in a table**. Rename it `List CardholderSummary`. Table = `CardholderSummary`.
3. Action 2: **List rows present in a table**, rename `List Transactions`. Table = `Transactions`. Settings ▸ Pagination ON, threshold 5000. Put this BEFORE the loop, so the sheet is read once.
4. Action 3: **Control** ▸ **Apply to each**. Select an output = `value` from `List CardholderSummary`. Everything below goes inside this loop.
5. Inside the loop, action 3a: **Data Operation** ▸ **Filter array**, rename `Filter this cardholder`. From = `value` (of `List Transactions`). Switch to advanced mode and enter:
   `@equals(item()?['Last Four'], items('Apply_to_each')?['Last Four'])`
6. Action 3b: **Data Operation** ▸ **Create HTML table**, rename `Build transaction table`. From = `Body` of `Filter this cardholder`. Columns = Custom, with these pairs:
   - Date = `item()?['Date']`
   - Merchant = `item()?['Merchant']`
   - Category = `item()?['Category']`
   - City = `item()?['City']`
   - Amount GBP = `formatNumber(item()?['Amount GBP'], 'N2', 'en-GB')`
7. Action 3c: **Office 365 Outlook** ▸ **Send an email (V2)**, rename `Email cardholder`.
   - To = `items('Apply_to_each')?['Email']`
   - Subject = `concat('Your card transactions - ', items('Apply_to_each')?['Employee Name'])`
   - Body (use the `</>` code view): a greeting using `first(split(items('Apply_to_each')?['Employee Name'],' '))`, the count `items('Apply_to_each')?['Transactions']`, the total `formatNumber(items('Apply_to_each')?['Total GBP'],'N2','en-GB')`, and then `body('Build_transaction_table')`.
8. Save ▸ **Test** ▸ Manually ▸ Run flow. Open Run history and click each action to show Inputs and Outputs.

**What to say during the demo**
- 🗣 Say: "Read the list once, loop over ten people, filter each person's rows by card number, build a small table, send one email."
- Common failures to point at: "Table not found" (the table name must match and the file must be closed); only 256 rows (turn on pagination); an empty Filter array (Last Four type mismatch, text vs number); dates showing as 46054 (convert to text in Power Query first).
- Expected result: 10 emails, each listing only that person's rows. The 10 totals add up to the Q1 total of `Transactions`, which is **19,615.46**.

**Alternative: Word mail merge (one letter per person)**
- Word ▸ Mailings ▸ Select Recipients ▸ Use an Existing List ▸ `Card_Flow.xlsx`, sheet `CardholderSummary` ▸ insert merge fields (Employee Name, Transactions, Total GBP) ▸ Finish & Merge ▸ Send Email Messages (needs Outlook). Use Edit individual documents first to preview. The transaction list still needs a per-person table, which PivotTable pages (Step 11) can provide.

**Break / switch to Power BI Desktop at 15:00**

### Common errors and fixes
| What they see | Why | Fix |
|---|---|---|
| March statement will not load (CSV header error) | Three title lines above the header | Skip rows until the first cell reads `Date` (the M does this) |
| Priya Nair or Aisha Rahman has far fewer rows | Leading zero lost on Last Four | `Text.PadStart(text, 4, "0")` after converting to text |
| Totals are too high | Pending rows kept, or duplicate Bank Refs kept | Filter Status = Posted; Trim and UPPER the Bank Ref, then Remove Duplicates |
| Feb transactions land in the wrong month | Using the Posted date, not the transaction Date | Filter on the `Date` column |
| Amount is an error on a refund | '45.00 CR' text | Use `fnAmount` (CR = negative) |
| Foreign rows show 0 or blank GBP | Blank or 0 Amount, no FX match | Use Original Currency Amount × rate. Trim and UPPER the currency before merging |
| Employee Name is blank for some rows | Last Four not matched | Check the leading zero first, then the Employees file (`Card Last Four`) |
| Show Report Filter Pages is greyed out | Employee Name not in the Filters area | Drag it into Filters first |
| Power Automate: "Table not found" | Table name differs, or file still open | Rename the table to `Transactions` / `CardholderSummary`, save and close the file |
| Power Automate: empty Filter array | Last Four stored as a number in one place | Keep Last Four as text in both the Transactions table and the Employees lookup |


---

## Power BI — Project-cost dashboard ⏱ 15:00–17:00

> **Who asked:** "Is it better to have multiple, isolated powerquery documents – or is it better to have less, larger models?" (JE) · "How does Powerquery feed into PowerBI?" (JE) · "How do PowerQuery, PowerPivot and PowerBI work together?" (CP) · "What are Data Cubes in PowerBI" (CP)
> **Why it matters to them:** Labs 1–6 cleaned the Wasp stock, timesheets, Purchase Log, Xero invoices, budget and cards. This block puts that cleaned data into one Project Cost model, so the project totals, budget variance and stock value come from one place and agree with each other. It also shows them what they would have to keep working once the course is over.

**Files:**
- `data/00_Reference/Projects.xlsx`, `data/00_Reference/Employees.xlsx`
- `data/01_Wasp_Stock/Wasp_Stock_Export.xlsx`
- `data/02_Timesheets/Timesheet_2026-01.csv` to `Timesheet_2026-05.csv`, and `Timesheet_2026-06.csv` (in `_Later`, must be in `02_Timesheets` for the Jan–Jun checks)
- `data/03_Valuation/Purchase_Log.xlsx`, `Xero_Payable_Invoice_Detail.xlsx`, `Supplier_Name_Map.xlsx`
- `data/05_Budget_Assets/Project_Budget_2026.xlsx`
- The M files in `solutions/m/` (`lab1_wasp.m`, `lab3_purchase_log.m`, `lab3_xero.m`, `lab3_valuation.m`, `lab2_timesheets.m`, `lab5_unpivot_budget.m`, optional `lab2_project_month_summary.m`)

**Timing and period.** Every check in this block is **Jan–Jun 2026**, with June in the folder. Q1 and Q2 are asked in Block 0, so they are not repeated here.

| Time | Segment |
|---|---|
| 15:00–15:15 | Concepts (Steps 1–4) |
| 15:15–15:40 | Get data (Steps 5–10) |
| 15:40–16:00 | Star schema, Date table, relationships |
| 16:00–16:10 | ☕ 10-minute break |
| 16:10–16:30 | DAX measures (Q21, Q22, Q25) |
| 16:30–16:50 | Report page (Q23, Q24, Q26) |
| 16:50–16:55 | Refresh, publish, scheduled refresh |
| 16:55–17:00 | Close the day |

> **Trainer note on the June file.** The Jan–Jun answers need `Timesheet_2026-06.csv` in `02_Timesheets`. Lab 2 already brings June in, so check the students' folder at Step 5. If June is missing, copy it in from `_Later` at Step 5. The live "drop June in and refresh" demo is in the 16:50 segment. Do not hold June back, or Q21–Q24 will not match.

**Setup assumption.** The data folder is `C:\CallenLenz_Course\data`, as in the course setup. Power BI Desktop is Windows-only. Anyone on a Mac should pair up.

### Teach — step by step

**Step 1 – Put the three tools in one picture (concepts)**
- Click path: open Power BI Desktop with a blank report. Show the three icons on the left rail: Report view, Table view, Model view. Show Home ▸ Transform data (this opens Power Query).
- 🗣 Say: "Power Query gets and shapes one table at a time, and it runs when you press Refresh. Power Pivot is the data model. It holds the tables, the links between them and the calculations (DAX). Power BI Desktop is both of those plus the report canvas. Same engines, three places." This answers **R11** (CP): Power Query shapes, the model relates and calculates, Power BI shows it.
- Draw this on the whiteboard:
```text
 SOURCES --> POWER QUERY --> DATA MODEL (Power Pivot) --> REPORT (visuals)
 Excel,       get & shape     tables + relationships       charts, slicers
 CSV, SQL     (M language)    + measures (DAX)             Power BI Desktop
                              [same engine: VertiPaq]      --> Service (share, schedule)
```
- 🧹 Mess this fixes: none yet. Labs 1–6 fixed the mess. This step shows where the clean output goes.
- ✅ Checkpoint: ask the room. "Unpivot the budget: which tool?" (Power Query.) "Total Project Cost = Labour + PO?" (DAX.) "One slicer filters two tables?" (Relationship.)

**Step 2 – "What are Data Cubes in Power BI?" (R12)**
- 🗣 Say: "A cube is the old name for a set of dimensions and measures you can slice. A classic OLAP cube is built by IT and queried with MDX. Power BI does not use cubes. It uses a tabular semantic model: tables, relationships and DAX measures, held in memory. Our Project Cost model is the cube. Project, Employee, Part and Date are the dimensions. Labour Cost and PO Cost are the measures. The model is the cube, and every report is a different view of it."
- Translation table for the room:

| Cube word | Power BI word |
|---|---|
| Dimension | Dim table (DimProject) |
| Fact | Fact table (FactLabour) |
| Measure | DAX measure (Total Project Cost) |
| Hierarchy (Year > Month) | Hierarchy on DimDate |

- 🧹 Mess this fixes: none.
- ✅ Checkpoint: "Could you rebuild the cube as a model?" Yes: the model is built the same way as the lab queries.

**Step 3 – One big model or many small ones? (JE, R4)**
- 🗣 Say: "One focused model per subject area, built as a star schema, with shared dimensions. For Callen-Lenz that is one Project Cost model: labour, purchases, budget and stock all share Project, Employee, Part and Date. Split a model only when the audience, the security or the refresh time is different. Cards are the example: they hold personal spend, so they would sit in their own model with tighter access. Define each measure once, so two reports never give two answers." This is the Power BI answer to **R4** and to Q2 (which is asked in Block 0).
- 🧹 Mess this fixes: none. Lab 2 already put the timesheet names and codes in one shape, so the dimensions can be shared.
- ✅ Checkpoint: ask "Where is Total Project Cost defined?" Answer: once, as a measure in the model.

**Step 4 – How Power Query feeds Power BI (JE, R5)**
- 🗣 Say: "Power BI runs Power Query for you. Each query is a list of M steps. Refresh runs those steps again on the source files and loads the result into the model. So the M you wrote in Excel is the same M you paste here. Three ways to bring it in: paste the M (what we do today), import the queries from an Excel workbook, or load the cleaned output tables from a workbook. We paste the M so you can see every step." This is the answer to **R5**: Power Query fills the model tables, refresh replays the steps, and the model holds the result in memory.
- 🧹 Mess this fixes: none new. The M in `solutions/m/` contains the Lab 1–6 fixes.
- ✅ Checkpoint: the same numbers come out of the pasted M as in Labs 1–6 (see Step 7).

**Step 5 – Save the file and create the FolderPath parameter**
- Click path: File ▸ Save as ▸ `C:\CallenLenz_Course\CallenLenz_ProjectCost.pbix`. Then Home ▸ Transform data ▸ Home ▸ Manage Parameters ▸ New Parameter. Name `FolderPath`, Type `Text`, Current Value `C:\CallenLenz_Course\data\` (keep the trailing backslash). Click OK.
- Before you go on, check that `02_Timesheets` holds `Timesheet_2026-06.csv`. If not, copy it from `02_Timesheets\_Later` now.
- 🗣 Say: "A parameter is a named setting. Every query reads FolderPath, so if the folder moves, you change one value, not forty queries."
- 🧹 Mess this fixes: none. The folder must hold the whole `data/` tree (all sub-folders, including `06_Cards`, which Lab 6 uses).
- ✅ Checkpoint: FolderPath is listed in the Queries pane. Save the file.

**Step 6 – Paste the M queries (Blank Query ▸ Advanced Editor)**
- Click path, for each row of the table: Home ▸ New Source ▸ Blank Query ▸ Home ▸ Advanced Editor ▸ select all ▸ paste the file ▸ Done ▸ in Properties, rename the query ▸ Enter. Do them in this order. Later queries depend on earlier ones.

| # | Paste this file | Rename query to | Mess it fixes (from the mess logs) | ✅ Checkpoint |
|---|---|---|---|---|
| 1 | `solutions/m/lab1_wasp.m` | `Wasp_Clean` | Two title rows above the header (`Wasp_Stock_Export.xlsx`, A1:X1); repeated header at A31:X31; junk test row A20; duplicate part at A33:X33; "Total" footer at A75; #REF! and N/A costs (J17, J10); £ text costs | 60 distinct valid parts (check_totals_B Q3) |
| 2 | `solutions/m/lab3_purchase_log.m` | `PurchaseLog_Clean` | Five blank columns with blank headers (H1 sheet); four CANCELLED lines (e.g. AQ160); project codes such as `cl-p101`, `CL P101`, `CLP101` (G2); headers with line breaks | 200 lines after removing cancelled; all-date total £2,104,847.74 (check_totals_B L3_pl) |
| 3 | `solutions/m/lab3_xero.m` | `Xero_Clean` | Three credit-note lines (A82:H82 and two more); near-duplicate invoice lines (A58:H58); a block of UK text dates (A65); £ text prices (H14); supplier name variants (B6); "Total" footer (A163) | 150 invoice lines after dedupe; ex-VAT total £1,318,886.68 (Q10) |
| 4 | `solutions/m/lab3_valuation.m` | `Stock_Valuation` | Latest price rule (Xero, else Purchase Log, else Wasp): this is the Lab 3 KR rule, not a mess fix | 60 parts (Xero 40, Purchase Log 10, Wasp 10); £293,947.12 (Q11) |
| 5 | `solutions/m/lab2_timesheets.m` | `LabourLines` (ends at step "Final Labour Lines") | Hours as "7.5 hrs" (Jan); project and employee codes in lower case, with spaces or no hyphen; "Proj Code" header in March (`renamed_column`); US dates in April (355 rows); bare employee numbers in May (451 rows); year typo 2062 (May, line 148); duplicates; Total footer (Jan, line 456); unknown codes CL-P199 and CL-P100 (Feb, 6 rows) | 2,320 rows (Jan–Jun); Labour Hours 19,457.5 (Q9); Labour Cost £671,900.00 |
| 6 | `solutions/m/lab5_unpivot_budget.m` | `Budget_Unpivot` | Title rows above the header; blank spacer column C; Total column P; footnote row A15; Total row A13; `cl-p104` (A8); text budget `45,000.00` (F9); month headers as text `Jan-26` (D4) | 96 rows (8 projects x 12 months); full-year 2026 budget £3,705,000.00 |
| 7 (optional) | `solutions/m/lab2_project_month_summary.m` | `ProjectMonthSummary` | None (a Group By on `LabourLines`) | Connection only. Not in the model |

- Order matters: row 3 before row 4 (`Stock_Valuation` reads `Xero_Clean`, `Wasp_Clean` and `PurchaseLog_Clean`). Row 5 before row 7.
- `Wasp_Clean` has its own Stock Value column at Wasp cost. Never take Stock Value from it. `Stock_Valuation` recomputes Stock Value with the KR rule.
- If a query shows "Formula.Firewall: ... references other queries or steps", set File ▸ Options and settings ▸ Options ▸ Current File ▸ Privacy ▸ Ignore the Privacy Levels. This is common with `Stock_Valuation`.

**Step 7 – Check the totals before you go on**
- 🗣 Say: "We do not trust a model until the totals agree with the labs. Each check below should match the lab answer exactly."
- ✅ Checkpoint: `Stock_Valuation` total = £293,947.12 (Q11, same as Lab 3). `LabourLines` = 2,320 rows. `Budget_Unpivot` = 96 rows.
- 🧹 Mess this fixes: none new. This is the proof that Labs 1–6 and this model agree.

**Step 8 – Reference the model tables (shape the tables)**
- Click path: right-click a query ▸ Reference (creates a new query that builds on it). Rename each one exactly as shown. Load only the seven model tables (3 dimensions, 4 facts). Everything else is Connection only: right-click ▸ untick Enable load.

| Model table | Built from | Remove or rename | Columns kept |
|---|---|---|---|
| `DimProject` | Blank query reading `00_Reference\Projects.xlsx` (sheet `Sheet1`) | Remove `Budget 2026 GBP` (budget lives in FactBudget; keeping it double counts) | Project Code, Project Name, Customer, Project Manager, Status |
| `DimEmployee` | Same pattern, `Employees.xlsx` (sheet `Sheet1`) | Remove Email, Card Last Four, Manager (personal data, not needed) | Employee ID, Employee Name, Department, Grade |
| `DimPart` | Reference of `Wasp_Clean` | Do not keep Wasp's Stock Value | Part Number, Description, Category |
| `FactLabour` | Reference of `LabourLines` | Remove Employee Name, Project Name, Month, Weekly Hours, Standard Rate (keep Standard Rate only if you want an audit column) | Week Ending, Employee ID, Project Code, Task, Hours, OT Hours, Labour Cost |
| `FactPurchases` | Reference of `PurchaseLog_Clean` | Rename `Date` to `Order Date`; rename `Net Value` to `Line Value` | PO Number, Order Date, Project Code, Part Number, Qty, Net Unit Price, Line Value |
| `FactBudget` | Reference of `Budget_Unpivot` | Remove Project Name | Project Code, Month, Budget GBP |
| `FactStockValuation` | Reference of `Stock_Valuation` | Nothing | Part Number, Price Source, Stock Value |

- The DimProject query, if a student wants to type it instead of copying it (small M written for this block; it is not in `solutions/m/`):
```m
let
    Source = Excel.Workbook(File.Contents(FolderPath & "00_Reference\Projects.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    Removed = Table.RemoveColumns(Source, {"Budget 2026 GBP"})
in
    Removed
```
- Set the data types: Week Ending, Order Date and Month = **Date**; Hours, Labour Cost, Line Value, Budget GBP and Stock Value = **Fixed decimal number**; all codes = **Text**.
- Click Home ▸ Close & Apply. Save.
- 🧹 Mess this fixes: the Budget 2026 GBP column in `Projects.xlsx` would double count against FactBudget. This is a design choice, not a file mess.
- ✅ Checkpoint: seven model tables loaded: DimProject 8, DimEmployee 25, DimPart 60, FactLabour 2,320, FactPurchases 200, FactBudget 96, FactStockValuation 60 (the date table comes in Step 13).

**Step 9 – Watch the Purchase Log dates: the Year = 2026 gotcha**
- 🗣 Say: "The Purchase Log has orders from 2025 as well as 2026. If you add up everything, the PO figure is far too big for this year."
- Click path: in Report view, add a card with PO Cost (create the measure in Step 15 if needed). Without a filter it shows **£2,104,847.74**. That is every PO line in the file.
- 🧹 Mess this fixes: not a Lab 3 mess. This is the period trap. The fix is the Date table and a Year/Month filter (Steps 13 and 16).
- ✅ Checkpoint: with Year = 2026 and Month Number ≤ 6, PO Cost is **£608,095.02** (51 PO lines dated Jan–Jun 2026).

**Step 10 – Check the data types and the errors banner**
- Look for yellow or red banners on each query. Fix any errors before you go on. Common causes: a typo in FolderPath, an Excel file from the folder still open (`~$` lock files), or the privacy setting from Step 6.
- ✅ Checkpoint: no red error on any query after Close & Apply.

**Step 11 – Star schema: the model view**
- Click path: click the Model view icon (left rail, third icon). Arrange the tables into a star. The dimensions go around the outside, the facts in the middle.
```text
 DIMENSIONS (filter the numbers)            FACTS (the numbers)

 DimProject (8)  --Project Code-->  FactLabour (2,320)   Hours, OT Hours, Labour Cost
 DimProject      --Project Code-->  FactPurchases (200)  Qty, Net Unit Price, Line Value
 DimProject      --Project Code-->  FactBudget (96)      Budget GBP
 DimEmployee (25)--Employee ID--->  FactLabour
 DimPart (60)    --Part Number--->  FactPurchases
 DimPart         --Part Number--->  FactStockValuation (60)   Stock Value (a snapshot)
 DimDate (730)   --Date-->          FactLabour       (via Week Ending)
 DimDate         --Date-->          FactPurchases    (via Order Date)
 DimDate         --Date-->          FactBudget       (via Month)

 Facts never link to each other. They meet through the shared dimensions.
```
- 🗣 Say: "A star: the dimension is the small list of things you slice by (projects, people, parts, days). The fact is the big list of things you add up (hours, cost, budget). Every fact points to the same dimensions, so one slicer filters every table it reaches."
- 🧹 Mess this fixes: none. Two messes give this shape: the project codes (Lab 2, Lab 3, Lab 5) and the part numbers (Lab 1, Lab 3) must match exactly, which the Labs already enforce.
- ✅ Checkpoint: the diagram has 9 relationship lines.

**Step 12 – Relationships (Model view, one-to-many, single direction)**
- Click path: drag from the dimension key (the "one" side) to the fact column (the "many" side). Check the dialog: Cardinality **1 : *** (one-to-many), Cross filter direction **Single**, **Active** ticked. Click OK.

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

- 🗣 Say: "Filters flow downhill, from the dimension to the fact. Keep the arrow single. Never set it to both directions to make a report work. If Power BI offers a many-to-many link, stop: a dimension has a duplicate key somewhere (check DimProject and DimPart)."
- 🧹 Mess this fixes: duplicate or variant keys. If a dimension has two rows for one project code, the totals come out too high. The Labs stop that with dedupe steps (Lab 1 dedupe on Item Number; Lab 3 Project Code normalising).
- ✅ Checkpoint: 9 relationships, all 1 to *, all single direction. Stock value has no date link: that is correct, it is a snapshot. Say so out loud.

**Step 13 – Date table: DimDate, marked as a date table**
- Click path: Modeling ▸ New table. Paste the DAX below. Then select the Month column ▸ Column tools ▸ Sort by column ▸ Month Number.
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
- Then right-click DimDate in the Data pane ▸ Mark as date table ▸ Date column = `Date` ▸ OK.
- Note: the guide also gives a Power Query (M) version of the same table in `solutions/powerbi_build.md` section 3.2. Use one or the other, not both.
- 🗣 Say: "The date table makes the time functions work. The Year = 2026 and Month ≤ 6 filter lives here. Budget is held by month, so we slice budget by Month, not by a single day."
- 🧹 Mess this fixes: the Purchase Log 2025 orders (Step 9). The date table and filter remove them from the 2026 totals.
- ✅ Checkpoint: DimDate has **730 rows** (1 Jan 2025 to 31 Dec 2026). Mark as date table is ticked.

**Step 14 – Hide the clutter**
- Right-click (or the eye icon in Model view) to hide the key columns and the numeric columns we will replace with measures. Hide: FactLabour[Project Code, Employee ID, Week Ending], FactPurchases[Project Code, Part Number, Order Date], FactBudget[Project Code, Month], FactStockValuation[Part Number], and the raw numeric columns (Hours, OT Hours, Labour Cost, Line Value, Qty, Net Unit Price, Budget GBP, Stock Value).
- 🗣 Say: "Hiding a column makes people use the measure, so everyone gets the same sum. If someone drags a raw column in, they get an unchecked 'Sum of'."
- ✅ Checkpoint: the Data pane shows measures and dimension columns only.

☕ **10-minute break (16:00–16:10).**

**Step 15 – Measures table and DAX measures (16:10–16:30)**
- Click path: Home ▸ Enter data. Leave it empty, name the table `_Measures`, Load. Select `_Measures`, then Modeling ▸ New measure for each line below. Copy each line exactly.

```dax
Labour Hours = SUM ( FactLabour[Hours] )
Labour Cost = SUM ( FactLabour[Labour Cost] )
PO Cost = SUM ( FactPurchases[Line Value] )
Total Project Cost = [Labour Cost] + [PO Cost]
Budget = SUM ( FactBudget[Budget GBP] )
Variance = [Budget] - [Total Project Cost]
Variance % = DIVIDE ( [Variance], [Budget] )
Stock Value = SUM ( FactStockValuation[Stock Value] )
YTD Total Cost = TOTALYTD ( [Total Project Cost], DimDate[Date] )
```
Optional, for Q26:
```dax
Budget (safe) = IF ( ISFILTERED ( DimEmployee ), BLANK (), [Budget] )
```

| Measure | What it does, in one line |
|---|---|
| Labour Hours | Adds up the hours on every labour line. |
| Labour Cost | Adds up the cost on every labour line (including the overtime rule from Lab 2). |
| PO Cost | Adds up the Line Value on every purchase line (Qty x Net Unit Price, from Lab 3). |
| Total Project Cost | Labour Cost plus PO Cost: the one cost number for a project. |
| Budget | Adds up the monthly project budget. |
| Variance | Budget minus Total Project Cost. Positive means under budget. |
| Variance % | Variance divided by Budget. `DIVIDE` returns blank instead of an error if Budget is zero. |
| Stock Value | Adds up the stock value per part (snapshot, no date). |
| YTD Total Cost | Total Project Cost from 1 January up to the date on the visual. |
| Budget (safe) | Shows Budget only when no employee is selected (`ISFILTERED` checks for a slicer on DimEmployee). |

- Formats: select the measure ▸ Measure tools. Currency measures: GBP, 0 decimals on cards, 2 decimals on checks. Variance %: Percentage, 1 decimal.
- 🗣 Say: "Filter context. Each row of a table filters the model to that row's project. The relationships carry the filter to the fact tables. That is why the same formula gives a different number in each row."

- ✅ Checkpoint values (Jan–Jun 2026, Year = 2026 and Month Number ≤ 6):

| Measure | Value |
|---|---|
| Labour Hours | 19,457.5 |
| Labour Cost | £671,900.00 |
| PO Cost | £608,095.02 |
| Total Project Cost | £1,279,995.02 |
| Budget | £1,894,429.00 (full year £3,705,000.00) |
| Variance | £614,433.98 (derived: 614,433.98 ÷ 1,894,429.00 = 32.4%) |
| Stock Value (no date) | £293,947.12 |

- To set the period: the visual-level or page-level filter is DimDate[Year] = 2026 and DimDate[Month Number] ≤ 6 (Advanced filtering, "is less than or equal to").
- Common gotcha: PO Cost of £2,104,847.74 means the period filter is missing.

### 🙋 Student question Q21 – Highest total project cost
Build the model and the measure Total Project Cost = Labour Cost + PO Cost. With the date filter at Jan–Jun 2026 (June refreshed), which project has the highest Total Project Cost and what is it? (Answer: project code and amount.)

<details open><summary><b>Solution Q21</b></summary>

- Steps: 1. Add a table visual. 2. Put DimProject[Project Code] in the rows. 3. Put Total Project Cost in the values. 4. Add the filter DimDate[Year] = 2026 and DimDate[Month Number] ≤ 6. 5. Sort by Total Project Cost, descending.
- Check the per-project list: CL-P101 317,332.65 (Labour 117,182.60 + PO 200,150.05); CL-P102 297,733.28; CL-P107 190,414.35; CL-P104 120,135.09; CL-P105 100,808.86; CL-P103 99,247.94; CL-P106 100,401.69; CL-P108 53,921.15.
- **Answer: CL-P101, GBP 317,332.65** (Labour 117,182.60 + PO 200,150.05)
- ⚠️ If a student gets a PO Cost of £2,104,847.74 → they have not applied the 2026 filter. The 2025 orders are included.
</details>

### 🙋 Student question Q22 – Variance % for CL-P104
What is the Variance % for CL-P104 for Jan–Jun 2026, where Variance = Budget − Total Project Cost and Variance % = Variance / Budget? (Give as a percentage, 2dp.)

<details open><summary><b>Solution Q22</b></summary>

- Steps: 1. Add a card visual. 2. Put Variance % in it. 3. Set the Project Code slicer to CL-P104. 4. Keep the Jan–Jun filter.
- Working: Budget 284,408.00; Total Project Cost 120,135.09; Variance 164,272.91; 164,272.91 ÷ 284,408.00 = 57.76%. Positive means under budget.
- **Answer: 57.76%** (under budget)
- ⚠️ If a student gets 0.58 or 0.5776 → the measure is not formatted as Percentage (see the troubleshooting table). If they get 32.4% (0.324) → that is the all-projects figure; they did not select CL-P104.
</details>

- Stock Value (Step 15 list): its checkpoint is £293,947.12. It has no date filter. That is correct: it is a snapshot.
- YTD Total Cost: use the `TOTALYTD` form above, or the form that is easier to extend: `YTD Total Cost = CALCULATE ( [Total Project Cost], DATESYTD ( DimDate[Date] ) )`. It needs DimDate marked as a date table (Step 13).
- 🧹 Mess this fixes: none new.

### 🙋 Student question Q25 – YTD Total Cost at May and June (asked here because it needs the measure just built)
YTD Total Cost at the end of May 2026: what does the measure YTD Total Cost = TOTALYTD([Total Project Cost], DimDate[Date]) show for all projects when the visual is at May 2026? What is it for June (the next month)?

<details open><summary><b>Solution Q25</b></summary>

- Steps: 1. Add a line or matrix visual. 2. Put DimDate[Month] on the axis (sorted by Month Number). 3. Put YTD Total Cost in the values. 4. Read the May row, then the June row.
- Working for May: labour 568,743.75 + PO 515,973.50.
- June month alone = 195,277.77 (derived: 1,279,995.02 − 1,084,717.25).
- **Answer: May YTD GBP 1,084,717.25** (Labour 568,743.75 + PO 515,973.50). **June YTD GBP 1,279,995.02.**
- ⚠️ If a student sees 1,279,995.02 on the May row → the date axis is not sorted, or the visual is filtered to June. Check Month Number ≤ 6 and the axis.
</details>

**Step 16 – Build the one-page report (16:30–16:50)**
- Click path: Report view. View ▸ Page view ▸ Fit to page (16:9). Rename the page "Project Cost" (double-click the tab).
- Wireframe:
```text
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
+---------------------------------------+------------------------------------+------+
| Project table / matrix                | Stock Value by Price Source        |SLICER|
| Project | Cost | Budget | Var | Var%  |  (donut: Xero | Log | Wasp)        |Dept  |
| (conditional formatting on Var / Var%)|                                    |      |
+---------------------------------------+------------------------------------+------+
```
1. **KPI cards:** Visualizations ▸ Card. Drag Total Project Cost, Budget, Variance, Variance %, Stock Value, one per card. Format ▸ Callout value ▸ Display units: None.
2. **Clustered bar:** Y-axis = DimProject[Project Name]; X-axis = Total Project Cost and Budget. Sort descending by Total Project Cost.
3. **Line chart:** X-axis = DimDate[Month] (sorted by Month Number); values = Total Project Cost and Budget. Visual-level filter: Year = 2026.
4. **Table:** rows = DimProject[Project Code] and Project Name; values = Total Project Cost, Budget, Variance, Variance %. Conditional formatting on Variance: font colour, red if < 0, green if ≥ 0.

**Q23 goes here, using the table you just built.**

### 🙋 Student question Q23 – Largest overspend
Jan–Jun 2026: which project has the largest overspend (most negative Variance) and by how much? Which other project is also over budget?

<details open><summary><b>Solution Q23</b></summary>

- Steps: 1. In the table visual, sort by Variance, ascending. 2. The top row is the largest overspend. 3. Look for other rows with a negative Variance.
- Working: CL-P107 cost 190,414.35 against budget 163,384.00 = over by 27,030.35. CL-P101 cost 317,332.65 against budget 315,566.00 = over by 1,766.65. Only these two projects are over budget in H1.
- **Answer: CL-P107, overspent by GBP 27,030.35** (cost 190,414.35 vs budget 163,384.00). Also over budget: **CL-P101** by GBP 1,766.65.
- ⚠️ If a student names CL-P101 as the largest overspend → they have sorted by total cost, not by variance. CL-P101 has the highest cost (Q21), but it is only £1,766.65 over budget.
</details>

5. **Donut:** Legend = FactStockValuation[Price Source]; Values = Stock Value. Turn on data labels with value and %. Turn off the date slicer's effect on it: Format ▸ Edit interactions. The stock value is a snapshot.

**Q24 goes here, using the donut.**

### 🙋 Student question Q24 – Stock value by price source
Stock Value by Price Source (donut/bar chart on FactStockValuation). Give the value for each of Xero, Purchase Log and Wasp, and show that they add up to the Lab 3 total valuation (check_totals_B Q11).

<details open><summary><b>Solution Q24</b></summary>

- Steps: 1. Use the donut from step 5. 2. Read each slice's label. 3. Add the three values and compare with the Stock Value card.
- Working: 180,256.41 + 59,526.00 + 54,164.71 = 293,947.12.
- **Answer: Xero 180,256.41; Purchase Log 59,526.00; Wasp 54,164.71. Total 293,947.12 = Lab 3 Q11.**
- ⚠️ If the total is 303,342.30 → cancelled Purchase Log lines were kept (Lab 3 trap). If the total is 291,781.48 → Xero credit notes were kept (Lab 3 trap).
</details>

6. **Slicers:** (a) DimProject[Project Code] (dropdown or list); (b) DimEmployee[Department] (list); (c) DimDate[Month] (month level, not day level, because budget sits on the 1st of each month). Add a page-level filter Year = 2026 and Month Number ≤ 6 for the default view.
7. **Clean up:** Align and distribute the visuals. Lock objects (View ▸ Lock objects) before handing the page over.

**Q26 goes here, after the slicers are in.**

### 🙋 Student question Q26 – Why does the Budget card ignore the employee slicer?
A student adds a slicer on DimEmployee[Employee Name] and a card showing Budget (= SUM(FactBudget[Budget GBP])). Whatever employee they click, the card shows the same number. Why, and what are the choices to fix it?

<details open><summary><b>Solution Q26</b></summary>

- Steps: 1. Add the DimEmployee slicer. 2. Add a card with Budget. 3. Click different employees. The card does not change. 4. In Model view, check that DimEmployee has no line to FactBudget.
- **Answer:** there is no relationship path from DimEmployee to FactBudget. Budget is held by project and month, not by employee. A slicer only filters tables it is related to (DimEmployee to FactLabour), so FactBudget never gets the filter and the SUM ignores it.
- Choices: (1) accept it: "budget has no employee grain". It is the correct answer. Better to blank the card when an employee is picked, with `Budget (safe) = IF ( ISFILTERED ( DimEmployee ), BLANK (), [Budget] )`. (2) Allocate the budget to employees with a rule the business agrees (for example, hours share). Do not invent a rule. (3) Do NOT fix it with a bi-directional or many-to-many relationship. It hides the problem and gives wrong totals.
- The same logic explains why PO Cost does not react to a Department slicer: FactPurchases has no employee or department column.
- ⚠️ If a student turns on bi-directional filtering to make the card move → the totals become wrong. Undo it and use one of the three choices above.
</details>

**Teaching moment:** select Department = FINANCE. Labour Cost moves, PO Cost and Budget do not. That is correct: FactPurchases and FactBudget have no link to DimEmployee. Use this to lead into Q26.

**Step 17 – Refresh, publish and schedule (16:50–16:55)**
- 🗣 Say: "Refresh is the whole point of the morning. Nothing is copied or pasted. Each refresh runs the Lab M again on the files in the folder."
- Refresh demo (2 min): Home ▸ Refresh. Read the counts with the Q-checks: FactLabour 2,320; Labour Hours 19,457.5; Labour Cost £671,900.00; YTD Total Cost at June end £1,279,995.02 (Q25).
- Publish: sign in (top right) with the work account. Home ▸ Publish ▸ choose a workspace (for example "Finance – Project Cost", not "My workspace" if others must see it). Open the link. The report and the semantic model appear in the workspace.
- Sharing: the Share button needs a Pro, Premium Per User or Fabric licence for other people to open it. **Never use Publish to web for company data.**
- **Scheduled refresh: the gateway question.** The Service cannot reach `C:\CallenLenz_Course\data`. Choose one:
  - **On-premises data gateway**: installed on an always-on PC or server that can see the folder. Set it on the semantic model (Settings ▸ Gateway connection ▸ Windows credentials).
  - **SharePoint or OneDrive for Business (recommended)**: move the data to a library and switch the source to the SharePoint folder connector, for example `SharePoint.Files("https://tenant.sharepoint.com/sites/Finance", [ApiVersion = 15])`. No gateway. The owners drop the monthly timesheet into the library.
- Then: workspace ▸ semantic model ▸ Settings ▸ Scheduled refresh ▸ On ▸ times ▸ Apply. Turn on failure notifications.
- When you move the data, only the FolderPath parameter changes (and the Folder.Files call). That is why we parameterised.
- ✅ Checkpoint: the report opens in the browser from the workspace link.

**Step 18 – File and link tips (read this aloud, then see Appendix B)**
- Save as a .pbix on a local drive first, not a network share. Keep a version suffix (`_v1`, `_v2`).
- Do not rename or move the data folders. If you must, edit FolderPath once and refresh.
- Do not email the .pbix. It holds payroll-derived numbers. Share it through the Service.
- Back up the template (File ▸ Export ▸ Power BI template) before you experiment.
- Full checklist: Appendix B.

### Close the day (16:55–17:00)

Five minutes. Say the recap aloud. Ask one question per row.

| Client request | Where it was answered today |
|---|---|
| R1 Consolidating data from multiple sources | Lab 2 (folder of timesheets, reference lookups), Lab 3 (three sources into one valuation) |
| R2 Best practice in structuring data | Block 0; Lab 2 (one shape for every month) |
| R3 Control checks | Lab 2 control totals (Q7, Q8) |
| R4 Many isolated files or fewer, larger models? | Power BI Step 3: one star-schema model per subject area, measures defined once |
| R5 How Power Query feeds Power BI | Power BI Step 4: the M is replayed at refresh and the model holds the result |
| R6 Version control without losing the rules | Lab 5; Appendix B (one master, Period parameter, values-only snapshots) |
| R7 Fuzzy matching payer names | Lab 4 (Table.FuzzyNestedJoin and the Fuzzy_Transform_Table) |
| R8 Roll forward each month | Lab 5 (Period parameter) |
| R9 Split by person and mail merge | Lab 6 (Merge, split per cardholder, Power Automate demo) |
| R10 Valuation from 4 reports | Lab 3 (Wasp, Xero, Purchase Log); Q10 to Q12 |
| R11 How PQ, Power Pivot and Power BI work together | Block 0 (Q1); Power BI Step 1 |
| R12 What are data cubes in Power BI | Power BI Step 2: a semantic model is the cube |
| R13 Saving files and broken links | Lab 5; Appendix B |
| R14 What is the M language | Block 0; Appendix A (M cheat sheet) |

- Suggested next steps for Callen-Lenz:
  1. Dry-run the whole pack once in Excel 365 (Windows) and Power BI Desktop before the course. The M was checked by review and the figures by independent Python recomputation, but the M and DAX have not been run in Excel or Power BI.
  2. Decide where the data lives (SharePoint or OneDrive library, not a PC folder) so scheduled refresh works without a gateway.
  3. Name the owner of the one shared Project Cost model, and certify it once it is published.
  4. Decide the rule for budget by employee, if they want it (Q26). Do not invent one.
  5. Move cards (Lab 6) into a separate model with restricted access (Q2).
  6. Agree a monthly routine: copy the timesheet CSV in, Refresh, check the Q-checks.

### Common errors and fixes

| What they see | Why | Fix |
|---|---|---|
| PO Cost £2,104,847.74 | 2025 orders are in the Purchase Log and no period filter is set | Add DimDate[Year] = 2026 and Month Number ≤ 6 (Step 9, Step 13) |
| Same value for every project in a table | Missing or inactive relationship from DimProject to that fact | Check Model view: all nine relationships present and active |
| Totals look too high (about 2 or 3 times) | A many-to-many or duplicate key in a dimension, or wrong direction | Remove the link; check DimProject and DimPart for duplicate keys; keep single direction |
| Budget card ignores the employee slicer | No relationship from DimEmployee to FactBudget (correct) | Q26: accept it, or blank with ISFILTERED. Never bi-directional |
| Date slicer shows 2025 dates | DimDate range wider than the data | Add the page filter Year = 2026 |
| Months sort alphabetically (Apr, Aug, ...) | Month is not sorted by Month Number | Column tools ▸ Sort by column ▸ Month Number |
| Budget shows only on the 1st of each month | Day-level date slicer | Use a Month slicer |
| Variance % shows 0.5776 instead of 57.76% | Measure not formatted as a percentage | Measure tools ▸ Format ▸ Percentage |
| TOTALYTD returns blank | DimDate not marked as a date table, or visual not using DimDate[Date] | Mark as date table (Step 13) |
| "Formula.Firewall" | Privacy levels | File ▸ Options and settings ▸ Options ▸ Current File ▸ Privacy ▸ Ignore |
| Refresh fails after a folder move | FolderPath still points to the old folder, or credentials are stale | Edit FolderPath; clear data source permissions; sign in again |
| Stock Value is different from £293,947.12 | Cancelled lines or Xero credit notes were kept in Lab 3 | Re-check Lab 3 (Q11 traps: £303,342.30 or £291,781.48) |

---

## Appendix A — M language cheat sheet

M is the language Power Query writes for you. Every button click adds a step, and each step is one line of M. You can type M in the Advanced Editor. The 20 functions below are the ones used in today's solution M or in this block.

| Function | Plain meaning | Tiny example |
|---|---|---|
| `Text.Trim` | Removes spaces at the start and end. | `Text.Trim("  CL-P101 ")` gives `"CL-P101"` |
| `Text.Clean` | Removes hidden non-printing characters, such as line breaks. | `Text.Clean("Ab#(lf)c")` gives `"Abc"` |
| `Text.Upper` | Makes text upper case. | `Text.Upper("cl-p101")` gives `"CL-P101"` |
| `Text.PadStart` | Adds leading characters up to a set length. | `Text.PadStart("7", 4, "0")` gives `"0007"` (used to restore card last four `0937`) |
| `Text.Select` | Keeps only the characters you list. | `Text.Select("£1,250.00", {"0".."9", "."})` gives `"1250.00"` |
| `Text.BetweenDelimiters` | Returns the text between a start and an end marker. | `Text.BetweenDelimiters("Pump [Drawing 1234]", "[Drawing ", "]")` gives `"1234"` |
| `Date.FromText` | Turns text into a date, using a culture so day and month are read correctly. | `Date.FromText("05/02/2026", [Culture = "en-GB"])` gives 5 February 2026 |
| `Date.EndOfMonth` | Returns the last day of the month for a date. Not used in today's solution M; included for month-end logic. | `Date.EndOfMonth(#date(2026, 2, 10))` gives 28 February 2026 |
| `Number.From` | Converts a value to a number. Use `Number.FromText` with `"en-GB"` when the text has commas. | `Number.FromText("1,250.00", "en-GB")` gives `1250` |
| `if … then … else` | Picks one result or another for each row. | `if [Hours] > 37.5 then "OT" else "Normal"` |
| `try … otherwise` | Tries the expression. If it fails, returns the fallback instead of an error. | `try Number.FromText("N/A") otherwise null` gives `null` |
| `Table.Group` | Groups rows by key columns and adds totals. | `Table.Group(Src, {"Project Code"}, {{"Hours", each List.Sum([Hours]), type number}})` |
| `Table.NestedJoin` | Matches two tables on a key (like a lookup), keeping the match as a nested table. | `Table.NestedJoin(Lines, {"Project Code"}, Projects, {"Project Code"}, "Project", JoinKind.LeftOuter)` |
| `Table.FuzzyNestedJoin` | Matches on similar text, with a similarity threshold. Not available in Mac Excel. | `Table.FuzzyNestedJoin(Receipts, {"Payer Text"}, Customers, {"Customer"}, "Match", JoinKind.LeftOuter, [Threshold = 0.8])` |
| `Table.UnpivotOtherColumns` | Turns wide month columns into rows: one row per project per month. | `Table.UnpivotOtherColumns(Budget, {"Project Code"}, "Month Label", "Budget Text")` |
| `Table.Pivot` | Turns row values into columns, adding a total for each. | `Table.Pivot(Src, List.Distinct(Src[Month]), "Month", "Hours", List.Sum)` |
| `Table.Buffer` | Loads a table into memory so the row order stays fixed while sorting and deduping. | `Table.Buffer(Table.Sort(Src, {{"Date", Order.Descending}}))` |
| `Table.Distinct` | Keeps the first row for each key (removes duplicates). | `Table.Distinct(Src, {"Ref"})` |
| `Folder.Files` | Lists every file in a folder as a table (name, extension, content). | `Folder.Files(FolderPath & "02_Timesheets")` |
| `Excel.Workbook` | Opens an Excel file and returns a sheet, chosen by name. | `Excel.Workbook(File.Contents(path), true){[Item = "Sheet1", Kind = "Sheet"]}[Data]` |

---

## Appendix B — Saving files and broken links checklist (R13)

**Answer to R13 in one line:** a query breaks when the path, a sheet name or a column name it reads changes. It is fixed by one edit to FolderPath or the source settings, then Refresh.

**What breaks a query link**
- Renaming or moving the data folder, or a file in it (FolderPath no longer finds it).
- Renaming a sheet. The M reads sheets by name, for example `Item = "Sheet1"` or `Item = "Budget"`.
- Changing a column header that the M reads, for example "Proj Code" in place of "Project Code" (the March timesheet has this mess).
- Opening the file on another PC where the path is different (for example `C:\Users\name\OneDrive - Callen-Lenz\...`).
- An Excel lock file (`~$...xlsx`) or the source file open in Excel ("file in use").
- Stored credentials going stale, or privacy levels changing (Formula.Firewall).

**How to re-link**
1. The usual fix: Home ▸ Transform data ▸ Manage Parameters ▸ FolderPath. Change the path once. Then Home ▸ Refresh.
2. For one source: in Power Query Editor, Home ▸ Data source settings ▸ Change Source. Or click the gear icon next to the Source step.
3. Credentials: File ▸ Options and settings ▸ Data source settings ▸ Edit permissions ▸ Clear permissions. Refresh and sign in again.
4. Privacy: File ▸ Options and settings ▸ Options ▸ Current File ▸ Privacy ▸ Ignore the Privacy Levels checks (for this file only, for the course).

**OneDrive and SharePoint tips**
- A OneDrive sync path differs on each PC. Use the same library path for everyone, or move the data to SharePoint and use the SharePoint folder connector.
- Scheduled refresh in the Service cannot read a local PC folder. Use SharePoint or OneDrive for Business, or an on-premises data gateway (see the publish step in the Power BI block).
- Keep the owners' drop-in routine simple: one folder, one naming pattern (`Timesheet_YYYY-MM.csv`). The M filters on that pattern and ignores lock files.

**Privacy levels**
- Power Query will not combine sources with different privacy levels unless you set them. Set all sources to Organizational for the company files. Do not set cards or payroll data to Public.

**Version control: keep the rules, not the copies**
- Keep **one master** file with the Power Query rules. Do not save a copy per month with the rules inside each copy.
- Use a **Period parameter** (as in Lab 5 roll-forward). Change the number, refresh, and the same rules give the new month.
- Keep a **values-only monthly snapshot** if you need a record of a closed month. Copy the output as values into a dated file (for example `ProjectCost_2026-05_values.xlsx`). The snapshot does not change when the rules change.
- Save each version with a suffix (`_v1`, `_v2`) and keep the PBIT template for the rules without the data.

**Before you save or share (short checklist)**
- [ ] Saved on a local drive first, then copied to the share.
- [ ] FolderPath is the only path in the queries.
- [ ] No Excel file from the folder is open.
- [ ] Refresh runs without an error.
- [ ] Not emailed: the .pbix holds payroll-derived numbers.

---

## Appendix C — Mess catalogue index

Every deliberate mess in the data, grouped by lab and by type. Counts are the number of mess log entries. Each log entry can cover several cells in its file (the log notes "[N instances in file; first 3 logged]"). Source: `_build/mess_log_A.json` and `_build/mess_log_B.json`, 333 entries in total.

| Lab | Mess type | Entries | Example location (first logged) |
|---|---|---:|---|
| Lab 1 | `part_number_variant` | 5 | `Wasp_Stock_Export.xlsx` (Stock), A7 |
| Lab 1 | `na_value` | 4 | `Wasp_Stock_Export.xlsx` (Stock), J10 |
| Lab 1 | `blank_category_fill_down` | 4 | `Wasp_Stock_Export.xlsx` (Stock), Q5 |
| Lab 1 | `blank_row` | 4 | `Wasp_Stock_Export.xlsx` (Stock), A14:X14 |
| Lab 1 | `duplicate_row` | 4 | `Wasp_Stock_Export.xlsx` (Stock), A33:X33 |
| Lab 1 | `gbp_text` | 3 | `Wasp_Stock_Export.xlsx` (Stock), J11 |
| Lab 1 | `line_break` | 3 | `Wasp_Stock_Export.xlsx` (Stock), C8 |
| Lab 1 | `nbsp` | 3 | `Wasp_Stock_Export.xlsx` (Stock), C18 |
| Lab 1 | `trailing_space` | 3 | `Wasp_Stock_Export.xlsx` (Stock), C9 |
| Lab 1 | `title_row` | 2 | `Wasp_Stock_Export.xlsx` (Stock), A1:X1 |
| Lab 1 | `ref_error` | 2 | `Wasp_Stock_Export.xlsx` (Stock), J17 |
| Lab 1 | `junk_row` | 2 | `Wasp_Stock_Export.xlsx` (Stock), A20 |
| Lab 1 | `subtotal_row` | 1 | `Wasp_Stock_Export.xlsx` (Stock), A75 |
| Lab 1 | `repeated_header` | 1 | `Wasp_Stock_Export.xlsx` (Stock), A31:X31 |
| Lab 2 | `lowercase_project_code` | 6 | `Timesheet_2026-01.csv`, lines 3, 80, 189, 193, 218, 246, 247, 301, ... (15 rows) |
| Lab 2 | `task_case` | 6 | `Timesheet_2026-01.csv`, lines 5, 6, 12, 22, 32, 48, 71, 84, ... (42 rows) |
| Lab 2 | `spaced_project_code` | 6 | `Timesheet_2026-01.csv`, lines 7, 13, 34, 36, 77, 103, 124, 128, ... (16 rows) |
| Lab 2 | `name_extra_spaces` | 6 | `Timesheet_2026-01.csv`, lines 8, 27, 33, 54, 60, 83, 85, 88, ... (34 rows) |
| Lab 2 | `no_hyphen_project_code` | 6 | `Timesheet_2026-01.csv`, lines 21, 112, 120, 127, 152, 176, 205, 221, ... (17 rows) |
| Lab 2 | `leading_space_project_code` | 6 | `Timesheet_2026-01.csv`, lines 33, 47, 83, 113, 135, 139, 157, 171, ... (10 rows) |
| Lab 2 | `mixedcase_project_code` | 6 | `Timesheet_2026-01.csv`, lines 39, 64, 76, 79, 114, 186, 225, 300, ... (15 rows) |
| Lab 2 | `trailing_space_project_code` | 6 | `Timesheet_2026-01.csv`, lines 48, 73, 85, 90, 129, 200, 220, 238, ... (19 rows) |
| Lab 2 | `blank_line` | 6 | `Timesheet_2026-01.csv`, lines 250, 294, 455 (3 rows) |
| Lab 2 | `comma_only_line` | 6 | `Timesheet_2026-01.csv`, line 302 |
| Lab 2 | `total_row` | 6 | `Timesheet_2026-01.csv`, line 456 |
| Lab 2 | `employee_id_case_space` | 5 | `Timesheet_2026-01.csv`, lines 28, 83, 85, 126, 202, 252, 298, 317, ... (12 rows) |
| Lab 2 | `exact_duplicate_row` | 5 | `Timesheet_2026-01.csv`, line 178 |
| Lab 2 | `near_duplicate_row` | 2 | `Timesheet_2026-01.csv`, line 289 |
| Lab 2 | `hours_text_hrs` | 1 | `Timesheet_2026-01.csv`, lines 2, 3, 4, 5, 6, 7, 8, 9, ... (448 rows) |
| Lab 2 | `unknown_project_code` | 1 | `Timesheet_2026-02.csv`, lines 76, 80, 106, 111, 173, 316 (6 rows) |
| Lab 2 | `renamed_column` | 1 | `Timesheet_2026-03.csv`, line 1 |
| Lab 2 | `us_date_text` | 1 | `Timesheet_2026-04.csv`, lines 2, 3, 4, 5, 6, 7, 8, 9, ... (355 rows) |
| Lab 2 | `numeric_employee_id` | 1 | `Timesheet_2026-05.csv`, lines 2, 3, 4, 5, 6, 7, 8, 9, ... (451 rows) |
| Lab 2 | `year_typo_2062` | 1 | `Timesheet_2026-05.csv`, line 148 |
| Lab 2 | `title_rows_above_header` | 1 | `Control_Totals.xlsx` (Control), A1 |
| Lab 3 | `text_date` | 10 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), A65 |
| Lab 3 | `gbp_text` | 8 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), H14 |
| Lab 3 | `part_number_variant` | 8 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), E6 |
| Lab 3 | `supplier_variant` | 8 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), B6 |
| Lab 3 | `case_or_space_variant` | 6 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), D13 |
| Lab 3 | `blank_header_column` | 5 | `Purchase_Log.xlsx` (Purchase Log), H1 |
| Lab 3 | `line_break_header` | 4 | `Purchase_Log.xlsx` (Purchase Log), D1 |
| Lab 3 | `cancelled_line` | 4 | `Purchase_Log.xlsx` (Purchase Log), AQ160 |
| Lab 3 | `project_code_variant` | 4 | `Purchase_Log.xlsx` (Purchase Log), G2 |
| Lab 3 | `title_row` | 3 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), A1:H1 |
| Lab 3 | `trailing_space` | 3 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), B9 |
| Lab 3 | `nbsp` | 3 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), F6 |
| Lab 3 | `near_duplicate_row` | 3 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), A58:H58 |
| Lab 3 | `credit_note` | 3 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), A82:H82 |
| Lab 3 | `trailing_space_header` | 3 | `Purchase_Log.xlsx` (Purchase Log), K1 |
| Lab 3 | `composite_notes` | 3 | `Purchase_Log.xlsx` (Purchase Log), AQ2 |
| Lab 3 | `blank_po_number` | 3 | `Purchase_Log.xlsx` (Purchase Log), B8 |
| Lab 3 | `double_space` | 2 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), B78 |
| Lab 3 | `subtotal_row` | 1 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), A163:H163 |
| Lab 3 | `blank_row` | 1 | `Xero_Payable_Invoice_Detail.xlsx` (Xero), A162:H162 |
| Lab 4 | `trailing_space` | 5 | `Bank_Receipts_Q1.xlsx` (Receipts), C9 |
| Lab 4 | `noise_prefix` | 4 | `Bank_Receipts_Q1.xlsx` (Receipts), C2 |
| Lab 4 | `noise_suffix` | 4 | `Bank_Receipts_Q1.xlsx` (Receipts), C2 |
| Lab 4 | `gbp_text` | 4 | `Bank_Receipts_Q1.xlsx` (Receipts), D2 |
| Lab 4 | `case_variant` | 3 | `Bank_Receipts_Q1.xlsx` (Receipts), C35 |
| Lab 4 | `double_space` | 3 | `Bank_Receipts_Q1.xlsx` (Receipts), C7 |
| Lab 4 | `text_date` | 3 | `Bank_Receipts_Q1.xlsx` (Receipts), A6 |
| Lab 4 | `duplicate_row` | 3 | `Bank_Receipts_Q1.xlsx` (Receipts), A21:D21 |
| Lab 4 | `unmatched_payer` | 2 | `Bank_Receipts_Q1.xlsx` (Receipts), C71 |
| Lab 4 | `bracket_negative` | 1 | `Bank_Receipts_Q1.xlsx` (Receipts), D62 |
| Lab 4 | `refund` | 1 | `Bank_Receipts_Q1.xlsx` (Receipts), A62:D62 |
| Lab 4 | `transform_table` | 1 | `Fuzzy_Transform_Table.xlsx` (Transform), A2:B2 |
| Lab 5 | `subtotal_row_inside_data` | 8 | `Asset_Movements_2026.xlsx` (Movements), A15 |
| Lab 5 | `title_rows_above_header` | 4 | `Project_Budget_2026.xlsx` (Budget), A1 |
| Lab 5 | `blank_column` | 3 | `Project_Budget_2026.xlsx` (Budget), C4 |
| Lab 5 | `currency_symbol_text` | 2 | `Asset_Movements_2026.xlsx` (Movements), E10 |
| Lab 5 | `grand_total_row` | 2 | `Asset_Movements_2026.xlsx` (Movements), A49 |
| Lab 5 | `number_stored_as_text` | 1 | `Project_Budget_2026.xlsx` (Budget), F9 |
| Lab 5 | `month_header_text` | 1 | `Project_Budget_2026.xlsx` (Budget), D4 |
| Lab 5 | `total_column` | 1 | `Project_Budget_2026.xlsx` (Budget), P4 |
| Lab 5 | `total_row` | 1 | `Project_Budget_2026.xlsx` (Budget), A13 |
| Lab 5 | `footnote_row` | 1 | `Project_Budget_2026.xlsx` (Budget), A15 |
| Lab 5 | `lowercase_project_code` | 1 | `Project_Budget_2026.xlsx` (Budget), A8 |
| Lab 5 | `trailing_space_project_code` | 1 | `Project_Budget_2026.xlsx` (Budget), A10 |
| Lab 5 | `merged_cell_blank_class` | 1 | `Asset_Movements_2026.xlsx` (Movements), A6 |
| Lab 5 | `brackets_negative_text` | 1 | `Asset_Movements_2026.xlsx` (Movements), D25 |
| Lab 5 | `blank_month_no_values` | 1 | `Asset_Movements_2026.xlsx` (Movements), C29 |
| Lab 5 | `dash_for_zero` | 1 | `Asset_Movements_2026.xlsx` (Movements), D6 |
| Lab 5 | `blank_vs_zero_mixed` | 1 | `Asset_Movements_2026.xlsx` (Movements), C5 |
| Lab 5 | `month_11_blank_column` | 1 | `Asset_Movements_2026.xlsx` (Roll_Forward_Template), C3 |
| Lab 6 | `name_spelling_mismatch` | 5 | `Employee_Lookup.xlsx` (Employees), A3 |
| Lab 6 | `merchant_extra_spaces` | 3 | `Card_Statement_2026-01.csv`, lines 2, 5, 9, 14, 16, 19, 38, 43, ... (9 rows) |
| Lab 6 | `city_case` | 3 | `Card_Statement_2026-01.csv`, lines 2, 10, 13, 17, 20, 23, 26, 28, ... (15 rows) |
| Lab 6 | `blank_amount_foreign` | 3 | `Card_Statement_2026-01.csv`, line 4 |
| Lab 6 | `currency_description_case` | 3 | `Card_Statement_2026-01.csv`, lines 4, 5, 12, 15, 17, 23, 26, 27, ... (20 rows) |
| Lab 6 | `zero_amount` | 3 | `Card_Statement_2026-01.csv`, lines 6, 11, 81 (3 rows) |
| Lab 6 | `status_case` | 3 | `Card_Statement_2026-01.csv`, lines 9, 10, 15, 20, 51 (5 rows) |
| Lab 6 | `name_on_card_case` | 3 | `Card_Statement_2026-01.csv`, lines 10, 14, 18, 29, 41 (5 rows) |
| Lab 6 | `city_double_space` | 3 | `Card_Statement_2026-01.csv`, lines 22, 50, 62, 64, 69, 71 (6 rows) |
| Lab 6 | `currency_symbol_amount` | 3 | `Card_Statement_2026-01.csv`, lines 34, 54, 60, 68 (4 rows) |
| Lab 6 | `blank_amount_gbp` | 3 | `Card_Statement_2026-01.csv`, lines 37, 62, 76 (3 rows) |
| Lab 6 | `blank_line` | 3 | `Card_Statement_2026-01.csv`, line 40 |
| Lab 6 | `cr_suffix_refund` | 3 | `Card_Statement_2026-01.csv`, line 52 |
| Lab 6 | `pending_status` | 3 | `Card_Statement_2026-01.csv`, lines 83, 84, 85, 86, 87 (5 rows) |
| Lab 6 | `last_four_leading_zero_lost` | 2 | `Card_Statement_2026-01.csv`, lines 4, 12, 13, 18, 20, 27, 30, 36, ... (21 rows) |
| Lab 6 | `duplicate_across_statements` | 2 | `Card_Statement_2026-02.csv`, lines 2, 4, 5, 7 (4 rows) |
| Lab 6 | `near_duplicate_bank_ref` | 2 | `Card_Statement_2026-02.csv`, lines 3, 6 (2 rows) |
| Lab 6 | `title_rows_above_header` | 2 | `Card_Statement_2026-03.csv`, line 1 |
| Lab 6 | `header_case_trailing_space` | 1 | `Card_Statement_2026-02.csv`, line 1 |
| Lab 6 | `text_date_dd_mon_yyyy` | 1 | `Card_Statement_2026-02.csv`, lines 2, 3, 4, 5, 6, 7, 8, 9, ... (86 rows) |
| Lab 6 | `blank_amount_refund` | 1 | `Card_Statement_2026-02.csv`, line 14 |
| Lab 6 | `email_upper_case` | 1 | `Employee_Lookup.xlsx` (Employees), D13 |
| **All** | **107 lab/type groups (88 distinct mess types)** | **333** | |

---

## Appendix D — Answer key at a glance

| Q | Short answer |
|---|---|
| Q1 | (a) Power Query; (b) Power Pivot/DAX; (c) Power BI; (d) Power Query; (e) Model relationship |
| Q2 | One star-schema model per subject area, shared dimensions, measures defined once |
| Q3 | 60 distinct valid parts |
| Q4 | 8 parts "Out of stock" |
| Q5 | £258,431.64 stock value at Wasp item cost |
| Q6 | £28,779.27 labour cost for CL-P103, Jan–Mar |
| Q7 | Feb: 7.5 hours over control (3,006.5 vs 2,999.0); 6 unknown-code rows |
| Q8 | 6 rows with unknown codes; 7.5 hours |
| Q9 | 19,457.5 valid hours, Jan–Jun |
| Q10 | £1,318,886.68 Xero invoice value ex VAT |
| Q11 | £293,947.12 total stock valuation |
| Q12 | (a) £185.66; (b) from Xero, since the latest Xero line is a credit note |
| Q13 | £390,768.42 total Q1 receipts after dedupe and refund |
| Q14 | Flintwood Marine: £42,759.67 net, 13 payments |
| Q15 | Refs 10131 and 10132 unmatched |
| Q16 | £198,551.00 budget for CL-P105, H1 |
| Q17 | 501.5 hours on CL-P102 in Mar-26 |
| Q18 | Cost Cfwd Month 11: £262,200.00 |
| Q19 | £1,043.54 net Feb 2026 spend, Callum Reid |
| Q20 | Aisha Rahman; £2,352.92; 33 transactions |
| Q21 | CL-P101; £317,332.65 total project cost |
| Q22 | 57.76% variance for CL-P104 |
| Q23 | CL-P107 over by £27,030.35; CL-P101 also over (£1,766.65) |
| Q24 | Xero £180,256.41; Purchase Log £59,526.00; Wasp £54,164.71; total £293,947.12 |
| Q25 | May YTD £1,084,717.25; June YTD £1,279,995.02 |
| Q26 | Employee slicer cannot reach FactBudget; budget has no employee grain |
