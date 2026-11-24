# Writer brief: TRAINER_WALKTHROUGH.md

You are writing part of a **trainer walkthrough** for a 1-day in-person course (9:00–17:00): **"Power Query for Finance + Power BI intro"**, for the finance team at **Callen-Lenz Associates Ltd**.
- Callen-Lenz is a UK manufacturer. It uses GBP and VAT 20%.
- Its systems are Timesheets, Wasp (stock), the Purchase Log (POs) and Xero (accounts).
- The audience is beginners. A couple of them have dabbled, and they had one 1-hour webinar on tables, recording and renaming steps, removing blank rows and splitting columns.
- They asked for a **practical session using specific, relevant examples**.

The reader is the trainer (Nitish), who delivers this live. He follows the document top to bottom: what to teach first, then second, then the questions for students, with the **solution directly below each question**. There is NO separate student handout. Students only receive the `data/` folder.

## Source files you MUST use (do not invent numbers)
All paths are relative to the repo root, `/Users/tarunsachdeva/dev/callenlenz-pq-pbi-course`.
- `solutions/check_totals_A.json`, `check_totals_B.json` and `check_totals_C.json` hold every question's text, answer, definition, "how", trap_answers and intermediate checkpoints.
  - **Copy every figure exactly from these files.**
  - The question text students see should be the JSON `question` text. You may lightly tidy the wording, but you must keep every rule it states.
- `_build/mess_log_A.json` and `_build/mess_log_B.json` list every deliberate error in the data: file, location, what it is, and the teaching fix. Each teach step should name the specific mess it fixes, for example "row 37 has a repeated header", "the March file calls it 'Proj Code'".
- `solutions/m/*.m` is the solution M code. Paste the relevant M into solutions. Long M can go in a collapsible `<details><summary>Full M code</summary>` block.
- `solutions/m/lab6_power_automate.md` is the Power Automate demo.
- `solutions/powerbi_build.md` is the Power BI build guide.
- `data/` holds the actual lab files. Use `ls` on it to get exact file and sheet names. If you need to peek at a file, use `_build/.venv/bin/python` with pandas, or `head` for CSVs.

## Format for every block (use exactly this structure)
```
## <Block number>. <Title>  ⏱ <start–end>

> **Who asked:** "<verbatim client request>" (initials)
> **Why it matters to them:** one or two sentences in their world (projects, parts, month-end).

**Files:** `data/...`

### Teach — step by step
**Step 1 – <action>**
- Click path: Data ▸ Get Data ▸ From File ▸ From Workbook ▸ … (exact Excel 365 menu path)
- 🗣 Say: "<short sentence the trainer says, plain words>"
- 🧹 Mess this fixes: <specific item from mess_log, with location>
- ✅ Checkpoint: <intermediate figure from JSON, e.g. "68 rows after removing blanks">

**Step 2 – …**

### 🙋 Student question Qn – <short title>
<question text from JSON, stated fully, including its rules>

<details open><summary><b>Solution Qn</b></summary>

- Steps: numbered clicks
- M (if useful): ```m ... ```
- **Answer: <exact value from JSON>**
- ⚠️ If a student gets <trap value> → they <mistake> (from trap_answers)
</details>

### Common errors and fixes
| What they see | Why | Fix |
```
- Use simple English and short sentences. Use `▸` for menu paths.
- When a step uses an M function, explain the function in one plain line, for example "`Text.PadStart` adds leading zeros up to a set length".
- Questions are numbered Q1–Q26 across the whole day. Keep the numbering from the JSON.

## Client requests (verbatim) and where each is answered
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

The brief also lists the "most useful functions for accountants". Each must be taught hands-on somewhere:
- **Extract:** Lab 1 drawing number; Lab 3 Notes; Lab 4 payer.
- **Conditional Columns:** Lab 1 Stock Status; Lab 3 Price Source.
- **Merge Queries:** Labs 2, 3, 4, 6.
- **Data from Folder:** Lab 2.
- **Unpivot:** Lab 5.

## Day run sheet (the overall order)
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

## Setup
Students get `data/` copied to `C:\CallenLenz_Course\data`. Every query uses a `FolderPath` parameter set to that folder. Teach how to create the parameter in Lab 1, Step 1 (Manage Parameters ▸ New Parameter). Excel 365 for Windows is required, because fuzzy merge and some connectors are limited on Mac.

## Honesty notes to include where relevant
- The M code was checked by review and the figures by independent Python recomputation. It was NOT run in Excel, so the trainer must dry-run once before the day.
- Pure fuzzy matching at 0.8 is approximate. The exact Lab 4 answers depend on using the Fuzzy_Transform_Table.

## Output rules
- Write ONLY your assigned part file in `_build/walkthrough_parts/`. Create the directory if needed.
- Do not edit any other file.
- Do not git commit.
- When done, report: your file path, line count, the Q ids covered, and any number you could not find in the JSON. Never make one up.
