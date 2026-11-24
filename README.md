# Callen-Lenz Power Query + Power BI course pack

Private training pack for the one-day in-person course **"Power Query for Finance + Power BI intro"**, run for the finance team at Callen-Lenz Associates Ltd. The day runs 9:00–17:00. Students work on the sample files in `data/` through six Power Query labs, then build a Project Cost dashboard in Power BI.

This is a trainer's repo. Students receive only the `data/` folder.

## What is in the repo

All names, people, emails, customers, suppliers, part numbers and figures are fictional. Emails use the `@callenlenz-demo.example` domain.

## Folder map

```text
data/                         Lab files students receive (copy to C:\CallenLenz_Course\data)
  00_Reference/              Customers, Employees, Labour_Rates, Projects, Suppliers_Master
  01_Wasp_Stock/             Wasp stock export (Lab 1: messy export, title rows, duplicates)
  02_Timesheets/             Timesheet_2026-01..05.csv, Control_Totals.xlsx
    _Later/                  Drop-in: Timesheet_2026-06.csv (Lab 2 refresh with June)
  03_Valuation/              Purchase_Log, Xero_Payable_Invoice_Detail, Supplier_Name_Map (Lab 3)
  04_Receipts/               Bank_Receipts_Q1, Customer_List, Fuzzy_Transform_Table (Lab 4)
  05_Budget_Assets/          Project_Budget_2026, Asset_Movements_2026 (Lab 5)
    _Later/                  Drop-in: Asset_Movements_2026-11.xlsx (roll-forward month 11)
  06_Cards/                  Card statements, Employee_Lookup, Expense_Template, FX_Rates (Lab 6)

solutions/                    Trainer answers, never given to students
  check_totals_A.json        Q6–Q9, Q16–Q20 (answers, definitions, traps, checkpoints)
  check_totals_B.json        Q3–Q5, Q10–Q15
  check_totals_C.json        Q1–Q2, Q21–Q26
  m/                         Solution M code for each lab (paste into Advanced Editor)
    lab6_power_automate.md   Power Automate demo (trainer-led)
  powerbi_build.md           Click-by-click Power BI build guide (15:00–17:00)

_build/                       Generator, tests, mess logs and the walkthrough writing brief
  master.py                  Master data and clean truth tables (seed 42)
  build_A.py                 Writes data for Labs 2, 5, 6 (timesheets, budget/assets, cards)
  build_B.py                 Writes data for Labs 1, 3, 4 (Wasp, Xero/Purchase Log, receipts)
  build_C.py                 Rewrites solutions/check_totals_C.json from the clean model
  model_C.py                 Power BI star-schema figures used by build_C.py
  mess_log_A.json, mess_log_B.json     Every deliberate error: file, location, fix
  test_master.py, test_A.py, test_B.py, test_C.py   Self-tests (hash and figure checks)
  requirements.txt           Pinned Python packages (numpy, openpyxl, pandas)
  walkthrough_parts/         Parts of the trainer walkthrough (assembled into one document)
```

## Quick start for the trainer

1. Read `TRAINER_WALKTHROUGH.md` from the top. It is the document you teach from, with each block, its questions and the solution directly below each question. Each block lists the client request it answers, the files used and the checkpoint figure for every step.
2. Do the pre-course checklist below.
3. Dry-run the whole day once, in the same Excel and Power BI versions the students will use.

## Pre-course checklist

- [ ] **Excel 365 for Windows** on every student machine. Fuzzy merge (`Table.FuzzyNestedJoin`) is not available in Mac Excel. If a student has a Mac, pair them up.
- [ ] **Power BI Desktop** (free, Windows only) installed for the Power BI block.
- [ ] Copy the `data/` folder to `C:\CallenLenz_Course\data` on each machine. Keep the folder structure exactly as it is.
- [ ] The `FolderPath` parameter is created by the trainer in Lab 1 Step 1 (and again in the Power BI file). Value `C:\CallenLenz_Course\data\` with the trailing backslash.
- [ ] Copy the `_Later` files into their folders when the lab says so: `Timesheet_2026-06.csv` (Lab 2, and the Power BI Jan–Jun checks need it) and `Asset_Movements_2026-11.xlsx` (Lab 5).
- [ ] Test the Microsoft 365 mailbox for Lab 6 and the Power Automate demo. Swap the demo emails `@callenlenz-demo.example` for the trainer's test addresses before the session (see `solutions/m/lab6_power_automate.md`).

## How to regenerate the data

The data is generated from seed 42, so a rerun gives identical files. Run these from the repo root.

```sh
python3 -m venv _build/.venv
_build/.venv/bin/pip install -r _build/requirements.txt
_build/.venv/bin/python _build/master.py
_build/.venv/bin/python _build/build_A.py
_build/.venv/bin/python _build/build_B.py
_build/.venv/bin/python _build/build_C.py
```

Run the self-tests. They check the generated files and the figures in `solutions/`.

```sh
_build/.venv/bin/python _build/test_master.py
_build/.venv/bin/python _build/test_A.py
_build/.venv/bin/python _build/test_B.py
_build/.venv/bin/python _build/test_C.py
```

## Honesty note

- The solution M code was checked by review. It has **not been run** in Excel 365 or Power BI Desktop.
- The figures were recomputed independently in Python from the source data. They match the answers in `solutions/check_totals_*.json`.
- **Do a dry run once before the course.** Paste every M query in order, refresh, and check each checkpoint in the walkthrough.
- Fuzzy matching at a threshold of 0.8 is approximate. The exact Lab 4 answers depend on using `Fuzzy_Transform_Table.xlsx` as the transform table, as the solution does.
- A few Power BI figures in the walkthrough are derived by subtraction or division from the checked totals. They are labelled as derived.
