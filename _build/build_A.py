#!/usr/bin/env python
"""Builder A: Lab 2 (timesheets), Lab 5 (budget / assets), Lab 6 (cards).

Run from repo root:  _build/.venv/bin/python _build/build_A.py

Writes messy lab files, check totals (computed by mimicking the Power Query steps on the
messy files and asserting against clean truth), mess log and M code.
"""
from __future__ import annotations

import csv
import io
import json
import os
import re
import shutil
import sys
import zipfile
from collections import OrderedDict
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import master as M  # noqa: E402

ROOT = M.ROOT
DATA = os.path.join(ROOT, "data")
SOL = os.path.join(ROOT, "solutions")
D_TS = os.path.join(DATA, "02_Timesheets")
D_BA = os.path.join(DATA, "05_Budget_Assets")
D_CD = os.path.join(DATA, "06_Cards")
FIXED_DT = datetime(2026, 7, 1, 9, 0, 0)
MESS_LOG: list[dict] = []
INTER: "OrderedDict[str, dict]" = OrderedDict()
QUESTIONS: list[dict] = []


def rng(k: int) -> np.random.Generator:
    return np.random.default_rng([M.SEED, 7000 + k])


def rel(p: str) -> str:
    return os.path.relpath(p, ROOT).replace(os.sep, "/")


def inter(key: str, value, definition: str):
    INTER[key] = {"value": value, "definition": definition}


def ask(qid, lab, question, answer, unit, definition, how, **extra):
    d = {"id": qid, "lab": lab, "question": question, "answer": answer, "unit": unit,
         "definition": definition, "how": how}
    d.update(extra)
    QUESTIONS.append(d)


# ------------------------------------------------------------------ xlsx helpers
def save_wb(wb: Workbook, path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.properties.created = FIXED_DT
    wb.properties.modified = FIXED_DT
    wb.properties.creator = "Callen-Lenz Finance"
    wb.properties.lastModifiedBy = "Callen-Lenz Finance"
    wb.save(path)
    stabilise_zip(path)


def stabilise_zip(path: str):
    """Rewrite zip with fixed timestamps so xlsx bytes are reproducible."""
    with zipfile.ZipFile(path) as zin:
        items = [(i.filename, zin.read(i.filename)) for i in zin.infolist()]
    items = [(n, re.sub(rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2026-07-01T09:00:00Z", c) if n == "docProps/core.xml" else c) for n, c in items]
    items.sort(key=lambda x: x[0])
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in items:
            zi = zipfile.ZipInfo(name, date_time=(2026, 7, 1, 9, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            zout.writestr(zi, content)
    os.replace(tmp, path)


def xl_log(path, sheet, cell, mess, desc, fix, lab, evidence):
    MESS_LOG.append({"file": rel(path), "sheet": sheet, "location": cell, "mess_type": mess,
                     "description": desc, "teaching_fix": fix, "lab": lab, "evidence": evidence})


# ------------------------------------------------------------------ csv helpers
def _csv_line(cells) -> str:
    if cells is None:
        return ""
    buf = io.StringIO()
    csv.writer(buf, lineterminator="").writerow(cells)
    return buf.getvalue()


def write_csv(path: str, lines: list[dict], lab: str):
    """lines: [{'cells': [...] | None, 'tags': [(mess_type, desc, fix), ...]}]. Logs consolidated mess."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    agg: "OrderedDict[str, dict]" = OrderedDict()
    texts = []
    for i, ln in enumerate(lines, start=1):
        txt = _csv_line(ln["cells"])
        texts.append(txt)
        for mt, desc, fix in ln.get("tags", []):
            a = agg.setdefault(mt, {"lines": [], "ev": txt, "desc": desc, "fix": fix})
            a["lines"].append(i)
    with open(path, "w", newline="", encoding="utf-8") as f:
        for t in texts:
            f.write(t + "\r\n")
    for mt, a in agg.items():
        ls = a["lines"]
        shown = ", ".join(str(x) for x in ls[:8]) + (", ..." if len(ls) > 8 else "")
        loc = ("line " if len(ls) == 1 else "lines ") + shown + (f" ({len(ls)} rows)" if len(ls) > 1 else "")
        MESS_LOG.append({"file": rel(path), "sheet": "", "location": loc, "mess_type": mt,
                         "description": a["desc"], "teaching_fix": a["fix"], "lab": lab, "evidence": a["ev"]})


# ====================================================================================
# LAB 2 - TIMESHEETS
# ====================================================================================
TS_COLS = ["Week Ending", "Employee ID", "Employee Name", "Project Code", "Task", "Hours"]
TS = M.timesheets()
TS["m"] = [d.month for d in TS["Week Ending"]]
EMP = M.employees()
EMP_GRADE = EMP.set_index("Employee ID")["Grade"].to_dict()
EMP_NAME = EMP.set_index("Employee ID")["Employee Name"].to_dict()
RATES = M.labour_rates().set_index("Grade")
PROJ = M.projects()
VALID_CODES = set(PROJ["Project Code"])

T = {
    "renamed_column": ("Column header is 'Proj Code' instead of 'Project Code' in the March file (breaks the combine).",
                       "Rename the column to 'Project Code' inside the Transform Sample File function (MissingField.Ignore) so all files align."),
    "us_date_text": ("Week Ending stored as US month/day/year text (April file only); UK files use day/month/year.",
                     "Change Type > Using Locale > Date, English (United States) for April only (in the file function, choose culture by file)."),
    "lowercase_project_code": ("Project code in lower case (cl-p101).", "Text.Upper then standardise (fnProjectCode)."),
    "spaced_project_code": ("Project code with a space instead of hyphen (CL P101).", "Remove spaces/hyphens, rebuild as CL-Pnnn."),
    "no_hyphen_project_code": ("Project code with no hyphen (CLP101).", "Remove spaces/hyphens, rebuild as CL-Pnnn."),
    "trailing_space_project_code": ("Project code with trailing space.", "Transform > Format > Trim."),
    "leading_space_project_code": ("Project code with leading space.", "Transform > Format > Trim."),
    "mixedcase_project_code": ("Project code in mixed case (Cl-P101).", "Text.Upper then standardise."),
    "unknown_project_code": ("Booking against a project code that is not in Projects.xlsx (CL-P199 / CL-P100).",
                             "Merge with Projects using Left Anti to list them; Inner join (or filter) to remove them."),
    "numeric_employee_id": ("Employee ID is a bare number (7) instead of text E007 (May file).",
                            "Add prefix 'E' and pad to 3 digits (Text.PadStart) so it matches Employees.xlsx."),
    "employee_id_case_space": ("Employee ID in lower case or with a trailing space.", "Trim + Upper."),
    "hours_text_hrs": ("Hours stored as text with ' hrs' suffix (e.g. '7.5 hrs', January file).",
                       "Text.Select keeping digits and '.', then change type to Decimal Number."),
    "exact_duplicate_row": ("An entire booking line appears twice (identical text).",
                            "Home > Remove Rows > Remove Duplicates (after cleaning)."),
    "near_duplicate_row": ("Same booking appears twice but one copy has a lower-case code and trailing space in the name.",
                           "Clean (trim / upper) first, THEN Remove Duplicates so the near-duplicate becomes an exact one."),
    "total_row": ("Footer row labelled 'Total' with a hours sum at the bottom of the file.",
                  "Filter out rows where Week Ending = 'Total'."),
    "blank_line": ("Completely blank line inside the data.", "Remove Rows > Remove Blank Rows, or filter Week Ending <> ''."),
    "comma_only_line": ("Line containing only commas (empty row).", "Remove Blank Rows / filter empty Week Ending."),
    "year_typo_2062": ("Week Ending typed with year 2062 instead of 2026.", "Replace year 2062 with 2026 (Date.Year test + #date)."),
    "name_extra_spaces": ("Employee Name with trailing or double spaces.", "Transform > Format > Trim, collapse double spaces (or take the name from Employees)."),
    "task_case": ("Task in lower or upper case.", "Transform > Format > Capitalize Each Word."),
}


def tag(k):
    return (k, *T[k])


def code_variant(code: str, r) -> tuple[str, str]:
    opts = [("lowercase_project_code", code.lower()), ("spaced_project_code", code.replace("-", " ")),
            ("no_hyphen_project_code", code.replace("-", "")), ("trailing_space_project_code", code + " "),
            ("leading_space_project_code", " " + code), ("mixedcase_project_code", "Cl-" + code[3:])]
    return opts[int(r.integers(0, len(opts)))]


def render_ts_row(row, cfg, r, plain=False):
    tags = []
    we = row["Week Ending"]
    if row.get("we_override"):
        wes = row["we_override"]
        tags.append(tag("year_typo_2062"))
    elif cfg.get("us"):
        wes = we.strftime("%m/%d/%Y")
        tags.append(tag("us_date_text"))
    else:
        wes = we.strftime("%d/%m/%Y")
    eid = row["Employee ID"]
    if cfg.get("numeric_id"):
        eid = str(int(eid[1:]))
        tags.append(tag("numeric_employee_id"))
    elif not plain and r.random() < 0.03:
        eid = eid.lower() if r.random() < 0.5 else eid + " "
        tags.append(tag("employee_id_case_space"))
    name = row["Employee Name"]
    if not plain and r.random() < 0.07:
        name = name + " " if r.random() < 0.5 else name.replace(" ", "  ", 1)
        tags.append(tag("name_extra_spaces"))
    code = row["Project Code"]
    if not plain and r.random() < 0.22:
        mt, code = code_variant(code, r)
        tags.append(tag(mt))
    task = row["Task"]
    if not plain and r.random() < 0.10:
        task = task.lower() if r.random() < 0.5 else task.upper()
        tags.append(tag("task_case"))
    hrs = f"{row['Hours']:.1f}"
    if cfg.get("hrs"):
        hrs += " hrs"
        tags.append(tag("hours_text_hrs"))
    return {"cells": [wes, eid, name, code, task, hrs], "tags": tags}


def build_timesheet_file(month: int, path: str, cfg: dict, k: int):
    r = rng(10 + k)
    df = TS[TS["m"] == month].reset_index(drop=True)
    base = [df.iloc[i].to_dict() for i in range(len(df))]
    # year typo (May): pick one row
    if cfg.get("year_typo"):
        i = int(r.integers(5, len(base) - 5))
        base[i]["we_override"] = base[i]["Week Ending"].strftime("%d/%m/") + "2062"
    lines = [render_ts_row(b, cfg, r) for b in base]
    # force canonical rendering for near-dup source later; pick sources
    inserts = []
    n = len(lines)
    if cfg.get("exact_dup"):
        src = int(r.integers(10, n - 10))
        copy = {"cells": list(lines[src]["cells"]), "tags": [tag("exact_duplicate_row")]}
        inserts.append((src + int(r.integers(1, 30)), copy))
    if cfg.get("near_dup"):
        cand = [i for i in range(n) if all(t[0] in ("us_date_text", "hours_text_hrs", "numeric_employee_id") for t in lines[i]["tags"])]
        src = cand[int(r.integers(5, len(cand) - 5))]
        b = base[src]
        cells = list(lines[src]["cells"])
        cells[2] = b["Employee Name"] + " "
        cells[3] = b["Project Code"].lower()
        inserts.append((src + int(r.integers(2, 40)), {"cells": cells, "tags": [tag("near_duplicate_row")]}))
    if cfg.get("unknown"):
        fridays = sorted(set(df["Week Ending"]))
        hours = [1.5, 1.0, 2.0, 1.0, 1.0, 1.0]
        codes = ["CL-P199", "CL-P199", "CL-P100", "CL-P199", "CL-P100", "CL-P199"]
        bookers = sorted(set(df["Employee ID"]))
        for j in range(6):
            e = bookers[int(r.integers(0, len(bookers)))]
            wk = fridays[int(r.integers(0, len(fridays)))]
            row = {"Week Ending": wk, "Employee ID": e, "Employee Name": EMP_NAME[e], "Project Code": codes[j],
                   "Task": "Admin", "Hours": hours[j]}
            ln = render_ts_row(row, cfg, r, plain=True)
            if j == 4:
                ln["cells"][3] = "cl-p100"
            ln["tags"].append(tag("unknown_project_code"))
            inserts.append((int(r.integers(5, n)), ln))
    for pos, ln in sorted(inserts, key=lambda x: -x[0]):
        lines.insert(min(pos, len(lines)), ln)
    # blank lines
    for ln in ({"cells": None, "tags": [tag("blank_line")]}, {"cells": None, "tags": [tag("blank_line")]},
               {"cells": [""] * 6, "tags": [tag("comma_only_line")]}):
        lines.insert(int(r.integers(5, len(lines) - 5)), ln)
    total = sum(float(re.sub(r"[^0-9.]", "", ln["cells"][5])) for ln in lines if ln["cells"] and ln["cells"][0])
    lines.append({"cells": None, "tags": [tag("blank_line")]})
    lines.append({"cells": ["Total", "", "", "", "", f"{total:.1f}"], "tags": [tag("total_row")]})
    hdr = list(TS_COLS)
    htags = []
    if cfg.get("rename"):
        hdr[3] = "Proj Code"
        htags.append(tag("renamed_column"))
    lines.insert(0, {"cells": hdr, "tags": htags})
    write_csv(path, lines, "Lab 2")
    return total


def build_all_timesheets():
    cfgs = {
        1: dict(hrs=True, exact_dup=True, near_dup=True),
        2: dict(exact_dup=True, unknown=True),
        3: dict(rename=True, exact_dup=True),
        4: dict(us=True, exact_dup=True),
        5: dict(numeric_id=True, year_typo=True, exact_dup=True),
    }
    totals = {}
    for m, cfg in cfgs.items():
        totals[m] = build_timesheet_file(m, os.path.join(D_TS, f"Timesheet_2026-{m:02d}.csv"), cfg, m)
    totals[6] = build_timesheet_file(6, os.path.join(D_TS, "_Later", "Timesheet_2026-06.csv"), dict(near_dup=True), 6)
    return totals


# ---- mimic of the Power Query steps
def fn_code(x: str) -> str:
    s = re.sub(r"[^A-Z0-9]", "", x.upper())
    return "CL-P" + s[-3:] if len(s) >= 4 else s


def fn_eid(x: str) -> str:
    u = x.upper()
    if u == "":
        return ""
    return u if u.startswith("E") else "E" + u.zfill(3)


def pq_trim(x: str) -> str:
    return x.replace("\xa0", " ").strip()


def clean_timesheet_file(path: str) -> pd.DataFrame:
    name = os.path.basename(path)
    with open(path, newline="", encoding="utf-8") as f:
        rows = [r for r in csv.reader(f) if r]  # Power Query drops wholly empty lines
    hdr = [pq_trim(h) for h in rows[0]]
    hdr = ["Project Code" if h == "Proj Code" else h for h in hdr]
    n_raw = len(rows) - 1
    df = pd.DataFrame(rows[1:], columns=hdr)[TS_COLS]
    df = df.apply(lambda c: c.map(pq_trim))
    df = df[(df["Week Ending"] != "") & (df["Week Ending"].str.lower() != "total")].copy()
    n_after_removal = len(df)
    fmt = "%m/%d/%Y" if "2026-04" in name else "%d/%m/%Y"
    df["Week Ending"] = [datetime.strptime(x, fmt).date() for x in df["Week Ending"]]
    df["Employee ID"] = df["Employee ID"].map(fn_eid)
    df["Employee Name"] = df["Employee Name"].map(lambda s: " ".join(x for x in s.split(" ") if x))
    df["Project Code"] = df["Project Code"].map(fn_code)
    df["Task"] = df["Task"].str.title()
    df["Hours"] = df["Hours"].map(lambda s: float(re.sub(r"[^0-9.]", "", s)))
    df.attrs["raw_data_lines"] = n_raw
    df.attrs["after_removal"] = n_after_removal
    return df


def fix_year(d: date) -> date:
    return date(2026, d.month, d.day) if d.year == 2062 else d


def labour_lines(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["Grade"] = d["Employee ID"].map(EMP_GRADE)
    d["Rate"] = d["Grade"].map(RATES["Standard Hourly Rate GBP"])
    d["Mult"] = d["Grade"].map(RATES["Overtime Multiplier"])
    d["Weekly Hours"] = d.groupby(["Employee ID", "Week Ending"])["Hours"].transform("sum")
    d["OT Hours"] = np.where(d["Weekly Hours"] > 37.5, (d["Weekly Hours"] - 37.5) * d["Hours"] / d["Weekly Hours"], 0.0)
    d["Labour Cost"] = d["Hours"] * d["Rate"] + d["OT Hours"] * d["Rate"] * (d["Mult"] - 1)
    d["Month"] = [date(x.year, x.month, 1) for x in d["Week Ending"]]
    return d


def load_timesheets(include_june: bool):
    paths = [os.path.join(D_TS, f"Timesheet_2026-{m:02d}.csv") for m in range(1, 6)]
    if include_june:
        paths.append(os.path.join(D_TS, "_Later", "Timesheet_2026-06.csv"))
    parts = [clean_timesheet_file(p) for p in paths]
    raw_lines = sum(p.attrs["raw_data_lines"] for p in parts)
    after_removal = sum(p.attrs["after_removal"] for p in parts)
    comb = pd.concat(parts, ignore_index=True)
    comb["Week Ending"] = comb["Week Ending"].map(fix_year)
    ded = comb.drop_duplicates().reset_index(drop=True)
    valid = ded[ded["Project Code"].isin(VALID_CODES)].reset_index(drop=True)
    unknown = ded[~ded["Project Code"].isin(VALID_CODES)].reset_index(drop=True)
    return dict(combined=comb, deduped=ded, valid=valid, unknown=unknown, raw_lines=raw_lines, after_removal=after_removal)


def write_control_totals():
    truth = TS.copy()
    truth["Month"] = [date(x.year, x.month, 1) for x in truth["Week Ending"]]
    g = truth.groupby("Month")["Hours"].agg(["sum", "count"])
    wb = Workbook()
    ws = wb.active
    ws.title = "Control"
    ws["A1"] = "Timesheet system - control totals (hours booked to valid projects, after system de-duplication)"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "Source: Timesheet system report TS-110, run 03/07/2026. Month = month of Week Ending."
    for c, h in enumerate(["Month", "System Hours", "System Lines"], start=1):
        ws.cell(4, c, h).font = Font(bold=True)
    for i, (mo, row) in enumerate(g.iterrows(), start=5):
        ws.cell(i, 1, mo).number_format = "mmm-yy"
        ws.cell(i, 2, float(row["sum"])).number_format = "#,##0.0"
        ws.cell(i, 3, int(row["count"]))
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14
    p = os.path.join(D_TS, "Control_Totals.xlsx")
    save_wb(wb, p)
    xl_log(p, "Control", "A1", "title_rows_above_header", "Two title/source rows and a blank row sit above the header (header on row 4).",
           "Remove Top Rows (3) then Use First Row as Headers.", "Lab 2", "Timesheet system - control totals (hours booked to valid projects, after system de-duplication)")
    return g


# ====================================================================================
# LAB 5 - BUDGET and ASSETS
# ====================================================================================
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def build_budget():
    bud = M.budget()
    bud["mi"] = [d.month for d in bud["Month"]]
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget"
    p = os.path.join(D_BA, "Project_Budget_2026.xlsx")
    ws["A1"] = "Callen-Lenz Associates Limited - Project Budget 2026"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "GBP, ex VAT. Approved by SLT 12/12/2025 - v3 (FP&A)"
    hdr = ["Project Code", "Project Name", None] + [f"{m}-26" for m in MONTHS] + ["Total"]
    for c, h in enumerate(hdr, start=1):
        cell = ws.cell(4, c, h)
        cell.font = Font(bold=True)
        if h and c >= 4 and h != "Total":
            cell.data_type = "s"
    projs = list(M.projects()["Project Code"])
    names = M.projects().set_index("Project Code")["Project Name"].to_dict()
    colsum = [0.0] * 12
    for i, code in enumerate(projs):
        r = 5 + i
        shown = code
        if code == "CL-P104":
            shown = "cl-p104"
        if code == "CL-P106":
            shown = "CL-P106 "
        ws.cell(r, 1, shown)
        ws.cell(r, 2, names[code])
        vals = [float(bud[(bud["Project Code"] == code) & (bud["mi"] == m)]["Budget GBP"].iloc[0]) for m in range(1, 13)]
        for m in range(12):
            colsum[m] += vals[m]
            c = ws.cell(r, 4 + m, vals[m])
            c.number_format = "#,##0.00"
            if code == "CL-P105" and m == 2:
                c.value = f"{vals[m]:,.2f}"
                c.data_type = "s"
                xl_log(p, "Budget", c.coordinate, "number_stored_as_text", "A budget cell holds the number as text ('45,000.00' style) so SUM/Unpivot see text.",
                       "Replace commas and Change Type to Decimal Number (or Number.FromText) after unpivot.", "Lab 5", c.value)
        ws.cell(r, 16, sum(vals)).number_format = "#,##0.00"
    tr = 5 + len(projs)
    ws.cell(tr, 1, "Total").font = Font(bold=True)
    for m in range(12):
        ws.cell(tr, 4 + m, colsum[m]).number_format = "#,##0.00"
    ws.cell(tr, 16, sum(colsum)).number_format = "#,##0.00"
    ws.cell(tr + 2, 1, "* CL-P108 closed 31 Mar 2026 - budget retained for reporting only")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 32
    for c in range(4, 17):
        ws.column_dimensions[ws.cell(4, c).column_letter].width = 12
    xl_log(p, "Budget", "A1", "title_rows_above_header", "Title and subtitle rows above the header (header on row 4).", "Remove Top Rows (3), Use First Row as Headers.", "Lab 5", "Callen-Lenz Associates Limited - Project Budget 2026")
    xl_log(p, "Budget", "C4", "blank_column", "Blank spacer column C with no header (becomes Column3).", "Remove Columns (Column3).", "Lab 5", "<blank>")
    xl_log(p, "Budget", "D4", "month_header_text", "Month headers are text like 'Jan-26' (wide layout).", "Unpivot Other Columns, then convert 'Jan-26' text to a date.", "Lab 5", "Jan-26")
    xl_log(p, "Budget", "P4", "total_column", "Total column at the right of the month columns.", "Remove the Total column BEFORE unpivoting (else it becomes a 'month').", "Lab 5", "Total")
    xl_log(p, "Budget", f"A{tr}", "total_row", "Total row at the bottom of the data.", "Filter out rows where Project Code does not start with CL-P.", "Lab 5", "Total")
    xl_log(p, "Budget", f"A{tr + 2}", "footnote_row", "Footnote row under the totals.", "Filter out with the Total row (keep only CL-P codes).", "Lab 5", "* CL-P108 closed 31 Mar 2026 - budget retained for reporting only")
    xl_log(p, "Budget", "A8", "lowercase_project_code", "Project code in lower case (cl-p104).", "Transform > Format > UPPERCASE.", "Lab 5", "cl-p104")
    xl_log(p, "Budget", "A10", "trailing_space_project_code", "Project code with trailing space (CL-P106 ).", "Transform > Format > Trim.", "Lab 5", "CL-P106 ")
    save_wb(wb, p)


ASSET_ORDER = list(M._ASSET_OPEN)  # noqa: SLF001
NOV_MOVES = {"Plant & Machinery": (7500.0, 0.0, 0.0), "Vehicles": (18000.0, 0.0, 0.0),
             "IT Equipment": (3200.0, 0.0, 600.0), "Fixtures & Fittings": (0.0, 900.0, 0.0)}


def fmt_br(v: float) -> str:
    return f"({v:,.2f})"


def write_movements_sheet(ws, p, moves_by_class: dict, periods: dict, lab="Lab 5", first_file=True, nov=False):
    ws["A1"] = "Callen-Lenz Associates Limited - Fixed Asset Register: cost movements 2026" if not nov else \
        "Callen-Lenz Associates Limited - Fixed Asset Register: cost movements - November 2026"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = "GBP. Source: Sage FA export (monthly)" if not nov else "GBP. Source: Sage FA export, November drop-in"
    for c, h in enumerate(["Asset Class", "Period", "Additions", "Disposals", "Write Offs", None, "Comment"], start=1):
        ws.cell(4, c, h).font = Font(bold=True)
    r = 5
    tot = [0.0, 0.0, 0.0]
    first_rows = {}
    rowmap = {}
    for cls in ASSET_ORDER:
        ct = [0.0, 0.0, 0.0]
        for k, per in enumerate(periods[cls]):
            a, d, w = moves_by_class[cls][per]
            ct = [ct[0] + a, ct[1] + d, ct[2] + w]
            if k == 0:
                ws.cell(r, 1, cls)
                first_rows[cls] = r
            ws.cell(r, 2, per)
            blank_month = (not nov) and cls == "IT Equipment" and per == 3
            rowmap[(cls, per)] = r
            # additions
            if blank_month:
                pass
            elif a == d == w == 0:
                ws.cell(r, 3, 0)
            elif a:
                ws.cell(r, 3, a).number_format = "#,##0.00"
            else:
                if per % 3 == 0:
                    pass  # blank
                else:
                    ws.cell(r, 3, 0)
            # disposals
            if blank_month:
                pass
            elif d:
                c = ws.cell(r, 4, fmt_br(d))
                c.data_type = "s"
                c.alignment = Alignment(horizontal="right")
            elif per % 2 == 0:
                c = ws.cell(r, 4, "-")
                c.data_type = "s"
                c.alignment = Alignment(horizontal="right")
            # write offs
            if blank_month:
                pass
            elif w:
                if (cls, per) == ("Plant & Machinery", 6):
                    c = ws.cell(r, 5, f"£{w:,.2f}")
                    c.data_type = "s"
                    xl_log(p, "Movements", c.coordinate, "currency_symbol_text", "Write Offs value stored as text with a £ sign and comma ('£1,500.00').",
                           "Remove £ and commas (Replace Values) then Change Type to Decimal Number.", lab, c.value)
                else:
                    ws.cell(r, 5, w).number_format = "#,##0.00"
            elif per in (1, 4, 7, 10) and not nov:
                ws.cell(r, 5, 0)
            if not nov and (cls, per) == ("Plant & Machinery", 2):
                ws.cell(r, 7, "CNC lathe - PO-55410")
            r += 1
        # subtotal row
        ws.cell(r, 1, f"Total {cls}").font = Font(bold=True)
        ws.cell(r, 3, ct[0]).number_format = "#,##0.00"
        c = ws.cell(r, 4, fmt_br(ct[1]) if ct[1] else "-")
        c.data_type = "s"
        ws.cell(r, 5, ct[2]).number_format = "#,##0.00"
        xl_log(p, "Movements", f"A{r}", "subtotal_row_inside_data", f"Subtotal row 'Total {cls}' sits inside the data block (no Period).",
               "Filter out rows where Period is null (after Fill Down).", lab, f"Total {cls}")
        tot = [tot[0] + ct[0], tot[1] + ct[1], tot[2] + ct[2]]
        r += 1
    ws.cell(r, 1, "Grand Total").font = Font(bold=True)
    ws.cell(r, 3, tot[0]).number_format = "#,##0.00"
    c = ws.cell(r, 4, fmt_br(tot[1]))
    c.data_type = "s"
    ws.cell(r, 5, tot[2]).number_format = "#,##0.00"
    xl_log(p, "Movements", f"A{r}", "grand_total_row", "Grand Total row at the bottom of the sheet.", "Filter out rows where Period is null.", lab, "Grand Total")
    xl_log(p, "Movements", "A1", "title_rows_above_header", "Title and source rows above the header (header on row 4).", "Remove Top Rows (3) and promote headers.", lab,
           ws["A1"].value)
    xl_log(p, "Movements", "F4", "blank_column", "Empty spacer column F between data and Comment.", "Select only the 5 needed columns.", lab, "<blank>")
    if not nov:
        xl_log(p, "Movements", "A6", "merged_cell_blank_class", "Asset Class only shown on the first row of each class (merged-cell style); the rest are blank.",
               "Transform > Fill > Down on Asset Class.", lab, "<blank>")
    return rowmap


def build_assets():
    A = M.assets()
    moves = {cls: {} for cls in ASSET_ORDER}
    for _, row in A.iterrows():
        moves[row["Asset Class"]][int(row["Month"])] = (float(row["Additions"]), float(row["Disposals"]), float(row["Write Offs"]))
    p = os.path.join(D_BA, "Asset_Movements_2026.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Movements"
    periods = {cls: list(range(1, 11)) for cls in ASSET_ORDER}
    rowmap = write_movements_sheet(ws, p, moves, periods)
    rr = rowmap[("Vehicles", 10)]
    xl_log(p, "Movements", f"D{rr}", "brackets_negative_text", "Disposals stored as text in brackets, e.g. '(25,000.00)'.",
           "Replace '(' ')' ',' then negate / Number.Abs; Change Type to number.", "Lab 5", ws.cell(rr, 4).value)
    rr = rowmap[("IT Equipment", 3)]
    xl_log(p, "Movements", f"C{rr}", "blank_month_no_values", "One month (IT Equipment, Period 3) has Period but completely blank amounts (blank, not 0).",
           "Replace Values null -> 0 before summing.", "Lab 5", "<blank>")
    for rr in range(5, ws.max_row + 1):
        if ws.cell(rr, 4).value == "-":
            xl_log(p, "Movements", f"D{rr}", "dash_for_zero", "Zero disposals shown as '-' text instead of 0.", "Replace '-' with 0 / treat as 0 (fnNum).", "Lab 5", "-")
            break
    for rr in range(5, ws.max_row + 1):
        if ws.cell(rr, 3).value == 0 and ws.cell(rr, 2).value:
            xl_log(p, "Movements", f"C{rr}", "blank_vs_zero_mixed", "Zero additions appear as 0 on some rows and as blank cells on others.", "Replace null with 0 on Additions / Disposals / Write Offs.", "Lab 5", "0")
            break
    # Opening cost
    wo = wb.create_sheet("Opening_Cost")
    wo["A1"] = "Opening cost at 1 January 2026 by asset class (GBP)"
    wo["A1"].font = Font(bold=True)
    wo.cell(3, 1, "Asset Class").font = Font(bold=True)
    wo.cell(3, 2, "Opening Cost at 1 Jan 2026").font = Font(bold=True)
    op = M.assets_opening()
    for i, (cls, v) in enumerate(zip(op["Asset Class"], op["Opening Cost"]), start=4):
        if cls == "IT Equipment":
            c = wo.cell(i, 2, f"£{v:,.2f}")
            c.data_type = "s"
            xl_log(p, "Opening_Cost", c.coordinate, "currency_symbol_text", "Opening cost for IT Equipment stored as text '£38,000.00'.", "Strip £ and comma, change type.", "Lab 5", c.value)
        else:
            wo.cell(i, 2, float(v)).number_format = "#,##0.00"
        wo.cell(i, 1, cls)
    xl_log(p, "Opening_Cost", "A1", "title_rows_above_header", "Title row and blank row above the header (header on row 3).", "Remove Top Rows (2) then promote headers.", "Lab 5", "Opening cost at 1 January 2026 by asset class (GBP)")
    wo.column_dimensions["A"].width = 24
    wo.column_dimensions["B"].width = 28
    # Roll forward template
    wt = wb.create_sheet("Roll_Forward_Template")
    wt["A1"] = "Cost roll-forward - all asset classes (GBP)"
    wt["A1"].font = Font(bold=True, size=12)
    wt["A3"] = ""
    wt["B3"] = "Month 10"
    wt["C3"] = "Month 11"
    for c in ("B3", "C3"):
        wt[c].font = Font(bold=True)
        wt[c].alignment = Alignment(horizontal="right")
    rf = M.assets_rollforward()
    m10 = rf[rf["Month"] == 10][["Cost BF", "Additions", "Disposals", "Write Offs", "Cost CF"]].sum()
    lines = [("Cost Bfwd", m10["Cost BF"]), ("Additions", m10["Additions"]), ("Disposals", -m10["Disposals"]),
             ("Write offs", -m10["Write Offs"]), ("Cost Cfwd", m10["Cost CF"])]
    for i, (lab_, v) in enumerate(lines, start=4):
        wt.cell(i, 1, lab_).font = Font(bold=(lab_ in ("Cost Bfwd", "Cost Cfwd")))
        wt.cell(i, 2, float(v)).number_format = "#,##0.00;(#,##0.00);-"
        wt.cell(i, 3).number_format = "#,##0.00;(#,##0.00);-"
    wt.column_dimensions["A"].width = 18
    wt.column_dimensions["B"].width = 16
    wt.column_dimensions["C"].width = 16
    xl_log(p, "Roll_Forward_Template", "C3", "month_11_blank_column", "Template has Month 10 filled and Month 11 left blank - the lab fills it by refreshing with the November drop-in and Period = 11.",
           "Parameter Period + roll-forward query; paste the Month 11 output into column C.", "Lab 5", "Month 11")
    wb.move_sheet("Roll_Forward_Template", offset=0)
    save_wb(wb, p)
    # November drop-in
    pn = os.path.join(D_BA, "_Later", "Asset_Movements_2026-11.xlsx")
    wbn = Workbook()
    wsn = wbn.active
    wsn.title = "Movements"
    moves_n = {cls: {11: NOV_MOVES[cls]} for cls in ASSET_ORDER}
    write_movements_sheet(wsn, pn, moves_n, {cls: [11] for cls in ASSET_ORDER}, nov=True)
    save_wb(wbn, pn)
    return moves


def asset_rollforward_mimic(period: int, include_nov: bool):
    """Mimic lab5_rollforward.m by reading the xlsx files back."""
    def fnnum(x):
        if x is None:
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x).strip()
        neg = (s.startswith("(") and s.endswith(")")) or s.startswith("-")
        dg = re.sub(r"[^0-9.]", "", s)
        return 0.0 if dg == "" else (-1 if neg else 1) * float(dg)

    files = [os.path.join(D_BA, "Asset_Movements_2026.xlsx")]
    if include_nov:
        files.append(os.path.join(D_BA, "_Later", "Asset_Movements_2026-11.xlsx"))
    recs = []
    for f in files:
        ws = load_workbook(f, data_only=True)["Movements"]
        rows = list(ws.iter_rows(values_only=True))[3:]
        hdr = list(rows[0])
        cls = None
        for row in rows[1:]:
            d = dict(zip(hdr, row))
            if d["Asset Class"] is not None:
                cls = d["Asset Class"]
            if d["Period"] is None:
                continue
            recs.append((cls, int(d["Period"]), abs(fnnum(d["Additions"])), abs(fnnum(d["Disposals"])), abs(fnnum(d["Write Offs"]))))
    wo = load_workbook(files[0], data_only=True)["Opening_Cost"]
    opening = sum(fnnum(r[1]) for r in list(wo.iter_rows(values_only=True))[3:] if r[0])
    df = pd.DataFrame(recs, columns=["cls", "p", "a", "d", "w"])
    e = df[df["p"] < period]
    t = df[df["p"] == period]
    bf = opening + e["a"].sum() - e["d"].sum() - e["w"].sum()
    out = {"Cost Bfwd": bf, "Additions": t["a"].sum(), "Disposals": -t["d"].sum(), "Write offs": -t["w"].sum()}
    out["Cost Cfwd"] = sum(out.values())
    return {k: round(float(v), 2) for k, v in out.items()}, df, opening


def budget_unpivot_mimic() -> pd.DataFrame:
    ws = load_workbook(os.path.join(D_BA, "Project_Budget_2026.xlsx"), data_only=True)["Budget"]
    rows = list(ws.iter_rows(values_only=True))[3:]
    hdr = [h if h is not None else "Column3" for h in rows[0]]
    recs = []
    for row in rows[1:]:
        d = dict(zip(hdr, row))
        code = "" if d["Project Code"] is None else str(d["Project Code"]).strip().upper()
        if not code.startswith("CL-P"):
            continue
        for k, v in d.items():
            if k in ("Project Code", "Project Name", "Column3", "Total"):
                continue
            if v is None:
                num = 0.0
            elif isinstance(v, (int, float)):
                num = float(v)
            else:
                num = float(re.sub(r"[^0-9.]", "", str(v)) or 0)
            mon = datetime.strptime("1 " + k[:3] + " 20" + k.strip()[-2:], "%d %b %Y").date()
            recs.append((code, mon, num))
    return pd.DataFrame(recs, columns=["Project Code", "Month", "Budget GBP"])


# ====================================================================================
# LAB 6 - CARDS
# ====================================================================================
CARD_STD = ["Date", "Posted date", "Status", "Merchant", "Amount", "Category", "City", "Last four", "Name on card", "MCC",
            "Trans Type", "Bank Ref", "Original Currency Amount", "Merchant Currency Description", "Merchant Country"]
CUR_DESC = {"GBP": "Pound Sterling", "USD": "US Dollar", "EUR": "Euro"}
CARDS = M.cards()
CARDS["m"] = [d.month for d in CARDS["Date"]]
HOLDER = EMP[EMP["Card Last Four"] != ""].set_index("Card Last Four")["Employee Name"].to_dict()
HOLDER_ID = EMP[EMP["Card Last Four"] != ""].set_index("Card Last Four")["Employee ID"].to_dict()
CT = {
    "blank_amount_gbp": ("Amount blank on a GBP row.", "Amount GBP = Original Currency Amount (GBP rate 1.00)."),
    "blank_amount_refund": ("Amount blank on a refund row (no CR amount shown).", "Amount GBP = Original Currency Amount (already negative for a refund) x rate 1.00."),
    "zero_amount": ("Amount is 0.00 although the transaction has a non-zero original amount.", "Treat 0 like blank: Original Currency Amount x FX rate."),
    "blank_amount_foreign": ("Amount blank on a non-GBP row (card bank did not convert).", "Amount GBP = Original Currency Amount x rate from FX_Rates.xlsx (merge on currency description)."),
    "cr_suffix_refund": ("Refund amounts written as '45.00 CR' text.", "Detect trailing CR, strip it and make the number negative."),
    "currency_symbol_amount": ("Amount stored as text with £ and thousands comma ('£1,250.00').", "Remove £ and comma, change type to Decimal Number."),
    "city_double_space": ("City with extra spaces (e.g. 'NEW  YORK', trailing space).", "Trim + collapse double spaces, then Capitalize Each Word."),
    "city_case": ("City in upper or lower case.", "Transform > Format > Capitalize Each Word."),
    "last_four_leading_zero_lost": ("Last four stored as a number so the leading 0 is lost (937 instead of 0937).", "Text.PadStart(text, 4, \"0\") after converting to text."),
    "merchant_extra_spaces": ("Merchant with leading/trailing or double spaces.", "Trim + collapse spaces."),
    "pending_status": ("Row has Status 'Pending' (not yet posted; no posted date).", "Filter Status = Posted."),
    "status_case": ("Status in upper case (POSTED) rather than 'Posted'.", "Compare with Text.Upper, or capitalise before filtering."),
    "duplicate_across_statements": ("Same Bank Ref appears at the end of one statement and the start of the next (month-end overlap).", "Remove Duplicates on Bank Ref only (after cleaning it)."),
    "near_duplicate_bank_ref": ("Overlapping transaction whose Bank Ref differs only by case/trailing space.", "Trim + Upper the Bank Ref first, then Remove Duplicates on Bank Ref."),
    "title_rows_above_header": ("Statement has title lines above the header row (header on line 4).", "Skip rows until the row where first cell = 'Date' (or Remove Top Rows)."),
    "header_case_trailing_space": ("Headers have trailing spaces / different case versus other statements (e.g. 'Last four ').", "Rename columns by position (all statements have the same 15 columns)."),
    "text_date_dd_mon_yyyy": ("Dates written as '05-Feb-2026' text in the February statement; others are dd/mm/yyyy.", "Date.FromText with en-GB culture handles both; or Change Type Using Locale."),
    "name_on_card_case": ("Name on card in lower or mixed case.", "Not needed for the join - match on Last Four (discussion point)."),
    "currency_description_case": ("Currency description in upper case or with trailing space.", "Upper + Trim both sides before merging with FX_Rates."),
    "blank_line": ("Blank line (all empty fields).", "Remove Blank Rows / filter Bank Ref <> ''."),
}


MATERIAL = {"blank_amount_gbp", "zero_amount", "blank_amount_foreign", "cr_suffix_refund", "currency_symbol_amount"}


def ctag(k):
    return (k, *CT[k])


def render_card(row, fi, r, flags, overlap=False, ref_variant=None):
    """flags: dict amount_mode in {None,'blank','zero'}"""
    tags = []
    fmts = {1: "%d/%m/%Y", 2: "%d-%b-%Y", 3: "%d/%m/%Y"}
    f = fmts[fi]
    ds = row["Date"].strftime(f)
    ps = row["Posted Date"].strftime(f)
    if fi == 2:
        tags.append(ctag("text_date_dd_mon_yyyy"))
    status = "Posted"
    if not overlap and r.random() < 0.08:
        status = "POSTED"
        tags.append(ctag("status_case"))
    merch = row["Merchant"]
    if not overlap and r.random() < 0.06:
        v = int(r.integers(0, 3))
        merch = [merch + " ", " " + merch, merch.replace(" ", "  ", 1) if " " in merch else merch + "  "][v]
        tags.append(ctag("merchant_extra_spaces"))
    cur = row["Original Currency"]
    mode = flags
    amt = f"{row['Amount GBP']:.2f}"
    is_ref = row["Trans Type"] == "Refund"
    if is_ref and mode != "blank":
        amt = f"{abs(row['Amount GBP']):.2f} CR"
        tags.append(ctag("cr_suffix_refund"))
    if mode == "blank":
        amt = ""
        tags.append(ctag("blank_amount_refund" if is_ref else ("blank_amount_gbp" if cur == "GBP" else "blank_amount_foreign")))
    elif mode == "zero":
        amt = "0.00"
        tags.append(ctag("zero_amount"))
    elif not is_ref and row["Amount GBP"] >= 100 and not overlap and r.random() < 0.18:
        amt = f"£{row['Amount GBP']:,.2f}"
        tags.append(ctag("currency_symbol_amount"))
    city = row["City"]
    if not overlap:
        x = r.random()
        if city == "New York":
            city = "NEW  YORK"
            tags.append(ctag("city_double_space"))
        elif x < 0.10:
            city = city.upper()
            tags.append(ctag("city_case"))
        elif x < 0.17:
            city = city.lower()
            tags.append(ctag("city_case"))
        elif x < 0.22:
            city = city + " "
            tags.append(ctag("city_double_space"))
    lf = row["Last Four"]
    if lf.startswith("0") and (fi == 1 or (fi == 2 and r.random() < 0.5)):
        lf = str(int(lf))
        tags.append(ctag("last_four_leading_zero_lost"))
    nm = row["Name on Card"]
    if not overlap and r.random() < 0.10:
        nm = nm.lower() if r.random() < 0.5 else nm.title()
        tags.append(ctag("name_on_card_case"))
    cd = CUR_DESC[cur]
    if not overlap and r.random() < 0.15:
        cd = cd.upper() if r.random() < 0.5 else cd + " "
        tags.append(ctag("currency_description_case"))
    ref = row["Bank Ref"]
    if ref_variant == "lower":
        ref = ref.lower() + " "
        tags.append(ctag("near_duplicate_bank_ref"))
    elif overlap:
        tags.append(ctag("duplicate_across_statements"))
    cells = [ds, ps, status, merch, amt, row["Category"], city, lf, nm, str(int(row["MCC"])), row["Trans Type"], ref,
             f"{row['Original Currency Amount']:.2f}", cd, row["Merchant Country"]]
    return {"cells": cells, "tags": tags}


def build_cards():
    r = rng(50)
    pend_merch = [("Trainline", "Travel", 4011, "London"), ("Shell", "Fuel", 5541, "Poole"), ("Premier Inn", "Accommodation", 7011, "Leeds"),
                  ("Pret", "Subsistence", 5814, "London"), ("Screwfix", "Materials", 5200, "Southampton")]
    holders = list(HOLDER)
    prev_rows = None
    flagged = set()
    for fi in (1, 2, 3):
        rr = np.random.default_rng([M.SEED, 7000, 60 + fi])
        df = CARDS[CARDS["m"] == fi].reset_index(drop=True)
        n = len(df)
        gbp_idx = [i for i in range(n) if df.loc[i, "Original Currency"] == "GBP" and df.loc[i, "Trans Type"] == "Purchase"]
        picks = rr.choice(gbp_idx, size=5, replace=False)
        modes = {int(picks[0]): "blank", int(picks[1]): "blank", int(picks[2]): "blank", int(picks[3]): "zero", int(picks[4]): "zero"}
        foreign = [i for i in range(n) if df.loc[i, "Original Currency"] != "GBP"]
        seq = {1: ["blank", None, "zero"], 2: ["blank", "zero", None], 3: ["blank", None]}[fi]
        for k, i in enumerate(foreign):
            if seq[k]:
                modes[i] = seq[k]
        refunds = [i for i in range(n) if df.loc[i, "Trans Type"] == "Refund"]
        if fi == 2:
            modes[refunds[0]] = "blank"
        lines = []
        if fi == 3:
            t1 = ["Callen-Lenz Associates Limited - Corporate Card Statement"] + [""] * 14
            t2 = ["Statement period: 01/03/2026 to 31/03/2026"] + [""] * 14
            lines.append({"cells": t1, "tags": [ctag("title_rows_above_header")]})
            lines.append({"cells": t2, "tags": []})
            lines.append({"cells": [""] * 15, "tags": []})
        hdr = list(CARD_STD)
        htags = []
        if fi == 2:
            hdr = [h + " " if h in ("Amount", "Last four", "Posted date") else h for h in hdr]
            hdr[8] = "NAME ON CARD"
            htags.append(ctag("header_case_trailing_space"))
        if fi == 3:
            hdr = ["Date", "Posted Date", "Status", "Merchant", "Amount", "Category", "City", "Last Four", "Name on Card", "MCC", "Trans Type",
                   "Bank Ref", "Original Currency Amount", "Merchant Currency Description", "Merchant Country"]
        lines.append({"cells": hdr, "tags": htags})
        # overlap from previous statement
        if prev_rows is not None:
            ov = prev_rows.tail(6).reset_index(drop=True)
            for j in range(len(ov)):
                lines.append(render_card(ov.loc[j], fi, rr, None, overlap=True, ref_variant="lower" if j in (1, 4) else None))
        body_start = len(lines)
        for i in range(n):
            ln = render_card(df.loc[i], fi, rr, modes.get(i))
            if modes.get(i) or any(t[0] in MATERIAL for t in ln["tags"]):
                flagged.add((fi, df.loc[i, "Bank Ref"]))
            lines.append(ln)
        # one blank line in the middle
        lines.insert(body_start + int(rr.integers(10, n - 10)), {"cells": [""] * 15, "tags": [ctag("blank_line")]})
        # pending rows at the end
        last = (date(2026, fi + 1, 1) - timedelta(days=1)).day
        for j in range(5):
            m_ = pend_merch[int(rr.integers(0, len(pend_merch)))]
            h = holders[int(rr.integers(0, len(holders)))]
            d_ = date(2026, fi, last - int(rr.integers(0, 3)))
            lf = h if not (h.startswith("0") and fi == 1) else str(int(h))
            cells = [d_.strftime("%d/%m/%Y" if fi != 2 else "%d-%b-%Y"), "", "Pending", m_[0], f"{float(rr.uniform(8, 160)):.2f}", m_[1], m_[3], lf,
                     "", str(m_[2]), "Purchase", f"PEND{int(rr.integers(100000, 999999))}", "", "Pound Sterling", "GB"]
            # fill name and orig amount
            cells[8] = (EMP[EMP["Card Last Four"] == h]["Employee Name"].iloc[0].split(" ", 1))
            cells[8] = f"{cells[8][0][0]} {cells[8][1]}".upper()
            cells[12] = cells[4]
            lines.append({"cells": cells, "tags": [ctag("pending_status")]})
        write_csv(os.path.join(D_CD, f"Card_Statement_2026-{fi:02d}.csv"), lines, "Lab 6")
        prev_rows = df
    return flagged


def build_card_refs():
    # FX rates
    wb = Workbook()
    ws = wb.active
    ws.title = "FX_Rates"
    ws["A1"] = "FX rates used to restate foreign-currency card spend (fixed for Q1 2026)"
    ws["A1"].font = Font(bold=True)
    ws["A2"] = "Source: Treasury. Rate = GBP per 1 unit of currency."
    for c, h in enumerate(["Currency Code", "Currency Description", "GBP per 1 unit"], start=1):
        ws.cell(4, c, h).font = Font(bold=True)
    for i, (code, desc) in enumerate(CUR_DESC.items(), start=5):
        ws.cell(i, 1, code)
        ws.cell(i, 2, desc)
        ws.cell(i, 3, M.FX_RATES[code]).number_format = "0.0000"
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 16
    p = os.path.join(D_CD, "FX_Rates.xlsx")
    save_wb(wb, p)
    xl_log(p, "FX_Rates", "A1", "title_rows_above_header", "Title and source rows above the header (header on row 4).", "Remove Top Rows (3), promote headers.", "Lab 6",
           "FX rates used to restate foreign-currency card spend (fixed for Q1 2026)")
    # Employee lookup (client 3C layout) with name mess
    wb = Workbook()
    ws = wb.active
    ws.title = "Employees"
    for c, h in enumerate(["Employee", "Department", "Manager", "Email"], start=1):
        ws.cell(1, c, h).font = Font(bold=True)
    mess = {"E002": "Rajesh Patell", "E005": "Priya Nayar", "E006": "Reid, Callum", "E007": "O. Hartley", "E012": "Aisha Rahman ",
            "E015": "Steve  Mallory", "E020": "Angus Mcbride", "E021": "Fiona Sutherlnd", "E022": "Marcus Low", "E024": "paul danvers"}
    i = 2
    p = os.path.join(D_CD, "Employee_Lookup.xlsx")
    shown = {}
    for _, e in EMP.iterrows():
        nm = mess.get(e["Employee ID"], e["Employee Name"])
        ws.cell(i, 1, nm)
        ws.cell(i, 2, e["Department"].title())
        ws.cell(i, 3, e["Manager"])
        ws.cell(i, 4, e["Email"] if e["Employee ID"] != "E012" else e["Email"].upper())
        if e["Employee ID"] in mess:
            shown[e["Employee ID"]] = (i, nm)
        i += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 44
    for eid, desc in (("E002", "typo in surname"), ("E006", "'Surname, First' order"), ("E007", "initial instead of first name"), ("E015", "double space"),
                      ("E024", "all lower case")):
        row, nm = shown[eid]
        xl_log(p, "Employees", f"A{row}", "name_spelling_mismatch", f"Employee name spelled differently from Employees.xlsx / card file ({desc}).",
               "Do NOT join on name: join cards to Employees on Last Four (this is a discussion point on why keys beat names).", "Lab 6", nm)
    row, nm = shown["E012"]
    xl_log(p, "Employees", f"D{row}", "email_upper_case", "Email in upper case.", "Transform > Format > lowercase.", "Lab 6", EMP[EMP["Employee ID"] == "E012"]["Email"].iloc[0].upper())
    save_wb(wb, p)
    # Expense template (client 3B header layout)
    wb = Workbook()
    ws = wb.active
    ws.title = "Expense Claim"
    ws["A1"] = "Callen-Lenz Associates Limited - Expense Claim Form"
    ws["A1"].font = Font(bold=True, size=13)
    for rr_, lab_ in enumerate(["Employee name:", "Department:", "Manager:", "Card last four:", "Claim period:"], start=3):
        ws.cell(rr_, 1, lab_).font = Font(bold=True)
    hdr = ["Date", "Merchant", "Category", "Project Code", "Description", "Net GBP", "VAT GBP", "Gross GBP", "Receipt attached (Y/N)"]
    for c, h in enumerate(hdr, start=1):
        cell = ws.cell(10, c, h)
        cell.font = Font(bold=True)
    ws.cell(21, 7, "Total").font = Font(bold=True)
    for c, w in enumerate([12, 24, 16, 14, 30, 12, 12, 12, 22], start=1):
        ws.column_dimensions[ws.cell(10, c).column_letter].width = w
    save_wb(wb, os.path.join(D_CD, "Expense_Template.xlsx"))


# ---- mimic of lab6_cards.m
def card_amount(x: str):
    up = x.upper()
    refund = up.endswith("CR") or up.startswith("-") or (up.startswith("(") and up.endswith(")"))
    dg = re.sub(r"[^0-9.]", "", up)
    if dg == "":
        return None
    return (-1 if refund else 1) * float(dg)


def clean_cards():
    paths = [os.path.join(D_CD, f"Card_Statement_2026-{m:02d}.csv") for m in (1, 2, 3)]
    parts = []
    raw_rows = 0
    for p in paths:
        with open(p, newline="", encoding="utf-8") as f:
            rows = [[pq_trim(c) for c in r] for r in csv.reader(f) if r]
        hi = [i for i, r in enumerate(rows) if r[0].lower() == "date"][0]
        parts.append(pd.DataFrame(rows[hi + 1:], columns=CARD_STD_POS))
    comb = pd.concat(parts, ignore_index=True)
    raw_rows = len(comb)
    nb = comb[comb["Bank Ref"] != ""].copy()
    posted = nb[nb["Status"].str.upper() == "POSTED"].copy()
    posted["Bank Ref"] = posted["Bank Ref"].str.upper()
    dd = posted.drop_duplicates("Bank Ref").copy()
    dd["Posted"] = [parse_uk(x) if x else None for x in dd["Posted Date"]]
    dd["Date"] = [parse_uk(x) for x in dd["Date"]]
    dd["City"] = dd["City"].map(lambda s: " ".join(x for x in s.split(" ") if x).title())
    dd["Last Four"] = dd["Last Four"].str.zfill(4)
    dd["Orig"] = dd["Original Currency Amount"].map(card_amount)
    dd["Stmt"] = dd["Amount"].map(card_amount)
    dd["Cur"] = dd["Merchant Currency Description"].map(lambda s: " ".join(x for x in s.split(" ") if x).upper())
    rate = {v.upper(): M.FX_RATES[k] for k, v in CUR_DESC.items()}
    dd["Rate"] = dd["Cur"].map(rate)
    dd["Amount GBP"] = [s if (not pd.isna(s) and s != 0) else round(o * rt, 2) for s, o, rt in zip(dd["Stmt"], dd["Orig"], dd["Rate"])]
    dd["Holder"] = dd["Last Four"].map(HOLDER)
    dd["Month"] = [date(x.year, x.month, 1) for x in dd["Date"]]
    return dict(raw_rows=raw_rows, nonblank=len(nb), pending_removed=len(nb) - len(posted), dup_removed=len(posted) - len(dd), clean=dd.reset_index(drop=True),
                posted_rows=len(posted))


CARD_STD_POS = ["Date", "Posted Date", "Status", "Merchant", "Amount", "Category", "City", "Last Four", "Name on Card", "MCC", "Trans Type",
                "Bank Ref", "Original Currency Amount", "Merchant Currency Description", "Merchant Country"]


def parse_uk(s: str) -> date:
    for f in ("%d/%m/%Y", "%d-%b-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(s)


# ====================================================================================
# MAIN
# ====================================================================================
def main():
    for d in (D_TS, os.path.join(D_TS, "_Later"), D_BA, os.path.join(D_BA, "_Later"), D_CD, os.path.join(SOL, "m")):
        os.makedirs(d, exist_ok=True)
    ts_totals = build_all_timesheets()
    ctl = write_control_totals()
    build_budget()
    build_assets()
    flagged = build_cards()
    build_card_refs()

    # ------------------------------------------------ LAB 2 checks
    truth = TS.copy()
    tl = labour_lines(truth)
    for label, jun in (("jan_may", False), ("jan_jun", True)):
        res = load_timesheets(jun)
        got = res["valid"].copy()
        mmax = 6 if jun else 5
        exp = truth[truth["m"] <= mmax][TS_COLS].copy()
        k = lambda d: sorted(zip(d["Week Ending"], d["Employee ID"], d["Employee Name"], d["Project Code"], d["Task"], d["Hours"]))  # noqa: E731
        assert k(got) == k(exp), f"cleaned timesheets != truth ({label})"
        assert abs(got["Hours"].sum() - exp["Hours"].sum()) < 1e-9
        if jun:
            res_all = res
        else:
            res_5 = res
    # stages
    r5, r6 = res_5, res_all
    inter("L2_files_before_June", 5, "Timesheet_2026-01..05.csv in 02_Timesheets (June sits in _Later until the refresh exercise)")
    inter("L2_rows_after_combine_raw_lines", r5["raw_lines"],
          "Data lines (below header) in the 5 files that contain something: includes Total rows (5) and comma-only lines (5); truly empty lines are skipped by Power Query - approximate until confirmed on the day")
    inter("L2_rows_after_removing_blank_and_total", r5["after_removal"], "Rows left after filtering blank lines and the 'Total' row in every file (still has duplicates and unknown-code rows)")
    n_dups = len(r5["combined"]) - len(r5["deduped"])
    inter("L2_duplicates_removed", n_dups, "Rows removed by Remove Duplicates after cleaning (5 exact + 1 near-duplicate in Jan = 6)")
    inter("L2_rows_after_dedupe", len(r5["deduped"]), "Rows after Remove Duplicates (still includes the unknown-code rows)")
    inter("L2_unknown_code_rows_Jan_May", len(r5["unknown"]), "Rows whose cleaned Project Code is not in Projects.xlsx (all in the February file)")
    inter("L2_rows_valid_Jan_May", len(r5["valid"]), "Rows left after removing unknown codes = clean truth rows for Jan-May")
    assert len(r5["valid"]) == int((truth["m"] <= 5).sum())
    inter("L2_rows_valid_Jan_Jun", len(r6["valid"]), "As above but after copying June into the folder and refreshing")
    deduped = r5["deduped"].copy()
    deduped["Month"] = [date(x.year, x.month, 1) for x in deduped["Week Ending"]]
    deduped["valid"] = deduped["Project Code"].isin(VALID_CODES)
    ctl_hours = ctl["sum"].to_dict()
    ctl_rows = ctl["count"].to_dict()
    diffs = {}
    chk = []
    for mo, g in deduped.groupby("Month"):
        diff = round(g["Hours"].sum() - ctl_hours[mo], 3)
        unk = round(g.loc[~g["valid"], "Hours"].sum(), 3)
        diffs[mo.month] = diff
        chk.append((mo.month, round(g["Hours"].sum(), 1), ctl_hours[mo], diff, unk, int((~g["valid"]).sum())))
    off = [c for c in chk if abs(c[3]) > 1e-9]
    assert len(off) == 1 and off[0][0] == 2 and abs(off[0][3] - 7.5) < 1e-9 and abs(off[0][4] - 7.5) < 1e-9 and off[0][5] == 6, off
    for c in chk:
        inter(f"L2_loaded_hours_after_dedupe_month_{c[0]:02d}", c[1], f"Hours in month {c[0]} after cleaning + dedupe (before removing unknown codes); control total {c[2]}; difference {c[3]}")
    inter("L2_control_off_month", "Feb", "Only month where loaded (deduped) hours differ from Control_Totals")
    for mo, g in sorted((mo, g) for mo, g in deduped.groupby("Month")):
        pass
    ll5 = labour_lines(r5["valid"])
    ll6 = labour_lines(r6["valid"])
    # cross-check against truth costs
    for ll, mmax in ((ll5, 5), (ll6, 6)):
        t_ = tl[tl["m"] <= mmax]
        assert abs(ll["Labour Cost"].sum() - t_["Labour Cost"].sum()) < 1e-6
    q1 = ll5[ll5["Month"] <= date(2026, 3, 1)]
    q6 = round(float(q1[q1["Project Code"] == "CL-P103"]["Labour Cost"].sum()), 2)
    q6_hours = float(q1[q1["Project Code"] == "CL-P103"]["Hours"].sum())
    q6_ot = round(float(q1[q1["Project Code"] == "CL-P103"]["OT Hours"].sum()), 2)
    t_q1 = tl[(tl["m"] <= 3) & (tl["Project Code"] == "CL-P103")]
    assert abs(t_q1["Labour Cost"].sum() - q6) < 0.006
    # ---- trap answers (wrong-but-common results, for trainer diagnosis)
    q1m = date(2026, 3, 1)
    def p103(df):
        d_ = labour_lines(df)
        d_ = d_[(d_["Month"] <= q1m) & (d_["Project Code"] == "CL-P103")]
        return d_
    trap6_unknown_first = round(float(p103(r5["deduped"])["Labour Cost"].sum()), 2)  # OT computed with unknown rows still in
    trap6_dups = round(float(p103(r5["combined"][r5["combined"]["Project Code"].isin(VALID_CODES)])["Labour Cost"].sum()), 2)
    x103 = ll5[(ll5["Month"] <= q1m) & (ll5["Project Code"] == "CL-P103")]
    trap6_noot = round(float((x103["Hours"] * x103["Rate"]).sum()), 2)
    ask("Q6", "Lab 2", "Combine the January-May timesheets and clean them (headers, dates, codes, hours, Total/blank rows, duplicates, unknown codes), then merge Employees and Labour_Rates. One Week Ending has a typo year 2062; correct it to 2026. Overtime rule: for each employee and Week Ending, hours above 37.5 are overtime, counted AFTER removing duplicates and unknown-code rows. Each line gets its share of that overtime in proportion to its hours. Labour cost per line = Hours x Standard Rate + that line's OT hours x Standard Rate x (OT Multiplier - 1). What is the total labour cost for CL-P103 for Jan-Mar (month of Week Ending)?",
        q6, "GBP",
        "Labour Cost per line = Hours x Standard Rate + OT Hours(line) x Standard Rate x (OT Multiplier - 1). Weekly Hours = sum of an employee's hours for a Week Ending (after removing unknown-code rows and duplicates). OT Hours(line) = max(0, Weekly Hours - 37.5) x Hours / Weekly Hours (pro-rata by hours). Month = month of Week Ending. Rounded to 2dp at the end.",
        "Folder combine with fnLoadTimesheet, Remove Duplicates, drop unknown codes, merge Employees (Grade) and Labour_Rates, Group By Employee+Week for Weekly Hours, add OT Hours and Labour Cost, filter CL-P103 & Jan-Mar, sum.",
        hours=q6_hours, ot_hours=q6_ot,
        trap_answers={"overtime_computed_before_removing_unknown_codes": trap6_unknown_first, "duplicates_kept": trap6_dups,
                      "no_overtime_premium_(hours_x_rate_only)": trap6_noot, "year_2062_not_fixed": "affects May weeks only for the month grouping; Q6 (Jan-Mar) unchanged"})
    inter("L2_CLP103_Q1_hours", q6_hours, "Hours booked to CL-P103 with Week Ending in Jan-Mar (clean)")
    inter("L2_CLP103_Q1_ot_hours", q6_ot, "Overtime hours allocated to CL-P103 in Q1 (pro-rata rule)")
    ask("Q7", "Lab 2", "Compare loaded hours per month (after cleaning and removing duplicates, BEFORE removing unknown project codes) with Control_Totals.xlsx. Which month fails and by how many hours, and why?",
        "Feb: loaded 7.5 hours MORE than control (3006.5 vs 2999.0); explained by 6 unknown-code rows", "text",
        "Per-month sum of Hours of the deduped table (Week Ending month) minus 'System Hours' in Control_Totals.xlsx. Control = clean truth hours for valid projects. Only Feb is non-zero: +7.5 hours (also +6 lines) = the six rows booked to CL-P199 (4 rows) and CL-P100 (2 rows).",
        "Group deduped table by Month; merge with Control; Difference = Loaded - System; then Left Anti merge with Projects to see the 6 rows (7.5 h) that explain it.",
        trap_answers={"unknown_codes_removed_before_control_check": "every month shows difference 0.0 (Feb looks fine, the discrepancy is hidden); loaded Feb = 2999.0 = control"},
        month="Feb", difference_hours=7.5, loaded_hours=3006.5, control_hours=2999.0, difference_lines=6)
    unk = r5["unknown"]
    assert len(unk) == 6 and abs(unk["Hours"].sum() - 7.5) < 1e-9
    assert (unk["Week Ending"].map(lambda d: d.month) == 2).all()
    ask("Q8", "Lab 2", "How many timesheet rows carry an unknown project code (not in Projects.xlsx), and how many hours are on them?",
        "6 rows; 7.5 hours", "text",
        "After cleaning codes (upper, rebuilt as CL-Pnnn) and Remove Duplicates, rows whose Project Code is not in Projects.xlsx: CL-P199 (4 rows) and CL-P100 (2 rows), all in the February file.",
        "Merge with Projects, join kind Left Anti; count rows and sum Hours.", rows=6, hours=7.5)
    ytd = round(float(ll6["Hours"].sum()), 2)
    assert abs(ytd - float(truth["Hours"].sum())) < 1e-9
    ask("Q9", "Lab 2", "Copy Timesheet_2026-06.csv from the _Later folder into 02_Timesheets and refresh. Remember to remove duplicates and unknown-code rows, and that one Week Ending has a typo year 2062; correct it to 2026. What are the total valid hours year to date (Jan-Jun)?",
        ytd, "hours",
        "Sum of Hours of cleaned, de-duplicated rows with valid project codes for all six monthly files (Week Ending Jan 2 - Jun 26 2026). June adds 2,987.5 hours.",
        "Drop file in folder, Refresh; the combine picks it up (culture: UK dates) and the result table sums to the total.",
        june_hours=round(float(ll6[ll6["Month"] == date(2026, 6, 1)]["Hours"].sum()), 2), labour_cost_ytd=round(float(ll6["Labour Cost"].sum()), 2))
    inter("L2_labour_cost_total_Jan_May", round(float(ll5["Labour Cost"].sum()), 2), "Total labour cost incl. overtime, valid rows Jan-May")
    inter("L2_labour_cost_total_Jan_Jun", round(float(ll6["Labour Cost"].sum()), 2), "Total labour cost incl. overtime, valid rows Jan-Jun")
    inter("L2_ot_hours_total_Jan_Jun", round(float(ll6["OT Hours"].sum()), 2), "Overtime hours, Jan-Jun (sum of weekly hours above 37.5)")
    inter("L2_project_month_groups_Jan_Jun", int(ll6.groupby(["Project Code", "Month"]).ngroups), "Rows in the Group By project-month output")
    inter("L2_rows_after_dedupe_with_June", len(r6["deduped"]), "Rows after Remove Duplicates with June included")
    inter("L2_duplicates_removed_with_June", len(r6["combined"]) - len(r6["deduped"]), "6 above + 1 near-duplicate in June")

    # ------------------------------------------------ LAB 5 checks
    ub = budget_unpivot_mimic()
    tb = M.budget()
    tb["Month"] = tb["Month"].map(lambda d: date(d.year, d.month, 1))
    a = ub.sort_values(["Project Code", "Month"]).reset_index(drop=True)
    b = tb.sort_values(["Project Code", "Month"]).reset_index(drop=True)
    assert len(a) == len(b) == 96
    assert (a["Project Code"] == b["Project Code"]).all() and (a["Month"] == b["Month"]).all()
    assert np.allclose(a["Budget GBP"], b["Budget GBP"])
    q16 = round(float(a[(a["Project Code"] == "CL-P105") & (a["Month"] <= date(2026, 6, 1))]["Budget GBP"].sum()), 2)
    inter("L5_budget_rows_after_unpivot", 96, "8 projects x 12 months")
    inter("L5_budget_total_all_projects", round(float(a["Budget GBP"].sum()), 2), "Sum of Budget GBP after unpivot = sum of Projects.xlsx Budget 2026")
    assert abs(a["Budget GBP"].sum() - M.projects()["Budget 2026 GBP"].sum()) < 0.005
    ask("Q16", "Lab 5", "Unpivot the wide budget (remove title rows, spacer/Total columns and Total/footnote rows first; fix lower-case code and the text number). What is the total budget for CL-P105 for H1 (Jan-Jun 2026)?",
        q16, "GBP", "Sum of Budget GBP for Project Code CL-P105 with Month Jan-26 .. Jun-26 after unpivot (CL-P105 March is stored as text and must be converted).",
        "Skip 3 rows, promote headers, remove Column3 + Total, filter Project Code starts with CL-P, Unpivot Other Columns, convert 'Jan-26' to date, filter H1, sum.")
    # pivot hours (jan-jun)
    piv = ll6.groupby(["Project Code", "Month"])["Hours"].sum().unstack(fill_value=0.0)
    q17_proj, q17_mon = "CL-P102", date(2026, 3, 1)
    q17 = round(float(piv.loc[q17_proj, q17_mon]), 2)
    top = piv.sum(axis=1).sort_values(ascending=False)
    inter("L5_pivot_grand_total_hours", round(float(piv.values.sum()), 2), "Grand total of the pivoted hours table (Jan-Jun, valid rows)")
    inter("L5_pivot_top_project", f"{top.index[0]} ({top.iloc[0]:.1f} h)", "Project with the most hours Jan-Jun")
    inter("L5_pivot_rows", int(len(piv)), "Rows in the pivot (projects with hours)")
    ask("Q17", "Lab 5", f"Pivot the cleaned hours (from Lab 2 incl. June) so Project Code is the rows and months Jan-Jun are columns. How many hours were booked to {q17_proj} in Mar-26?",
        q17, "hours", "Sum of Hours for Project Code CL-P102 with Week Ending in March 2026 (valid, de-duplicated rows).",
        "LabourLines -> keep Project Code, Month, Hours -> month name -> Pivot Column (Hours, Sum) -> read CL-P102 / Mar.")
    # roll forward
    rf10, adf10, opening = asset_rollforward_mimic(10, False)
    truth_rf = M.assets_rollforward()
    t10 = truth_rf[truth_rf["Month"] == 10].sum(numeric_only=True)
    assert abs(rf10["Cost Bfwd"] - t10["Cost BF"]) < 0.005 and abs(rf10["Additions"] - t10["Additions"]) < 0.005
    assert abs(rf10["Disposals"] + t10["Disposals"]) < 0.005 and abs(rf10["Write offs"] + t10["Write Offs"]) < 0.005
    assert abs(rf10["Cost Cfwd"] - t10["Cost CF"]) < 0.005
    assert rf10 == {"Cost Bfwd": 250000.0, "Additions": 10000.0, "Disposals": -25000.0, "Write offs": 0.0, "Cost Cfwd": 235000.0}
    rf11, _, _ = asset_rollforward_mimic(11, True)
    nov = [sum(NOV_MOVES[c][i] for c in NOV_MOVES) for i in range(3)]
    assert abs(rf11["Cost Bfwd"] - 235000.0) < 0.005
    exp_cf = round(235000.0 + nov[0] - nov[1] - nov[2], 2)
    assert abs(rf11["Cost Cfwd"] - exp_cf) < 0.005
    inter("L5_assets_opening_total", round(float(opening), 2), "Opening cost 1 Jan 2026 across the four classes (IT Equipment stored as text)")
    inter("L5_rollforward_month10", rf10, "Month 10 roll-forward (Disposals and Write offs shown negative); matches client example")
    inter("L5_rollforward_month11", rf11, "Month 11 roll-forward after adding the November drop-in and Period = 11")
    inter("L5_movement_rows_before_nov", int(len(adf10[adf10['p'] <= 10])), "Class-period rows after Fill Down and removing subtotal rows (4 classes x 10 periods)")
    assert len(adf10) == 40
    ask("Q18", "Lab 5", "Build the roll-forward with a Period parameter. Check Month 10 (Bfwd 250,000.00 / Additions 10,000.00 / Disposals (25,000.00) / Write offs 0 / Cfwd 235,000.00). Then copy Asset_Movements_2026-11.xlsx from _Later into 05_Budget_Assets, set Period = 11 and refresh. (The 1 January opening cost per asset class is on the Opening_Cost sheet of Asset_Movements_2026.xlsx.) What is Cost Cfwd for Month 11?",
        rf11["Cost Cfwd"], "GBP",
        "Cost Bfwd(period) = opening cost at 1 Jan 2026 (238,000.00) + all Additions - Disposals - Write offs in earlier periods. Cost Cfwd = Bfwd + Additions - Disposals - Write offs for the period, all four asset classes combined. Month 11: Bfwd 235,000.00, Additions 28,700.00, Disposals (900.00), Write offs (600.00).",
        "Combine movement files, Fill Down Asset Class, remove subtotal rows, clean text numbers, filter Period < p and = p, add up.",
        month11=rf11)

    # ------------------------------------------------ LAB 6 checks
    cc = clean_cards()
    cl = cc["clean"]
    truth_c = CARDS.copy()
    truth_c["Date"] = truth_c["Date"].map(lambda d: d)
    assert len(cl) == len(truth_c) == 240, (len(cl), len(truth_c))
    m_ = cl.merge(truth_c, on="Bank Ref", suffixes=("", "_t"))
    assert len(m_) == 240
    assert (m_["Date"] == m_["Date_t"]).all()
    assert np.allclose(m_["Amount GBP"], m_["Amount GBP_t"], atol=0.0051), m_.loc[~np.isclose(m_["Amount GBP"], m_["Amount GBP_t"], atol=0.0051), ["Bank Ref", "Amount GBP", "Amount GBP_t"]]
    assert (m_["City"] == m_["City_t"]).all() and (m_["Last Four"] == m_["Last Four_t"]).all()
    inter("L6_rows_after_combine", cc["raw_rows"], "Data rows under the header in the 3 statements (incl. pending, month-end overlaps, blank-field rows)")
    inter("L6_rows_after_removing_blank_lines", cc["nonblank"], "After removing lines with no Bank Ref")
    inter("L6_pending_rows_removed", cc["pending_removed"], "Rows with Status Pending (5 per statement)")
    inter("L6_duplicate_bank_refs_removed", cc["dup_removed"], "Rows removed by dedupe on Bank Ref (6 from Jan on the Feb statement + 6 from Feb on the Mar statement)")
    inter("L6_rows_final", len(cl), "Rows after cleaning = truth cards (80 per month)")
    inter("L6_rows_amount_blank_or_zero_before_fix", int(sum(1 for s in cl["Stmt"] if pd.isna(s) or s == 0)), "Rows needing the blank/0 Amount rule (after dedupe)")
    inter("L6_non_gbp_rows", int((cl["Cur"] != "POUND STERLING").sum()), "Foreign-currency rows (USD 0.79, EUR 0.86)")
    inter("L6_refund_rows", int((cl["Amount GBP"] < 0).sum()), "Refund rows (negative Amount GBP)")
    inter("L6_total_net_spend_Q1", round(float(cl["Amount GBP"].sum()), 2), "Net Amount GBP all cardholders Jan-Mar after dedupe and FX")
    assert abs(cl["Amount GBP"].sum() - truth_c["Amount GBP"].sum()) < 0.02
    feb = cl[cl["Month"] == date(2026, 2, 1)]
    by_h = feb.groupby("Holder")["Amount GBP"].sum()
    # holder with most flagged rows in Feb
    fl = {}
    for _, row in truth_c[truth_c["m"] == 2].iterrows():
        if (2, row["Bank Ref"]) in flagged or row["Original Currency"] != "GBP" or row["Trans Type"] == "Refund":
            fl[HOLDER[row["Last Four"]]] = fl.get(HOLDER[row["Last Four"]], 0) + 1
    X = sorted(fl.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
    q19 = round(float(by_h[X]), 2)
    assert abs(q19 - float(truth_c[(truth_c["m"] == 2) & (truth_c["Last Four"].map(HOLDER) == X)]["Amount GBP"].sum())) < 0.011
    posted_feb = round(float(cl[(cl["Holder"] == X) & cl["Posted"].map(lambda d: d is not None and d.month == 2)]["Amount GBP"].sum()), 2)
    EMP_CARD = {v: k for k, v in HOLDER.items()}
    lost = cl[cl["Last Four"].str.startswith("0")]
    trap19 = {"by_posted_date_instead_of_transaction_date": posted_feb,
              "leading_zeros_not_restored": (0.0 if EMP_CARD[X].startswith("0") else q19),
              "note_leading_zeros": f"Cardholders with a leading-zero card: Priya Nair (0937) and Aisha Rahman (0184); {len(lost)} rows would fail to match. {X} has no leading zero, so this trap does not change Q19 (it changes their totals and Q20)."}
    ask("Q19", "Lab 6", f"Combine the three card statements, keep Posted rows only, dedupe on Bank Ref, and match to Employees on Last Four (restore leading zeros: 937 becomes 0937). Amount rule: if Amount is blank or 0, use Original Currency Amount x the rate in FX_Rates.xlsx (GBP = 1.00); an amount ending 'CR' is a refund and must be negative. What is the total GBP spend (net of refunds) for {X} using the transaction Date (not the Posted date) in February 2026?",
        q19, "GBP",
        "Sum of Amount GBP where Date is in Feb 2026 and Employees.Card Last Four (4 digits, leading zeros restored) maps to the cardholder. Amount GBP = statement Amount, or if blank/0: Original Currency Amount x rate (GBP 1.00, USD 0.79, EUR 0.86). 'CR' amounts are negative. Pending rows removed; duplicates by Bank Ref removed (keep first).",
        "lab6_cards.m; filter Employee Name and Month = Feb, sum Amount GBP.", cardholder=X, trap_answers=trap19)
    q1h = cl.groupby("Holder")["Amount GBP"].agg(["sum", "count"]).sort_values("sum", ascending=False)
    topn, topv, topc = q1h.index[0], round(float(q1h.iloc[0]["sum"]), 2), int(q1h.iloc[0]["count"])
    inter("L6_Q1_net_spend_by_cardholder", {k: round(float(v), 2) for k, v in q1h["sum"].items()}, "Net Q1 GBP per cardholder (name from Employees via Last Four)")
    inter("L6_Q1_transactions_by_cardholder", {k: int(v) for k, v in q1h["count"].items()}, "Transactions per cardholder, Q1, after dedupe/pending removal")
    ask("Q20", "Lab 6", "Create the function fnCardholder (lab6_fn_cardholder.m) and invoke it for the cardholder with the highest net Q1 spend. Who is it, how much did they spend (net GBP) and how many transactions? Hint: two cards start with 0 (0937, 0184); if Last Four lost its leading zero those cardholders will show missing rows.",
        f"{topn}; {topv:.2f}; {topc} transactions", "text",
        "Net Amount GBP and row count per cardholder over Jan-Mar 2026 after the Q19 cleaning (Posted only, dedupe on Bank Ref, FX/blank rule).",
        "fnCardholder(name) -> Total of Amount GBP and Table.RowCount; or Group By Employee Name on Transactions to confirm the maximum.",
        name=topn, amount=topv, transactions=topc)

    # ------------------------------------------------ outputs
    for fn in sorted(os.listdir(os.path.join(HERE, "m_src_A"))):
        shutil.copyfile(os.path.join(HERE, "m_src_A", fn), os.path.join(SOL, "m", fn))
    out = {"builder": "A", "questions": QUESTIONS, "intermediate": INTER,
           "meta": {"overtime_rule": "per employee per Week Ending: weekly hours above 37.5 are overtime; Labour Cost(line) = Hours x Rate + OT Hours(line) x Rate x (Multiplier-1); OT Hours(line) = (Weekly - 37.5) x Hours / Weekly",
                    "fx_rates": M.FX_RATES, "seed": M.SEED}}

    def default(o):
        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return float(o)
        if isinstance(o, (date, datetime)):
            return o.isoformat()
        raise TypeError(type(o))

    with open(os.path.join(SOL, "check_totals_A.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False, default=default)
        f.write("\n")
    with open(os.path.join(HERE, "mess_log_A.json"), "w", encoding="utf-8") as f:
        json.dump(MESS_LOG, f, indent=2, ensure_ascii=False)
        f.write("\n")
    by_lab: dict = {}
    for e in MESS_LOG:
        by_lab.setdefault(e["lab"], set()).add(e["mess_type"])
    print("mess types per lab:", {k: len(v) for k, v in sorted(by_lab.items())})
    for q in QUESTIONS:
        print(q["id"], "->", q["answer"])
    print("BUILD A OK")


if __name__ == "__main__":
    main()
