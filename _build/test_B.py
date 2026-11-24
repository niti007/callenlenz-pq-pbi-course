#!/usr/bin/env python
"""Self-test for Builder B. Run: _build/.venv/bin/python _build/test_B.py  (from repo root)."""
import hashlib
import json
import os
import re
import subprocess
import sys

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = os.path.join(HERE, ".venv", "bin", "python")
sys.path.insert(0, HERE)
import build_B as B  # noqa: E402

FILES = ["data/01_Wasp_Stock/Wasp_Stock_Export.xlsx", "data/03_Valuation/Xero_Payable_Invoice_Detail.xlsx",
         "data/03_Valuation/Purchase_Log.xlsx", "data/03_Valuation/Supplier_Name_Map.xlsx",
         "data/04_Receipts/Bank_Receipts_Q1.xlsx", "data/04_Receipts/Customer_List.xlsx", "data/04_Receipts/Fuzzy_Transform_Table.xlsx"]
MS = ["lab1_wasp.m", "lab3_xero.m", "lab3_purchase_log.m", "lab3_valuation.m", "lab4_receipts_fuzzy.m", "lab4_customer_summary.m"]
fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)
        print("FAIL:", msg)


def digest():
    h = hashlib.sha256()
    for f in FILES:
        wb = load_workbook(os.path.join(ROOT, f))
        for ws in wb:
            h.update(ws.title.encode())
            for row in ws.iter_rows(values_only=True):
                h.update(repr(row).encode())
        h.update(repr((wb.properties.created, wb.properties.modified)).encode())
    for f in ["solutions/check_totals_B.json", "_build/mess_log_B.json", "_build/truth_B/valuation.csv"] + ["solutions/m/" + m for m in MS]:
        h.update(open(os.path.join(ROOT, f), "rb").read())
    return h.hexdigest()


def run_build():
    r = subprocess.run([PY, os.path.join(HERE, "build_B.py")], cwd=ROOT, capture_output=True, text=True)
    check(r.returncode == 0, "build_B.py failed: " + r.stderr[-800:])


# ---- 1. determinism + truth assertions (build_B asserts internally)
run_build()
d1 = digest()
run_build()
d2 = digest()
check(d1 == d2, "build is not deterministic (content digest differs between runs)")

# ---- 2. mess log verification
STR = lambda v: isinstance(v, str)
PRED = {
    "title_row": lambda vs: any(STR(v) for v in vs),
    "part_number_variant": lambda vs: any(STR(v) and B.norm_part(v) and B.norm_part(v) != v for v in vs),
    "gbp_text": lambda vs: any(STR(v) and "£" in v for v in vs),
    "na_value": lambda vs: any(v == "N/A" for v in vs),
    "ref_error": lambda vs: any(v == "#REF!" for v in vs),
    "line_break": lambda vs: any(STR(v) and "\n" in v for v in vs),
    "nbsp": lambda vs: any(STR(v) and "\u00a0" in v for v in vs),
    "trailing_space": lambda vs: any(STR(v) and v != v.strip() for v in vs),
    "double_space": lambda vs: any(STR(v) and "  " in v for v in vs),
    "blank_category_fill_down": lambda vs: all(v is None for v in vs),
    "junk_row": lambda vs: any(v in ("TEST-ITEM", "ZZ SAMPLE") for v in vs),
    "subtotal_row": lambda vs: vs[0] == "Total",
    "blank_row": lambda vs: all(v is None for v in vs),
    "repeated_header": lambda vs: vs[0] == "Item Number",
    "blank_header_column": lambda vs: vs[0] is None,
    "line_break_header": lambda vs: STR(vs[0]) and "\n" in vs[0],
    "trailing_space_header": lambda vs: STR(vs[0]) and vs[0].endswith(" "),
    "blank_po_number": lambda vs: vs[0] is None,
    "cancelled_line": lambda vs: STR(vs[0]) and "CANCELLED" in vs[0].upper(),
    "composite_notes": lambda vs: STR(vs[0]) and " / Jira " in vs[0],
    "credit_note": lambda vs: vs[2] == "Payable Credit Note" and vs[6] < 0,
    "near_duplicate_row": None,
    "duplicate_row": None,
    "supplier_variant": lambda vs: STR(vs[0]) and vs[0].strip() not in STD,
    "text_date": lambda vs: any(STR(v) and re.match(r"^\d\d[/-]", v) for v in vs),
    "project_code_variant": lambda vs: STR(vs[0]) and B.norm_proj(vs[0]) != vs[0],
    "case_or_space_variant": lambda vs: STR(vs[0]) and (vs[0] != vs[0].strip() or vs[0] != vs[0].upper()),
    "bracket_negative": lambda vs: any(STR(v) and v.startswith("(") for v in vs),
    "case_variant": lambda vs: STR(vs[0]) and vs[0] == vs[0].lower() and vs[0] != vs[0].upper(),
    "noise_prefix": lambda vs: STR(vs[0]) and vs[0].lower().startswith(("recd from", "bacs")),
    "noise_suffix": lambda vs: STR(vs[0]) and " - INV" in vs[0],
    "refund": lambda vs: STR(vs[3]) and vs[3].startswith("("),
    "unmatched_payer": lambda vs: STR(vs[0]) and (vs[0] == "CASH DEPOSIT" or "J SMITH" in vs[0]),
    "transform_table": lambda vs: all(v is not None for v in vs),
}
import master as M  # noqa: E402
STD = {n for n in M.suppliers()["Standard Supplier Name"]}
STD = {x for x in STD}
SUPV = {v for vs in M.SUPPLIER_VARIANTS.values() for v in vs}
STD_ALL = STD  # supplier_variant pred: not a standard name (variants may also include trailing-space noise)

log = json.load(open(os.path.join(ROOT, "_build", "mess_log_B.json")))
check(len(log) > 100, "mess log too small")
wbs = {}
for it in log:
    for k in ("file", "sheet", "location", "mess_type", "description", "teaching_fix", "lab"):
        check(k in it and it[k], f"log item missing {k}: {it}")
    f = it["file"]
    wbs.setdefault(f, load_workbook(os.path.join(ROOT, f)))
    ws = wbs[f][it["sheet"]]
    c1, r1, c2, r2 = range_boundaries(it["location"])
    vals = [ws.cell(row=r, column=c).value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
    t = it["mess_type"]
    check(t in PRED, f"no predicate for mess_type {t}")
    p = PRED.get(t)
    if t == "near_duplicate_row":
        col_e = [ws.cell(row=r, column=5).value for r in range(1, r1)]
        key = B.norm_part(vals[4])
        check(any(B.norm_part(x) == key for x in col_e), f"near dup has no earlier original at {it['location']}")
    elif t == "supplier_variant":
        check(STR(vals[0]) and B.clean_text(vals[0]) not in STD, f"supplier_variant not a variant at {it['location']}")
    elif p:
        check(bool(p(vals)), f"mess '{t}' not verifiable at {f}!{it['sheet']}!{it['location']} -> {vals[:4]}")
    if t == "duplicate_row" and f.endswith("Wasp_Stock_Export.xlsx"):
        key = B.norm_part(vals[0])
        check(any(B.norm_part(ws.cell(row=r, column=1).value) == key for r in range(4, r1)), f"wasp dup has no earlier original {it['location']}")
    if t == "duplicate_row" and f.endswith("Bank_Receipts_Q1.xlsx"):
        check(any(ws.cell(row=r, column=2).value == vals[1] for r in range(2, r1)), "receipt dup lacks earlier original")

# distinct mess types per lab
for lab in ("Lab 1", "Lab 3", "Lab 4"):
    n = len({i["mess_type"] for i in log if i["lab"] == lab})
    check(n >= 6, f"{lab} has only {n} distinct mess types")

# ---- 3. questions answered
ct = json.load(open(os.path.join(ROOT, "solutions", "check_totals_B.json")))
for q in ("Q3", "Q4", "Q5", "Q10", "Q11", "Q12", "Q13", "Q14", "Q15"):
    e = ct.get(q)
    check(e is not None, f"{q} missing")
    if e:
        for k in ("id", "lab", "question", "answer", "unit", "definition", "how"):
            check(k in e and e[k] not in (None, ""), f"{q} missing {k}")
        check(e["unit"] in ("GBP", "rows", "parts", "text"), f"{q} bad unit")
for k in ("L1_rows_after_promote", "L1_rows_after_dedupe", "L3_parts_by_source", "L3_valuation_table", "L4_rows_after_dedupe"):
    check(k in ct, f"checkpoint {k} missing")
check(ct["Q3"]["answer"] == 60, "Q3 != 60")
check(ct["L3_valuation_table"]["value"]["count"] == 60, "valuation count")

# ---- 4. layout checks
ws = wbs.get("data/03_Valuation/Purchase_Log.xlsx") or load_workbook(os.path.join(ROOT, FILES[2]))
ws = ws["Purchase Log"]
hdr = [c.value for c in ws[1]]
check(len(hdr) == 48 and sum(h is None for h in hdr) == 5, "PL must have 48 cols with 5 blank headers")
check(all(ws.cell(row=r, column=c).value is None for c in [i + 1 for i, h in enumerate(hdr) if h is None] for r in range(2, ws.max_row + 1)), "blank PL columns must be empty")
wasp = load_workbook(os.path.join(ROOT, FILES[0]))["Stock"]
check(wasp.max_column == 24 and [c.value for c in wasp[3]][0] == "Item Number", "wasp layout")

# ---- 5. M files
def strip_comments(s):
    out = []
    for line in s.split("\n"):
        i, inq = 0, False
        res = ""
        while i < len(line):
            ch = line[i]
            if ch == '"':
                inq = not inq
            if not inq and line[i:i + 2] == "//":
                break
            res += ch
            i += 1
        out.append(res)
    return "\n".join(out)


for m in MS:
    s = open(os.path.join(ROOT, "solutions", "m", m), encoding="utf-8").read()
    c = strip_comments(s)
    check(c.count('"') % 2 == 0, f"{m}: unbalanced quotes")
    for a, b in ("()", "[]", "{}"):
        check(c.count(a) == c.count(b), f"{m}: unbalanced {a}{b}")
    toks = re.findall(r'\b(let|in)\b', re.sub(r'"[^"]*"', '""', c))
    check(toks.count("let") == toks.count("in") and toks[0] == "let" and toks[-1] == "in", f"{m}: let/in mismatch {toks}")
    check(c.rstrip().split("\n")[-1].strip() != "in" and not c.rstrip().endswith(","), f"{m}: bad ending")
    if m != "lab3_valuation.m" and m != "lab4_customer_summary.m":
        check("FolderPath" in c and "File.Contents" in c, f"{m}: FolderPath/File.Contents missing")
    check("C:\\" not in c, f"{m}: hard-coded path")
    check(s.lstrip().startswith("//"), f"{m}: no leading comment")
    # step names unique at top level: lines like '    Name = ' at indent 4
    names = re.findall(r"^    ([A-Za-z_][A-Za-z0-9_]*) =", c, re.M)
    check(len(names) == len(set(names)), f"{m}: duplicate step names")
    # final 'in' identifier must exist as a step
    final = c.rstrip().split("\n")[-1].strip()
    check(final in names, f"{m}: final result '{final}' is not a defined step")
fz = open(os.path.join(ROOT, "solutions/m/lab4_receipts_fuzzy.m")).read()
for k in ("Table.FuzzyNestedJoin", "Threshold = 0.8", "IgnoreCase = true", "IgnoreSpace = true", "TransformationTable"):
    check(k in fz, f"fuzzy M missing {k}")

print("\nTEST SUMMARY:", "ALL PASSED" if not fails else f"{len(fails)} FAILURES", f"({len(log)} mess-log items verified)")
sys.exit(1 if fails else 0)
