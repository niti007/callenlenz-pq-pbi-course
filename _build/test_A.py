#!/usr/bin/env python
"""Self-test for Builder A. Run from repo root: _build/.venv/bin/python _build/test_A.py"""
import hashlib, json, os, re, subprocess, sys
from datetime import date
import numpy as np
import pandas as pd
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import master as M

ROOT = M.ROOT
PY = sys.executable
FAILS = []


def check(cond, msg):
    if not cond:
        FAILS.append(msg)
        print("FAIL:", msg)


def out_files():
    res = []
    for d in ("02_Timesheets", "05_Budget_Assets", "06_Cards"):
        for dp, _, fs in os.walk(os.path.join(ROOT, "data", d)):
            res += [os.path.join(dp, f) for f in fs]
    res += [os.path.join(ROOT, "solutions", "check_totals_A.json"), os.path.join(HERE, "mess_log_A.json")]
    mdir = os.path.join(ROOT, "solutions", "m")
    res += [os.path.join(mdir, f) for f in os.listdir(os.path.join(HERE, "m_src_A"))]
    return sorted(res)


def hashes():
    return {f: hashlib.sha256(open(f, "rb").read()).hexdigest() for f in out_files()}


def run_build():
    r = subprocess.run([PY, os.path.join(HERE, "build_A.py")], cwd=ROOT, capture_output=True, text=True)
    check(r.returncode == 0, "build_A.py failed: " + r.stderr[-800:])
    return r.stdout


# 1. determinism
run_build()
h1 = hashes()
run_build()
h2 = hashes()
check(h1 == h2, "outputs differ between two runs: " + ", ".join(k for k in h1 if h1[k] != h2.get(k)))
print(f"determinism: {len(h1)} files hashed twice")

ck = json.load(open(os.path.join(ROOT, "solutions", "check_totals_A.json")))
qs = {q["id"]: q for q in ck["questions"]}

# 2. questions
for qid in ["Q6", "Q7", "Q8", "Q9", "Q16", "Q17", "Q18", "Q19", "Q20"]:
    check(qid in qs, f"missing {qid}")
    if qid in qs:
        q = qs[qid]
        for k in ("lab", "question", "answer", "unit", "definition", "how"):
            check(k in q and q[k] not in (None, ""), f"{qid} missing {k}")
check(len(ck["intermediate"]) >= 30, "too few intermediate checkpoints")

# 3. independent recompute from truth
ts = M.timesheets()
ts["m"] = [d.month for d in ts["Week Ending"]]
emp = M.employees().set_index("Employee ID")
rt = M.labour_rates().set_index("Grade")
ts["rate"] = ts["Employee ID"].map(emp["Grade"]).map(rt["Standard Hourly Rate GBP"])
ts["wk"] = ts.groupby(["Employee ID", "Week Ending"])["Hours"].transform("sum")
ts["cost"] = [h * r + max(0, w - 37.5) * (h / w) * r * 0.5 for h, r, w in zip(ts["Hours"], ts["rate"], ts["wk"])]
q6 = round(ts[(ts["m"] <= 3) & (ts["Project Code"] == "CL-P103")]["cost"].sum(), 2)
check(abs(qs["Q6"]["answer"] - q6) < 0.011, f"Q6 {qs['Q6']['answer']} vs truth {q6}")
check(abs(qs["Q9"]["answer"] - ts["Hours"].sum()) < 1e-6, "Q9 vs truth")
b = M.budget()
b["mi"] = [d.month for d in b["Month"]]
q16 = b[(b["Project Code"] == "CL-P105") & (b["mi"] <= 6)]["Budget GBP"].sum()
check(abs(qs["Q16"]["answer"] - q16) < 0.005, "Q16 vs truth")
q17 = ts[(ts["Project Code"] == "CL-P102") & (ts["m"] == 3)]["Hours"].sum()
check(abs(qs["Q17"]["answer"] - q17) < 1e-6, "Q17 vs truth")
check(abs(qs["Q18"]["answer"] - 262200.0) < 0.005, "Q18 value")
c = M.cards()
c["m"] = [d.month for d in c["Date"]]
holder = emp[emp["Card Last Four"] != ""].set_index("Card Last Four")["Employee Name"]
c["h"] = c["Last Four"].map(holder)
x = qs["Q19"]["cardholder"]
check(abs(qs["Q19"]["answer"] - c[(c["m"] == 2) & (c["h"] == x)]["Amount GBP"].sum()) < 0.011, "Q19 vs truth")
g = c.groupby("h")["Amount GBP"].agg(["sum", "count"]).sort_values("sum", ascending=False)
check(qs["Q20"]["name"] == g.index[0] and abs(qs["Q20"]["amount"] - g.iloc[0]["sum"]) < 0.011 and qs["Q20"]["transactions"] == g.iloc[0]["count"], "Q20 vs truth")
check(ck["intermediate"]["L2_control_off_month"]["value"] == "Feb", "control off month")

# 4. control totals: only Feb off
wb = load_workbook(os.path.join(ROOT, "data/02_Timesheets/Control_Totals.xlsx"), data_only=True)
rows = [r for r in wb["Control"].iter_rows(min_row=5, values_only=True) if r[0]]
check(len(rows) == 6, "control rows")
for r in rows:
    tru = ts[ts["m"] == r[0].month]["Hours"].sum()
    check(abs(r[1] - tru) < 1e-9, f"control {r[0]} != truth")

# 5. mess log verification
log = json.load(open(os.path.join(HERE, "mess_log_A.json")))
need = {"file", "sheet", "location", "mess_type", "description", "teaching_fix", "lab"}
cache = {}
for e in log:
    check(need <= set(e), f"mess_log keys {e.get('mess_type')}")
    path = os.path.join(ROOT, e["file"])
    check(os.path.exists(path), f"mess_log file missing {e['file']}")
    if not os.path.exists(path):
        continue
    if path.endswith(".csv"):
        if path not in cache:
            cache[path] = open(path, newline="", encoding="utf-8").read().split("\r\n")
        m = re.match(r"lines? (\d+)", e["location"])
        check(bool(m), f"bad csv location {e['location']}")
        if m:
            n = int(m.group(1))
            check(cache[path][n - 1] == e["evidence"], f"{e['file']} line {n} != evidence for {e['mess_type']}")
    else:
        if path not in cache:
            cache[path] = load_workbook(path, data_only=True)
        wbk = cache[path]
        check(e["sheet"] in wbk.sheetnames, f"sheet {e['sheet']} missing in {e['file']}")
        if e["sheet"] in wbk.sheetnames:
            v = wbk[e["sheet"]][e["location"]].value
            got = "<blank>" if v is None else str(v)
            check(got == e["evidence"], f"{e['file']}!{e['sheet']}!{e['location']} = {got!r} != {e['evidence']!r} ({e['mess_type']})")
by_lab = {}
for e in log:
    by_lab.setdefault(e["lab"], set()).add(e["mess_type"])
for lab in ("Lab 2", "Lab 5", "Lab 6"):
    check(len(by_lab.get(lab, ())) >= 6, f"{lab} has <6 mess types")
print("mess types per lab:", {k: len(v) for k, v in by_lab.items()}, "| log entries:", len(log))

# 6. files exist
for f in ["02_Timesheets/Timesheet_2026-0%d.csv" % i for i in range(1, 6)] + ["02_Timesheets/_Later/Timesheet_2026-06.csv", "02_Timesheets/Control_Totals.xlsx",
          "05_Budget_Assets/Project_Budget_2026.xlsx", "05_Budget_Assets/Asset_Movements_2026.xlsx", "05_Budget_Assets/_Later/Asset_Movements_2026-11.xlsx",
          "06_Cards/Card_Statement_2026-01.csv", "06_Cards/Card_Statement_2026-02.csv", "06_Cards/Card_Statement_2026-03.csv",
          "06_Cards/FX_Rates.xlsx", "06_Cards/Employee_Lookup.xlsx", "06_Cards/Expense_Template.xlsx"]:
    check(os.path.exists(os.path.join(ROOT, "data", f)), f"missing {f}")
check("Roll_Forward_Template" in load_workbook(os.path.join(ROOT, "data/05_Budget_Assets/Asset_Movements_2026.xlsx")).sheetnames, "Roll_Forward_Template sheet")


# 7. M files
def strip_m(s):
    out, i, n, instr = [], 0, len(s), False
    while i < n:
        ch = s[i]
        if instr:
            if ch == '"':
                if i + 1 < n and s[i + 1] == '"':
                    i += 2
                    continue
                instr = False
            i += 1
            continue
        if ch == '"':
            instr = True
            out.append('""')
            i += 1
        elif s.startswith("//", i):
            while i < n and s[i] != "\n":
                i += 1
        elif s.startswith("/*", i):
            i = s.index("*/", i) + 2
        else:
            out.append(ch)
            i += 1
    return "".join(out), instr


M_FILES = ["lab2_timesheets.m", "lab2_project_month_summary.m", "lab2_control_checks.m", "lab5_unpivot_budget.m", "lab5_pivot_hours.m", "lab5_rollforward.m", "lab6_cards.m", "lab6_fn_cardholder.m"]
for f in M_FILES:
    p = os.path.join(ROOT, "solutions", "m", f)
    check(os.path.exists(p), f"missing {f}")
    s = open(p).read()
    code, open_str = strip_m(s)
    check(not open_str, f"{f}: unterminated string")
    st = []
    pairs = {")": "(", "]": "[", "}": "{"}
    ok = True
    for ch in code:
        if ch in "([{":
            st.append(ch)
        elif ch in ")]}":
            if not st or st.pop() != pairs[ch]:
                ok = False
    check(ok and not st, f"{f}: unbalanced brackets")
    words = re.findall(r"(?<![\w.#])(let|in)(?![\w.])", code)
    check(words.count("let") == words.count("in") and words.count("let") >= 1, f"{f}: let/in mismatch {words.count('let')}/{words.count('in')}")
    # step names referenced in final 'in' exist
    m = re.search(r"\bin\s*\n\s*(#\"[^\"]*\"|\w+)\s*$", s.split("//")[0] if False else re.sub(r"//[^\n]*", "", s).strip())
    check(bool(m), f"{f}: cannot find final expression")
    if m:
        nm = m.group(1)
        check(re.search(re.escape(nm) + r"\s*=", s) is not None, f"{f}: final step {nm} not defined")
    # no hard-coded drive paths
    check(not re.search(r"[A-Z]:\\\\?Users", s), f"{f}: hard-coded path")
check(os.path.exists(os.path.join(ROOT, "solutions", "m", "lab6_power_automate.md")), "missing power automate md")
check(open(os.path.join(ROOT, "solutions/m/lab6_fn_cardholder.m")).read().split("//")[-1].lstrip().startswith("(Cardholder as text) as table =>") or "(Cardholder as text) as table =>" in open(os.path.join(ROOT, "solutions/m/lab6_fn_cardholder.m")).read(), "fn signature")
pa = open(os.path.join(ROOT, "solutions/m/lab6_power_automate.md")).read()
for kw in ("List rows present in a table", "Apply to each", "Filter array", "Create HTML table", "Send an email (V2)", "Show Report Filter Pages", "mail merge", "item()?['Last Four']"):
    check(kw.lower() in pa.lower(), f"PA md missing {kw}")

print("FAILURES:", len(FAILS))
sys.exit(1 if FAILS else 0)
