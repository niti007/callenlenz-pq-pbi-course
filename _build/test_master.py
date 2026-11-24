"""Self-tests for master.py. Run from repo root: _build/.venv/bin/python _build/test_master.py"""
import glob
import hashlib
import os
import subprocess
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
import master as m

ROOT = m.ROOT
PY = sys.executable
fails = []


def check(name, cond, info=""):
    print(("PASS " if cond else "FAIL ") + name + (f"  [{info}]" if info else ""))
    if not cond:
        fails.append(name)


def hashes():
    return {os.path.basename(p): hashlib.md5(open(p, "rb").read()).hexdigest() for p in sorted(glob.glob(m.TRUTH_DIR + "/*.csv"))}


# determinism: run script twice
subprocess.run([PY, os.path.join(ROOT, "_build/master.py")], cwd=ROOT, check=True, capture_output=True)
h1 = hashes()
subprocess.run([PY, os.path.join(ROOT, "_build/master.py")], cwd=ROOT, check=True, capture_output=True)
h2 = hashes()
check("determinism (byte-identical truth CSVs)", h1 == h2 and len(h1) == 9, f"{len(h1)} files")

P, E, S, C = m.projects(), m.employees(), m.suppliers(), m.customers()
parts, ts, po, xe = m.parts(), m.timesheets(), m.purchase_orders(), m.xero_invoices()
rc, bu, asx, cd = m.receipts(), m.budget(), m.assets(), m.cards()

check("row counts ref", (len(P), len(E), len(S), len(C), len(parts)) == (8, 25, 15, 15, 60))
for name, df, cols in [("Project Code", P, ["Project Code"]), ("Employee ID", E, ["Employee ID"]), ("Supplier ID", S, ["Supplier ID"]),
                       ("Customer ID", C, ["Customer ID"]), ("Part Number", parts, ["Part Number"]),
                       ("PO line", po, ["PO Number", "Part Number"]), ("Invoice+line", xe, ["Invoice Number", "Item Code"]),
                       ("Receipt Ref", rc, ["Ref"]), ("Card Bank Ref", cd, ["Bank Ref"])]:
    check(f"unique {name}", not df.duplicated(cols).any())

# referential integrity
check("RI timesheets project", set(ts["Project Code"]) <= set(P["Project Code"]))
check("RI timesheets employee", set(ts["Employee ID"]) <= set(E["Employee ID"]))
check("RI PO project/supplier/part", set(po["Project Code"]) <= set(P["Project Code"]) and set(po["Supplier ID"]) <= set(S["Supplier ID"])
      and set(po["Part Number"]) <= set(parts["Part Number"]))
check("RI parts supplier", set(parts["Default Supplier ID"]) <= set(S["Supplier ID"]))
check("RI xero", set(xe["Contact"]) <= set(S["Supplier ID"]) and set(xe["Item Code"]) <= set(parts["Part Number"]))
check("RI receipts customer", set(rc["Customer ID"]) <= set(C["Customer ID"]))
check("RI budget project", set(bu["Project Code"]) == set(P["Project Code"]))
check("RI project PM is employee", set(P["Project Manager"]) <= set(E["Employee Name"]))
check("RI employee manager", set(E["Manager"]) - {"Board of Directors"} <= set(E["Employee Name"]))
bad = 0
for p in glob.glob(m.TRUTH_DIR + "/*.csv"):
    t = open(p).read()
    bad += ("CL-P199" in t) + ("CL-P100" in t)
check("CL-P199/CL-P100 absent from truth", bad == 0)

# budget
bs = bu.groupby("Project Code")["Budget GBP"].sum().round(2)
ok = all(abs(bs[r["Project Code"]] - r["Budget 2026 GBP"]) < 0.005 for _, r in P.iterrows())
check("budget sums to Projects budget", ok and len(bu) == 96)

# assets
rf = m.assets_rollforward()
t10 = rf[rf["Month"] == 10][["Cost BF", "Additions", "Disposals", "Write Offs", "Cost CF"]].sum().round(2).to_dict()
check("assets Oct totals", t10 == {"Cost BF": 250000.0, "Additions": 10000.0, "Disposals": 25000.0, "Write Offs": 0.0, "Cost CF": 235000.0}, str(t10))
chain = True
for cls, g in rf.groupby("Asset Class"):
    g = g.sort_values("Month")
    chain &= all(abs(a - b) < 0.005 for a, b in zip(g["Cost BF"].iloc[1:], g["Cost CF"].iloc[:-1]))
    chain &= (g[["Cost BF", "Cost CF"]] >= 0).all().all()
check("assets BF=prev CF, no negative cost", chain)
check("assets opening BF Jan = opening", abs(rf[rf["Month"] == 1]["Cost BF"].sum() - m.assets_opening()["Opening Cost"].sum()) < 0.005)

# xero/po/coverage
mm = xe.merge(po[["PO Number", "Part Number", "Order Date", "Net Unit Price", "Qty"]], left_on=["Reference", "Item Code"],
              right_on=["PO Number", "Part Number"], how="left")
check("every invoice line ties to a PO line", mm["Order Date"].notna().all())
check("invoice date > PO date", (pd.to_datetime(mm["Invoice Date"]) > pd.to_datetime(mm["Order Date"])).all())
inv_p, po_p = set(xe["Item Code"]), set(po["Part Number"])
n_inv, n_po_only, n_none = len(inv_p), len(po_p - inv_p), len(set(parts["Part Number"]) - po_p)
check("coverage 40/10/10", (n_inv, n_po_only, n_none) == (40, 10, 10), f"invoiced={n_inv} PO-only={n_po_only} neither={n_none}")
check("PO ~200 lines, xero ~150", len(po) == 200 and len(xe) == 150, f"po={len(po)} xero={len(xe)}")
dr = po.groupby("Part Number")["Net Unit Price"].nunique()
check("price drift present", (dr[dr.index.isin(po_p)] > 1).mean() > 0.9)
check("PO dates in range", po["Order Date"].min() >= pd.Timestamp("2025-01-01").date() and po["Order Date"].max() <= pd.Timestamp("2026-06-30").date())
check("shared PR numbers exist", po.groupby("PR Number")["PO Number"].nunique().max() > 1 or po.groupby("PR Number").size().max() > 1)

# timesheets
wk = ts.groupby(["Week Ending", "Employee ID"])["Hours"].sum()
check("weekly hours 30-45", wk.between(30, 45).all())
check("some overtime weeks", (wk > 37.5).sum() > 50, f"{(wk > 37.5).sum()} emp-weeks > 37.5")
mon = pd.to_datetime(ts["Week Ending"]).dt.month.value_counts().sort_index()
check("monthly rows ~350-450", mon.between(340, 470).all(), str(mon.to_dict()))
check("week endings are Fridays", pd.to_datetime(ts["Week Ending"]).dt.weekday.eq(4).all() and ts["Week Ending"].nunique() == 26)
check("no dup (week,emp,proj,task)", not ts.duplicated(["Week Ending", "Employee ID", "Project Code", "Task"]).any())

# cards
ch = E[E["Card Last Four"] != ""]
check("10 cardholders", len(ch) == 10 and ch["Card Last Four"].str.len().eq(4).all())
check("a card starts with 0", ch["Card Last Four"].str.startswith("0").any())
check("cards 240 rows, 80/month", len(cd) == 240 and pd.to_datetime(cd["Date"]).dt.month.value_counts().eq(80).all())
check("~8 non-GBP", (cd["Original Currency"] != "GBP").sum() == 8)
check("cards last four map to cardholders (all 10 used)", set(cd["Last Four"]) == set(ch["Card Last Four"]))
check("receipts 120 in Q1", len(rc) == 120 and pd.to_datetime(rc["Date"]).between("2026-01-01", "2026-03-31").all() and rc["Ref"].min() == 10010)

# variants
check("variants 3-5 per ID", all(3 <= len(v) <= 5 for v in m.SUPPLIER_VARIANTS.values()) and all(3 <= len(v) <= 5 for v in m.CUSTOMER_VARIANTS.values())
      and set(m.SUPPLIER_VARIANTS) == set(S["Supplier ID"]) and set(m.CUSTOMER_VARIANTS) == set(C["Customer ID"]))

print("\nSUMMARY:", "ALL PASS" if not fails else f"FAILED: {fails}")
sys.exit(1 if fails else 0)
