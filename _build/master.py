"""Callen-Lenz course data: master data and CLEAN truth tables (Phase 1).

Importable without side effects. Files are only written by write_all() / __main__.

Public API (all return fresh DataFrame copies, deterministic, cached internally):
    projects(), employees(), labour_rates(), suppliers(), customers()
    parts(), timesheets(), purchase_orders(), xero_invoices(), receipts(),
    budget(), assets(), assets_opening(), assets_rollforward(), cards()
    SUPPLIER_VARIANTS, CUSTOMER_VARIANTS, FX_RATES, SEED
    write_all()
"""
from __future__ import annotations

import os
from datetime import date, timedelta
from functools import lru_cache

import numpy as np
import pandas as pd

SEED = 42
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
REF_DIR = os.path.join(ROOT, "data", "00_Reference")
TRUTH_DIR = os.path.join(ROOT, "_build", "truth")

FX_RATES = {"GBP": 1.00, "USD": 0.79, "EUR": 0.86}  # GBP per 1 unit of currency


def _rng(k: int) -> np.random.Generator:
    return np.random.default_rng([SEED, k])


def _d(x) -> date:
    return x.date() if isinstance(x, pd.Timestamp) else x


# ----------------------------------------------------------------------------
# Reference data
# ----------------------------------------------------------------------------
_EMP = [
    # id, name, dept, grade, manager
    ("E001", "Margaret Ellison", "SLT", "G5", "Board of Directors"),
    ("E002", "Rajesh Patel", "SLT", "G5", "Margaret Ellison"),
    ("E003", "Hannah Brooks", "FINANCE", "G4", "Margaret Ellison"),
    ("E004", "Tom Whitaker", "FINANCE", "G2", "Hannah Brooks"),
    ("E005", "Priya Nair", "PROCUREMENT", "G3", "Hannah Brooks"),
    ("E006", "Callum Reid", "PROCUREMENT", "G2", "Priya Nair"),
    ("E007", "Oliver Hartley", "ENGINEERING", "G4", "Rajesh Patel"),
    ("E008", "Sophie Marsh", "ENGINEERING", "G3", "Oliver Hartley"),
    ("E009", "Daniel Okafor", "ENGINEERING", "G3", "Oliver Hartley"),
    ("E010", "Emily Chen", "ENGINEERING", "G2", "Oliver Hartley"),
    ("E011", "Liam Garside", "ENGINEERING", "G2", "Oliver Hartley"),
    ("E012", "Aisha Rahman", "ENGINEERING", "G3", "Oliver Hartley"),
    ("E013", "Ben Thornton", "ENGINEERING", "G1", "Oliver Hartley"),
    ("E014", "Chloe Davies", "ENGINEERING", "G2", "Oliver Hartley"),
    ("E015", "Steve Mallory", "PRODUCTION", "G3", "Rajesh Patel"),
    ("E016", "Karen Fox", "PRODUCTION", "G2", "Steve Mallory"),
    ("E017", "Dean Hollis", "PRODUCTION", "G1", "Steve Mallory"),
    ("E018", "Jamie Walsh", "PRODUCTION", "G1", "Steve Mallory"),
    ("E019", "Nadia Kowalski", "PRODUCTION", "G2", "Steve Mallory"),
    ("E020", "Angus McBride", "MARITIME", "G4", "Rajesh Patel"),
    ("E021", "Fiona Sutherland", "MARITIME", "G3", "Angus McBride"),
    ("E022", "Marcus Lowe", "MARITIME", "G2", "Angus McBride"),
    ("E023", "Ellie Hargreaves", "MARITIME", "G2", "Angus McBride"),
    ("E024", "Paul Danvers", "SUPPORT", "G3", "Rajesh Patel"),
    ("E025", "Yasmin Qureshi", "SUPPORT", "G2", "Paul Danvers"),
]
# 10 cardholders -> card last four (text; two start with 0)
_CARDS = {
    "E002": "4821", "E005": "0937", "E006": "7316", "E007": "5502", "E012": "0184",
    "E015": "3349", "E020": "6275", "E021": "8110", "E022": "2693", "E024": "9047",
}
# 20 booking employees: all ENGINEERING, PRODUCTION, MARITIME, SUPPORT + E006
_BOOKERS = [e[0] for e in _EMP if e[2] in ("ENGINEERING", "PRODUCTION", "MARITIME", "SUPPORT")] + ["E006"]
_BOOKERS = sorted(_BOOKERS)

_CUST = [
    # id, name, abbr1, abbr2
    ("C01", "Flintwood Marine Ltd", "FLINTWD", "MARINE"),
    ("C02", "Rabbitson Engineering", "RABBITSON", "ENG"),
    ("C03", "Duckworth Defence Systems", "DUCKWORTH", "DEF SYS"),
    ("C04", "Pennywhistle Offshore Ltd", "PENNYWHSTL", "OFFSHORE"),
    ("C05", "Bramblecombe Aerospace", "BRAMBLCMB", "AERO"),
    ("C06", "Thistlewood Naval Ltd", "THISTLWD", "NAVAL"),
    ("C07", "Haddockson Fisheries Tech", "HADDOCKSN", "FISH TECH"),
    ("C08", "Cobblestone Marine Ltd", "COBBLSTN", "MARINE"),
    ("C09", "Wigglesworth Robotics", "WIGGLSWTH", "ROBOTICS"),
    ("C10", "Gooseberry Subsea Ltd", "GOOSEBRY", "SUBSEA"),
    ("C11", "Mulberry Harbour Systems", "MULBERRY", "HARBOUR SYS"),
    ("C12", "Tumbleweed Dynamics", "TUMBLWD", "DYN"),
    ("C13", "Ironbridge Maritime Ltd", "IRONBRDG", "MARITIME"),
    ("C14", "Peppercorn Defence Ltd", "PEPPRCRN", "DEFENCE"),
    ("C15", "Saltmarsh Propulsion Ltd", "SALTMRSH", "PROP"),
]

_SUPP = [
    ("S01", "Apex Fasteners Ltd", "Fasteners"),
    ("S02", "Boltwright Supplies Limited", "Fasteners"),
    ("S03", "Kestrel Fixings Ltd", "Fasteners"),
    ("S04", "Northgate Composites Limited", "Composites"),
    ("S05", "Tideline Composites Ltd", "Composites"),
    ("S06", "Solent Marine Electronics Ltd", "Electronics"),
    ("S07", "Meridian Circuits Limited", "Electronics"),
    ("S08", "Harbourlight Sensors Ltd", "Electronics"),
    ("S09", "Precision Turned Components Ltd", "Machined Parts"),
    ("S10", "Ironside Machining Limited", "Machined Parts"),
    ("S11", "Wessex CNC Ltd", "Machined Parts"),
    ("S12", "Copperfield Cables Ltd", "Cables"),
    ("S13", "Sterling Wire & Cable Limited", "Cables"),
    ("S14", "Brightway Industrial Supplies Ltd", "Consumables"),
    ("S15", "Marlow Workshop Supplies Limited", "Consumables"),
]

_PROJ = [
    ("CL-P101", "Maritime Sensor Pod", "C01", "Oliver Hartley", "2025-01-13", "Active", 640000),
    ("CL-P102", "Autonomous Survey Vessel", "C10", "Angus McBride", "2024-09-02", "Active", 920000),
    ("CL-P103", "Subsea Cable Winch Assembly", "C04", "Daniel Okafor", "2025-03-17", "Active", 480000),
    ("CL-P104", "Composite Radome Housing", "C03", "Sophie Marsh", "2025-05-05", "Active", 560000),
    ("CL-P105", "Naval Mast Camera Gimbal", "C06", "Aisha Rahman", "2025-06-23", "Active", 380000),
    ("CL-P106", "Thruster Control Cabinet", "C15", "Steve Mallory", "2025-08-11", "Active", 275000),
    ("CL-P107", "Coastal Patrol RIB Console", "C14", "Fiona Sutherland", "2025-10-06", "Active", 330000),
    ("CL-P108", "Legacy Buoy Telemetry Refit", "C11", "Rajesh Patel", "2024-04-15", "Closed", 120000),
]
_CLOSED = "CL-P108"
_CLOSED_AFTER = date(2026, 3, 31)  # no bookings / orders to closed project after this


@lru_cache(maxsize=None)
def _projects() -> pd.DataFrame:
    cn = {c[0]: c[1] for c in _CUST}
    return pd.DataFrame(
        [(p[0], p[1], cn[p[2]], p[3], pd.Timestamp(p[4]).date(), p[5], float(p[6])) for p in _PROJ],
        columns=["Project Code", "Project Name", "Customer", "Project Manager", "Start Date", "Status", "Budget 2026 GBP"],
    )


@lru_cache(maxsize=None)
def _employees() -> pd.DataFrame:
    rows = []
    for eid, name, dept, grade, mgr in _EMP:
        f, l = name.lower().split(" ", 1)
        email = f"{f}.{l.replace(' ', '')}@callenlenz-demo.example"
        rows.append((eid, name, dept, grade, mgr, email, _CARDS.get(eid, "")))
    return pd.DataFrame(rows, columns=["Employee ID", "Employee Name", "Department", "Grade", "Manager", "Email", "Card Last Four"])


@lru_cache(maxsize=None)
def _labour_rates() -> pd.DataFrame:
    return pd.DataFrame(
        {"Grade": ["G1", "G2", "G3", "G4", "G5"],
         "Standard Hourly Rate GBP": [22.50, 30.00, 40.00, 52.50, 65.00],
         "Overtime Multiplier": [1.5] * 5}
    )


@lru_cache(maxsize=None)
def _suppliers() -> pd.DataFrame:
    return pd.DataFrame(_SUPP, columns=["Supplier ID", "Standard Supplier Name", "Category"])


@lru_cache(maxsize=None)
def _customers() -> pd.DataFrame:
    return pd.DataFrame([(c[0], c[1]) for c in _CUST], columns=["Customer ID", "Customer Name"])


# ----------------------------------------------------------------------------
# Variants
# ----------------------------------------------------------------------------
def _swap(word: str, i: int) -> str:
    if len(word) <= i + 1:
        return word
    return word[:i] + word[i + 1] + word[i] + word[i + 2:]


def _supplier_variants() -> dict:
    out = {}
    for sid, name, _ in _SUPP:
        words = name.split(" ")
        suffix = words[-1]  # Ltd or Limited
        base = " ".join(words[:-1])
        other = "Limited" if suffix == "Ltd" else "Ltd"
        bw = base.split(" ")
        typo = " ".join([_swap(bw[0], 2)] + bw[1:]) if len(bw[0]) > 4 else " ".join([bw[0]] + [_swap(bw[1], 2)] + bw[2:])
        cand = [
            f"{base} {other}",
            f"{base.upper()} {suffix.upper()}.",
            f"{base} Ltd.",
            f"{typo} {suffix}",
            " ".join(bw[:2]) if len(bw) > 2 else base,
        ]
        seen, vs = {name}, []
        for c in cand:
            if c not in seen:
                seen.add(c)
                vs.append(c)
        out[sid] = vs
    return out


def _customer_variants() -> dict:
    out = {}
    for i, (cid, name, a1, a2) in enumerate(_CUST):
        words = name.split(" ")
        first = words[0]
        full_up = name.upper()
        typo = " ".join([_swap(first, 2)] + words[1:])
        cand = [
            f"Recd from {a1} {a2}",
            f"BACS {a1} {a2}",
            f"{first} {a2.title()} - INV {2200 + 7 * i + 3}",
            full_up,
            typo,
        ]
        seen, vs = {name}, []
        for c in cand:
            if c not in seen:
                seen.add(c)
                vs.append(c)
        out[cid] = vs
    return out


SUPPLIER_VARIANTS = _supplier_variants()
CUSTOMER_VARIANTS = _customer_variants()


# ----------------------------------------------------------------------------
# Parts
# ----------------------------------------------------------------------------
_PART_DESC = {
    "Fasteners": ["M6 Hex Bolt - SS316", "M8 Socket Cap Screw - A4", "M10 Nyloc Nut - SS316", "M5 Countersunk Screw - A2",
                  "M12 Flat Washer - SS316", "M6 Eyebolt - Marine Grade", "Self-Tapping Screw 4.8mm - A4",
                  "Rivet Nut M6 - Aluminium", "Stud Bolt M10x60 - SS316", "Cable Gland Locknut M20"],
    "Composites": ["Carbon Fibre Sheet 2mm 1m x 1m", "Glass Fibre Mat 450gsm", "Epoxy Resin System 5kg", "Kevlar Cloth 200gsm",
                   "Foam Core Panel 10mm", "Gelcoat White 2.5kg", "Peel Ply Roll 50m", "Vacuum Bagging Film 25m",
                   "Honeycomb Panel 15mm", "Radome Laminate Blank"],
    "Electronics": ["Marine GPS Receiver Module", "IMU Sensor Board 9-axis", "24V DC-DC Converter", "Ethernet Switch 8-port IP67",
                    "Sonar Transducer 200kHz", "Relay Module 4-channel", "LiPo Battery Pack 14.8V", "Waterproof Connector 8-pin",
                    "Camera Module 4K Marine", "PCB Control Board Rev C"],
    "Machined Parts": ["Aluminium Mounting Bracket 6082", "Titanium Shaft 25mm", "Stainless Flange DN50", "Delrin Bearing Housing",
                       "Brass Bushing 20mm", "Gimbal Ring Machined 316", "Thruster Mount Plate", "Winch Drum Core 6082",
                       "Pod End Cap Anodised", "Gearbox Spacer Set"],
    "Cables": ["Tinned Copper Cable 2.5mm 3-core", "Cat6 Armoured Cable", "Coax RG58 Marine", "Subsea Umbilical 12-core",
               "Shielded Twisted Pair 4-pair", "Power Cable 16mm Black", "Fibre Optic Patch Lead LC-LC",
               "Multicore Control Cable 7-core", "Battery Cable 35mm Red", "Sensor Cable M12 5m"],
    "Consumables": ["Loctite 243 50ml", "Marine Sealant 300ml", "Abrasive Discs 125mm Pack", "Cable Ties 300mm Pack 100",
                    "Heat Shrink Assorted Kit", "Isopropyl Alcohol 5L", "Nitrile Gloves Box 100", "Masking Tape 50m",
                    "Solder Wire 1kg", "Cutting Fluid 5L"],
}
_CAT_UNITS = {"Fasteners": ["Each", "Pack"], "Composites": ["Sheet", "Kg", "Roll"], "Electronics": ["Each"],
              "Machined Parts": ["Each"], "Cables": ["Metre", "Roll"], "Consumables": ["Each", "Pack", "Box"]}
_CAT_COST = {"Fasteners": (10, 60), "Composites": (80, 600), "Electronics": (60, 900), "Machined Parts": (50, 700),
             "Cables": (20, 200), "Consumables": (10, 80)}
_CAT_QTY = {"Fasteners": (100, 500), "Composites": (2, 30), "Electronics": (5, 50), "Machined Parts": (5, 60),
            "Cables": (50, 300), "Consumables": (20, 100)}
_CATS = list(_PART_DESC)


@lru_cache(maxsize=None)
def _parts() -> pd.DataFrame:
    r = _rng(1)
    sup_by_cat = {c: [s[0] for s in _SUPP if s[2] == c] for c in _CATS}
    rows = []
    n = 0
    for cat in _CATS:
        for desc in _PART_DESC[cat]:
            n += 1
            lo, hi = _CAT_COST[cat]
            cost = round(float(r.uniform(lo, hi)), 2)
            drawing = f"D-{1000 + n * 7 + int(r.integers(0, 5))}"
            units = _CAT_UNITS[cat]
            rows.append({
                "Part Number": f"PRT-{n:04d}",
                "Description": f"{desc} [Drawing {drawing}]",
                "Category": cat,
                "Default Supplier ID": str(r.choice(sup_by_cat[cat])),
                "Stocking Unit": units[int(r.integers(0, len(units)))],
                "Wasp Item Cost": cost,
            })
    df = pd.DataFrame(rows)
    idx = r.permutation(60)
    avail = np.zeros(60, dtype=int)
    for k, i in enumerate(idx):
        if k < 8:
            avail[i] = 0
        elif k < 16:
            avail[i] = int(r.integers(1, 6))
        else:
            avail[i] = int(r.integers(6, 41))
    df["Total Available"] = avail
    df["Total In House"] = avail + r.integers(0, 9, 60)
    df["On Order"] = r.integers(0, 31, 60)
    df["Committed"] = np.minimum(r.integers(0, 9, 60), df["Total In House"].to_numpy())
    df["Location"] = [f"{chr(65 + int(a))}-{int(b):02d}-{int(c):02d}" for a, b, c in
                      zip(r.integers(0, 6, 60), r.integers(1, 13, 60), r.integers(1, 9, 60))]
    df["Site"] = np.where(r.random(60) < 0.7, "Poole Main", "Marine Yard")
    return df[["Part Number", "Description", "Category", "Default Supplier ID", "Stocking Unit", "Wasp Item Cost",
               "Total Available", "Total In House", "On Order", "Committed", "Location", "Site"]]


# ----------------------------------------------------------------------------
# Timesheets
# ----------------------------------------------------------------------------
_TASKS = ["Design", "Build", "Test", "Install", "Admin"]
_DEPT_TASKS = {"ENGINEERING": ["Design", "Test", "Admin"], "PRODUCTION": ["Build", "Test"],
               "MARITIME": ["Install", "Test"], "SUPPORT": ["Install", "Admin"], "PROCUREMENT": ["Admin"]}


@lru_cache(maxsize=None)
def _timesheets() -> pd.DataFrame:
    r = _rng(2)
    emp = _employees().set_index("Employee ID")
    codes = [p[0] for p in _PROJ]
    weeks = pd.date_range("2026-01-02", "2026-06-26", freq="7D")
    pools = {}
    for e in _BOOKERS:
        pools[e] = [codes[i] for i in sorted(r.choice(len(codes), size=4, replace=False))]
    rows = []
    for we in weeks:
        wd = we.date()
        for e in _BOOKERS:
            total = 37.5 if r.random() < 0.45 else float(r.integers(60, 91)) / 2.0  # 30.0..45.0
            L = 4 if r.random() < 0.55 else 5
            pool = [c for c in pools[e] if not (c == _CLOSED and wd > _CLOSED_AFTER)]
            pref = _DEPT_TASKS[emp.loc[e, "Department"]]
            combos = [(c, t) for c in pool for t in _TASKS]
            w = np.array([3.0 if t in pref else 1.0 for _, t in combos])
            pick = r.choice(len(combos), size=L, replace=False, p=w / w.sum())
            units = int(round(total * 2))
            alloc = np.full(L, 2)  # min 1.0 h each
            extra = r.multinomial(units - 2 * L, np.ones(L) / L)
            alloc = alloc + extra
            for k, ci in enumerate(sorted(pick, key=lambda i: combos[i])):
                c, t = combos[ci]
                rows.append((wd, e, emp.loc[e, "Employee Name"], c, t, alloc[k] / 2.0))
    return pd.DataFrame(rows, columns=["Week Ending", "Employee ID", "Employee Name", "Project Code", "Task", "Hours"])


# ----------------------------------------------------------------------------
# Purchase orders and Xero invoices
# ----------------------------------------------------------------------------
_PO_START = date(2025, 1, 6)
_PO_END = date(2026, 6, 26)
_INV_CUTOFF = date(2026, 5, 20)   # POs for invoiced parts are dated on/before this
_INV_END = date(2026, 6, 30)
_DEPTS = ["ENGINEERING", "PRODUCTION", "MARITIME", "SUPPORT"]


def _bizday(r, lo: date, hi: date) -> date:
    while True:
        d = lo + timedelta(days=int(r.integers(0, (hi - lo).days + 1)))
        if d.weekday() < 5:
            return d


@lru_cache(maxsize=None)
def _po_and_xero():
    r = _rng(3)
    parts = _parts()
    pids = list(parts["Part Number"])
    perm = [pids[i] for i in r.permutation(60)]
    invoiced, po_only, neither = perm[:40], perm[40:50], perm[50:]
    # line counts: 40*4+5 = 165 ; 10*3+5 = 35 -> 200
    n_lines = {}
    bump = set(r.choice(40, size=5, replace=False).tolist())
    for i, p in enumerate(invoiced):
        n_lines[p] = 4 + (1 if i in bump else 0)
    bump2 = set(r.choice(10, size=5, replace=False).tolist())
    for i, p in enumerate(po_only):
        n_lines[p] = 3 + (1 if i in bump2 else 0)
    pinfo = parts.set_index("Part Number")

    lines = []
    for p, n in n_lines.items():
        hi = _INV_CUTOFF if p in invoiced else _PO_END
        for _ in range(n):
            lines.append({"part": p, "sup": pinfo.loc[p, "Default Supplier ID"], "date": _bizday(r, _PO_START, hi)})
    lines.sort(key=lambda x: (x["sup"], x["date"], x["part"]))

    # group into orders (one PO per order): same supplier, within 10 days of group start, distinct parts, <=4 lines
    orders = []
    cur = None
    for ln in lines:
        if (cur and cur["sup"] == ln["sup"] and (ln["date"] - cur["date"]).days <= 10
                and len(cur["lines"]) < 4 and ln["part"] not in [x["part"] for x in cur["lines"]]):
            cur["lines"].append(ln)
        else:
            cur = {"sup": ln["sup"], "date": ln["date"], "lines": [ln]}
            orders.append(cur)
    orders.sort(key=lambda o: (o["date"], o["sup"]))

    projs = [p[0] for p in _PROJ]
    sup_terms = {s[0]: str(r.choice(["30 days", "30 days", "60 days", "Pro-forma"])) for s in _SUPP}
    pr_no, prev = 10100, None
    jira = 2100
    po_no = 55000
    for o in orders:
        po_no += int(r.integers(1, 4))
        o["po"] = f"PO-{po_no}"
        if prev is not None and r.random() < 0.3 and (o["date"] - prev["date"]).days <= 14:
            o["pr"], o["project"], o["dept"], o["jira"] = prev["pr"], prev["project"], prev["dept"], prev["jira"]
        else:
            pr_no += int(r.integers(1, 4))
            o["pr"] = f"PR-{pr_no}"
            ok = [c for c in projs if not (c == _CLOSED and o["date"] > _CLOSED_AFTER)]
            o["project"] = str(r.choice(ok))
            o["dept"] = str(r.choice(_DEPTS))
            jira += int(r.integers(1, 6))
            o["jira"] = f"CL-{jira}"
        o["terms"] = sup_terms[o["sup"]]
        prev = o

    # price drift per part
    base, drift = {}, {}
    for p in pids:
        base[p] = float(pinfo.loc[p, "Wasp Item Cost"]) * float(r.uniform(0.8, 1.2))
        drift[p] = float(r.choice([-1, 1])) * float(r.uniform(0.04, 0.12))
    t0 = date(2025, 1, 1)
    po_rows = []
    for o in orders:
        for ln in o["lines"]:
            p = ln["part"]
            yrs = (o["date"] - t0).days / 365.0
            price = round(max(0.5, base[p] * (1 + drift[p] * yrs) * (1 + float(r.normal(0, 0.01)))), 2)
            lo, hi = _CAT_QTY[pinfo.loc[p, "Category"]]
            qty = int(r.integers(lo, hi + 1))
            po_rows.append({
                "Order Date": o["date"], "PR Number": o["pr"], "PO Number": o["po"], "Part Number": p,
                "Description": pinfo.loc[p, "Description"], "Qty": qty, "Net Unit Price": price,
                "Supplier ID": o["sup"], "Project Code": o["project"], "Department": o["dept"],
                "Payment Terms": o["terms"], "Jira": o["jira"],
            })
    po = pd.DataFrame(po_rows).sort_values(["Order Date", "PO Number", "Part Number"], kind="stable").reset_index(drop=True)

    # invoices: 150 of the 165 lines belonging to invoiced parts, >=1 per invoiced part
    cand = po[po["Part Number"].isin(invoiced)]
    chosen = set()
    for p in invoiced:
        sub = cand[cand["Part Number"] == p]
        chosen.add(int(r.choice(sub.index.to_numpy())))
    rest = [i for i in cand.index if i not in chosen]
    chosen.update(int(i) for i in r.choice(rest, size=150 - len(chosen), replace=False))
    inv = po.loc[sorted(chosen)].copy()
    inv_date = {}
    for po_no_, _ in inv.groupby("PO Number"):
        od = po.loc[po["PO Number"] == po_no_, "Order Date"].iloc[0]
        inv_date[po_no_] = min(od + timedelta(days=int(r.integers(5, 46))), _INV_END)
    inv["Invoice Date"] = inv["PO Number"].map(inv_date)
    nums = {}
    for k, (d, pn) in enumerate(sorted((v, k_) for k_, v in inv_date.items())):
        nums[pn] = f"INV-{4001 + k}"
    inv["Invoice Number"] = inv["PO Number"].map(nums)
    xero = pd.DataFrame({
        "Invoice Date": inv["Invoice Date"], "Contact": inv["Supplier ID"], "Invoice Number": inv["Invoice Number"],
        "Reference": inv["PO Number"], "Item Code": inv["Part Number"], "Description": inv["Description"],
        "Quantity": inv["Qty"], "Unit Price ex VAT": inv["Net Unit Price"],
    }).sort_values(["Invoice Date", "Invoice Number", "Item Code"], kind="stable").reset_index(drop=True)
    return po, xero


# ----------------------------------------------------------------------------
# Receipts, budget, assets, cards
# ----------------------------------------------------------------------------
@lru_cache(maxsize=None)
def _receipts() -> pd.DataFrame:
    r = _rng(4)
    days = [d.date() for d in pd.bdate_range("2026-01-02", "2026-03-31")]
    cids = [c[0] for c in _CUST]
    w = r.uniform(0.5, 2.0, len(cids))
    w = w / w.sum()
    n = 120
    dts = sorted(days[int(i)] for i in r.integers(0, len(days), n))
    cust = r.choice(cids, size=n, p=w)
    amt = np.round(r.choice([1, 2, 5, 10, 15, 25], size=n) * 250 + r.integers(0, 100, n) * 12.5 + r.integers(0, 100, n) / 100.0, 2)
    df = pd.DataFrame({"Date": dts, "Ref": np.arange(10010, 10010 + n), "Customer ID": cust, "Amount": amt})
    return df


@lru_cache(maxsize=None)
def _budget() -> pd.DataFrame:
    r = _rng(5)
    seas = np.array([0.9, 0.9, 1.0, 1.0, 1.1, 1.1, 1.0, 0.8, 1.1, 1.1, 1.0, 1.0])
    rows = []
    for code, *_rest, bud in _PROJ:
        total_c = int(bud * 100)
        w = seas * r.uniform(0.8, 1.2, 12)
        if code == _CLOSED:
            w = w * np.array([1, 1, 1, .2, .2, .2, .1, .1, .1, .1, .1, .1])
        c = np.floor(w / w.sum() * total_c).astype(np.int64)
        c = (c // 100) * 100  # whole pounds
        c[11] += total_c - int(c.sum())
        assert int(c.sum()) == total_c
        for m in range(12):
            rows.append((code, date(2026, m + 1, 1), c[m] / 100.0))
    return pd.DataFrame(rows, columns=["Project Code", "Month", "Budget GBP"])


_ASSET_OPEN = {"Plant & Machinery": 120000.0, "Vehicles": 65000.0, "IT Equipment": 38000.0, "Fixtures & Fittings": 15000.0}
# month -> (additions, disposals, write_offs)
_ASSET_MOVES = {
    "Plant & Machinery": {2: (8000, 0, 0), 3: (0, 5000, 0), 6: (0, 0, 1500), 8: (6000, 0, 0), 9: (0, 3000, 0), 10: (4000, 0, 0)},
    "Vehicles": {4: (10000, 0, 0), 7: (0, 10000, 0), 10: (0, 25000, 0)},
    "IT Equipment": {1: (5000, 0, 0), 4: (0, 2000, 0), 6: (4500, 0, 0), 8: (0, 0, 800), 10: (6000, 0, 0)},
    "Fixtures & Fittings": {2: (0, 0, 300), 3: (2500, 0, 0), 9: (0, 1400, 0)},
}


@lru_cache(maxsize=None)
def _assets() -> pd.DataFrame:
    rows = []
    for cls in _ASSET_OPEN:
        for m in range(1, 11):
            a, d, w = _ASSET_MOVES[cls].get(m, (0, 0, 0))
            rows.append((cls, m, float(a), float(d), float(w)))
    return pd.DataFrame(rows, columns=["Asset Class", "Month", "Additions", "Disposals", "Write Offs"])


@lru_cache(maxsize=None)
def _assets_opening() -> pd.DataFrame:
    return pd.DataFrame(list(_ASSET_OPEN.items()), columns=["Asset Class", "Opening Cost"])


@lru_cache(maxsize=None)
def _assets_rollforward() -> pd.DataFrame:
    a = _assets()
    out = []
    for cls, g in a.groupby("Asset Class", sort=False):
        bf = _ASSET_OPEN[cls]
        for _, row in g.sort_values("Month").iterrows():
            cf = round(bf + row["Additions"] - row["Disposals"] - row["Write Offs"], 2)
            out.append((cls, int(row["Month"]), bf, row["Additions"], row["Disposals"], row["Write Offs"], cf))
            bf = cf
    return pd.DataFrame(out, columns=["Asset Class", "Month", "Cost BF", "Additions", "Disposals", "Write Offs", "Cost CF"])


# merchant, category, MCC, cities, weight, (low, high)
_UK_MERCH = [
    ("Trainline", "Travel", 4011, ["London"], 12, (15, 180)),
    ("Premier Inn", "Accommodation", 7011, ["Birmingham", "Manchester", "Leeds", "Bristol", "Glasgow", "Portsmouth", "Aberdeen"], 9, (65, 280)),
    ("Travelodge", "Accommodation", 7011, ["Southampton", "Cardiff", "Newcastle", "Plymouth"], 5, (55, 220)),
    ("Shell", "Fuel", 5541, ["Poole", "Southampton", "Bournemouth", "Reading", "Exeter"], 14, (38, 98)),
    ("Screwfix", "Materials", 5200, ["Poole", "Southampton", "Christchurch"], 9, (8, 260)),
    ("RS Components", "Materials", 5065, ["Corby"], 6, (15, 420)),
    ("Microsoft", "Software", 5734, ["London"], 5, (8, 95)),
    ("Pret", "Subsistence", 5814, ["London", "Southampton", "Bristol", "Birmingham"], 14, (3, 14)),
    ("Costa Coffee", "Subsistence", 5814, ["Poole", "Reading", "Cardiff"], 6, (3, 12)),
    ("Greene King Inn", "Subsistence", 5812, ["Winchester", "Salisbury", "Portsmouth"], 5, (14, 55)),
]
_FOREIGN_MERCH = [
    ("Hilton Garden Inn Hamburg", "Accommodation", 7011, "Hamburg", "DE", "EUR", (90, 260)),
    ("Courtyard by Marriott Boston", "Accommodation", 7011, "Boston", "US", "USD", (140, 380)),
    ("Ibis Rotterdam Centrum", "Accommodation", 7011, "Rotterdam", "NL", "EUR", (70, 190)),
    ("Deutsche Bahn", "Travel", 4011, "Hamburg", "DE", "EUR", (25, 120)),
    ("Uber", "Travel", 4121, "New York", "US", "USD", (12, 75)),
    ("Brasserie Les Halles", "Subsistence", 5812, "Paris", "FR", "EUR", (22, 85)),
    ("Shell", "Fuel", 5541, "Rotterdam", "NL", "EUR", (45, 110)),
    ("Delta Air Lines", "Travel", 3058, "Atlanta", "US", "USD", (180, 640)),
]


@lru_cache(maxsize=None)
def _cards() -> pd.DataFrame:
    r = _rng(6)
    emp = _employees().set_index("Employee ID")
    holders = list(_CARDS)
    w = np.array([m[4] for m in _UK_MERCH], float)
    w /= w.sum()
    nonlocal_months = [1, 1, 1, 2, 2, 2, 3, 3]  # 8 non-GBP rows
    rows = []
    for month in (1, 2, 3):
        last = (date(2026, month + 1, 1) - timedelta(days=1)).day
        n_for = nonlocal_months.count(month)
        n_ref = 1 if month != 2 else 2  # 4 refunds in total
        n_uk = 80 - n_for - n_ref
        for _ in range(n_uk):
            m = _UK_MERCH[int(r.choice(len(_UK_MERCH), p=w))]
            lo, hi = m[5]
            rows.append(dict(month=month, last=last, h=holders[int(r.integers(0, 10))], merch=m[0], cat=m[1], mcc=m[2],
                             city=str(r.choice(m[3])), country="GB", cur="GBP", orig=round(float(r.uniform(lo, hi)), 2), tt="Purchase"))
        for _ in range(n_for):
            m = _FOREIGN_MERCH[int(r.integers(0, len(_FOREIGN_MERCH)))]
            rows.append(dict(month=month, last=last, h=holders[int(r.integers(0, 10))], merch=m[0], cat=m[1], mcc=m[2],
                             city=m[3], country=m[4], cur=m[5], orig=round(float(r.uniform(*m[6])), 2), tt="Purchase"))
        for _ in range(n_ref):
            m = [x for x in _UK_MERCH if x[0] in ("Screwfix", "RS Components", "Trainline")][int(r.integers(0, 3))]
            rows.append(dict(month=month, last=last, h=holders[int(r.integers(0, 10))], merch=m[0], cat=m[1], mcc=m[2],
                             city=str(r.choice(m[3])), country="GB", cur="GBP",
                             orig=-round(float(r.uniform(10, 120)), 2), tt="Refund"))
    refs = r.choice(90_000_000, size=len(rows), replace=False) + 10_000_000
    out = []
    for row, ref in zip(rows, refs):
        dt = date(2026, row["month"], int(r.integers(1, row["last"] + 1)))
        posted = dt + timedelta(days=int(r.integers(0, 4)))
        e = emp.loc[row["h"]]
        first, last_name = e["Employee Name"].split(" ", 1)
        amt = round(row["orig"] * FX_RATES[row["cur"]], 2)
        out.append({
            "Date": dt, "Posted Date": posted, "Merchant": row["merch"], "Amount GBP": amt, "Category": row["cat"],
            "City": row["city"], "Last Four": e["Card Last Four"], "Name on Card": f"{first[0]} {last_name}".upper(),
            "MCC": row["mcc"], "Trans Type": row["tt"], "Bank Ref": f"BK{int(ref)}",
            "Original Currency Amount": row["orig"], "Original Currency": row["cur"], "Merchant Country": row["country"],
        })
    df = pd.DataFrame(out).sort_values(["Date", "Bank Ref"], kind="stable").reset_index(drop=True)
    return df


# ----------------------------------------------------------------------------
# Public getters (return copies)
# ----------------------------------------------------------------------------
def projects(): return _projects().copy()
def employees(): return _employees().copy()
def labour_rates(): return _labour_rates().copy()
def suppliers(): return _suppliers().copy()
def customers(): return _customers().copy()
def parts(): return _parts().copy()
def timesheets(): return _timesheets().copy()
def purchase_orders(): return _po_and_xero()[0].copy()
def xero_invoices(): return _po_and_xero()[1].copy()
def receipts(): return _receipts().copy()
def budget(): return _budget().copy()
def assets(): return _assets().copy()
def assets_opening(): return _assets_opening().copy()
def assets_rollforward(): return _assets_rollforward().copy()
def cards(): return _cards().copy()


# ----------------------------------------------------------------------------
# Writers (only under __main__ / explicit call)
# ----------------------------------------------------------------------------
def _write_xlsx(df: pd.DataFrame, path: str, money_cols=(), date_cols=()):
    from openpyxl.styles import Font
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Sheet1")
        ws = xw.sheets["Sheet1"]
        for c in ws[1]:
            c.font = Font(bold=True)
        for i, col in enumerate(df.columns, start=1):
            width = max(len(str(col)), *(len(str(v)) for v in df[col])) + 2
            ws.column_dimensions[ws.cell(1, i).column_letter].width = min(width, 45)
            for row in range(2, len(df) + 2):
                cell = ws.cell(row, i)
                if col in money_cols:
                    cell.number_format = "#,##0.00"
                if col in date_cols:
                    cell.number_format = "dd/mm/yyyy"
                if isinstance(cell.value, str):
                    cell.data_type = "s"
                if cell.value == "":
                    cell.value = None


def write_all():
    os.makedirs(REF_DIR, exist_ok=True)
    os.makedirs(TRUTH_DIR, exist_ok=True)
    _write_xlsx(projects(), os.path.join(REF_DIR, "Projects.xlsx"), money_cols=("Budget 2026 GBP",), date_cols=("Start Date",))
    _write_xlsx(employees(), os.path.join(REF_DIR, "Employees.xlsx"))
    _write_xlsx(labour_rates(), os.path.join(REF_DIR, "Labour_Rates.xlsx"), money_cols=("Standard Hourly Rate GBP",))
    _write_xlsx(suppliers(), os.path.join(REF_DIR, "Suppliers_Master.xlsx"))
    _write_xlsx(customers(), os.path.join(REF_DIR, "Customers.xlsx"))
    tables = {
        "parts": parts(), "timesheets": timesheets(), "purchase_orders": purchase_orders(),
        "xero_invoices": xero_invoices(), "receipts": receipts(), "budget": budget(), "assets": assets(),
        "assets_opening": assets_opening(), "cards": cards(),
    }
    for name, df in tables.items():
        df.to_csv(os.path.join(TRUTH_DIR, f"{name}.csv"), index=False, float_format="%.2f", date_format="%Y-%m-%d")
    return {k: len(v) for k, v in tables.items()}


if __name__ == "__main__":
    for k, v in write_all().items():
        print(f"{k}: {v} rows")
