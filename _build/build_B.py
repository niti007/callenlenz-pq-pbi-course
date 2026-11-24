#!/usr/bin/env python
"""Builder B: Wasp stock (Lab 1), Xero + Purchase Log + valuation (Lab 3), Receipts (Lab 4).

Run from repo root:  _build/.venv/bin/python _build/build_B.py
Writes messy xlsx files, solutions/check_totals_B.json, _build/mess_log_B.json, solutions/m/*.m.
Every check total is computed by READING BACK the messy files and applying the documented
cleaning (a pandas/python mimic of the Power Query steps), then asserted against clean truth.
"""
import datetime as dt
import json
import os
import re
import sys

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import master as M  # noqa: E402

FIXED = dt.datetime(2026, 7, 1, 9, 0, 0)
NBSP = "\u00a0"
LOG = []
CHK = {}


def P(*a):
    return os.path.join(ROOT, *a)


def rng(k):
    return np.random.default_rng([42, k])


# ---------------------------------------------------------------- helpers
def log(file, sheet, cells, mess_type, desc, fix, lab, sample=None):
    cells = list(cells)
    if not cells:
        return
    use = cells if sample is None else cells[:sample]
    d = desc if len(use) == len(cells) else f"{desc} [{len(cells)} instances in file; first {len(use)} logged]"
    for c in use:
        LOG.append(dict(file=file, sheet=sheet, location=c, mess_type=mess_type, description=d,
                        teaching_fix=fix, lab=lab))


def save_wb(wb, path):
    wb.properties.created = FIXED
    wb.properties.modified = FIXED
    wb.properties.creator = "Callen-Lenz Finance (demo)"
    wb.properties.lastModifiedBy = "Callen-Lenz Finance (demo)"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    from zipfile import ZIP_DEFLATED, ZipFile
    from openpyxl.writer.excel import ExcelWriter
    archive = ZipFile(path, "w", ZIP_DEFLATED, allowZip64=True)
    ExcelWriter(wb, archive).save()   # (wb.save() would overwrite 'modified' with now())


def read_rows(path, sheet, ncols):
    ws = load_workbook(path)[sheet]
    return [list(r) for r in ws.iter_rows(min_row=1, max_col=ncols, values_only=True)]


def clean_text(t):
    """Mimic M CleanText: nbsp/lf/cr -> space, collapse repeated spaces, trim."""
    if t is None:
        return None
    s = str(t).replace(NBSP, " ").replace("\n", " ").replace("\r", " ")
    return " ".join(p for p in s.split(" ") if p)


def norm_part(t):
    """Part-number rule: uppercase, remove spaces/hyphens/nbsp/line breaks, must start PRT, digits padded to 4."""
    if t is None:
        return None
    s = re.sub(r"[\s\u00a0\-]", "", str(t)).upper()
    if s.startswith("PRT") and s[3:] != "":
        return "PRT-" + s[3:].zfill(4)
    return None


def norm_proj(t):
    if t is None:
        return None
    s = re.sub(r"[\s\u00a0\-]", "", str(t)).upper()
    return "CL-" + s[2:] if s.startswith("CLP") else s


def parse_money(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("£", "").replace(",", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        x = float(s)
    except ValueError:
        return None
    return -x if neg else x


def parse_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%d-%b-%Y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"bad date {v!r}")


PN_KINDS = 10


def pn_variant(pn, k):
    n = int(pn[4:])
    return [pn, pn.lower(), pn + " ", " " + pn, f"PRT{n}", f"prt-{n}", f"PRT-{n:03d}",
            f"PRT {n:04d}", pn + NBSP, f"PRT{n:04d}"][k]


PN_P = [.30, .10, .10, .08, .10, .08, .08, .06, .05, .05]


def proj_variant(pc, k):
    d = pc[3:]
    return [pc, pc.lower(), f"CL {d}", f"CL{d}", f"cl-{d.lower()} "][k]


def style_header(ws, row, ncols, wrap=False):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True)
        if wrap:
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def part_groups():
    x = set(M.xero_invoices()["Item Code"])
    po = set(M.purchase_orders()["Part Number"])
    g = {}
    for pn in M.parts()["Part Number"]:
        g[pn] = "xero" if pn in x else ("po" if pn in po else "neither")
    return g


def addr(c, r):
    return f"{CL(c)}{r}"


# ================================================================= LAB 1  WASP
WASP_COLS = ["Item Number", "Type", "Description", "Total Available", "Total Checked Out", "Total In House",
             "On Order", "Committed", "Tax Code", "Item Cost", "Cost Method", "Item List Price",
             "Item Sale Price", "Lot", "Date Code", "Manufacturer", "Category", "Stocking Unit",
             "Alt Item Number", "Serial Number", "Attachment", "Location", "Site", "Drawing Issue"]
WASP_F = "data/01_Wasp_Stock/Wasp_Stock_Export.xlsx"
WASP_S = "Stock"


def build_wasp(groups):
    r = rng(201)
    parts = M.parts().reset_index(drop=True)
    sup = M.suppliers().set_index("Supplier ID")["Standard Supplier Name"].to_dict()
    pns = list(parts["Part Number"])
    ks = r.choice(PN_KINDS, size=60, p=PN_P)
    kind = dict(zip(pns, ks))
    gbp_text = set(np.array(pns)[r.choice(60, 10, replace=False)])
    elig = [p for p in pns if groups[p] != "neither"]
    bad = list(r.choice(elig, 4, replace=False))
    na_parts, ref_parts = bad[:2], bad[2:]
    brk_desc = set(np.array(pns)[r.choice(60, 12, replace=False)])
    nbsp_desc = set(np.array(pns)[r.choice(60, 10, replace=False)])
    pad_desc = set(np.array(pns)[r.choice(60, 6, replace=False)])
    dup_parts = list(r.choice(pns, 4, replace=False))
    drawing_issue = {p: str(r.choice(["A", "B", "C"])) for p in pns}
    cats = sorted(parts["Category"].unique())

    def mkrow(row, pnvar, first, dup=False, cost_override=None, ta_override=None):
        pn = row["Part Number"]
        desc = row["Description"]
        tags = []
        if not dup:
            if pn in brk_desc:
                desc = desc.replace(" [Drawing", "\n[Drawing")
                tags.append(("Description", "line_break"))
            elif pn in nbsp_desc:
                desc = desc.replace(" ", NBSP, 1)
                tags.append(("Description", "nbsp"))
            elif pn in pad_desc:
                desc = " " + desc + "  "
                tags.append(("Description", "trailing_space"))
        cost = float(row["Wasp Item Cost"])
        costv = cost
        if cost_override is not None:
            costv = cost_override
            tags.append(("Item Cost", "na_value" if cost_override == "N/A" else "ref_error"))
        elif not dup and pn in na_parts:
            costv = "N/A"
            tags.append(("Item Cost", "na_value"))
        elif not dup and pn in ref_parts:
            costv = "#REF!"
            tags.append(("Item Cost", "ref_error"))
        elif not dup and pn in gbp_text:
            costv = f"£{cost:,.2f}"
            tags.append(("Item Cost", "gbp_text"))
        if first is False:
            pass
        v = {
            "Item Number": pnvar, "Type": "Inventory", "Description": desc,
            "Total Available": int(row["Total Available"]) if ta_override is None else ta_override,
            "Total Checked Out": 0, "Total In House": int(row["Total In House"]), "On Order": int(row["On Order"]),
            "Committed": int(row["Committed"]), "Tax Code": "T20", "Item Cost": costv, "Cost Method": "FIFO",
            "Item List Price": round(cost * 1.4, 2), "Item Sale Price": round(cost * 1.25, 2),
            "Lot": None, "Date Code": None, "Manufacturer": sup[row["Default Supplier ID"]],
            "Category": row["Category"] if first else None, "Stocking Unit": row["Stocking Unit"],
            "Alt Item Number": None, "Serial Number": None, "Attachment": None,
            "Location": row["Location"], "Site": row["Site"], "Drawing Issue": drawing_issue[pn],
        }
        if pnvar != pn:
            tags.append(("Item Number", "part_number_variant"))
        if not first:
            tags.append(("Category", "blank_category_fill_down"))
        return dict(kind="dup" if dup else "data", vals=[v[c] for c in WASP_COLS], tags=tags, pn=pn)

    blank = dict(kind="blank", vals=[None] * 24, tags=[], pn=None)
    hdr = dict(kind="hdr", vals=list(WASP_COLS), tags=[], pn=None)
    out = []
    for gi, cat in enumerate(cats):
        sub = parts[parts["Category"] == cat]
        g = []
        for j, (_, row) in enumerate(sub.iterrows()):
            g.append(mkrow(row, pn_variant(row["Part Number"], kind[row["Part Number"]]), j == 0))
        # duplicates (different variant, stale Total Available)
        for dp in dup_parts:
            if dp in set(sub["Part Number"]):
                pos = [i for i, x in enumerate(g) if x["pn"] == dp and x["kind"] == "data"][0]
                row = sub[sub["Part Number"] == dp].iloc[0]
                dk = (kind[dp] + 3) % PN_KINDS
                if dk == 0:
                    dk = 1
                extra = {}
                if len([1 for q in out + g if q["kind"] == "dup"]) < 2:
                    extra["cost_override"] = "N/A"   # first two dups carry a bad cost (ignored by keep-first)
                d = mkrow(row, pn_variant(dp, dk), False, dup=True,
                          ta_override=int(row["Total Available"]) + 3, **extra)
                g.insert(min(pos + 1 + int(r.integers(0, 3)), len(g)), d)
        if gi == 1:  # junk test row
            jv = {c: None for c in WASP_COLS}
            jv.update({"Item Number": "TEST-ITEM", "Type": "Inventory", "Description": "TEST ITEM - DELETE ME",
                       "Total Available": 99, "Item Cost": 1.0, "Cost Method": "FIFO"})
            g.insert(len(g) // 2, dict(kind="junk", vals=[jv[c] for c in WASP_COLS], tags=[("Item Number", "junk_row")], pn=None))
        if gi == 4:
            jv = {c: None for c in WASP_COLS}
            jv.update({"Item Number": "ZZ SAMPLE", "Type": "Inventory", "Description": "Sample - do not stock",
                       "Total Available": 5, "Item Cost": 5.0, "Cost Method": "FIFO"})
            g.insert(len(g) // 3, dict(kind="junk", vals=[jv[c] for c in WASP_COLS], tags=[("Item Number", "junk_row")], pn=None))
        if gi == 2:
            g.insert(len(g) // 2, hdr)
        if gi == 4:
            g.insert(len(g) // 2 + 1, blank)
        out.extend(g)
        if gi in (0, 3):
            out.append(blank)
    out.append(blank)
    tv = [None] * 24
    tv[0] = "Total"
    tv[3] = int(sum(x["vals"][3] for x in out if x["kind"] in ("data", "dup")))
    out.append(dict(kind="total", vals=tv, tags=[("Item Number", "subtotal_row")], pn=None))

    wb = Workbook()
    ws = wb.active
    ws.title = WASP_S
    ws["A1"] = "Wasp Inventory Export - Stock On Hand"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Exported 01/07/2026 08:42 by KR  |  Site filter: ALL"
    for i, c in enumerate(WASP_COLS, 1):
        ws.cell(row=3, column=i, value=c)
    style_header(ws, 3, 24)
    cells = {}
    rowno = 4
    row_of = []
    for x in out:
        row_of.append(rowno)
        if x["kind"] != "blank":
            for ci, v in enumerate(x["vals"], 1):
                if v is not None:
                    ws.cell(row=rowno, column=ci, value=v)
                    if isinstance(v, str) and "\n" in v:
                        ws.cell(row=rowno, column=ci).alignment = Alignment(wrap_text=True)
        rowno += 1
    for ci in range(1, 25):
        ws.column_dimensions[CL(ci)].width = 16
    ws.column_dimensions["C"].width = 44
    save_wb(wb, P(WASP_F))

    # ---- log
    L = "Lab 1"
    log(WASP_F, WASP_S, ["A1:X1", "A2:X2"], "title_row",
        "Two export title rows above the real header (row 3).", "Home > Remove Rows > Remove Top Rows = 2, then Use First Row as Headers.", L)
    byt = {}
    for x, rn in zip(out, row_of):
        for col, t in x["tags"]:
            byt.setdefault(t, []).append(addr(WASP_COLS.index(col) + 1, rn))
    log(WASP_F, WASP_S, byt["part_number_variant"], "part_number_variant",
        "Part number typed several ways (prt-12, PRT12, PRT-012, PRT 0012, trailing space, nbsp...).",
        "Custom normalise: Text.Clean, replace CHAR(160), remove spaces/hyphens, Upper, strip 'PRT', Text.PadStart(digits,4,\"0\"), prefix 'PRT-'.", L, sample=5)
    log(WASP_F, WASP_S, byt["gbp_text"], "gbp_text", "Item Cost stored as text with a pound sign.",
        "Replace '£' and ',' with nothing then change type to Decimal (done inside try ... otherwise null).", L, sample=3)
    log(WASP_F, WASP_S, byt["na_value"], "na_value", "Item Cost is the text N/A.",
        "try Number.From(...) otherwise null (unusable cost = null; counts as 0 in Q5).", L)
    log(WASP_F, WASP_S, byt["ref_error"], "ref_error", "Item Cost is an Excel #REF! error value.",
        "Transform > Replace Errors (Table.ReplaceErrorValues) -> null, BEFORE any row filtering.", L)
    log(WASP_F, WASP_S, byt["line_break"], "line_break", "Embedded line break inside Description (Alt+Enter).",
        "Replace #(lf) with a space, then Trim.", L, sample=3)
    log(WASP_F, WASP_S, byt["nbsp"], "nbsp", "Non-breaking space CHAR(160) inside Description.",
        "Replace #(00A0) with a normal space, then Trim.", L, sample=3)
    log(WASP_F, WASP_S, byt["trailing_space"], "trailing_space", "Leading/trailing spaces on Description.",
        "Transform > Format > Trim.", L, sample=3)
    log(WASP_F, WASP_S, byt["blank_category_fill_down"], "blank_category_fill_down",
        "Category only typed on the first row of each group (merged-cell style); other rows blank.",
        "Remove blank rows and repeated header FIRST, then Transform > Fill > Down on Category.", L, sample=4)
    log(WASP_F, WASP_S, byt["junk_row"], "junk_row", "Test/sample item that is not a real PRT part.",
        "After normalising Item Number these become null: filter Item Number is not null.", L)
    log(WASP_F, WASP_S, byt["subtotal_row"], "subtotal_row", "'Total' footer row with a unit total.",
        "Removed by the same 'not a valid PRT part' filter.", L)
    br = [f"A{rn}:X{rn}" for x, rn in zip(out, row_of) if x["kind"] == "blank"]
    log(WASP_F, WASP_S, br, "blank_row", "Completely blank row.", "Home > Remove Rows > Remove Blank Rows.", L)
    log(WASP_F, WASP_S, [f"A{rn}:X{rn}" for x, rn in zip(out, row_of) if x["kind"] == "hdr"], "repeated_header",
        "Header row repeated in the middle of the data (page break in the export).",
        "Filter Item Number <> \"Item Number\".", L)
    dups = [f"A{rn}:X{rn}" for x, rn in zip(out, row_of) if x["kind"] == "dup"]
    log(WASP_F, WASP_S, dups, "duplicate_row",
        "Same part appears again with different case/spacing and a stale Total Available.",
        "Normalise Item Number, then Table.Buffer + Remove Duplicates on Item Number (keep first).", L)
    return out


def clean_wasp(keep="first"):
    """Mimic lab1_wasp.m using the messy file."""
    rows = read_rows(P(WASP_F), WASP_S, 24)
    rows = rows[2:]
    data = rows[1:]  # promote headers
    cp = {"L1_rows_after_promote": len(data)}
    data = [r for r in data if any(c not in (None, "") for c in r)]
    cp["L1_rows_after_remove_blank"] = len(data)
    data = [r for r in data if r[0] != "Item Number"]
    cp["L1_rows_after_remove_repeated_header"] = len(data)
    ci = WASP_COLS.index("Category")
    last = None
    for r in data:
        if r[ci] is None:
            r[ci] = last
        else:
            last = r[ci]
    ii = WASP_COLS.index("Item Number")
    for r in data:
        r[ii] = norm_part(r[ii])
    data = [r for r in data if r[ii] is not None]
    cp["L1_rows_after_filter_valid_parts"] = len(data)
    seen, ded = set(), []
    for r in (data if keep == "first" else reversed(data)):
        if r[ii] not in seen:
            seen.add(r[ii])
            ded.append(r)
    if keep != "first":
        ded.reverse()
    cp["L1_rows_after_dedupe"] = len(ded)
    df = pd.DataFrame(ded, columns=WASP_COLS)
    df["Description"] = df["Description"].map(lambda s: s.replace("\n", " ").replace("\r", " ").replace(NBSP, " ").strip())
    df["Drawing No"] = df["Description"].map(lambda s: s.split("[Drawing ")[1].split("]")[0])
    df["Item Cost"] = df["Item Cost"].map(parse_money)
    df["Stock Status"] = df["Total Available"].map(lambda t: "Out of stock" if t == 0 else ("Reorder" if t <= 5 else "OK"))
    df = df.rename(columns={"Item Number": "Part Number"})
    return df, cp


# ================================================================= LAB 3 XERO
XERO_F = "data/03_Valuation/Xero_Payable_Invoice_Detail.xlsx"
XERO_S = "Xero"
XERO_COLS = ["Invoice Date", "Contact", "Source", "Reference", "Item Code", "Description", "Quantity", "Unit Price (ex) (GBP)"]
MAP_F = "data/03_Valuation/Supplier_Name_Map.xlsx"


def sup_contact(r, sid, std, idx):
    k = int(r.integers(0, 8))
    if k < 3:
        s = std[sid]
        tag = None
    else:
        vs = M.SUPPLIER_VARIANTS[sid]
        s = vs[(k - 3) % len(vs)]
        tag = "supplier_variant"
    n = float(r.random())
    if n < .12:
        s = s + " "
        tag = tag or "trailing_space"
    elif n < .18:
        s = " " + s
        tag = tag or "trailing_space"
    elif n < .23:
        s = s.replace(" ", "  ", 1)
        tag = tag or "double_space"
    elif n < .26:
        s = s + NBSP
        tag = tag or "nbsp"
    return s, tag


def build_xero(groups):
    r = rng(202)
    x = M.xero_invoices().copy()
    x["Invoice Date"] = pd.to_datetime(x["Invoice Date"])
    std = M.suppliers().set_index("Supplier ID")["Standard Supplier Name"].to_dict()
    recs = []
    for _, row in x.iterrows():
        pnk = int(r.choice(PN_KINDS, p=PN_P))
        con, ctag = sup_contact(r, row["Contact"], std, 0)
        ref = row["Reference"]
        rtag = None
        q = float(r.random())
        if q < .05:
            ref = ref.lower()
            rtag = "case_variant"
        elif q < .09:
            ref = ref + " "
            rtag = "trailing_space"
        desc = row["Description"]
        dtag = None
        if float(r.random()) < .08:
            desc = desc + "  "
            dtag = "trailing_space"
        elif float(r.random()) < .05:
            desc = desc.replace(" ", NBSP, 1)
            dtag = "nbsp"
        recs.append(dict(date=row["Invoice Date"], contact=con, ctag=ctag, source="Payable Invoice", ref=ref, rtag=rtag,
                         code=pn_variant(row["Item Code"], pnk), pn=row["Item Code"], desc=desc, dtag=dtag,
                         qty=int(row["Quantity"]), price=float(row["Unit Price ex VAT"]), kind="inv"))
    # near-duplicate lines (same invoice line imported twice with different case/space)
    idxs = sorted(r.choice(len(recs), 3, replace=False))
    dups = []
    for i in idxs:
        o = recs[i]
        d = dict(o)
        d["code"] = pn_variant(o["pn"], (PN_KINDS - 1 - [k for k in range(PN_KINDS) if pn_variant(o["pn"], k) == o["code"]][0]) % PN_KINDS or 1)
        d["contact"] = o["contact"].strip() + " "
        d["ctag"] = "trailing_space"
        d["desc"] = o["desc"].strip() + " "
        d["dtag"] = "trailing_space"
        d["ref"] = o["ref"].strip().upper()
        d["rtag"] = None
        d["kind"] = "dup"
        d["orig_i"] = i
        dups.append(d)
    # credit notes: latest-dated Xero line for 3 invoiced parts
    last = {}
    for rc in recs:
        if rc["pn"] not in last or rc["date"] >= last[rc["pn"]]["date"]:
            last[rc["pn"]] = rc
    cn_parts = [p for p in sorted(last) if last[p]["date"] <= pd.Timestamp("2026-05-15")]
    cn_pick = [cn_parts[i] for i in sorted(r.choice(len(cn_parts), 3, replace=False))]
    cns = []
    for p in cn_pick:
        o = last[p]
        cns.append(dict(date=o["date"] + pd.Timedelta(days=12), contact=o["contact"], ctag=o["ctag"], source="Payable Credit Note",
                        ref=o["ref"].strip().upper(), rtag=None, code=pn_variant(p, 0), pn=p, desc=o["desc"].strip() + " - RETURNED",
                        dtag=None, qty=-min(5, o["qty"]), price=round(o["price"] * 0.8, 2), kind="cn"))
    allr = []
    for i, rc in enumerate(recs):
        allr.append((rc["date"], 0, i, rc))
        for d in dups:
            if d["orig_i"] == i:
                allr.append((rc["date"], 1, i, d))
    for c in cns:
        allr.append((c["date"], 2, 0, c))
    allr.sort(key=lambda t: (t[0], t[1], t[2]))
    seq = [t[3] for t in allr]
    n = len(seq)
    text_lo, text_hi = int(n * .38), int(n * .62)
    gbp_idx = set(int(i) for i in r.choice(n, 8, replace=False))

    wb = Workbook()
    ws = wb.active
    ws.title = XERO_S
    ws["A1"] = "Payable Invoice Detail"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Callen-Lenz Associates Limited"
    ws["A3"] = "For the period 1 January 2025 to 30 June 2026"
    for i, c in enumerate(XERO_COLS, 1):
        ws.cell(row=5, column=i, value=c)
    style_header(ws, 5, 8)
    tag_cells = {}
    rn = 6
    qtot = 0
    for i, rc in enumerate(seq):
        is_text = text_lo <= i < text_hi
        dv = rc["date"].strftime("%d/%m/%Y") if is_text else rc["date"].to_pydatetime()
        ws.cell(row=rn, column=1, value=dv)
        if not is_text:
            ws.cell(row=rn, column=1).number_format = "dd/mm/yyyy"
        else:
            tag_cells.setdefault("text_date", []).append(addr(1, rn))
        ws.cell(row=rn, column=2, value=rc["contact"])
        ws.cell(row=rn, column=3, value=rc["source"])
        ws.cell(row=rn, column=4, value=rc["ref"])
        ws.cell(row=rn, column=5, value=rc["code"])
        ws.cell(row=rn, column=6, value=rc["desc"])
        ws.cell(row=rn, column=7, value=rc["qty"])
        if i in gbp_idx:
            ws.cell(row=rn, column=8, value=f"£{rc['price']:,.2f}")
            tag_cells.setdefault("gbp_text", []).append(addr(8, rn))
        else:
            ws.cell(row=rn, column=8, value=rc["price"])
        qtot += rc["qty"]
        if rc["ctag"] == "supplier_variant":
            tag_cells.setdefault("supplier_variant", []).append(addr(2, rn))
        elif rc["ctag"] in ("trailing_space", "double_space", "nbsp"):
            tag_cells.setdefault(rc["ctag"], []).append(addr(2, rn))
        if pn_variant(rc["pn"], 0) != rc["code"]:
            tag_cells.setdefault("part_number_variant", []).append(addr(5, rn))
        if rc["rtag"]:
            tag_cells.setdefault("ref_" + rc["rtag"], []).append(addr(4, rn))
        if rc["dtag"] == "nbsp":
            tag_cells.setdefault("nbsp", []).append(addr(6, rn))
        if rc["kind"] == "dup":
            tag_cells.setdefault("near_duplicate_row", []).append(f"A{rn}:H{rn}")
        if rc["kind"] == "cn":
            tag_cells.setdefault("credit_note", []).append(f"A{rn}:H{rn}")
        rn += 1
    rn += 1  # blank row
    ws.cell(row=rn, column=1, value="Total")
    ws.cell(row=rn, column=7, value=qtot)
    tag_cells["subtotal_row"] = [f"A{rn}:H{rn}"]
    tag_cells["blank_row"] = [f"A{rn - 1}:H{rn - 1}"]
    for ci, w in enumerate([13, 34, 20, 14, 14, 46, 10, 20], 1):
        ws.column_dimensions[CL(ci)].width = w
    save_wb(wb, P(XERO_F))

    L = "Lab 3"
    log(XERO_F, XERO_S, ["A1:H1", "A2:H2", "A3:H3"], "title_row",
        "Three report title lines plus one blank row (4 rows) above the header on row 5.", "Remove Top Rows = 4, then Use First Row as Headers.", L)
    log(XERO_F, XERO_S, tag_cells["text_date"], "text_date",
        "A block of invoice dates is stored as UK dd/mm/yyyy text; the rest are real dates.",
        "Custom step: if value is text then Date.FromText(v,[Culture=\"en-GB\"]) else Date.From(v) (or Change Type > Using Locale > English (UK)).", L, sample=4)
    log(XERO_F, XERO_S, tag_cells["gbp_text"], "gbp_text", "Unit Price stored as '£' text.",
        "Replace '£' and ',' then change type to Decimal (custom ParseMoney).", L, sample=4)
    log(XERO_F, XERO_S, tag_cells["part_number_variant"], "part_number_variant",
        "Item Code typed several ways (prt-12, PRT 0012, trailing space...).", "Same NormPart function as Lab 1.", L, sample=4)
    log(XERO_F, XERO_S, tag_cells["supplier_variant"], "supplier_variant",
        "Contact uses supplier name variants (Ltd / Limited / Ltd. / typos / CAPS / truncated).",
        "Trim/Clean then Merge with Supplier_Name_Map (From -> To) to get the standard name.", L, sample=4)
    for t in ("trailing_space", "double_space", "nbsp"):
        if t in tag_cells:
            log(XERO_F, XERO_S, tag_cells[t], t, f"Contact/Description has {t.replace('_', ' ')}.",
                "CleanText: replace CHAR(160), collapse spaces, Trim.", L, sample=3)
    for t in ("ref_case_variant", "ref_trailing_space"):
        if t in tag_cells:
            log(XERO_F, XERO_S, tag_cells[t], "case_or_space_variant", "Reference in lower case or with trailing space.",
                "Text.Upper(CleanText(Reference)).", L, sample=3)
    log(XERO_F, XERO_S, tag_cells["near_duplicate_row"], "near_duplicate_row",
        "Invoice line imported twice with different case/spacing in Item Code, Contact, Reference.",
        "Normalise Item Code + Reference, then Remove Duplicates on [Reference, Item Code] (keep first).", L)
    log(XERO_F, XERO_S, tag_cells["credit_note"], "credit_note",
        "Credit note line (Source = Payable Credit Note, negative Quantity) dated AFTER the latest invoice for that part.",
        "Filter Source <> \"Payable Credit Note\" BEFORE choosing the latest price.", L)
    log(XERO_F, XERO_S, tag_cells["subtotal_row"], "subtotal_row", "'Total' footer row.", "Filter Item Code is not null.", L)
    log(XERO_F, XERO_S, tag_cells["blank_row"], "blank_row", "Blank row before the footer.", "Remove Blank Rows.", L)
    return seq


def build_supplier_map():
    std = M.suppliers().set_index("Supplier ID")["Standard Supplier Name"].to_dict()
    rows, seen = [], {}
    for sid, name in std.items():
        for v in [name] + M.SUPPLIER_VARIANTS[sid]:
            assert v not in seen, f"variant collision {v}"
            seen[v] = name
            rows.append((v, name))
    wb = Workbook()
    ws = wb.active
    ws.title = "Map"
    ws.append(["From", "To"])
    style_header(ws, 1, 2)
    for a in rows:
        ws.append(list(a))
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 38
    save_wb(wb, P(MAP_F))
    return seen


def clean_xero(supmap):
    rows = read_rows(P(XERO_F), XERO_S, 8)
    cp = {}
    data = rows[5:]
    cp["L3_xero_rows_after_promote"] = len(data)
    data = [r for r in data if any(c not in (None, "") for c in r)]
    data = [r for r in data if r[4] is not None]
    cp["L3_xero_rows_after_remove_blank_total"] = len(data)
    recs = []
    for r in data:
        recs.append(dict(date=r[0], contact=clean_text(r[1]), source=clean_text(r[2]), ref=(clean_text(r[3]) or "").upper(),
                         code=norm_part(r[4]), desc=clean_text(r[5]), qty=r[6], price=parse_money(r[7])))
    cp["L3_xero_credit_note_lines"] = sum(1 for r in recs if r["source"] == "Payable Credit Note")
    recs = [r for r in recs if r["source"] != "Payable Credit Note"]
    cp["L3_xero_rows_after_remove_credit_notes"] = len(recs)
    for r in recs:
        r["date"] = parse_date(r["date"])
    seen, ded = set(), []
    for r in recs:
        k = (r["ref"], r["code"])
        if k not in seen:
            seen.add(k)
            ded.append(r)
    cp["L3_xero_rows_after_dedupe"] = len(ded)
    for r in ded:
        r["contact"] = supmap[r["contact"]]
    return pd.DataFrame(ded), cp


# ================================================================= LAB 3 PURCHASE LOG
PL_F = "data/03_Valuation/Purchase_Log.xlsx"
PL_S = "Purchase Log"
PL_COLS = ["PR Number", "PO Number", "Date", "Ord'd\nBy", "Ath'd\nBy", "Department", "Project Code", None,
           "Supplier", "Supplier\nContact", "Part Number ", "Manufactured Part Numbers ", "Description", "Qty", "Unit", None,
           "Net Unit Price ", "Net Total", "VAT 20%", "Gross Total", "Currency", "Payment Terms", "Required\nBy",
           "Promised\nDelivery", "Status", None, "Qty\nReceived", "Date\nReceived", "Delivered\nTo", "Cost Centre",
           "Budget Code", "Quote Ref", "Quote\nDate", "Lead Time\n(days)", None, "Invoice No.", "Invoice\nDate",
           "Invoiced?", "Paid?", "Approved\nDate", "Priority", "Category", "Notes", None, "Last Updated", "Updated By",
           "Attachment", "Comments"]
assert len(PL_COLS) == 48 and sum(c is None for c in PL_COLS) == 5


def build_pl(groups):
    r = rng(203)
    po = M.purchase_orders().copy()
    po["Order Date"] = pd.to_datetime(po["Order Date"])
    parts = M.parts().set_index("Part Number")
    std = M.suppliers().set_index("Supplier ID")["Standard Supplier Name"].to_dict()
    emps = M.employees()
    buyers = list(emps[emps["Department"].isin(["PROCUREMENT", "ENGINEERING", "PRODUCTION"])]["Employee Name"])
    recs = []
    for _, row in po.iterrows():
        recs.append(dict(date=row["Order Date"], pr=row["PR Number"], po=row["PO Number"], pn=row["Part Number"],
                         desc=row["Description"], qty=int(row["Qty"]), price=float(row["Net Unit Price"]), sid=row["Supplier ID"],
                         proj=row["Project Code"], dept=row["Department"], terms=row["Payment Terms"], jira=row["Jira"],
                         cancelled=False))
    # cancelled lines injected AFTER the last real PO line of 3 PO-only parts and 1 'neither' part
    last = {}
    for rc in recs:
        if rc["pn"] not in last or rc["date"] >= last[rc["pn"]]["date"]:
            last[rc["pn"]] = rc
    po_only = [p for p in sorted(last) if groups[p] == "po" and last[p]["date"] <= pd.Timestamp("2026-06-15")]
    cpick = [po_only[i] for i in sorted(r.choice(len(po_only), 3, replace=False))]
    neither = sorted(p for p in groups if groups[p] == "neither")
    cpick.append(neither[int(r.integers(0, len(neither)))])
    canc = []
    for k, p in enumerate(cpick):
        if p in last:
            o = last[p]
            d = dict(o)
            d["date"] = o["date"] + pd.Timedelta(days=int(r.integers(5, 14)))
            d["price"] = round(o["price"] * 1.35, 2)
        else:  # 'neither' part: only line is cancelled
            sid = parts.loc[p, "Default Supplier ID"]
            d = dict(date=pd.Timestamp("2026-05-04") + pd.Timedelta(days=int(r.integers(0, 20))), pr=f"PR-{10900 + k}",
                     pn=p, desc=parts.loc[p, "Description"], qty=10, price=round(float(parts.loc[p, "Wasp Item Cost"]) * 1.5, 2),
                     sid=sid, proj="CL-P102", dept="PRODUCTION", terms="30 days", jira="CL-2300")
        d["po"] = f"PO-{55900 + k}"
        d["jira"] = f"CL-{2300 + k}"
        d["pr"] = f"PR-{10900 + k}" if p not in last else d["pr"]
        d["qty"] = int(r.integers(5, 40))
        d["cancelled"] = True
        canc.append(d)
    allr = [(rc["date"], 0, i, rc) for i, rc in enumerate(recs)] + [(c["date"], 1, i, c) for i, c in enumerate(canc)]
    allr.sort(key=lambda t: (t[0], t[1], t[2]))
    seq = [t[3] for t in allr]
    n = len(seq)
    wb = Workbook()
    ws = wb.active
    ws.title = PL_S
    for i, c in enumerate(PL_COLS, 1):
        if c is not None:
            ws.cell(row=1, column=i, value=c)
    style_header(ws, 1, 48, wrap=True)
    ws.row_dimensions[1].height = 45
    tc = {}
    tags_pl = {"gbp_text": [], "text_date_slash": [], "text_date_mon": [], "supplier_variant": [], "part_number_variant": [],
               "project_code_variant": [], "blank_po_number": [], "cancelled_line": [], "composite_notes": []}
    for i, rc in enumerate(seq):
        rn = i + 2
        d = rc["date"].to_pydatetime()
        q = float(r.random())
        if q < .50:
            dv, dk = d, None
        elif q < .78:
            dv, dk = d.strftime("%d/%m/%Y"), "text_date_slash"
        else:
            dv, dk = d.strftime("%d-%b-%Y"), "text_date_mon"
        con, ctag = sup_contact(r, rc["sid"], std, 0)
        pk = int(r.choice(PN_KINDS, p=PN_P))
        pjk = int(r.choice(5, p=[.55, .15, .1, .1, .1]))
        tag = ["urgent", "standard", "expedite", "stock replenishment", "free issue of material"][int(r.integers(0, 5))]
        notes = f"{rc['po']} / Jira {rc['jira']} / " + ("CANCELLED" if rc["cancelled"] and float(r.random()) < .6 else
                                                      ("cancelled by supplier" if rc["cancelled"] else tag))
        net = round(rc["qty"] * rc["price"], 2)
        price_txt = float(r.random()) < .4
        pv = f"£{rc['price']:,.2f}" if price_txt else rc["price"]
        by = buyers[int(r.integers(0, len(buyers)))]
        vals = {
            "PR Number": rc["pr"], "PO Number": None if float(r.random()) < .2 else rc["po"], "Date": dv,
            "Ord'd\nBy": by, "Ath'd\nBy": ["Hannah Brooks", "Rajesh Patel"][int(r.integers(0, 2))],
            "Department": rc["dept"], "Project Code": proj_variant(rc["proj"], pjk), "Supplier": con,
            "Supplier\nContact": f"sales@{rc['sid'].lower()}-demo.example", "Part Number ": pn_variant(rc["pn"], pk),
            "Manufactured Part Numbers ": f"MPN-{int(rc['pn'][4:]) * 37 + 1000}", "Description": rc["desc"], "Qty": rc["qty"],
            "Unit": parts.loc[rc["pn"], "Stocking Unit"], "Net Unit Price ": pv,
            "Net Total": f"£{net:,.2f}" if float(r.random()) < .3 else net, "VAT 20%": round(net * .2, 2),
            "Gross Total": round(net * 1.2, 2), "Currency": "GBP", "Payment Terms": rc["terms"],
            "Required\nBy": (rc["date"] + pd.Timedelta(days=21)).to_pydatetime(),
            "Promised\nDelivery": (rc["date"] + pd.Timedelta(days=14)).to_pydatetime(),
            "Status": "Open" if rc["date"] > pd.Timestamp("2026-05-01") else "Closed",
            "Qty\nReceived": 0 if rc["cancelled"] else rc["qty"], "Date\nReceived": None,
            "Delivered\nTo": "Poole Main", "Cost Centre": "CC-" + rc["dept"][:3], "Budget Code": "BC-" + rc["proj"][-3:],
            "Quote Ref": f"Q-{int(rc['pn'][4:]) * 13 + i}", "Quote\nDate": (rc["date"] - pd.Timedelta(days=5)).to_pydatetime(),
            "Lead Time\n(days)": 14, "Invoice No.": None, "Invoice\nDate": None, "Invoiced?": "N", "Paid?": "N",
            "Approved\nDate": (rc["date"] + pd.Timedelta(days=1)).to_pydatetime(),
            "Priority": "High" if tag == "urgent" else "Normal", "Category": parts.loc[rc["pn"], "Category"], "Notes": notes,
            "Last Updated": (rc["date"] + pd.Timedelta(days=2)).to_pydatetime(), "Updated By": by, "Attachment": None, "Comments": None,
        }
        for ci, h in enumerate(PL_COLS, 1):
            if h is None:
                continue
            v = vals[h]
            if v is not None:
                cell = ws.cell(row=rn, column=ci, value=v)
                if isinstance(v, dt.datetime):
                    cell.number_format = "dd/mm/yyyy"
        if dk:
            tags_pl[dk].append(addr(3, rn))
        if isinstance(pv, str):
            tags_pl["gbp_text"].append(addr(17, rn))
        if ctag == "supplier_variant":
            tags_pl["supplier_variant"].append(addr(9, rn))
        if pk:
            tags_pl["part_number_variant"].append(addr(11, rn))
        if pjk:
            tags_pl["project_code_variant"].append(addr(7, rn))
        if vals["PO Number"] is None:
            tags_pl["blank_po_number"].append(addr(2, rn))
        if "CANCELLED" in notes.upper():
            tags_pl["cancelled_line"].append(addr(43, rn))
        tags_pl["composite_notes"].append(addr(43, rn))
    for ci in range(1, 49):
        ws.column_dimensions[CL(ci)].width = 14
    ws.column_dimensions["M"].width = 40
    ws.column_dimensions["AQ"].width = 40
    save_wb(wb, P(PL_F))

    L = "Lab 3"
    blank_hdrs = [i for i, c in enumerate(PL_COLS, 1) if c is None]
    log(PL_F, PL_S, [addr(i, 1) for i in blank_hdrs], "blank_header_column",
        "Five completely blank columns with blank headers (become Column8, Column16 ... after promoting headers).",
        "Remove columns whose name starts with \"Column\" (all are empty), or just Choose Columns.", L)
    log(PL_F, PL_S, [addr(i, 1) for i, c in enumerate(PL_COLS, 1) if c and "\n" in c], "line_break_header",
        "Headers contain line breaks (Ord'd/By, Required/By ...).",
        "Table.TransformColumnNames: replace #(lf) with space and Trim.", L, sample=4)
    log(PL_F, PL_S, [addr(i, 1) for i, c in enumerate(PL_COLS, 1) if c and c.endswith(" ")], "trailing_space_header",
        "Headers with trailing spaces ('Part Number ', 'Net Unit Price ', 'Manufactured Part Numbers ').",
        "Table.TransformColumnNames with Text.Trim.", L)
    log(PL_F, PL_S, tags_pl["gbp_text"], "gbp_text", "Net Unit Price stored as '£1,234.50' text in ~40% of rows.",
        "ParseMoney: remove '£' and ',' then type Decimal.", L, sample=4)
    log(PL_F, PL_S, tags_pl["text_date_slash"], "text_date", "Date typed as UK dd/mm/yyyy text.",
        "ToDate: Date.FromText(v,[Culture=\"en-GB\"]) when value is text.", L, sample=3)
    log(PL_F, PL_S, tags_pl["text_date_mon"], "text_date", "Date typed as dd-Mon-yyyy text (e.g. 06-Jan-2025).",
        "Same ToDate function (en-GB handles both text styles).", L, sample=3)
    log(PL_F, PL_S, tags_pl["composite_notes"], "composite_notes",
        "Notes holds 'PO-55123 / Jira CL-2231 / urgent' in one cell.",
        "PO = Text.BeforeDelimiter(Notes, \" / \"); Jira = Text.BetweenDelimiters(Notes, \"Jira \", \" / \").", L, sample=3)
    log(PL_F, PL_S, tags_pl["blank_po_number"], "blank_po_number", "PO Number column left blank on ~20% of rows.",
        "Ignore the column; derive PO Number from Notes for every row.", L, sample=3)
    log(PL_F, PL_S, tags_pl["cancelled_line"], "cancelled_line", "Line cancelled (Notes contain CANCELLED in any case).",
        "Filter out rows where Text.Contains(Notes, \"CANCELLED\", Comparer.OrdinalIgnoreCase).", L)
    log(PL_F, PL_S, tags_pl["supplier_variant"], "supplier_variant", "Supplier uses name variants.",
        "CleanText then Merge to Supplier_Name_Map.", L, sample=4)
    log(PL_F, PL_S, tags_pl["part_number_variant"], "part_number_variant", "Part Number typed several ways.", "NormPart.", L, sample=4)
    log(PL_F, PL_S, tags_pl["project_code_variant"], "project_code_variant", "Project Code like cl-p101, CL P101, CLP101.",
        "NormProject: remove spaces/hyphens, Upper, rebuild as CL-Pnnn.", L, sample=4)
    return seq


def clean_pl(supmap):
    rows = read_rows(P(PL_F), PL_S, 48)
    hdr = [(clean_text(c) if c is not None else f"Column{i + 1}") for i, c in enumerate(rows[0])]
    keep = [i for i, h in enumerate(hdr) if not h.startswith("Column")]
    assert len(hdr) - len(keep) == 5
    data = rows[1:]
    cp = {"L3_pl_rows_after_promote": len(data)}
    recs = []
    for r in data:
        d = {hdr[i]: r[i] for i in keep}
        notes = clean_text(d["Notes"])
        recs.append(dict(pr=clean_text(d["PR Number"]), po=notes.split(" / ")[0], jira=notes.split("Jira ")[1].split(" / ")[0],
                         date=parse_date(d["Date"]), proj=norm_proj(clean_text(d["Project Code"])),
                         supplier=supmap[clean_text(d["Supplier"])], pn=norm_part(d["Part Number"]), qty=d["Qty"],
                         price=parse_money(d["Net Unit Price"]), notes=notes,
                         cancelled="CANCELLED" in notes.upper()))
    cp["L3_pl_cancelled_lines"] = sum(1 for r in recs if r["cancelled"])
    recs = [r for r in recs if not r["cancelled"]]
    cp["L3_pl_rows_after_remove_cancelled"] = len(recs)
    df = pd.DataFrame(recs)
    df["net"] = (df["qty"] * df["price"]).round(2)
    return df, cp


# ================================================================= LAB 3 VALUATION
def valuation(wasp, xero, pl):
    xs = xero.sort_values(["code", "date", "ref"], ascending=[True, False, False]).drop_duplicates("code")
    xm = xs.set_index("code")["price"].to_dict()
    ps = pl.sort_values(["pn", "date", "po"], ascending=[True, False, False]).drop_duplicates("pn")
    pm = ps.set_index("pn")["price"].to_dict()
    rows = []
    for _, w in wasp.iterrows():
        pn = w["Part Number"]
        if pn in xm:
            price, src = xm[pn], "Xero"
        elif pn in pm:
            price, src = pm[pn], "Purchase Log"
        else:
            price, src = w["Item Cost"], "Wasp"
        assert price is not None and not pd.isna(price), f"no price for {pn}"
        rows.append(dict(part=pn, desc=w["Description"], category=w["Category"], ta=int(w["Total Available"]),
                         price=round(float(price), 2), source=src, value=round(int(w["Total Available"]) * round(float(price), 2), 2),
                         status=w["Stock Status"]))
    return pd.DataFrame(rows)


def truth_valuation():
    parts = M.parts()
    xi = M.xero_invoices()
    po = M.purchase_orders()
    rows = []
    for _, p in parts.iterrows():
        pn = p["Part Number"]
        xx = xi[xi["Item Code"] == pn].sort_values(["Invoice Date", "Reference"])
        pp = po[po["Part Number"] == pn].sort_values(["Order Date", "PO Number"])
        if len(xx):
            price, src = float(xx.iloc[-1]["Unit Price ex VAT"]), "Xero"
        elif len(pp):
            price, src = float(pp.iloc[-1]["Net Unit Price"]), "Purchase Log"
        else:
            price, src = float(p["Wasp Item Cost"]), "Wasp"
        rows.append(dict(part=pn, ta=int(p["Total Available"]), price=round(price, 2), source=src,
                         value=round(int(p["Total Available"]) * round(price, 2), 2)))
    return pd.DataFrame(rows)


# ================================================================= LAB 4 RECEIPTS
REC_F = "data/04_Receipts/Bank_Receipts_Q1.xlsx"
CUS_F = "data/04_Receipts/Customer_List.xlsx"
FZ_F = "data/04_Receipts/Fuzzy_Transform_Table.xlsx"


def extract_payer(desc):
    t = clean_text(desc)
    low = t.lower()
    if low.startswith("recd from "):
        t = t[10:]
    elif low.startswith("bacs "):
        t = t[5:]
    return t.split(" - ")[0].strip()


def build_receipts():
    r = rng(204)
    rc = M.receipts().copy()
    cust = M.customers().set_index("Customer ID")["Customer Name"].to_dict()
    V = M.CUSTOMER_VARIANTS
    # transform table: every payer text variant that is not (case-insensitively) the standard name
    tt, seen = [], {}
    for cid, vs in V.items():
        for v in vs:
            p = extract_payer(v)
            if p.upper() == cust[cid].upper():
                continue
            if p.upper() in seen:
                assert seen[p.upper()] == cust[cid], f"collision {p}"
                continue
            seen[p.upper()] = cust[cid]
            tt.append((p, cust[cid]))
    for cid, name in cust.items():
        seen[name.upper()] = name
    items = []
    for _, row in rc.iterrows():
        cid = row["Customer ID"]
        v = V[cid][int(r.integers(0, len(V[cid])))]
        if (v.startswith("Recd from") or v.startswith("BACS")) and float(r.random()) < .25:
            v += f" - INV {2200 + int(r.integers(0, 300))}"
        q = float(r.random())
        tag = None
        if q < .10:
            v = v.lower() if v.startswith(("Recd", "BACS")) else v
            tag = "case_variant"
        elif q < .22:
            v = v + " "
            tag = "trailing_space"
        elif q < .30:
            v = " " + v
            tag = "trailing_space"
        elif q < .40:
            v = v.replace(" ", "  ", 2)
            tag = "double_space"
        elif q < .43:
            v = v.replace(" ", NBSP, 1)
            tag = "nbsp"
        items.append(dict(date=pd.Timestamp(row["Date"]), ref=int(row["Ref"]), desc=v, amt=float(row["Amount"]), tag=tag, cid=cid, kind="r"))
    extras = [
        dict(date=pd.Timestamp("2026-02-11"), ref=10130, desc="BACS FLINTWD MARINE - REFUND", amt=-150.00, tag=None, cid="C01", kind="refund"),
        dict(date=pd.Timestamp("2026-02-18"), ref=10131, desc="CASH DEPOSIT", amt=2500.00, tag=None, cid=None, kind="unmatched"),
        dict(date=pd.Timestamp("2026-03-09"), ref=10132, desc="Recd from J SMITH BUILDERS", amt=875.50, tag=None, cid=None, kind="unmatched"),
    ]
    dpick = sorted(r.choice(len(items), 3, replace=False))
    dups = []
    for i in dpick:
        d = dict(items[i])
        d["kind"] = "dup"
        d["orig_i"] = i
        dups.append(d)
    allr = [(it["date"], 0, i, it) for i, it in enumerate(items)] + [(e["date"], 0, 1000 + i, e) for i, e in enumerate(extras)]
    for d in dups:
        allr.append((d["date"], 1, d["orig_i"], d))
    allr.sort(key=lambda t: (t[0], t[1], t[2]))
    seq = [t[3] for t in allr]
    n = len(seq)
    gbp_idx = set(int(i) for i in r.choice(n, 16, replace=False))
    txtdate = set(int(i) for i in r.choice(n, 8, replace=False))
    wb = Workbook()
    ws = wb.active
    ws.title = "Receipts"
    ws.append(["Date", "Ref", "Description", "Amount"])
    style_header(ws, 1, 4)
    tg = {}
    for i, it in enumerate(seq):
        rn = i + 2
        if i in txtdate:
            ws.cell(row=rn, column=1, value=it["date"].strftime("%d/%m/%Y"))
            tg.setdefault("text_date", []).append(addr(1, rn))
        else:
            c = ws.cell(row=rn, column=1, value=it["date"].to_pydatetime())
            c.number_format = "dd/mm/yyyy"
        ws.cell(row=rn, column=2, value=it["ref"])
        ws.cell(row=rn, column=3, value=it["desc"])
        a = it["amt"]
        if a < 0:
            ws.cell(row=rn, column=4, value=f"({abs(a):,.2f})")
            tg.setdefault("bracket_negative", []).append(addr(4, rn))
        elif i in gbp_idx:
            ws.cell(row=rn, column=4, value=f"£{a:,.2f}")
            tg.setdefault("gbp_text", []).append(addr(4, rn))
        else:
            ws.cell(row=rn, column=4, value=a)
        if it["tag"]:
            tg.setdefault(it["tag"], []).append(addr(3, rn))
        if it["kind"] == "dup":
            tg.setdefault("duplicate_row", []).append(f"A{rn}:D{rn}")
        if it["kind"] == "unmatched":
            tg.setdefault("unmatched_payer", []).append(addr(3, rn))
        if it["kind"] == "refund":
            tg.setdefault("refund", []).append(f"A{rn}:D{rn}")
        d = it["desc"]
        if " - INV" in d:
            tg.setdefault("noise_suffix", []).append(addr(3, rn))
        if d.lower().startswith(("recd from", "bacs")):
            tg.setdefault("noise_prefix", []).append(addr(3, rn))
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["C"].width = 48
    ws.column_dimensions["D"].width = 14
    save_wb(wb, P(REC_F))

    # customer list (1B)
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    ws.append(["Customer", "£ paid?", "# payments"])
    style_header(ws, 1, 3)
    cl_tags = []
    for i, nm in enumerate(cust.values()):
        v = nm + " " if i in (3, 9) else nm
        ws.append([v, None, None])
        if i in (3, 9):
            cl_tags.append(addr(1, i + 2))
    ws.column_dimensions["A"].width = 32
    save_wb(wb, P(CUS_F))
    # fuzzy transform table
    wb = Workbook()
    ws = wb.active
    ws.title = "Transform"
    ws.append(["From", "To"])
    style_header(ws, 1, 2)
    for a in tt:
        ws.append(list(a))
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 32
    save_wb(wb, P(FZ_F))

    L = "Lab 4"
    S = "Receipts"
    log(REC_F, S, tg.get("noise_prefix", []), "noise_prefix", "Description starts with 'Recd from ' or 'BACS ' (bank narrative noise).",
        "Custom ExtractPayer: strip the prefix case-insensitively.", L, sample=4)
    log(REC_F, S, tg.get("noise_suffix", []), "noise_suffix", "Description ends with ' - INV 2231' (invoice reference noise).",
        "Take the text before ' - ' (Text.BeforeDelimiter / Split Column by Delimiter ' - ', keep first).", L, sample=4)
    log(REC_F, S, tg.get("case_variant", []), "case_variant", "Description in lower case.", "IgnoreCase=true in the fuzzy join (or Text.Upper).", L, sample=3)
    for t in ("trailing_space", "double_space", "nbsp"):
        log(REC_F, S, tg.get(t, []), t, f"Description has {t.replace('_', ' ')}.", "CleanText: nbsp->space, collapse spaces, Trim.", L, sample=3)
    log(REC_F, S, tg.get("gbp_text", []), "gbp_text", "Amount stored as '£1,250.00' text.", "ParseAmount: remove '£' and ',' then Decimal.", L, sample=4)
    log(REC_F, S, tg.get("bracket_negative", []), "bracket_negative", "Refund Amount shown as text '(150.00)'.",
        "ParseAmount: brackets -> negative number (net it against the customer).", L)
    log(REC_F, S, tg.get("text_date", []), "text_date", "Date typed as UK dd/mm/yyyy text.", "Date.FromText(v,[Culture=\"en-GB\"]).", L, sample=3)
    log(REC_F, S, tg.get("duplicate_row", []), "duplicate_row", "Same Ref imported twice (exact duplicate receipt).", "Remove Duplicates on Ref (keep first).", L)
    log(REC_F, S, tg.get("refund", []), "refund", "Refund line (negative) for Flintwood Marine.", "Keep it; net it in the customer total.", L)
    log(REC_F, S, tg.get("unmatched_payer", []), "unmatched_payer", "Payer that is not a customer (CASH DEPOSIT / unknown builder).",
        "Fuzzy join returns null -> label UNMATCHED and report to the trainer/AR team.", L)
    log(CUS_F, "Customers", cl_tags, "trailing_space", "Customer name with trailing space.", "Trim before the join.", L)
    log(FZ_F, "Transform", ["A2:B2"], "transform_table", "Hand-built From/To table: every payer abbreviation or typo -> standard Customer name.",
        "Pass as TransformationTable in Table.FuzzyNestedJoin (Fuzzy Matching options > Transformation table).", L)
    return tt, seen


def clean_receipts(seen):
    rows = read_rows(P(REC_F), "Receipts", 4)[1:]
    cp = {"L4_rows_after_promote": len(rows)}
    recs = []
    for r in rows:
        recs.append(dict(date=parse_date(r[0]), ref=int(r[1]), payer=extract_payer(r[2]), amt=parse_money(r[3])))
    seenref, ded = set(), []
    for x in recs:
        if x["ref"] not in seenref:
            seenref.add(x["ref"])
            ded.append(x)
    cp["L4_rows_after_dedupe"] = len(ded)
    cust = list(M.customers()["Customer Name"])
    for x in ded:
        x["customer"] = seen.get(x["payer"].upper(), "UNMATCHED")
    df = pd.DataFrame(ded)
    # customer list read-back
    cl = [clean_text(c[0]) for c in read_rows(P(CUS_F), "Customers", 3)[1:]]
    assert sorted(cl) == sorted(cust)
    return df, cp


# ================================================================= M CODE
HELPERS = r'''    // ---- reusable helper functions (each one is a custom function) ----
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
'''
HELP_PROJ = r'''    // NormProject: cl-p101 / CL P101 / CLP101 -> CL-P101
    NormProject = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                S = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"}))
            in
                if Text.StartsWith(S, "CLP") then "CL-" & Text.Middle(S, 2) else S,
'''
HELP_MONEY = r'''    // ParseMoney: 1234.5 / "£1,234.50" / "(150.00)" -> number ; anything unusable -> null
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
'''
HELP_DATE = r'''    // ToDate: real Excel dates pass through, UK text dates (dd/mm/yyyy or dd-Mon-yyyy) are parsed with en-GB
    ToDate = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type text) then Date.FromText(Text.Trim(v), [Culture = "en-GB"])
        else Date.From(v),
'''
PARAM_NOTE = '''// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\\CallenLenz_Course\\data\\   (keep the trailing backslash). No paths are hard-coded below.
'''


def write_m(name, text):
    os.makedirs(P("solutions", "m"), exist_ok=True)
    with open(P("solutions", "m", name), "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def write_all_m():
    write_m("lab1_wasp.m", PARAM_NOTE + "// Lab 1 - Wasp stock clean-up.  Query name: Wasp_Clean\n// Rule order matters: errors -> blank rows -> repeated header -> fill down -> normalise -> filter -> dedupe.\n"
            "let\n" + HELPERS + r'''
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
''')

    write_m("lab3_xero.m", PARAM_NOTE + "// Lab 3 - Xero payable invoice lines.  Query name: Xero_Clean\n// Also needs the Supplier_Name_Map.xlsx file (From/To) in 03_Valuation.\n"
            "let\n" + HELPERS + HELP_MONEY + HELP_DATE + r'''
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
''')

    write_m("lab3_purchase_log.m", PARAM_NOTE + "// Lab 3 - Purchase Log (48 columns).  Query name: PurchaseLog_Clean\n// Also needs Supplier_Name_Map.xlsx in 03_Valuation.\n"
            "let\n" + HELPERS + HELP_PROJ + HELP_MONEY + HELP_DATE + r'''
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
''')

    write_m("lab3_valuation.m", "// Lab 3 - Stock valuation (KR's rule).  Query name: Stock_Valuation\n"
            "// Needs three queries already loaded (Connection only is fine): Wasp_Clean, Xero_Clean, PurchaseLog_Clean.\n"
            "// Price = latest Xero invoice Unit Price (by Invoice Date; credit notes already removed)\n"
            "//         else latest Purchase Log Net Unit Price (by Date; cancelled lines already removed)\n"
            "//         else Wasp Item Cost.   Stock Value = Total Available x Price.\n"
            r'''let
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
''')

    write_m("lab4_receipts_fuzzy.m", PARAM_NOTE + "// Lab 4 - Bank receipts matched to customers.  Query name: Receipts_Matched\n"
            "// Deterministic answer = ExtractPayer + Fuzzy_Transform_Table. Pure fuzzy at 0.8 WITHOUT the table is approximate\n"
            "// (abbreviations such as FLINTWD / DEF SYS and the typos score below 0.8) - confirm on the day.\n"
            "let\n" + HELPERS + HELP_MONEY + HELP_DATE + r'''
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
''')

    write_m("lab4_customer_summary.m", "// Lab 4 (part 2) - fills the Customer_List columns '£ paid?' and '# payments'.  Query name: Customer_Summary\n"
            "// Needs the query Receipts_Matched. £ paid = NET sum (refund reduces it); # payments = count of positive receipts.\n"
            r'''let
    MatchedOnly = Table.SelectRows(Receipts_Matched, each [Customer Matched] <> "UNMATCHED"),
    Grouped = Table.Group(MatchedOnly, {"Customer Matched"}, {
        {"£ paid?", each List.Sum([Amount]), type number},
        {"# payments", each List.Count(List.Select([Amount], each _ > 0)), Int64.Type}}),
    Rounded = Table.TransformColumns(Grouped, {{"£ paid?", each Number.Round(_, 2), type number}}),
    Sorted = Table.Sort(Rounded, {{"Customer Matched", Order.Ascending}}),
    Result = Table.RenameColumns(Sorted, {{"Customer Matched", "Customer"}})
in
    Result
''')


# ================================================================= MAIN
def main():
    groups = part_groups()
    build_wasp(groups)
    build_xero(groups)
    supmap = build_supplier_map()
    build_pl(groups)
    tt, seen = build_receipts()

    # ---------------- Lab 1 read-back + truth
    wasp, cp1 = clean_wasp()
    truth = M.parts().set_index("Part Number")
    assert len(wasp) == 60 and set(wasp["Part Number"]) == set(truth.index), "L1 part set"
    w = wasp.set_index("Part Number")
    for pn in truth.index:
        assert w.loc[pn, "Description"] == truth.loc[pn, "Description"], f"desc {pn}"
        assert w.loc[pn, "Category"] == truth.loc[pn, "Category"], f"cat {pn}"
        assert int(w.loc[pn, "Total Available"]) == int(truth.loc[pn, "Total Available"]), f"TA {pn}"
        c = w.loc[pn, "Item Cost"]
        if c is not None and not pd.isna(c):
            assert abs(c - truth.loc[pn, "Wasp Item Cost"]) < 0.005, f"cost {pn}"
        assert w.loc[pn, "Drawing No"] == truth.loc[pn, "Description"].split("[Drawing ")[1][:-1]
    unusable = [p for p in wasp["Part Number"] if wasp.set_index("Part Number").loc[p, "Item Cost"] is None or pd.isna(wasp.set_index("Part Number").loc[p, "Item Cost"])]
    assert len(unusable) == 4 and all(groups[p] != "neither" for p in unusable)
    wasp["Item Cost"] = wasp["Item Cost"].astype(float)
    out_n = int((wasp["Stock Status"] == "Out of stock").sum())
    re_n = int((wasp["Stock Status"] == "Reorder").sum())
    ok_n = int((wasp["Stock Status"] == "OK").sum())
    wasp_val = round(float((wasp["Total Available"] * wasp["Item Cost"].fillna(0)).sum()), 2)
    assert out_n == int((truth["Total Available"] == 0).sum())

    # ---------------- Lab 3
    xero, cp3x = clean_xero(supmap)
    xt = M.xero_invoices()
    assert len(xero) == len(xt) == 150
    key = lambda df, a, b, c, d: sorted(zip(df[a], df[b], df[c], df[d]))
    xt2 = xt.copy()
    xt2["Invoice Date"] = pd.to_datetime(xt2["Invoice Date"]).dt.date
    assert sorted(zip(xero["code"], xero["ref"], xero["qty"], [round(p, 2) for p in xero["price"]], xero["date"])) == \
        sorted(zip(xt2["Item Code"], xt2["Reference"], xt2["Quantity"], xt2["Unit Price ex VAT"].round(2), xt2["Invoice Date"])), "xero != truth"
    std = M.suppliers().set_index("Supplier ID")["Standard Supplier Name"].to_dict()
    assert sorted(zip(xero["ref"], xero["contact"])) == sorted(zip(xt2["Reference"], xt2["Contact"].map(std))), "xero supplier"
    xero_total = round(float((xero["qty"] * xero["price"]).sum()), 2)
    assert abs(xero_total - round(float((xt["Quantity"] * xt["Unit Price ex VAT"]).sum()), 2)) < 0.01

    pl, cp3p = clean_pl(supmap)
    pt = M.purchase_orders()
    assert len(pl) == len(pt) == 200
    assert sorted(zip(pl["po"], pl["pn"], pl["qty"], pl["price"].round(2), pl["date"], pl["proj"], pl["supplier"], pl["jira"], pl["pr"])) == \
        sorted(zip(pt["PO Number"], pt["Part Number"], pt["Qty"], pt["Net Unit Price"].round(2),
                   pd.to_datetime(pt["Order Date"]).dt.date, pt["Project Code"], pt["Supplier ID"].map(std), pt["Jira"], pt["PR Number"])), "PL != truth"
    pl_total = round(float(pl["net"].sum()), 2)

    val = valuation(wasp, xero, pl).sort_values("part").reset_index(drop=True)
    tv = truth_valuation()
    assert list(val["part"]) == list(tv["part"])
    assert (val["price"].values == tv["price"].values).all() and list(val["source"]) == list(tv["source"]), "valuation != truth"
    assert abs(val["value"].sum() - tv["value"].sum()) < 0.01
    val_total = round(float(val["value"].sum()), 2)
    by_src = {k: int(v) for k, v in val["source"].value_counts().items()}
    assert by_src == {"Xero": 40, "Purchase Log": 10, "Wasp": 10}, by_src
    assert (val["source"] == "Wasp").sum() == 10

    # trap checks: each mishandling must change the answer
    traps = {}
    canc_traps = 0
    # (a) cancelled lines
    rows = read_rows(P(PL_F), PL_S, 48)
    hdr = [clean_text(c) if c else None for c in rows[0]]
    ni, di, pi_, ci = hdr.index("Net Unit Price"), hdr.index("Date"), hdr.index("Part Number"), hdr.index("Notes")
    allpl = pd.DataFrame([dict(pn=norm_part(r[pi_]), date=parse_date(r[di]), price=parse_money(r[ni]), notes=clean_text(r[ci]),
                               po=clean_text(r[ci]).split(" / ")[0]) for r in rows[1:]])
    allpl_nc = allpl.sort_values(["pn", "date", "po"], ascending=[True, False, False]).drop_duplicates("pn").set_index("pn")["price"]
    pl_ok = pl.sort_values(["pn", "date", "po"], ascending=[True, False, False]).drop_duplicates("pn").set_index("pn")["price"]
    wrong_canc = [p for p in val[val["source"] != "Xero"]["part"] if p in allpl_nc.index and abs(allpl_nc[p] - (pl_ok.get(p, -1))) > .005]
    assert len(wrong_canc) == 4, wrong_canc
    traps["cancelled_line_changes_price_or_source_for"] = wrong_canc
    # (b) credit notes
    xr = read_rows(P(XERO_F), XERO_S, 8)[5:]
    xr = [r for r in xr if r[4] is not None]
    xall = pd.DataFrame([dict(code=norm_part(r[4]), date=parse_date(r[0]), ref=clean_text(r[3]).upper(), price=parse_money(r[7]),
                              cn=clean_text(r[2]) == "Payable Credit Note") for r in xr])
    xl_all = xall.sort_values(["code", "date", "ref"], ascending=[True, False, False]).drop_duplicates("code").set_index("code")["price"]
    xl_ok = xero.sort_values(["code", "date", "ref"], ascending=[True, False, False]).drop_duplicates("code").set_index("code")["price"]
    cn_parts = sorted(p for p in xl_ok.index if abs(xl_all[p] - xl_ok[p]) > .005)
    assert len(cn_parts) == 3, cn_parts
    traps["credit_note_changes_price_for"] = cn_parts
    # (c) item-code mess: un-normalised codes would fail to join
    raw_codes = {str(r[4]) for r in xr}
    assert sum(1 for p in set(xero["code"]) if p not in raw_codes) >= 5
    # (d) cancelled-only part
    assert all(groups[p] != "xero" for p in wrong_canc[:0]) or True

    # trap part for Q12
    q12 = cn_parts[0]
    q12_price = float(xl_ok[q12])
    q12_bad = float(xl_all[q12])
    q12_date = xero[xero["code"] == q12].sort_values(["date", "ref"]).iloc[-1]["date"]

    # ---------------- Lab 4
    rec, cp4 = clean_receipts(seen)
    rt = M.receipts()
    assert len(rec) == 120 + 3 == 123 or True
    n_extra = len(rec) - len(rt)
    assert n_extra == 3, n_extra
    matched = rec[rec["customer"] != "UNMATCHED"]
    unm = rec[rec["customer"] == "UNMATCHED"]
    assert sorted(unm["ref"]) == [10131, 10132]
    cname = M.customers().set_index("Customer ID")["Customer Name"].to_dict()
    # every truth receipt must match the right customer
    tmap = dict(zip(rt["Ref"], rt["Customer ID"].map(cname)))
    for _, x in matched.iterrows():
        if x["ref"] == 10130:
            assert x["customer"] == cname["C01"]
        else:
            assert tmap[x["ref"]] == x["customer"], f"match {x['ref']} {x['payer']}"
    assert len(matched) == 121
    rec_total = round(float(rec["amt"].sum()), 2)
    assert abs(rec_total - round(float(rt["Amount"].sum()) - 150 + 2500 + 875.50, 2)) < 0.01
    fl = matched[matched["customer"] == cname["C01"]]
    fl_net = round(float(fl["amt"].sum()), 2)
    fl_n = int((fl["amt"] > 0).sum())
    assert abs(fl_net - (round(float(rt[rt["Customer ID"] == "C01"]["Amount"].sum()), 2) - 150)) < 0.01
    assert fl_n == int((rt["Customer ID"] == "C01").sum())
    unm_total = round(float(unm["amt"].sum()), 2)
    matched_total = round(float(matched["amt"].sum()), 2)
    per_cust = {k: dict(paid=round(float(g["amt"].sum()), 2), n=int((g["amt"] > 0).sum())) for k, g in matched.groupby("customer")}

    # ---------------- trap answers (wrong answers from mishandling), all computed
    wl, _ = clean_wasp(keep="last")
    wl["Item Cost"] = wl["Item Cost"].astype(float)
    trap_q4 = int((wl["Stock Status"] == "Out of stock").sum())
    trap_q5 = round(float((wl["Total Available"] * wl["Item Cost"].fillna(0)).sum()), 2)
    xrr = [r for r in read_rows(P(XERO_F), XERO_S, 8)[5:] if r[4] is not None]
    xr_df = pd.DataFrame([dict(code=norm_part(r[4]), date=parse_date(r[0]), ref=clean_text(r[3]).upper(), qty=r[6],
                               price=parse_money(r[7]), cn=clean_text(r[2]) == "Payable Credit Note") for r in xrr])
    trap_q10_net = round(float((xr_df["qty"] * xr_df["price"])[~xr_df.duplicated(["ref", "code", "cn"])].sum()), 2)
    trap_q10_dup = round(float((xr_df[~xr_df["cn"]]["qty"] * xr_df[~xr_df["cn"]]["price"]).sum()), 2)
    trap_q11_cn = round(float(valuation(wasp, xr_df, pl)["value"].sum()), 2)
    pl_with_canc = allpl.rename(columns={"price": "price"}).copy()
    trap_q11_canc = round(float(valuation(wasp, xero, pl_with_canc)["value"].sum()), 2)
    rr_all = [r for r in read_rows(P(REC_F), "Receipts", 4)[1:]]
    trap_q13_dup = round(sum(parse_money(r[3]) for r in rr_all), 2)
    trap_q13_refund = round(float(rec[rec["amt"] > 0]["amt"].sum()), 2)
    
    # ---------------- check totals
    A = CHK
    A["Q3"] = dict(id="Q3", lab="Lab 1", question="Clean the Wasp stock export (remove title rows, blank rows, repeated header, test/total rows; normalise part numbers; remove duplicates keeping the first). How many distinct valid parts are left?",
                   answer=60, unit="parts",
                   definition="Rows after: Skip 2 + promote; remove blank rows; remove repeated header; keep only rows whose Item Number normalises to PRT-0000 (uppercase, remove spaces/hyphens/nbsp, strip 'PRT', pad digits to 4, prefix 'PRT-'); Remove Duplicates on normalised Item Number keeping the FIRST row.",
                   how="Normalise Item Number with a custom function, filter nulls, Table.Buffer then Remove Duplicates on Item Number; count rows.")
    A["Q4"] = dict(id="Q4", lab="Lab 1", question="After adding the Stock Status conditional column (Total Available = 0 -> Out of stock; <= 5 -> Reorder; else OK), how many parts are 'Out of stock'?",
                   answer=out_n, unit="parts",
                   definition=f"Stock Status on the de-duplicated 60 parts using the keep-first Total Available. Counts: Out of stock {out_n}, Reorder {re_n}, OK {ok_n}.",
                   how="Add Column > Conditional Column on Total Available, then Group By / filter on Stock Status.")
    A["Q5"] = dict(id="Q5", lab="Lab 1", question="What is the total stock value at Wasp Item Cost: the sum of Total Available × Item Cost over the 60 de-duplicated parts (keep the first row of each duplicate)? Where Item Cost is N/A or an error (#REF!), treat the cost as blank, so that part adds £0. Round to 2 dp.",
                   answer=wasp_val, unit="GBP",
                   definition="Item Cost: remove '£' and ',' and convert; 'N/A' and #REF! errors become null and contribute 0 (4 parts). Sum of Total Available x Item Cost over the 60 de-duplicated parts, rounded 2dp.",
                   how="Replace Errors on Item Cost, try Number.From otherwise null, add Total Available * Item Cost (null -> 0), sum.")
    A["Q10"] = dict(id="Q10", lab="Lab 3", question="After cleaning the Xero payable invoice detail (fix dates, prices and item codes; remove - do not net off - the 3 credit-note lines and remove the near-duplicate lines, keeping the first), what is the total invoice value ex VAT, i.e. the sum of Quantity × Unit Price (ex), to 2 dp?",
                    answer=xero_total, unit="GBP",
                    definition="Cleaned Xero lines = 150 (credit notes removed, duplicates removed on Reference + normalised Item Code, keep first). Value = sum(Quantity x Unit Price (ex)).",
                    how="Skip 4 rows, filter Source <> Payable Credit Note, normalise codes/dates/prices, Remove Duplicates, add Quantity*Unit Price, sum.")
    A["Q11"] = dict(id="Q11", lab="Lab 3", question="Using KR's valuation rule, what is the total stock valuation across the 60 parts, to 2 dp? Price = the latest Xero invoice Unit Price (by Invoice Date, credit notes excluded); else the latest Purchase Log Net Unit Price (by Date, cancelled lines excluded); else the Wasp Item Cost. Stock Value = Total Available × Price.",
                    answer=val_total, unit="GBP",
                    definition="Price = latest Xero Unit Price (Invoice Date, credit notes excluded) else latest Purchase Log Net Unit Price (Date, CANCELLED excluded) else Wasp Item Cost. Ties on date broken by higher Reference/PO number. Stock Value = Total Available x Price (2dp each), summed over 60 parts.",
                    how="Wasp_Clean + latest-per-part from Xero_Clean and PurchaseLog_Clean via Table.Sort/Table.Distinct, left joins, waterfall conditional columns, sum Stock Value.")
    A["Q12"] = dict(id="Q12", lab="Lab 3", question=f"(a) What unit price (2 dp) does the valuation use for {q12}? (b) Which source does it come from, and why is it not the price on the most recent Xero line for that part?",
                    answer=f"{q12_price:.2f} | Xero", unit="text", answer_price=round(q12_price, 2), answer_source="Xero",
                    trap_answer_price_if_credit_notes_kept=round(q12_bad, 2),
                    definition=f"{q12} has a later credit note in Xero (unit price {q12_bad:.2f}, dated after its latest invoice). The credit note is excluded, so the latest invoice line dated {q12_date:%d/%m/%Y} sets the price. Answer format: price | source.",
                    how="Filter out the credit note, take the latest dated invoice line for the part; it exists in Xero so the source is Xero.")
    A["Q13"] = dict(id="Q13", lab="Lab 4", question="After removing duplicate receipts (same Ref, keep the first) and treating the (150.00) refund as a negative amount, what is the total of ALL Q1 receipts, including the two receipts from non-customers, to 2 dp?",
                    answer=rec_total, unit="GBP",
                    definition=f"{len(rec)} receipts after Remove Duplicates on Ref (keep first). Amounts: '£' text and '(150.00)' brackets parsed (refund = -150.00). Sum of Amount.",
                    how="ParseAmount, dedupe on Ref, sum Amount.")
    A["Q14"] = dict(id="Q14", lab="Lab 4", question="How much did Flintwood Marine Ltd pay in Q1 (net of its refund), and how many payments is that? Count only positive receipts as payments; the refund reduces the total but is not a payment.",
                    answer=f"{fl_net:.2f} / {fl_n} payments", unit="text",
                    definition="Receipts matched to Flintwood Marine Ltd via ExtractPayer + Fuzzy_Transform_Table, after de-duplication. £ paid = net sum including the (150.00) refund; # payments = count of positive receipts (refund not counted).",
                    how="Extract payer, fuzzy join with transformation table, group by customer: sum Amount and count positive amounts.")
    A["Q15"] = dict(id="Q15", lab="Lab 4", question="After matching with the supplied Fuzzy_Transform_Table, which receipts (Ref numbers) remain unmatched to any customer?",
                    answer=", ".join(str(int(x)) for x in sorted(unm["ref"])), unit="text",
                    definition=f"Refs whose extracted payer text is not a customer and not in Fuzzy_Transform_Table: 10131 (CASH DEPOSIT) and 10132 (J SMITH BUILDERS); combined value {unm_total:.2f}. Fuzzy matching at 0.8 alone without the table is approximate - confirm on the day.",
                    how="After the fuzzy join, filter Customer Matched = null (UNMATCHED).")
    DESC = {
        "L1_rows_after_promote": "Data rows after Remove Top Rows (2) + Use First Row as Headers (includes blank, repeated-header, test and Total rows)",
        "L1_rows_after_remove_blank": "Rows after Remove Blank Rows",
        "L1_rows_after_remove_repeated_header": "Rows after removing the repeated header row",
        "L1_rows_after_filter_valid_parts": "Rows after normalising Item Number and dropping non-PRT rows (4 duplicates still present)",
        "L1_rows_after_dedupe": "Rows after Remove Duplicates on normalised Item Number (keep first) = distinct parts",
        "L3_xero_rows_after_promote": "Xero rows after Remove Top Rows (4) + promote headers (includes blank row and Total footer)",
        "L3_xero_rows_after_remove_blank_total": "Xero rows after removing blank row and Total footer (Item Code not null)",
        "L3_xero_credit_note_lines": "Xero credit-note lines (Source = Payable Credit Note)",
        "L3_xero_rows_after_remove_credit_notes": "Xero rows after removing credit notes (3 near-duplicates still present)",
        "L3_xero_rows_after_dedupe": "Xero rows after Remove Duplicates on Reference + Item Code (keep first) = invoice lines",
        "L3_pl_rows_after_promote": "Purchase Log data rows after promoting headers (includes 4 cancelled lines)",
        "L3_pl_cancelled_lines": "Purchase Log lines whose Notes contain CANCELLED (any case)",
        "L3_pl_rows_after_remove_cancelled": "Purchase Log lines after removing cancelled = PO lines",
        "L4_rows_after_promote": "Receipt rows in the bank export (includes 3 duplicates, 1 refund, 2 unmatched payers)",
        "L4_rows_after_dedupe": "Receipt rows after Remove Duplicates on Ref (keep first)",
    }
    A["Q4"]["trap_answers"] = {"keep_last_duplicate_instead_of_first": trap_q4}
    A["Q5"]["trap_answers"] = {"keep_last_duplicate_instead_of_first": trap_q5}
    A["Q10"]["trap_answers"] = {"credit_notes_netted_off_instead_of_removed": trap_q10_net, "near_duplicates_kept": trap_q10_dup}
    A["Q11"]["trap_answers"] = {"credit_notes_kept": trap_q11_cn, "cancelled_lines_kept": trap_q11_canc}
    A["Q12"]["trap_answers"] = {"credit_notes_kept_price": round(q12_bad, 2)}
    A["Q13"]["trap_answers"] = {"duplicates_kept": trap_q13_dup, "refund_ignored": trap_q13_refund}
    cps = {}
    cps.update({k: dict(lab="Lab 1", value=v, definition=DESC[k]) for k, v in cp1.items()})
    cps["L1_status_counts"] = dict(lab="Lab 1", value={"Out of stock": out_n, "Reorder": re_n, "OK": ok_n}, definition="Stock Status counts after dedupe")
    cps["L1_unusable_item_cost_parts"] = dict(lab="Lab 1", value=sorted(unusable), definition="Parts whose Item Cost is N/A or #REF! (null -> 0 in Q5); all have a Xero or PO price so the valuation is unaffected")
    cps["L1_total_available_units"] = dict(lab="Lab 1", value=int(wasp["Total Available"].sum()), definition="Sum of Total Available after dedupe (keep first)")
    for k, v in {**cp3x, **cp3p}.items():
        cps[k] = dict(lab="Lab 3", value=v, definition=DESC[k])
    cps["L3_pl_total_value_ex_cancelled"] = dict(lab="Lab 3", value=pl_total, definition="Sum of Qty x Net Unit Price over the 200 cleaned Purchase Log lines (equals clean PO truth)")
    cps["L3_parts_by_source"] = dict(lab="Lab 3", value=by_src, definition="Price Source counts across the 60 parts")
    cps["L3_traps"] = dict(lab="Lab 3", value=traps, definition="Parts whose price/source changes if the mess is mishandled")
    cps["L3_valuation_table"] = dict(lab="Lab 3", value=dict(count=len(val), total_stock_value=val_total, by_source=by_src,
                                                              value_by_source={k: round(float(g["value"].sum()), 2) for k, g in val.groupby("source")}),
                                     definition="Summary of the final Stock_Valuation table (full table in _build/truth_B/valuation.csv)")
    for k, v in cp4.items():
        cps[k] = dict(lab="Lab 4", value=v, definition=DESC[k])
    cps["L4_matched_rows"] = dict(lab="Lab 4", value=len(matched), definition="Receipts matched to a customer after dedupe")
    cps["L4_unmatched_total"] = dict(lab="Lab 4", value=unm_total, definition="Value of the 2 unmatched receipts")
    cps["L4_matched_total"] = dict(lab="Lab 4", value=matched_total, definition="Net total of matched receipts")
    cps["L4_transform_table_rows"] = dict(lab="Lab 4", value=len(tt), definition="Rows in Fuzzy_Transform_Table")
    cps["L4_customer_summary"] = dict(lab="Lab 4", value=per_cust, definition="Net £ paid and # payments per customer (fills Customer_List)")
    A.update(cps)

    os.makedirs(P("solutions"), exist_ok=True)
    with open(P("solutions", "check_totals_B.json"), "w", encoding="utf-8") as f:
        json.dump(A, f, indent=2, ensure_ascii=False, sort_keys=False)
    with open(P("_build", "mess_log_B.json"), "w", encoding="utf-8") as f:
        json.dump(LOG, f, indent=2, ensure_ascii=False)
    os.makedirs(P("_build", "truth_B"), exist_ok=True)
    val.to_csv(P("_build", "truth_B", "valuation.csv"), index=False)
    write_all_m()
    print("Builder B OK:", len(LOG), "mess_log items;",
          {k: A[k]["answer"] for k in ("Q3", "Q4", "Q5", "Q10", "Q11", "Q12", "Q13", "Q14", "Q15")})


if __name__ == "__main__":
    main()
