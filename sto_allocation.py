"""
Daily STO Allocation builder.

Reads the five daily input files from the `input` folder and writes the
completed STO Allocation workbook (Data + WMS% + Summary sheets) to `output`.

    python sto_allocation.py
    python sto_allocation.py --input "C:/path/to/input" --output "C:/path/to/output"

Input files are recognised by name (case-insensitive):
    STO allocation export (Algonomy) : contains "STO_Allocation_Orders" (and not "Final")
    Article report                   : contains "Article_Report"
    Store list                       : contains "Store_list"
    OOS root cause report            : contains "Root_Cause"
    DC BOH                           : contains "DC_BOH"
(spaces, dashes and underscores are treated the same)
"""
import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

FILE_PATTERNS = {
    "export": "sto_allocation_orders",
    "article": "article_report",
    "store": "store_list",
    "oos": "root_cause",
    "boh": "dc_boh",
}

DATA_COLUMNS = [
    "Article (SKU)", "Article Description", "Warehouse", "Warehouse Description",
    "Store", "Store Description", "Store LEGS", "Store Delivery Date",
    "Planned Quantity (Selling Units)", "Overwrite", "Planned Quantity (Buying Units)",
    "Unit Buying Price", "Order Value", "Vendor Code", "Allocation Effective Date",
    "Ordering Cycle", "Cap based on vertical facing", "WMS %", "SUF",
    "Final Allocation Qty", "Department", "Category", "SUB CAT", "Child SKU?",
    "COMBINE Article, Store & EFFECTIVE DATE", "LEGS", "Region", "Combine",
    "SOH", "GIT", "DC Pending Today", "DC Pending Pick", "PUF", "by PUF",
    "DL", "OB", "TOP 40 OOS", "Total SOH", "OOS RCA",
]
COL = {name: i for i, name in enumerate(DATA_COLUMNS)}


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def norm(s):
    return re.sub(r"[\s\-]+", "_", str(s).strip().lower())


def find_inputs(folder: Path):
    files = [f for f in folder.iterdir()
             if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsb", ".xls", ".xlsm")
             and not f.name.startswith("~$")]
    found = {}
    for key, pat in FILE_PATTERNS.items():
        hits = [f for f in files if pat in norm(f.name)]
        if key == "export":
            hits = [f for f in hits if "final" not in norm(f.name)]
        if not hits:
            sys.exit(f"ERROR: no {key} file found in {folder} (name must contain '{pat}')")
        if len(hits) > 1:
            hits.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            log(f"WARNING: several {key} files found, using newest: {hits[0].name}")
        found[key] = hits[0]
    return found


def read(path, sheet=0, header=None):
    return pd.read_excel(path, sheet_name=sheet, header=header, engine="calamine", dtype=object)


def to_int(series):
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def to_num(series):
    return pd.to_numeric(series, errors="coerce")


def header_row(df, first_value, must_contain):
    for i in range(min(len(df), 100)):
        row = [str(v).strip() for v in df.iloc[i].tolist()]
        if first_value in row and must_contain in row:
            return i
    raise ValueError(f"could not find header row containing '{first_value}'")


# ---------------------------------------------------------------- loaders
def load_export(path):
    raw = read(path)
    h = header_row(raw, "Article (SKU)", "Store")
    meta = " ".join(str(v) for v in raw.iloc[:h, 0].dropna())
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[h]]
    df = df.loc[:, [c for c in df.columns if c and c != "nan"]]
    df = df[to_num(df["Article (SKU)"]).notna()].reset_index(drop=True)
    m = re.search(r"Plan Created Date:\s*(\d{1,2}-\w{3}-\d{4})", meta)
    plan_date = datetime.strptime(m.group(1), "%d-%b-%Y") if m else datetime.now()
    return df, plan_date


def flag(series):
    """'X' if the article report column is ticked, else 0 (as in the template)."""
    return series.astype(str).str.strip().str.upper().eq("X").map({True: "X", False: 0}).astype(object)


def load_article(path):
    df = read(path, sheet="Article Report", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    # Columns by position as used in the manual process (AH, AJ, AL, CU),
    # cross-checked against the header names.
    expected = {33: "Department", 35: "Category", 37: "Sub-category", 98: "Article Category"}
    for pos, name in expected.items():
        if df.columns[pos] != name:
            log(f"WARNING: Article report column {pos + 1} is '{df.columns[pos]}', expected '{name}'")
    out = pd.DataFrame({
        "art": to_int(df["Article Number"]),
        "Department": df.iloc[:, 33],
        "Category": df.iloc[:, 35],
        "SUB CAT": df.iloc[:, 37],
        "Child SKU?": df.iloc[:, 98],
        "SUF": to_num(df["Delivery Rate against Basic UoM (SUF/IN1)"]),
        "PUF": to_num(df["Ordering UOM Rate against Basic UoM (PUF)"]),
        "DL": flag(df["Key Line Item"]),
        "OB": flag(df["Corporate Brand"]),
    })
    return out.dropna(subset=["art"]).drop_duplicates("art")


def load_store(path):
    df = read(path, sheet=0, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    cycle_col = next(c for c in df.columns if c.lower().startswith("ordering cycle"))
    out = pd.DataFrame({
        "st": to_int(df["Store code"]),
        "Ordering Cycle": df[cycle_col],
        "LEGS": df["NEW LEGS"],
        "Region": df["Site Hierarchy L2"],
    })
    return out.dropna(subset=["st"]).drop_duplicates("st")


def load_oos(path):
    raw = read(path, sheet="OOS ROOT CAUSE REPORT")
    h = header_row(raw, "Article No", "Site No")
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[h]]
    art, site = to_int(df["Article No"]), to_int(df["Site No"])
    ok = art.notna() & site.notna()
    return set(art[ok].astype(str) + site[ok].astype(str))


def load_boh(path):
    df = read(path, sheet=0, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    # "Available to be pick" = Available - Stock In Stage & Others
    atp = to_num(df["Available"]).fillna(0) - to_num(df["Stock In Stage & Others"]).fillna(0)
    out = pd.DataFrame({"art": to_int(df["Item"]), "DC BOH": atp})
    return out.dropna(subset=["art"]).groupby("art", as_index=False)["DC BOH"].sum()


# ---------------------------------------------------------------- logic
def mround(x, m):
    """Excel MROUND (half away from zero; multiple 0 -> 0)."""
    x = np.asarray(x, float)
    m = np.asarray(m, float)
    safe = np.where((m == 0) | np.isnan(m), 1, m)
    r = np.sign(x) * np.floor(np.abs(x) / safe + 0.5 + 1e-9) * safe
    return np.where((m == 0), 0, r)


def build(exp, art, store, oos_keys, boh):
    d = pd.DataFrame()
    d["Article (SKU)"] = to_int(exp["Article (SKU)"])
    d["Article Description"] = exp["Article Description"]
    d["Warehouse"] = exp["Warehouse"].astype(str)
    d["Warehouse Description"] = exp["Warehouse Description"]
    d["Store"] = to_int(exp["Store"])
    d["Store Description"] = exp["Store Description"]
    d["Store LEGS"] = exp["Store LEGS"]
    d["Store Delivery Date"] = exp["Store Delivery Date"].astype(str)
    d["Planned Quantity (Selling Units)"] = to_num(exp["Planned Quantity (Selling Units)"])
    d["Overwrite"] = d["Planned Quantity (Selling Units)"]
    d["Planned Quantity (Buying Units)"] = to_num(exp["Planned Quantity (Buying Units)"])
    d["Unit Buying Price"] = to_num(exp["Unit Buying Price"])
    d["Order Value"] = to_num(exp["Order Value"])
    d["Vendor Code"] = exp["Vendor Code"].astype(str)
    d["Allocation Effective Date"] = exp["Allocation Effective Date"].astype(str)

    d = d.merge(store.rename(columns={"st": "Store"}), on="Store", how="left")
    d = d.merge(art.rename(columns={"art": "Article (SKU)"}), on="Article (SKU)", how="left")
    d["Cap based on vertical facing"] = 0

    # WMS% sheet: pivot of Overwrite by article + DC BOH
    wms = (d.groupby("Article (SKU)", as_index=False)
             .agg(desc=("Article Description", "first"), total=("Overwrite", "sum"))
             .sort_values("Article (SKU)"))
    wms = wms.merge(boh.rename(columns={"art": "Article (SKU)"}), on="Article (SKU)", how="left")
    wms["DC BOH"] = wms["DC BOH"].fillna(0)
    with np.errstate(all="ignore"):
        wms["WMS%"] = np.minimum(wms["DC BOH"] / wms["total"], 1.0)
    d["WMS %"] = d["Article (SKU)"].map(wms.set_index("Article (SKU)")["WMS%"])

    jr = d["Overwrite"] * d["WMS %"]
    with np.errstate(all="ignore"):
        d["by PUF"] = jr / d["PUF"]
    d["Final Allocation Qty"] = np.where(d["by PUF"] >= 0.8, mround(jr, d["PUF"]), mround(jr, d["SUF"]))

    d["Combine"] = d["Article (SKU)"].astype(str) + d["Store"].astype(str)
    d["COMBINE Article, Store & EFFECTIVE DATE"] = d["Combine"] + "_" + d["Allocation Effective Date"]
    in_oos = d["Combine"].isin(oos_keys)
    d["OOS RCA"] = np.where(in_oos, d["Combine"], None)
    d["Total SOH"] = np.where(in_oos, 0, 1)
    for c in ["SOH", "GIT", "DC Pending Today", "DC Pending Pick", "TOP 40 OOS"]:
        d[c] = None
    return d[DATA_COLUMNS], wms


# ---------------------------------------------------------------- writer
def clean(v):
    if v is None or v is pd.NA:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v)
    return v


def write_workbook(path, data, wms):
    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    bold = wb.add_format({"bold": True})
    hdr = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "text_wrap": True})
    pct = wb.add_format({"num_format": "0%"})
    num = wb.add_format({"num_format": "#,##0"})

    # ---- WMS% (written first so it is the first tab, as in the template)
    ws_w = wb.add_worksheet("WMS%")
    ws_s = wb.add_worksheet("Summary")
    ws_d = wb.add_worksheet("Data")

    w_hdr = ["Article (SKU)", "Article Description", "Sum of Overwrite", None,
             "Article (SKU)", "Article Description", "Sum of Overwrite", "DC BOH", "WMS%"]
    for c, v in enumerate(w_hdr):
        if v:
            ws_w.write(2, c, v, hdr)
    last_w = 3 + len(wms)  # 1-based last row
    for i, r in enumerate(wms.itertuples(index=False)):
        row = 3 + i
        a, desc, total, boh, pctv = int(r[0]), r.desc, float(r.total), float(r[3]), r[4]
        ws_w.write_row(row, 0, [a, desc, total])
        ws_w.write_row(row, 4, [a, desc])
        ws_w.write_formula(row, 6, f"=_xlfn.XLOOKUP(E{row + 1},A$4:A${last_w},C$4:C${last_w})", None, total)
        ws_w.write_number(row, 7, boh)
        ws_w.write_formula(row, 8, f"=MIN(H{row + 1}/G{row + 1},100%)", pct,
                           clean(pctv) if pd.notna(pctv) else "#DIV/0!")
    ws_w.set_column(0, 0, 13); ws_w.set_column(1, 1, 34); ws_w.set_column(2, 2, 10)
    ws_w.set_column(4, 4, 13); ws_w.set_column(5, 5, 34); ws_w.set_column(6, 8, 10)
    ws_w.freeze_panes(3, 0)

    # ---- Summary (pivot-style tables, values)
    q = "Final Allocation Qty"
    blocks = [
        (0, data.groupby("Article Description")[q].sum()),
        (3, data.groupby("Store Description")[q].sum()),
    ]
    by_dept = data.groupby("Department")[q].sum().sort_values(ascending=False)
    by_region = data.groupby("Region")[q].sum().sort_values(ascending=False)
    grand = float(data[q].sum())
    ws_s.write_number(1, 1, grand, num)
    ws_s.write_number(1, 4, grand, num)
    # constant_memory requires writing rows in order, so build a grid first
    grid = {}
    for col, s in blocks:
        s = s.sort_values(ascending=False)
        grid[(2, col)] = ("Row Labels", hdr); grid[(2, col + 1)] = ("Sum of Final Allocation Qty", hdr)
        for i, (k, v) in enumerate(s.items()):
            grid[(3 + i, col)] = (k, None); grid[(3 + i, col + 1)] = (float(v), num)
        grid[(3 + len(s), col)] = ("Grand Total", bold); grid[(3 + len(s), col + 1)] = (grand, bold)
    r0 = 2
    for s in (by_dept, by_region):
        grid[(r0, 6)] = ("Row Labels", hdr); grid[(r0, 7)] = ("Sum of Final Allocation Qty", hdr)
        for i, (k, v) in enumerate(s.items()):
            grid[(r0 + 1 + i, 6)] = (k, None); grid[(r0 + 1 + i, 7)] = (float(v), num)
        grid[(r0 + 1 + len(s), 6)] = ("Grand Total", bold); grid[(r0 + 1 + len(s), 7)] = (grand, bold)
        r0 = r0 + len(s) + 3
    for (r, c) in sorted(grid):
        v, f = grid[(r, c)]
        ws_s.write(r, c, clean(v), f)
    ws_s.set_column(0, 0, 36); ws_s.set_column(1, 1, 14); ws_s.set_column(3, 3, 36)
    ws_s.set_column(4, 4, 14); ws_s.set_column(6, 6, 24); ws_s.set_column(7, 7, 14)

    # ---- Data
    n = len(data)
    last = n + 2  # 1-based last data row
    ws_d.write_formula(0, COL[q], f"=SUBTOTAL(9,T3:T{last})", num, grand)
    for c, h in enumerate(DATA_COLUMNS):
        ws_d.write(1, c, h, hdr)
    # Pre-split columns into (col, values) so the row loop does as little as possible.
    fcols = {COL["Overwrite"], COL["WMS %"], COL[q], COL["COMBINE Article, Store & EFFECTIVE DATE"],
             COL["Combine"], COL["by PUF"]}
    static = []
    for c, name in enumerate(DATA_COLUMNS):
        if c in fcols:
            continue
        col = [clean(v) for v in data[name].tolist()]
        if all(v is None for v in col):
            continue
        static.append((c, col))
    ov = [clean(v) for v in data["Overwrite"].tolist()]
    wv = [clean(v) for v in data["WMS %"].tolist()]
    fv = [clean(v) for v in data[q].tolist()]
    bv = [clean(v) for v in data["by PUF"].tolist()]
    k1 = data["COMBINE Article, Store & EFFECTIVE DATE"].tolist()
    k2 = data["Combine"].tolist()
    wnum, wstr, wf = ws_d.write_number, ws_d.write_string, ws_d.write_formula
    cO, cR, cT, cY, cAB, cAH = (COL["Overwrite"], COL["WMS %"], COL[q],
                                COL["COMBINE Article, Store & EFFECTIVE DATE"], COL["Combine"], COL["by PUF"])
    for i in range(n):
        row = i + 2
        x = row + 1  # Excel row number
        for c, col in static:
            v = col[i]
            if v is None:
                continue
            if isinstance(v, str):
                wstr(row, c, v)
            else:
                wnum(row, c, v)
        wf(row, cO, f"=I{x}", None, ov[i])
        wf(row, cR, f"=_xlfn.XLOOKUP(A{x},'WMS%'!E$4:E${last_w},'WMS%'!I$4:I${last_w})", pct,
           wv[i] if wv[i] is not None else "#N/A")
        wf(row, cT, f"=IF(AH{x}>=0.8,MROUND(J{x}*R{x},AG{x}),MROUND(J{x}*R{x},S{x}))",
           None, fv[i] if fv[i] is not None else "#N/A")
        wf(row, cY, f'=CONCATENATE(AB{x},"_",O{x})', None, k1[i])
        wf(row, cAB, f"=A{x}&E{x}", None, k2[i])
        b = bv[i]
        wf(row, cAH, f"=J{x}*R{x}/AG{x}", None, b if b is not None and np.isfinite(b) else "#DIV/0!")
    ws_d.autofilter(1, 0, last - 1, len(DATA_COLUMNS) - 1)
    ws_d.freeze_panes(2, 0)
    ws_d.set_column(0, 0, 12); ws_d.set_column(1, 1, 32); ws_d.set_column(5, 5, 28)
    wb.close()


def main():
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description="Build the daily STO Allocation workbook")
    ap.add_argument("--input", default=here / "input", type=Path)
    ap.add_argument("--output", default=here / "output", type=Path)
    args = ap.parse_args()
    t0 = time.time()

    files = find_inputs(args.input)
    for k, f in files.items():
        log(f"{k:8s}: {f.name}")

    log("Reading STO allocation export ...");  exp, plan_date = load_export(files["export"])
    log(f"  {len(exp):,} allocation rows (plan date {plan_date:%d-%b-%Y})")
    log("Reading article report ...");         art = load_article(files["article"])
    log("Reading store list ...");             store = load_store(files["store"])
    log("Reading OOS RCA ...");                oos = load_oos(files["oos"])
    log("Reading DC BOH ...");                 boh = load_boh(files["boh"])

    log("Building ...")
    data, wms = build(exp, art, store, oos, boh)

    # exception report
    miss_art = sorted(set(data.loc[data["PUF"].isna(), "Article (SKU)"].dropna().astype(int)))
    miss_st = sorted(set(data.loc[data["Region"].isna(), "Store"].dropna().astype(int)))
    no_boh = int((~wms["Article (SKU)"].isin(boh["art"])).sum())
    if miss_art:
        log(f"WARNING: {len(miss_art)} article(s) not in article report (blank SUF/PUF -> errors): "
            f"{miss_art[:10]}{' ...' if len(miss_art) > 10 else ''}")
    if miss_st:
        log(f"WARNING: {len(miss_st)} store(s) not in store list: {miss_st[:20]}")
    if no_boh:
        log(f"NOTE: {no_boh} article(s) not in DC BOH file -> DC BOH = 0 (WMS% 0%)")

    args.output.mkdir(parents=True, exist_ok=True)
    out = args.output / f"{plan_date:%m.%d}_Final_STO_Allocation_Orders_Article_SKU.xlsx"
    log(f"Writing {out.name} ...")
    write_workbook(out, data, wms)

    q = data["Final Allocation Qty"]
    print("\n================ SUMMARY ================")
    print(f"Plan date            : {plan_date:%d-%b-%Y}")
    print(f"Rows / articles / stores : {len(data):,} / {data['Article (SKU)'].nunique():,} / {data['Store'].nunique():,}")
    print(f"Total Overwrite qty  : {data['Overwrite'].sum():,.0f}")
    print(f"Final Allocation Qty : {q.sum():,.0f}")
    print(f"Articles WMS% < 100% : {(wms['WMS%'] < 1).sum():,}  (0%: {(wms['WMS%'] == 0).sum():,})")
    print("By department:")
    for k, v in data.groupby("Department")[q.name].sum().sort_values(ascending=False).items():
        print(f"   {k:28s}{v:>12,.0f}")
    print("By region:")
    for k, v in data.groupby("Region")[q.name].sum().sort_values(ascending=False).items():
        print(f"   {k:28s}{v:>12,.0f}")
    print("=========================================")
    log(f"Done in {time.time() - t0:.0f}s -> {out}")


if __name__ == "__main__":
    main()
