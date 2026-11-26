"""
Compare your manually finished STO file with the Python output and write a
small report you can share (compare_report.xlsx).

    python compare.py                      (uses the folders below)
    python compare.py MANUAL_FILE PYTHON_FILE

Default files:
    manual : newest Excel file in the  compare  folder
    python : newest .xlsb/.xlsx in the  output  folder
"""
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

COLS = ["Article (SKU)", "Article Description", "Store", "Store Description", "LEGS", "Region",
        "Department", "DL", "OB", "Overwrite", "WMS %", "Final Allocation Qty", "Total SOH"]
NUM = ["Overwrite", "WMS %", "Final Allocation Qty"]


def newest(folder, exts):
    files = [f for f in Path(folder).glob("*") if f.suffix.lower() in exts and not f.name.startswith("~$")]
    if not files:
        sys.exit(f"ERROR: no Excel file found in {folder}")
    return max(files, key=lambda f: f.stat().st_mtime)


def read_data(path):
    print(f"Reading {path.name} ...", flush=True)
    raw = pd.read_excel(path, sheet_name="Data", header=None, engine="calamine", dtype=object)
    h = next(i for i in range(10) if "Article (SKU)" in [str(v).strip() for v in raw.iloc[i]])
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[h]]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df = df[pd.to_numeric(df["Article (SKU)"], errors="coerce").notna()]
    for c in COLS:
        if c not in df.columns:
            df[c] = None
    df = df[COLS].copy()
    df["Article (SKU)"] = pd.to_numeric(df["Article (SKU)"]).astype("int64")
    df["Store"] = pd.to_numeric(df["Store"], errors="coerce").astype("Int64")
    for c in NUM:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["DL"] = df["DL"].astype(str).str.strip().replace({"0.0": "0", "nan": "", "None": ""})
    df["OB"] = df["OB"].astype(str).str.strip().replace({"0.0": "0", "nan": "", "None": ""})
    df["key"] = df["Article (SKU)"].astype(str) + "_" + df["Store"].astype(str)
    return df.drop_duplicates("key")


def main():
    here = Path(__file__).resolve().parent
    if len(sys.argv) >= 3:
        manual, python = Path(sys.argv[1]), Path(sys.argv[2])
    else:
        manual = newest(here / "compare", {".xlsb", ".xlsx", ".xlsm"})
        python = newest(here / "output", {".xlsb", ".xlsx"})
    print(f"MANUAL : {manual}\nPYTHON : {python}\n")
    t0 = time.time()
    m, p = read_data(manual), read_data(python)

    only_m = m[~m.key.isin(p.key)]
    only_p = p[~p.key.isin(m.key)]
    both = m.merge(p, on="key", suffixes=(" (manual)", " (python)"))
    diffs = {}
    for c in NUM:
        a, b = both[f"{c} (manual)"], both[f"{c} (python)"]
        diffs[c] = ~(np.isclose(a.fillna(-999), b.fillna(-999), atol=1e-6))

    summary = pd.DataFrame([
        ["Rows", len(m), len(p)],
        ["Stores", m.Store.nunique(), p.Store.nunique()],
        ["Articles", m["Article (SKU)"].nunique(), p["Article (SKU)"].nunique()],
        ["Sum of Overwrite", m.Overwrite.sum(), p.Overwrite.sum()],
        ["Sum of Final Allocation Qty", m["Final Allocation Qty"].sum(), p["Final Allocation Qty"].sum()],
        ["Final Qty = 0 rows", int((m["Final Allocation Qty"] == 0).sum()), int((p["Final Allocation Qty"] == 0).sum())],
        ["Rows only in this file", len(only_m), len(only_p)],
        ["Rows in both: Overwrite differs", int(diffs["Overwrite"].sum()), ""],
        ["Rows in both: WMS % differs", int(diffs["WMS %"].sum()), ""],
        ["Rows in both: Final Qty differs", int(diffs["Final Allocation Qty"].sum()), ""],
    ], columns=["Check", "Manual", "Python"])

    def by(df, cols):
        if df.empty:
            return pd.DataFrame(columns=cols + ["Rows", "Overwrite", "Final Allocation Qty"])
        return (df.fillna({c: "(blank)" for c in cols})
                  .groupby(cols, dropna=False)
                  .agg(Rows=("key", "size"), Overwrite=("Overwrite", "sum"),
                       **{"Final Allocation Qty": ("Final Allocation Qty", "sum")})
                  .reset_index().sort_values("Rows", ascending=False))

    def totals(col):
        a = m.groupby(col, dropna=False)["Final Allocation Qty"].sum().rename("Manual")
        b = p.groupby(col, dropna=False)["Final Allocation Qty"].sum().rename("Python")
        t = pd.concat([a, b], axis=1).fillna(0)
        t["Difference"] = t["Python"] - t["Manual"]
        return t.reset_index()

    fd = both[diffs["Final Allocation Qty"]].copy()
    fd["Final diff"] = fd["Final Allocation Qty (python)"] - fd["Final Allocation Qty (manual)"]
    fd_cols = ["Article (SKU) (manual)", "Article Description (manual)", "Store (manual)",
               "Store Description (manual)", "LEGS (manual)", "LEGS (python)", "DL (manual)", "OB (manual)",
               "Total SOH (manual)", "Total SOH (python)",
               "Overwrite (manual)", "Overwrite (python)", "WMS % (manual)", "WMS % (python)",
               "Final Allocation Qty (manual)", "Final Allocation Qty (python)", "Final diff"]
    wms = (both[diffs["WMS %"]]
           .groupby(["Article (SKU) (manual)", "Article Description (manual)"])
           .agg(**{"WMS % manual": ("WMS % (manual)", "first"), "WMS % python": ("WMS % (python)", "first"),
                   "Rows": ("key", "size")}).reset_index())

    out = here / "compare_report.xlsx"
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        totals("Department").to_excel(xw, sheet_name="Totals by Dept", index=False)
        totals("Region").to_excel(xw, sheet_name="Totals by Region", index=False)
        by(only_m, ["Store", "Store Description", "LEGS", "Region"]).to_excel(xw, sheet_name="Only in MANUAL by store", index=False)
        by(only_m, ["LEGS", "Region", "DL", "OB"]).to_excel(xw, sheet_name="Only in MANUAL by type", index=False)
        by(only_p, ["Store", "Store Description", "LEGS", "Region"]).to_excel(xw, sheet_name="Only in PYTHON by store", index=False)
        by(only_p, ["LEGS", "Region", "DL", "OB"]).to_excel(xw, sheet_name="Only in PYTHON by type", index=False)
        wms.head(3000).to_excel(xw, sheet_name="WMS% differs (articles)", index=False)
        fd[fd_cols].head(3000).to_excel(xw, sheet_name="Final Qty differs (sample)", index=False)
        for ws in xw.sheets.values():
            ws.set_column(0, 20, 18)

    print("\n" + summary.to_string(index=False))
    print(f"\nReport written: {out}  ({time.time() - t0:.0f}s)")
    print("Please send compare_report.xlsx")


if __name__ == "__main__":
    main()
