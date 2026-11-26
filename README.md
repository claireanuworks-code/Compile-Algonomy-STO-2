# STO Allocation builder

Builds the daily **STO Allocation** workbook (Data, WMS% and Summary sheets) from the Algonomy export and the daily lookup files.

## One-time setup (Windows)
1. Install Python. During install, tick **"Add python.exe to PATH"**.
2. Double-click **`setup.bat`**. It installs pandas, python-calamine, xlsxwriter and pywin32. `run_sto.bat` also installs anything that's missing.
3. Put a previous finished STO allocation file (`..._Final_STO_Allocation_Orders_Article_SKU.xlsb`) in the **`input`** folder (or in `template`). The script uses it as the layout: the Data headers, formats and the pivot tables on WMS% and Summary. Its old data is replaced every run, and the template file itself is never changed.

## Daily use
1. Put the day's 5 files in the **`input`** folder and remove the previous day's files. The files are recognised by name:

   | File | Name must contain |
   |---|---|
   | Algonomy export | `STO_Allocation_Orders` (and not `Final`) |
   | Article report | `Article_Report` |
   | Store list | `Store_list` |
   | OOS RCA | `Root_Cause` |
   | DC BOH | `DC_BOH` |

   Spaces, dashes and underscores count as the same character.
2. Double-click **`run_sto.bat`**. It takes about 1–2 minutes.
3. The result appears in **`output`** as `MM.DD_Final_STO_Allocation_Orders_Article_SKU.xlsb`. The date comes from the export's *Plan Created Date*. The console also prints a short summary and warns about any stores or articles that are missing from the lookup files.

## What it does
**Data sheet:** the export columns plus the lookups below. Formula columns are written as live Excel formulas, the same ones as the template.

| Column | Source |
|---|---|
| J Overwrite | `=I` |
| P Ordering Cycle | Store list → `Ordering cycle (...)` |
| Q Cap based on vertical facing | 0 |
| R WMS % | `XLOOKUP(article, WMS%!E, WMS%!I)` |
| S SUF | Article report → `Delivery Rate against Basic UoM (SUF/IN1)` |
| T Final Allocation Qty | `IF(AH>=0.8, MROUND(J*R, AG), MROUND(J*R, S))` |
| U/V/W Department / Category / SUB CAT | Article report AH / AJ / AL |
| X Child SKU? | Article report CU (Article Category) |
| Y, AB keys | `article&store_effective-date`, `article&store` |
| Z LEGS, AA Region | Store list → `NEW LEGS`, `Site Hierarchy L2` |
| AG PUF | Article report → `Ordering UOM Rate against Basic UoM (PUF)` |
| AH by PUF | `J*R/AG` |
| AI DL / AJ OB | Article report → `Key Line Item` / `Corporate Brand` (`X` or 0) |
| AL Total SOH / AM OOS RCA | 0 + the key if article+store is in the OOS RCA report, otherwise 1 + blank |

SOH, GIT, DC Pending and TOP 40 OOS are left blank, as in the template.

**WMS% sheet:** Sum of Overwrite per article, DC BOH (`Available − Stock In Stage & Others` from the DC BOH file, 0 if the article is missing) and `WMS% = MIN(DC BOH / Sum, 100%)`.

**How the file is produced:** Python builds the data, then opens your template in a hidden copy of Excel. It replaces the Data rows (keeping your formats), fills the formulas down, points all 5 pivot tables (WMS% and Summary) at the new data, refreshes them, fills WMS% columns E–I, and saves as `.xlsb`. Excel must be installed. With no template, or with `python sto_allocation.py --simple`, it writes a plain `.xlsx` without pivots.

## Row deletion (before WMS% is calculated)
1. Keep Region **WM** and **EM** only.
2. Delete stores whose LEGS (from the store list) is **NEW**, **Reno - Total Closure** (other Reno stores, e.g. "Reno Alloc - OB DL Only", are kept), **CLOSED**, **E-Commerce**, **EXPO** or **NIL**.

The console shows how many rows each rule removed.

## Not automated yet
- Zeroing Final Allocation Qty on selected rows (done by hand in the 20 Nov file).
- **Manual LEGS edits** (e.g. "(Reno Alloc - OB DL Only)") are not applied. LEGS comes straight from the store list.
