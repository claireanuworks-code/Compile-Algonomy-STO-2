"""
Collect today's input files into the `input` folder.

    python fetch_inputs.py

For each file type it finds the most recent file in its source folder (by the
date in the file name, otherwise the modified date) and copies it into
`input`, replacing the older file of the same type. DC BOH is taken from the
newest Outlook e-mail with a "BOH" attachment.

Edit the SOURCES section below if a folder moves.
"""
import re
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

HOME = Path.home()
ONEDRIVE = HOME / "OneDrive - The Dairy Farm Company Ltd"

# ------------------------------------------------------------------ SOURCES
# label: (folder, words that must be in the file name, search sub-folders?)
SOURCES = {
    "Article Report": (ONEDRIVE / "MDM - PUBLIC - MDM" / "01. ARTICLE REPORT", ["article", "report"], True),
    "Store list": (ONEDRIVE / "Demand Planning - Documents" / "G. Store List", ["store", "list"], False),
    "OOS Root Cause": (ONEDRIVE / "Demand Planning - Documents" / "ZZZZ OOS RCA", ["root", "cause"], False),
    "STO Allocation Orders": (HOME / "Downloads", ["sto", "allocation", "orders"], False),
}
# DC BOH from Outlook: newest mail (last N days) with an attachment whose name contains these words
OUTLOOK_BOH_WORDS = ["boh"]
OUTLOOK_DAYS = 7
# --------------------------------------------------------------------------

EXCEL = {".xlsx", ".xlsb", ".xls", ".xlsm", ".csv"}
HERE = Path(__file__).resolve().parent
INPUT = HERE / "input"


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def norm(s):
    return re.sub(r"[\s\-_]+", " ", str(s).lower())


def name_date(name):
    """Date written in a file name, e.g. 20.11.2026, 11.20.2026, 2026-11-20, 20112026."""
    s = str(name)
    m = re.search(r"(20\d{2})[.\-_ ](\d{1,2})[.\-_ ](\d{1,2})", s)            # 2026-11-20
    if m:
        y, a, b = map(int, m.groups())
        for mo, d in ((a, b), (b, a)):
            try:
                return datetime(y, mo, d)
            except ValueError:
                pass
    m = re.search(r"(\d{1,2})[.\-_ ](\d{1,2})[.\-_ ](20\d{2})", s)            # 20.11.2026 / 11.20.2026
    if m:
        a, b, y = map(int, m.groups())
        options = [(b, a), (a, b)] if a > 12 else [(a, b), (b, a)] if b > 12 else [(b, a)]
        for mo, d in options:
            try:
                return datetime(y, mo, d)
            except ValueError:
                pass
    return None


def sort_key(f):
    d = name_date(f.name)
    mtime = datetime.fromtimestamp(f.stat().st_mtime)
    return (d or mtime.replace(hour=0, minute=0, second=0, microsecond=0), mtime)


def newest_file(folder, words, recursive):
    if not folder.exists():
        log(f"   folder not found: {folder}")
        return None
    it = folder.rglob("*") if recursive else folder.iterdir()
    files = [f for f in it if f.is_file() and f.suffix.lower() in EXCEL
             and not f.name.startswith("~$") and all(w in norm(f.name) for w in words)
             and "final" not in norm(f.name)]
    return max(files, key=sort_key) if files else None


def existing(words):
    return [f for f in INPUT.iterdir() if f.is_file() and all(w in norm(f.name) for w in words)]


def put(src_name, label, words, writer):
    """Remove older files of this type from input, then write the new one."""
    for old in existing(words):
        try:
            old.unlink()
        except Exception as e:
            log(f"   could not remove old {old.name}: {e}")
    name = src_name if all(w in norm(src_name) for w in words) else f"{label} - {src_name}"
    dest = INPUT / name
    writer(dest)
    return dest


def age_note(d):
    days = (datetime.now().date() - d.date()).days
    return "today" if days == 0 else "yesterday" if days == 1 else f"{days} days old"


def fetch_file(label, folder, words, recursive):
    log(f"{label}: searching {folder} ...")
    f = newest_file(folder, words, recursive)
    if f is None:
        log(f"   WARNING: no {label} file found -> keeping what is in input")
        return False
    same = INPUT / f.name
    if same.exists() and same.stat().st_size == f.stat().st_size:
        log(f"   already in input: {f.name}  ({age_note(sort_key(f)[0])})")
        return True
    put(f.name, label, words, lambda dest: shutil.copy2(f, dest))
    log(f"   copied {f.name}  ({age_note(sort_key(f)[0])})")
    return True


def fetch_boh_outlook():
    log("DC BOH: searching Outlook ...")
    try:
        import win32com.client
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        log(f"   WARNING: could not open Outlook ({e}) -> keeping what is in input")
        return False
    since = datetime.now() - timedelta(days=OUTLOOK_DAYS)
    best = None  # (received, attachment, folder name)

    def scan(folder, depth=0):
        nonlocal best
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            count = 0
            for item in items:
                count += 1
                if count > 400:
                    break
                try:
                    rt = item.ReceivedTime
                    rt = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
                except Exception:
                    continue
                if rt < since:
                    break
                if best and rt <= best[0]:
                    break
                try:
                    atts = item.Attachments
                except Exception:
                    continue
                for i in range(1, atts.Count + 1):
                    a = atts.Item(i)
                    fn = str(a.FileName)
                    if Path(fn).suffix.lower() in EXCEL and all(w in norm(fn) for w in OUTLOOK_BOH_WORDS):
                        best = (rt, a, folder.Name, str(getattr(item, "Subject", "")))
                        break
        except Exception:
            pass
        if depth < 2:
            try:
                for i in range(1, folder.Folders.Count + 1):
                    scan(folder.Folders.Item(i), depth + 1)
            except Exception:
                pass

    scan(ns.GetDefaultFolder(6))  # Inbox (+ its sub-folders)
    if best is None:
        log(f"   WARNING: no e-mail with a BOH attachment in the last {OUTLOOK_DAYS} days -> keeping what is in input")
        return False
    rt, att, folder_name, subject = best
    ext = Path(str(att.FileName)).suffix
    name = f"DC BOH {rt:%Y-%m-%d}{ext}"
    put(name, "DC BOH", ["dc", "boh"], lambda dest: att.SaveAsFile(str(dest)))
    log(f"   saved attachment '{att.FileName}' from '{subject}' ({folder_name}, received {rt:%d-%b %H:%M}, {age_note(rt)})")
    return True


def main():
    INPUT.mkdir(exist_ok=True)
    log(f"Collecting input files into {INPUT}")
    ok = True
    for label, (folder, words, rec) in SOURCES.items():
        ok &= fetch_file(label, folder, words, rec)
    ok &= fetch_boh_outlook()
    print()
    log("Files now in input:")
    for f in sorted(INPUT.iterdir()):
        if f.is_file() and f.suffix.lower() in EXCEL:
            log(f"   {f.name}   (modified {datetime.fromtimestamp(f.stat().st_mtime):%d-%b %H:%M})")
    if not ok:
        log("Some files were not found automatically - put them in the input folder by hand if needed.")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
