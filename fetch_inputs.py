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
# DC BOH:
#  1) first looks in DC_BOH_FOLDER (e.g. a OneDrive folder where a Power Automate flow
#     saves the e-mail attachment) - leave as None if you don't use one
#  2) otherwise searches classic Outlook (all mailboxes + group/shared mailboxes)
DC_BOH_FOLDER = None   # another folder to copy DC BOH from (None = not used)
DC_BOH_IN_INPUT = True  # Power Automate saves DC BOH straight into the input folder
OUTLOOK_BOH_WORDS = ["boh"]          # attachment name must contain these
OUTLOOK_SUBJECT_WORDS = ["dc", "boh"]  # ...or the e-mail subject must contain these
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
    if DC_BOH_IN_INPUT:
        log("DC BOH: saved into input by Power Automate ...")
        have = existing(["dc", "boh"])
        if have:
            f = max(have, key=sort_key)
            log(f"   using {f.name}  ({age_note(sort_key(f)[0])})")
            if len(have) > 1:
                log(f"   NOTE: {len(have)} DC BOH files in input - the build uses the newest one")
            return True
        log("   WARNING: no DC BOH file in input yet - check the Power Automate flow, trying Outlook ...")
    if DC_BOH_FOLDER:
        log(f"DC BOH: searching {DC_BOH_FOLDER} ...")
        f = newest_file(Path(DC_BOH_FOLDER), ["boh"], True)
        if f is not None:
            same = INPUT / f.name
            if not (same.exists() and same.stat().st_size == f.stat().st_size):
                put(f.name, "DC BOH", ["dc", "boh"], lambda dest: shutil.copy2(f, dest))
            log(f"   using {f.name}  ({age_note(sort_key(f)[0])})")
            return True
        log("   nothing found there, trying Outlook ...")

    log("DC BOH: searching Outlook ...")
    try:
        import win32com.client
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        stores = [ns.Folders.Item(i) for i in range(1, ns.Folders.Count + 1)]
    except Exception as e:
        log(f"   WARNING: could not open classic Outlook ({e}).")
        log("   The 'new Outlook' cannot be read by Python - switch it off (toggle at the top right)")
        log("   or save the attachment into input by hand.")
        return False
    since = datetime.now() - timedelta(days=OUTLOOK_DAYS)
    best = None  # (received, attachment, folder path, subject)
    scanned = []

    def matches(item):
        subj = norm(getattr(item, "Subject", ""))
        subj_ok = all(w in subj for w in OUTLOOK_SUBJECT_WORDS)
        atts = item.Attachments
        for i in range(1, atts.Count + 1):
            a = atts.Item(i)
            fn = str(a.FileName)
            if Path(fn).suffix.lower() in EXCEL and (all(w in norm(fn) for w in OUTLOOK_BOH_WORDS) or subj_ok):
                return a
        return None

    def scan(folder, path, depth):
        nonlocal best
        name = str(folder.Name)
        if norm(name) in ("deleted items", "junk email", "junk e mail", "outbox", "sent items", "drafts",
                          "calendar", "contacts", "tasks", "notes", "journal", "conversation history"):
            return
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            n = 0
            for item in items:
                n += 1
                if n > 300:
                    break
                try:
                    rt = item.ReceivedTime
                    rt = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
                except Exception:
                    continue
                if rt < since or (best and rt <= best[0]):
                    break
                try:
                    a = matches(item)
                except Exception:
                    continue
                if a is not None:
                    best = (rt, a, f"{path}/{name}", str(getattr(item, "Subject", "")))
                    break
            scanned.append(f"{path}/{name} ({n})")
        except Exception:
            pass
        if depth < 3:
            try:
                for i in range(1, folder.Folders.Count + 1):
                    scan(folder.Folders.Item(i), f"{path}/{name}", depth + 1)
            except Exception:
                pass

    for st in stores:  # your mailbox, plus any shared / group mailboxes added to Outlook
        scan(st, "", 0)
    if best is None:
        log(f"   WARNING: no e-mail with a BOH attachment in the last {OUTLOOK_DAYS} days.")
        log(f"   Mailboxes searched: {[str(s.Name) for s in stores]}")
        log(f"   Folders searched: {len(scanned)} (e.g. {scanned[:8]})")
        return False
    rt, att, where, subject = best
    ext = Path(str(att.FileName)).suffix
    name = f"DC BOH {rt:%Y-%m-%d %H%M}{ext}"
    put(name, "DC BOH", ["dc", "boh"], lambda dest: att.SaveAsFile(str(dest)))
    log(f"   saved '{att.FileName}' from '{subject}' ({where}, received {rt:%d-%b %H:%M}, {age_note(rt)})")
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
