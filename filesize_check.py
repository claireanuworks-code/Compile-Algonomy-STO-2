"""Show what makes an .xlsb file big.  python filesize_check.py [file ...]
Without arguments: checks the newest file in output and in any *template* folder."""
import sys
import zipfile
from pathlib import Path


def report(f):
    z = zipfile.ZipFile(f)
    total = sum(i.compress_size for i in z.infolist())
    print(f"\n== {f.name}   ({total / 1e6:.1f} MB)")
    for i in sorted(z.infolist(), key=lambda i: -i.compress_size)[:10]:
        print(f"   {i.compress_size / 1e6:7.1f} MB  {i.filename}")
    n = sum(1 for i in z.infolist() if "pivotCacheRecords" in i.filename)
    print(f"   pivot caches with saved data: {n}")


here = Path(__file__).resolve().parent
files = [Path(a) for a in sys.argv[1:]]
if not files:
    for d in [here / "output"] + [d for d in here.iterdir() if d.is_dir() and "template" in d.name.lower()]:
        xs = sorted(d.glob("*.xlsb"), key=lambda f: f.stat().st_mtime)
        if xs:
            files.append(xs[-1])
for f in files:
    report(f)
input("\nPress Enter to close")
