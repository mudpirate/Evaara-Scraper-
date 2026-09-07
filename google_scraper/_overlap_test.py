# /// script
# requires-python = ">=3.11"
# dependencies = ["rich", "openpyxl"]
# ///
"""
Overlap test: do 4 broad keywords find (almost) everything the other 11 find?

Usage:
    python _overlap_test.py <city> [district]
    e.g. python _overlap_test.py telgi vijaypur

Runs scrape.py twice for the city:
  SET-A (4 broad):  Mobile shop, Cell phone store, Electronics showroom,
                    Used mobile shop
  SET-B (11 niche): the remaining 11 keywords
Saves _overlap_parts/<city>_A.xlsx and <city>_B.xlsx, then compares with the
same dedupe key the final merge uses (title+street+city+phone) and prints:
  |A|, |B|, |union|, added_by_B, coverage of A, sample B-only titles.
"""
import os
import subprocess
import sys
from pathlib import Path

if sys.platform == 'win32':
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    os.environ.setdefault('PYTHONUTF8', '1')

SET_A = [
    "Mobile shop",
    "Cell phone store",
    "Electronics showroom",
    "Used mobile shop",
]
SET_B = [
    "Smartphone retailer",
    "Mobile phone dealer",
    "Mobile showroom",
    "Mobile accessories shop",
    "Smartphone accessories dealer",
    "Second-hand phone store",
    "Buy and sell cell phones",
    "Consumer electronics retailer",
    "Electronics store and home appliances",
    "Smart home appliances shop",
    "Electronics and smartphone shops",
]

HERE = Path(__file__).parent
PARTS = HERE / "_overlap_parts"
STATE = "karnataka"

if len(sys.argv) < 2:
    sys.exit("usage: python _overlap_test.py <city> [district]")
CITY = sys.argv[1]
DISTRICT = sys.argv[2] if len(sys.argv) > 2 else "vijaypur"


def run_set(keywords: list[str], out_file: Path) -> bool:
    env = os.environ.copy()
    env['KEEP_ALL'] = '1'
    env.setdefault('HEADLESS', '1')
    env.setdefault('PARALLEL_TABS', '3')
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUNBUFFERED'] = '1'
    place = f"{CITY}, {DISTRICT}, {STATE}"
    cmd = [sys.executable, str(HERE / 'scrape.py'), ", ".join(keywords),
           place, str(out_file)]
    print(f"\n=== SET ({len(keywords)} kws) in {place} -> {out_file.name}",
          flush=True)
    proc = subprocess.Popen(cmd, env=env, stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, text=True,
                            encoding='utf-8', errors='replace', bufsize=1)
    assert proc.stdout is not None
    for line in proc.stdout:
        print(line, end='', flush=True)
    proc.wait()
    return proc.returncode == 0 and out_file.exists()


def load_keys(path: Path) -> dict:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    ws = wb['shops']
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        r = (tuple(r) + ('',) * 24)[:24]
        key = (str(r[0]).lower().strip(), str(r[3]).strip(),
               str(r[4]).strip(), str(r[8]).strip())
        out[key] = str(r[0])
    wb.close()
    return out


def main() -> None:
    PARTS.mkdir(exist_ok=True)
    slug = ''.join(c if c.isalnum() else '_' for c in CITY.lower()).strip('_')
    fa, fb = PARTS / f"{slug}_A4.xlsx", PARTS / f"{slug}_B11.xlsx"
    oka = run_set(SET_A, fa) if not fa.exists() else True
    okb = run_set(SET_B, fb) if not fb.exists() else True
    if not (fa.exists() and fb.exists()):
        print(f"OVERLAP_INCOMPLETE city={CITY} A_ok={oka} B_ok={okb}", flush=True)
        sys.exit(1)
    A, B = load_keys(fa), load_keys(fb)
    union = set(A) | set(B)
    only_b = set(B) - set(A)
    only_a = set(A) - set(B)
    cov = (len(set(A) & union) / len(union) * 100) if union else 100.0
    print(f"\n===== OVERLAP city={CITY} =====", flush=True)
    print(f"A(4 broad)={len(A)}  B(11 niche)={len(B)}  union={len(union)}",
          flush=True)
    print(f"added_by_B_only={len(only_b)}  only_in_A={len(only_a)}", flush=True)
    print(f"COVERAGE_A={cov:.1f}%  (share of all shops the 4 broad find)",
          flush=True)
    if only_b:
        print("sample B-only titles:", flush=True)
        for k in list(only_b)[:15]:
            print(f"  + {B[k]}", flush=True)
    print(f"RESULT_COVERAGE {CITY} {cov:.1f} {len(A)} {len(B)} {len(union)}",
          flush=True)


if __name__ == "__main__":
    main()
