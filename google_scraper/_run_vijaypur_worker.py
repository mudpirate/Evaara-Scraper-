# /// script
# requires-python = ">=3.11"
# dependencies = ["rich", "openpyxl"]
# ///
"""
Parallel worker for Vijayapura multi-keyword scrape.

Usage:
    python _run_vijaypur_worker.py <group_line_file.txt> [--no-merge]

Runs scrape.py once per city in the group file, with all 3 keywords per
city. Part files go to the SHARED _vijaypur_parts/ dir (slug-named, so two
workers never overwrite each other). Finished cities are SKIPPED on restart.

Each worker MUST use its own browser profile (set PROFILE_DIR env) and its
own log file (caller redirects stdout). PARALLEL_TABS env controls
extraction concurrency (default 3, 5 recommended for speed).

With --no-merge (used when 2 workers run side by side) the final merge is
skipped — run _merge_vijaypur_final.py after ALL workers finish.
"""
import os
import re
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

KEYWORDS = [
    "Mobile shops",
    "Mobile accessories",
    "Electronics and appliances",
]

DISTRICT = "vijaypur"
STATE = "karnataka"
HERE = Path(__file__).parent
PARTS = HERE / "_vijaypur_parts"

if len(sys.argv) < 2:
    sys.exit("usage: python _run_vijaypur_worker.py <group_line_file.txt> [--no-merge]")

LINE_FILE = HERE / sys.argv[1]
NO_MERGE = '--no-merge' in sys.argv
TAG = Path(sys.argv[1]).stem  # e.g. _vijaypur_groupA


def cities_from_line_file() -> list[str]:
    text = LINE_FILE.read_text(encoding='utf-8-sig').strip()
    _, _, cities = text.partition(':')
    return [c.strip() for c in cities.split(',') if c.strip()]


def slug(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:40]


def run_city(city: str, out_file: Path) -> bool:
    env = os.environ.copy()
    env['KEEP_ALL'] = '1'
    env.setdefault('HEADLESS', '1')
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUNBUFFERED'] = '1'
    place = f"{city}, {DISTRICT}, {STATE}"
    kw_arg = ", ".join(KEYWORDS)
    cmd = [sys.executable, str(HERE / 'scrape.py'), kw_arg, place, str(out_file)]
    print(f"\n[{TAG}] === RUN [{city}]: {len(KEYWORDS)} keywords in {place}", flush=True)
    proc = subprocess.Popen(cmd, env=env, stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, text=True,
                            encoding='utf-8', errors='replace', bufsize=1)
    assert proc.stdout is not None
    for line in proc.stdout:
        print(f"[{TAG}] {line}", end='', flush=True)
    proc.wait()
    return proc.returncode == 0 and out_file.exists()


def shops_rows(path: Path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    if 'shops' not in wb.sheetnames:
        wb.close()
        return None
    ws = wb['shops']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    return rows


def main() -> None:
    cities = cities_from_line_file()
    print(f"[{TAG}] {len(cities)} cities x {len(KEYWORDS)} keywords = "
          f"{len(cities) * len(KEYWORDS)} queries. Parts: {PARTS}", flush=True)
    PARTS.mkdir(exist_ok=True)
    failed: list[str] = []
    for i, city in enumerate(cities, 1):
        out_file = PARTS / f"{slug(city)}.xlsx"
        if out_file.exists():
            try:
                rows = shops_rows(out_file)
            except Exception:
                rows = None
            if rows is not None:
                print(f"[{TAG}] [{i}/{len(cities)}] SKIP {city} (already done, "
                      f"{len(rows)} rows)", flush=True)
                continue
        print(f"[{TAG}] [{i}/{len(cities)}] {city}", flush=True)
        ok = run_city(city, out_file)
        if not ok:
            print(f"[{TAG}] !!! FAILED: {city}", flush=True)
            failed.append(city)
    if failed:
        print(f"[{TAG}] Cities that failed: {', '.join(failed)}", flush=True)
    if not NO_MERGE:
        import _merge_vijaypur_final as m
        m.merge()
    else:
        print(f"[{TAG}] done (--no-merge, final merge runs after all workers finish)",
              flush=True)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
