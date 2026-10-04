# /// script
# requires-python = ">=3.11"
# dependencies = ["rich", "openpyxl"]
# ///
"""
Vijayapura (vijaypur) per-city scraper — single keyword:
    "Mobile accessories and repair in <city>, vijaypur, karnataka"

Usage:
    python _run_mobile_acc_vijaypur.py
    python _run_mobile_acc_vijaypur.py --dry-run   (preview queries only)

- Cities are read from the "vijaypur:" line of districts.txt.
- One Excel part file per city is saved to _mobile_acc_vijaypur_parts/.
- Cities already done are SKIPPED on re-run (resumable after crash/captcha).
- At the end all parts are merged (duplicates removed) into ONE Excel:
    mobileaccessoriesvijapura.xlsx
  A "citySearched" column records which city each row came from.
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

KEYWORDS = ["Mobile accessories and repair"]
DISTRICT = "vijaypur"
STATE = "karnataka"
HERE = Path(__file__).parent
DISTRICTS_TXT = HERE / "districts.txt"
LINE_FILE = HERE / "_vijaypur_only.txt"
PARTS = HERE / "_mobile_acc_vijaypur_parts"
OUT = HERE / "mobileaccessoriesvijapura.xlsx"


def cities_from_districts_txt() -> list[str]:
    """Read the 'vijaypur: city, ...' line from districts.txt."""
    text = DISTRICTS_TXT.read_text(encoding='utf-8-sig')
    for line in text.replace('\r\n', '\n').split('\n'):
        line = line.strip().lstrip('\ufeff')
        if not line or line.startswith('#'):
            continue
        name, _, cities = line.partition(':')
        if name.strip().lower() == DISTRICT:
            return [c.strip() for c in cities.split(',') if c.strip()]
    # Fallback: pre-extracted per-district line file.
    text = LINE_FILE.read_text(encoding='utf-8-sig').strip()
    _, _, cities = text.partition(':')
    return [c.strip() for c in cities.split(',') if c.strip()]


def slug(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:40]


def run_city(city: str, out_file: Path) -> bool:
    """Run scrape.py once for one city. Returns True on success."""
    env = os.environ.copy()
    env['KEEP_ALL'] = '1'
    env.setdefault('HEADLESS', '1')
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUNBUFFERED'] = '1'
    place = f"{city}, {DISTRICT}, {STATE}"
    kw_arg = ", ".join(KEYWORDS)
    cmd = [sys.executable, str(HERE / 'scrape.py'), kw_arg, place, str(out_file)]
    print(f"\n=== RUN [{city}] -> {out_file.name}", flush=True)
    print(f"    - {kw_arg} in {place}", flush=True)
    proc = subprocess.Popen(cmd, env=env, stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, text=True,
                            encoding='utf-8', errors='replace', bufsize=1)
    assert proc.stdout is not None
    for line in proc.stdout:
        print(line, end='', flush=True)
    proc.wait()
    return proc.returncode == 0 and out_file.exists()


def shops_rows(path: Path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    if 'shops' not in wb.sheetnames:
        wb.close()
        return None  # not a valid part file -> re-run
    ws = wb['shops']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    return rows


def merge(parts: list[Path]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    import scrape as _scrape  # reuse CHAIN_NAMES + highlight logic

    base_cols = ['title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
                 'countryCode', 'website', 'phone'] + \
        [f'categories/{i}' for i in range(11)] + ['url', 'categoryName']
    cols = ['citySearched'] + base_cols
    seen: set = set()
    merged: list[list] = []
    for p in parts:
        rows = shops_rows(p)
        if not rows:
            continue
        city = p.stem.split('_', 1)[1] if '_' in p.stem else p.stem
        city = city.replace('_', ' ')
        for r in rows:
            r = (tuple(r) + ('',) * 25)[:len(base_cols)]
            key = (str(r[0]).lower().strip(), str(r[3]).strip(),
                   str(r[4]).strip(), str(r[8]).strip())
            if key in seen:
                continue
            seen.add(key)
            merged.append([city] + list(r))

    wb = Workbook()
    ws = wb.active
    ws.title = "shops"
    ws.append(cols)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color='DDDDDD')
        cell.alignment = Alignment(horizontal='center')

    phone_idx = cols.index('phone')
    green = PatternFill('solid', start_color=_scrape.CHAIN_FILL)
    yellow = PatternFill('solid', start_color=_scrape.NO_PHONE_FILL)
    for row in merged:
        ws.append(row)
        row_num = ws.max_row
        chain_match = _scrape.is_chain_match(str(row[1] or ''))
        phone_val = str(row[phone_idx] or '').strip()
        if chain_match:
            for idx in range(len(row)):
                if idx == phone_idx:
                    continue
                ws.cell(row=row_num, column=idx + 1).fill = green
        if phone_val:
            if chain_match:
                ws.cell(row=row_num, column=phone_idx + 1).fill = green
        else:
            ws.cell(row=row_num, column=phone_idx + 1).fill = yellow

    for col, w in {'A': 18, 'B': 36, 'C': 10, 'D': 12, 'E': 32, 'F': 18,
                   'G': 18, 'H': 12, 'I': 32, 'J': 18}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    wb.save(OUT)
    phones = sum(1 for row in merged if row[phone_idx])
    chains = sum(1 for row in merged if _scrape.is_chain_match(str(row[1] or '')))
    print(f"\n=== MERGED: {len(merged)} unique businesses ({phones} with phone, "
          f"{chains} chain stores highlighted) -> {OUT}", flush=True)


def main() -> None:
    dry = '--dry-run' in sys.argv
    cities = cities_from_districts_txt()
    total_q = len(cities) * len(KEYWORDS)
    print(f"Vijayapura (vijaypur) district: {len(cities)} cities x "
          f"{len(KEYWORDS)} keyword = {total_q} queries.", flush=True)
    print(f"Parts dir: {PARTS}", flush=True)
    print(f"Final output (ONE Excel): {OUT}", flush=True)
    if dry:
        for city in cities:
            place = f"{city}, {DISTRICT}, {STATE}"
            for kw in KEYWORDS:
                print(f"  - {kw} in {place}")
        print(f"\nDRY-RUN: {total_q} queries previewed, nothing scraped.")
        return
    PARTS.mkdir(exist_ok=True)
    failed: list[str] = []
    for i, city in enumerate(cities, 1):
        out_file = PARTS / f"{i:02d}_{slug(city)}.xlsx"
        if out_file.exists():
            rows = None
            try:
                rows = shops_rows(out_file)
            except Exception:
                rows = None
            if rows is not None:
                print(f"[{i}/{len(cities)}] SKIP {city} (already done, "
                      f"{len(rows)} rows)", flush=True)
                continue
        print(f"[{i}/{len(cities)}] {city}", flush=True)
        ok = run_city(city, out_file)
        if not ok:
            print(f"!!! FAILED: {city}", flush=True)
            failed.append(city)
    parts = sorted(PARTS.glob('*.xlsx'))
    if not parts:
        print("No part files produced — nothing to merge.", flush=True)
        sys.exit(1)
    merge(parts)
    if failed:
        print(f"Cities that failed: {', '.join(failed)}", flush=True)
        sys.exit(1)


if __name__ == "__main__":
    main()