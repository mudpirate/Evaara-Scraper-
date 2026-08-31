# /// script
# requires-python = ">=3.11"
# dependencies = ["rich", "openpyxl"]
# ///
"""
Resumable per-city Haveri district scraper.

Runs scrape.py once per city (city, haveri, karnataka). Each city is saved to
its own file in _haveri_parts/. Cities whose part file already exists are
SKIPPED, so the run can be resumed after a crash/captcha timeout without
losing finished cities. At the end all parts are merged (duplicates removed)
into universal_haveri_district.xlsx.
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

DISTRICT = "haveri"
KEYWORD = "mobile phones"
STATE = "karnataka"
HERE = Path(__file__).parent
PARTS = HERE / "_haveri_parts"
OUT = HERE / "universal_haveri_district.xlsx"
LINE_FILE = HERE / "_haveri_only.txt"


def cities_from_line_file() -> list[str]:
    text = LINE_FILE.read_text(encoding='utf-8').strip()
    _, _, cities = text.partition(':')
    return [c.strip() for c in cities.split(',') if c.strip()]


def slug(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:40]


def run_city(city: str, out_file: Path) -> bool:
    """Run scrape.py for one city. Returns True on success."""
    env = os.environ.copy()
    env['KEEP_ALL'] = '1'
    env.setdefault('HEADLESS', '0')
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUNBUFFERED'] = '1'
    place = f"{city}, {DISTRICT}, {STATE}"
    cmd = [sys.executable, str(HERE / 'scrape.py'), KEYWORD, place, str(out_file)]
    print(f"\n=== RUN: {KEYWORD} in {place}", flush=True)
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
        return []
    ws = wb['shops']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    return rows


def merge(parts: list[Path]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    import scrape as _scrape  # reuse the exact CHAIN_NAMES + highlight logic

    cols = ['title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
            'countryCode', 'website', 'phone'] + \
        [f'categories/{i}' for i in range(11)] + ['url', 'categoryName']
    seen: set = set()
    merged: list[tuple] = []
    for p in parts:
        for r in shops_rows(p):
            r = (tuple(r) + ('',) * 24)[:24]
            key = (str(r[0]).lower().strip(), str(r[3]).strip(),
                   str(r[4]).strip(), str(r[8]).strip())
            if key in seen:
                continue
            seen.add(key)
            merged.append(r)

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
    for r in merged:
        ws.append(list(r))
        row_num = ws.max_row
        chain_match = _scrape.is_chain_match(str(r[0] or ''))
        phone_val = str(r[phone_idx] or '').strip()
        if chain_match:
            for idx in range(len(r)):
                if idx == phone_idx:
                    continue
                ws.cell(row=row_num, column=idx + 1).fill = green
        if phone_val:
            if chain_match:
                ws.cell(row=row_num, column=phone_idx + 1).fill = green
        else:
            ws.cell(row=row_num, column=phone_idx + 1).fill = yellow

    for col, w in {'A': 36, 'B': 10, 'C': 12, 'D': 32, 'E': 18, 'F': 18,
                   'G': 12, 'H': 32, 'I': 18}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    wb.save(OUT)
    phones = sum(1 for r in merged if r[8])
    chains = sum(1 for r in merged if _scrape.is_chain_match(str(r[0] or '')))
    print(f"\n=== MERGED: {len(merged)} unique businesses ({phones} with phone, "
          f"{chains} chain stores highlighted) -> {OUT}", flush=True)


def main() -> None:
    cities = cities_from_line_file()
    PARTS.mkdir(exist_ok=True)
    print(f"Haveri district: {len(cities)} cities, resumable per-city run.",
          flush=True)
    failed: list[str] = []
    for i, city in enumerate(cities, 1):
        out_file = PARTS / f"{i:02d}_{slug(city)}.xlsx"
        if out_file.exists():
            try:
                done = len(shops_rows(out_file)) > 0
            except Exception:
                done = False
            if done:
                print(f"[{i}/{len(cities)}] SKIP {city} (already done)",
                      flush=True)
                continue
        print(f"[{i}/{len(cities)}] {city}", flush=True)
        ok = run_city(city, out_file)
        if not ok:
            print(f"!!! FAILED: {city}", flush=True)
            failed.append(city)
    parts = sorted(PARTS.glob('*.xlsx'))
    merge(parts)
    if failed:
        print(f"Cities that failed: {', '.join(failed)}", flush=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
