# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl"]
# ///
"""Final merge for Vijayapura: combine _vijaypur_parts/*.xlsx into ONE Excel
(vijaypura_all_keywords_all_cities.xlsx), duplicates removed.

Run after ALL workers finish:
    python _merge_vijaypur_final.py
"""
import sys
from pathlib import Path

HERE = Path(__file__).parent
PARTS = HERE / "_vijaypur_parts"
OUT = HERE / "vijaypura_all_keywords_all_cities.xlsx"


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


def merge() -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    import scrape as _scrape

    cols = ['title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
            'countryCode', 'website', 'phone'] + \
        [f'categories/{i}' for i in range(11)] + ['url', 'categoryName']
    parts = sorted(PARTS.glob('*.xlsx'))
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
    print(f"MERGED {len(merged)} unique from {len(parts)} parts "
          f"({phones} with phone, {chains} chains) -> {OUT}", flush=True)
    return OUT


if __name__ == "__main__":
    if not list(PARTS.glob('*.xlsx')):
        sys.exit("no part files to merge")
    merge()
