# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl"]
# ///
"""Merge all Solapur taluka / area workbooks into one deduplicated master.

Reads every scrape output in google_scraper/solapur_talukas (taluka runs plus
the North/South Solapur and Barshi masters), deduplicates on the Google place
ID found in each row's Maps URL, keeps the richest copy of every shop, fills
blank fields from the other copies, and highlights chain stores in green.

PhoneWale counts as a chain store (added to CHAIN_NAMES in scrape.py).

Usage:
    python merge_solapur_talukas.py [--out PATH]
"""

from __future__ import annotations

import argparse
import re
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from scrape import is_chain_match

BASE = Path(__file__).resolve().parent / 'solapur_talukas'
DEFAULT_OUT = BASE / 'solapur_talukas_merged_master.xlsx'

# Taluka runs first (primary coverage), then the city-area masters.
FILES: list[tuple[str, str]] = [
    ('Pandharpur',      'pandharpur_mobile_all.xlsx'),
    ('Malshiras',       'malshiras_mobile_all.xlsx'),
    ('Sangola',         'sangola_mobile_all.xlsx'),
    ('Akkalkot',        'akkalkot_mobile_all.xlsx'),
    ('Madha',           'madha_mobile_all.xlsx'),
    ('Mohol',           'mohol_mobile_all.xlsx'),
    ('Karmala',         'karmala_grid_all_3_keywords.xlsx'),
    ('Karmala (part)',  'karmala_mobile_all_partial.xlsx'),
    ('Mangalwedha',     'mangalwedha_mobile_all.xlsx'),
    ('North Solapur',   'north_solapur_mobile_master.xlsx'),
    ('South Solapur',   'south_solapur_mobile_master.xlsx'),
    ('Barshi',          'barshi_mobile_master.xlsx'),
]

STD_COLS = [
    'title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
    'countryCode', 'website', 'phone',
] + [f'categories/{i}' for i in range(11)] + [
    'url', 'categoryName', 'latitude', 'longitude',
]
EXTRA_COLS = ['chainStore', 'foundIn', 'sourceCount']

FIELDS_FOR_RICHNESS = (
    'phone', 'website', 'street', 'url', 'totalScore', 'reviewsCount',
    'categoryName', 'latitude', 'longitude',
)

PID_HEX = re.compile(r'!1s(0x[0-9a-fA-F]+:0x[0-9a-fA-F]+)')
PID_B64 = re.compile(r'!19s(ChIJ[A-Za-z0-9_\-]+)')
CID = re.compile(r'[?&]cid=(\d+)')
PHONEWALE = re.compile(r'phone[\s\-_]*wale', re.I)

CHAIN_FILL = 'C6EFCE'
NO_PHONE_FILL = 'FFEB9C'
HEADER_FILL = 'DDDDDD'


def norm(value) -> str:
    return re.sub(r'[^\w]+', ' ', str(value or '').strip().lower()).strip()


def place_id(row: dict) -> str | None:
    url = str(row.get('url') or '')
    m = PID_HEX.search(url) or PID_B64.search(url)
    if m:
        return m.group(1).lower()
    m = CID.search(url)
    if m:
        return 'cid:' + m.group(1)
    return None


def read_sheet(path: Path, sheet: str) -> list[dict]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            return []
        ws = wb[sheet]
        rows = ws.iter_rows(values_only=True)
        try:
            header = next(rows)
        except StopIteration:
            return []
        cols = [str(c or '').strip() for c in header]
        if 'title' not in cols:
            return []
        out = []
        for values in rows:
            if not values or all(v is None or str(v).strip() == '' for v in values):
                continue
            r = {c: values[i] if i < len(values) else None
                 for i, c in enumerate(cols) if c}
            r['categoryName'] = r.get('categoryName') or r.get('categories/0') or ''
            out.append(r)
        return out
    finally:
        wb.close()


def richness(row: dict) -> int:
    return sum(1 for f in FIELDS_FOR_RICHNESS if str(row.get(f) or '').strip())


def merge(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    groups: 'OrderedDict[str, list[dict]]' = OrderedDict()
    for r in rows:
        pid = r.get('_pid')
        key = pid or 'text:' + '|'.join(
            (norm(r.get('title')), norm(r.get('street')), norm(r.get('city')))
        )
        groups.setdefault(key, []).append(r)

    shops: list[dict] = []
    dupes: list[dict] = []
    for key, group in groups.items():
        # Richest copy first; ties keep input order (taluka files come first).
        group.sort(key=richness, reverse=True)
        base = dict(group[0])
        sources: list[str] = []
        for row in group:
            src = row['_source']
            if src not in sources:
                sources.append(src)
            for field in STD_COLS:
                if str(base.get(field) or '').strip() == '' and \
                        str(row.get(field) or '').strip() != '':
                    base[field] = row[field]
        base['foundIn'] = '; '.join(sources)
        base['sourceCount'] = len(sources)
        base['chainStore'] = 'Yes' if is_chain_match(base.get('title', '')) else ''
        shops.append(base)

        for row in group[1:]:
            d = dict(row)
            d['duplicateOf'] = base.get('title')
            d['keptIn'] = base['foundIn']
            dupes.append(d)
    return shops, dupes


def apply_row_style(ws, row_num: int, values: list, chain: bool,
                    phone_idx: int) -> None:
    green = PatternFill('solid', start_color=CHAIN_FILL)
    yellow = PatternFill('solid', start_color=NO_PHONE_FILL)
    if chain:
        for idx in range(len(values)):
            if idx != phone_idx:
                ws.cell(row=row_num, column=idx + 1).fill = green
    phone_val = str(values[phone_idx] or '').strip()
    if phone_val:
        if chain:
            ws.cell(row=row_num, column=phone_idx + 1).fill = green
    else:
        ws.cell(row=row_num, column=phone_idx + 1).fill = yellow


def write_sheet(ws, cols: list[str], rows: list[dict], extra: list[str]) -> None:
    header = cols + extra
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color=HEADER_FILL)
        cell.alignment = Alignment(horizontal='center')
    phone_idx = header.index('phone')
    for r in rows:
        values = [r.get(c, '') for c in header]
        ws.append(values)
        chain = str(r.get('chainStore') or '').strip().lower() == 'yes'
        apply_row_style(ws, ws.max_row, values, chain, phone_idx)

    widths = {
        'title': 38, 'totalScore': 10, 'reviewsCount': 12, 'street': 34,
        'city': 18, 'state': 16, 'countryCode': 11, 'website': 32,
        'phone': 18, 'url': 40, 'categoryName': 26, 'latitude': 12,
        'longitude': 12, 'chainStore': 11, 'foundIn': 44, 'sourceCount': 12,
    }
    for idx, name in enumerate(header, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(name, 16)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', type=Path, default=DEFAULT_OUT)
    args = ap.parse_args()

    all_rows: list[dict] = []
    all_filtered: list[dict] = []
    per_file: dict[str, dict] = {}

    for label, name in FILES:
        path = BASE / name
        if not path.exists():
            print(f'! missing: {name}')
            continue
        shops = read_sheet(path, 'shops')
        bad = read_sheet(path, 'filtered_out')
        for r in shops + bad:
            r['_source'] = label
            r['_file'] = name
            r['_pid'] = place_id(r)
        all_rows.extend(shops)
        all_filtered.extend(bad)
        per_file[label] = {'file': name, 'rows': len(shops), 'filtered': len(bad)}
        print(f'+ {label:15s} {name:42s} shops={len(shops):4d} filtered={len(bad):3d}')

    missing_pid = [r for r in all_rows if not r['_pid']]
    print(f'\ninput rows = {len(all_rows)}  (no place id: {len(missing_pid)})')

    shops, dupes = merge(all_rows)

    # Filtered rows are kept separate; drop ones already present in shops.
    kept_ids = {r['_pid'] for r in shops if r['_pid']}
    fgroups: 'OrderedDict[str, list[dict]]' = OrderedDict()
    for r in all_filtered:
        key = r['_pid'] or 'text:' + '|'.join(
            (norm(r.get('title')), norm(r.get('street')), norm(r.get('city')))
        )
        fgroups.setdefault(key, []).append(r)
    filtered: list[dict] = []
    for key, group in fgroups.items():
        if key in kept_ids:
            continue
        group.sort(key=richness, reverse=True)
        base = dict(group[0])
        base['chainStore'] = 'Yes' if is_chain_match(base.get('title', '')) else ''
        sources = []
        for row in group:
            if row['_source'] not in sources:
                sources.append(row['_source'])
        base['foundIn'] = '; '.join(sources)
        base['sourceCount'] = len(sources)
        filtered.append(base)

    # Summary numbers.
    ids = {r['_pid'] for r in all_rows if r['_pid']}
    print(f'unique places = {len(ids)}')
    print(f'duplicate rows removed = {len(dupes)} '
          f'({len(dupes) / max(len(all_rows), 1) * 100:.1f}%)')
    print(f'filtered_out (deduped, extras) = {len(filtered)}')

    wb = Workbook()
    ws = wb.active
    ws.title = 'shops'
    write_sheet(ws, STD_COLS, shops, EXTRA_COLS)

    ws2 = wb.create_sheet('duplicates')
    dup_extra = ['sourceFile', 'duplicateOf', 'keptIn']
    header = STD_COLS + dup_extra
    ws2.append(header)
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color=HEADER_FILL)
        cell.alignment = Alignment(horizontal='center')
    phone_idx = header.index('phone')
    for r in dupes:
        values = [
            r.get('_file', '') if c == 'sourceFile' else r.get(c, '')
            for c in header
        ]
        ws2.append(values)
        apply_row_style(ws2, ws2.max_row, values,
                        is_chain_match(r.get('title', '')), phone_idx)
    for idx, name in enumerate(header, start=1):
        widths = {'title': 38, 'street': 34, 'city': 18, 'phone': 18,
                  'url': 40, 'categoryName': 26, 'sourceFile': 40,
                  'duplicateOf': 38, 'keptIn': 44, 'latitude': 12,
                  'longitude': 12, 'website': 32}
        ws2.column_dimensions[get_column_letter(idx)].width = widths.get(name, 16)
    ws2.freeze_panes = 'A2'
    ws2.auto_filter.ref = ws2.dimensions

    if filtered:
        ws3 = wb.create_sheet('filtered_out')
        write_sheet(ws3, STD_COLS, filtered, EXTRA_COLS)

    # Summary sheet.
    ws4 = wb.create_sheet('summary')
    ws4.append(['source', 'file', 'shops rows', 'filtered rows',
                'distinct places', 'also in other files', 'only in this file'])
    for cell in ws4[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color=HEADER_FILL)

    pid_count: dict[str, int] = {}
    for r in all_rows:
        if r['_pid']:
            pid_count[r['_pid']] = pid_count.get(r['_pid'], 0) + 1
    per_ids: dict[str, set] = {}
    for r in all_rows:
        per_ids.setdefault(r['_source'], set()).add(r['_pid'])
    exclusive_total = sum(1 for c in pid_count.values() if c == 1)
    for label, _name in FILES:
        if label not in per_file:
            continue
        info = per_file[label]
        file_ids = per_ids.get(label, set())
        only = sum(1 for p in file_ids if pid_count.get(p) == 1)
        shared = sum(1 for p in file_ids if pid_count.get(p, 0) > 1)
        ws4.append([label, info['file'], info['rows'], info['filtered'],
                    len(file_ids), shared, only])
    ws4.append(['TOTAL', '', len(all_rows), len(all_filtered), len(ids),
                len(dupes), exclusive_total])
    ws4.append([])
    phonewale_in = sum(1 for r in all_rows if PHONEWALE.search(str(r.get('title') or '')))
    phonewale_out = sum(1 for r in shops if PHONEWALE.search(str(r.get('title') or '')))
    ws4.append(['metric', 'value'])
    ws4.append(['total input rows', len(all_rows)])
    ws4.append(['duplicate rows removed', len(dupes)])
    ws4.append(['duplicate rate', f'{len(dupes) / max(len(all_rows), 1) * 100:.1f}%'])
    ws4.append(['unique shops in merged file', len(shops)])
    ws4.append(['PhoneWale rows in input', phonewale_in])
    ws4.append(['PhoneWale shops after dedupe', phonewale_out])
    ws4.append(['chain-store rows marked (green)', sum(
        1 for r in shops if r.get('chainStore') == 'Yes')])
    ws4.column_dimensions['A'].width = 34
    ws4.column_dimensions['B'].width = 40
    for col in ('C', 'D', 'E', 'F', 'G'):
        ws4.column_dimensions[col].width = 18

    out = args.out
    wb.save(out)
    print(f'\nPhoneWale: {phonewale_in} input rows -> {phonewale_out} unique shops')
    print(f'chain-store rows marked green: '
          f"{sum(1 for r in shops if r.get('chainStore') == 'Yes')}")
    print(f'saved -> {out}')


if __name__ == '__main__':
    main()
