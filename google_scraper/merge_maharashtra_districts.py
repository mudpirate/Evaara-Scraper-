# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl"]
# ///
"""Merge all Maharashtra district workbooks into one deduplicated master.

Reads `<slug>_mobile_all.xlsx` for every district listed in
maharashtra/districts.json, deduplicates on the Google place ID found in each
row's Maps URL, keeps the richest copy of every shop, fills blank fields from
the other copies, and highlights chain stores in green.

Rows rejected by the category filter stay in a separate `filtered_out` sheet.

Usage:
    python merge_maharashtra_districts.py
    python merge_maharashtra_districts.py --out maharashtra/state_master.xlsx
    python merge_maharashtra_districts.py --only pune,nashik,thane
"""

from __future__ import annotations

import argparse
import json
import re
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from scrape import is_chain_match

ROOT = Path(__file__).resolve().parent
BASE = ROOT / 'maharashtra_districts'
PLAN = ROOT / 'maharashtra' / 'districts.json'
DEFAULT_OUT = BASE / 'maharashtra_merged_master.xlsx'

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
    for group in groups.values():
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
    if str(values[phone_idx] or '').strip():
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
        apply_row_style(ws, ws.max_row, values,
                        str(r.get('chainStore') or '').lower() == 'yes',
                        phone_idx)

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


def load_feeders(only: set[str]) -> list[tuple[str, str, Path]]:
    if not PLAN.exists():
        raise SystemExit(f'Missing {PLAN}; run fetch_maharashtra_districts.py')
    plan = json.loads(PLAN.read_text(encoding='utf-8'))
    feeders = []
    for district in plan['districts']:
        if only and district['slug'] not in only:
            continue
        path = BASE / f"{district['slug']}_mobile_all.xlsx"
        feeders.append((district['name'], district['slug'], path))
    return feeders


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', type=Path, default=DEFAULT_OUT)
    ap.add_argument('--only', default='',
                    help='comma-separated district names or slugs')
    args = ap.parse_args()

    only = {s.strip().lower().replace(' ', '_')
            for s in args.only.split(',') if s.strip()}
    feeders = load_feeders(only)
    if not feeders:
        print('No districts matched --only.')
        return 1

    all_rows: list[dict] = []
    all_filtered: list[dict] = []
    per_file: dict[str, dict] = {}
    summary_order: list[str] = []
    missing: list[str] = []

    for label, slug, path in feeders:
        sources: list[tuple[str, Path]] = []
        if path.exists():
            sources.append((label, path))
        else:
            missing.append(f'{label} ({path.name})')
        # Optional third-keyword pass, produced by -AccessoriesOnly.
        accessories = BASE / f'{slug}_accessories.xlsx'
        if accessories.exists():
            sources.append((f'{label} (accessories)', accessories))

        for source_label, source_path in sources:
            shops = read_sheet(source_path, 'shops')
            bad = read_sheet(source_path, 'filtered_out')
            outside = read_sheet(source_path, 'outside_district')
            for r in shops + bad:
                r['_source'] = source_label
                r['_file'] = source_path.name
                r['_pid'] = place_id(r)
            all_rows.extend(shops)
            all_filtered.extend(bad)
            per_file[source_label] = {
                'file': source_path.name, 'slug': slug,
                'rows': len(shops), 'filtered': len(bad),
                'outside': len(outside),
            }
            summary_order.append(source_label)
            print(f'+ {source_label:28s} shops={len(shops):4d} '
                  f'filtered={len(bad):3d} outside={len(outside):3d}')

    if missing:
        print(f'\n! missing workbooks ({len(missing)}):')
        for item in missing:
            print(f'   {item}')
    if not all_rows:
        print('\nNo rows to merge yet; run some districts first.')
        return 1

    print(f'\ninput rows = {len(all_rows)} '
          f'(no place id: {sum(1 for r in all_rows if not r["_pid"])})')

    shops, dupes = merge(all_rows)

    # Filtered rows stay separate; drop ones already kept as shops.
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

    ids = {r['_pid'] for r in all_rows if r['_pid']}
    print(f'unique shops = {len(shops)}')
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
        write_sheet(wb.create_sheet('filtered_out'), STD_COLS, filtered,
                    EXTRA_COLS)

    ws4 = wb.create_sheet('summary')
    ws4.append(['district', 'file', 'shops rows', 'filtered rows',
                'outside rows', 'distinct places', 'also in other files',
                'only in this file'])
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
    for label in summary_order:
        info = per_file[label]
        file_ids = per_ids.get(label, set())
        only_here = sum(1 for p in file_ids if pid_count.get(p) == 1)
        shared = sum(1 for p in file_ids if pid_count.get(p, 0) > 1)
        ws4.append([label, info['file'], info['rows'], info['filtered'],
                    info['outside'], len(file_ids), shared, only_here])

    outside_total = sum(info['outside'] for info in per_file.values())
    exclusive_total = sum(1 for c in pid_count.values() if c == 1)
    ws4.append(['TOTAL', '', len(all_rows), len(all_filtered), outside_total,
                len(ids), len(dupes), exclusive_total])
    ws4.append([])
    ws4.append(['metric', 'value'])
    ws4.append(['districts merged', len(per_file)])
    ws4.append(['districts missing', len(missing)])
    ws4.append(['outside-district rows excluded', outside_total])
    ws4.append(['total input rows', len(all_rows)])
    ws4.append(['duplicate rows removed', len(dupes)])
    ws4.append(['duplicate rate',
                f'{len(dupes) / max(len(all_rows), 1) * 100:.1f}%'])
    ws4.append(['unique shops in merged file', len(shops)])
    ws4.append(['chain-store rows marked (green)',
                sum(1 for r in shops if r.get('chainStore') == 'Yes')])
    ws4.append(['rows with no phone (yellow)',
                sum(1 for r in shops if not str(r.get('phone') or '').strip())])
    ws4.append(['shops with coordinates',
                sum(1 for r in shops
                    if str(r.get('latitude') or '').strip())])
    ws4.column_dimensions['A'].width = 36
    ws4.column_dimensions['B'].width = 44
    for col in ('C', 'D', 'E', 'F', 'G', 'H'):
        ws4.column_dimensions[col].width = 18

    args.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    print(f'\nchain-store rows marked green: '
          f"{sum(1 for r in shops if r.get('chainStore') == 'Yes')}")
    print(f'saved -> {args.out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
