# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl"]
# ///
"""Split a district workbook into in-district shops and border spill-over.

The scraper clips search cells to the district polygon, but a Maps search at a
border cell still returns shops from the neighbouring district or state. Older
workbooks were filtered by the rectangular bbox only, so they can contain
Belagavi (KA), Sangli, Sindhudurg, etc.

This script tests every row's coordinates against the district polygon and
rewrites the workbook with:

    shops             only rows inside the district
    outside_district  rows outside, with a sourceSheet column
    filtered_out      category-rejected rows (inside only)
    summary           counts, including the KA/MA split

Usage:
    python clean_district_polygon.py kolhapur --apply
    python clean_district_polygon.py kolhapur solapur
    python clean_district_polygon.py kolhapur --out cleaned.xlsx
"""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from grid import load_polygon_geojson
from scrape import is_chain_match

ROOT = Path(__file__).resolve().parent
BASE = ROOT / 'maharashtra_districts'
PLAN = ROOT / 'maharashtra' / 'districts.json'

CHAIN_FILL = 'C6EFCE'
NO_PHONE_FILL = 'FFEB9C'
HEADER_FILL = 'DDDDDD'


def read_table(path: Path, sheet: str) -> tuple[list[str], list[dict]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            return [], []
        rows = wb[sheet].iter_rows(values_only=True)
        try:
            header = [str(c or '').strip() for c in next(rows)]
        except StopIteration:
            return [], []
        out = []
        for values in rows:
            if not values or all(v is None or str(v).strip() == '' for v in values):
                continue
            out.append({c: values[i] if i < len(values) else None
                        for i, c in enumerate(header) if c})
        return [h for h in header if h], out
    finally:
        wb.close()


def write_table(ws, header: list[str], rows: list[dict],
                extra: tuple[str, ...] = ()) -> None:
    full = header + [c for c in extra if c not in header]
    ws.append(full)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color=HEADER_FILL)
        cell.alignment = Alignment(horizontal='center')

    phone_idx = full.index('phone') if 'phone' in full else None
    green = PatternFill('solid', start_color=CHAIN_FILL)
    yellow = PatternFill('solid', start_color=NO_PHONE_FILL)
    for row in rows:
        values = [row.get(c, '') for c in full]
        ws.append(values)
        chain = is_chain_match(row.get('title', '') or '')
        if chain:
            for idx in range(len(values)):
                if idx != phone_idx:
                    ws.cell(row=ws.max_row, column=idx + 1).fill = green
        if phone_idx is not None:
            phone = str(values[phone_idx] or '').strip()
            if phone:
                if chain:
                    ws.cell(row=ws.max_row, column=phone_idx + 1).fill = green
            else:
                ws.cell(row=ws.max_row, column=phone_idx + 1).fill = yellow

    widths = {'title': 38, 'street': 34, 'city': 18, 'state': 16,
              'website': 32, 'phone': 18, 'url': 40, 'categoryName': 26,
              'latitude': 12, 'longitude': 12, 'sourceSheet': 14,
              'countryCode': 12, 'chainStore': 11, 'foundIn': 44}
    for idx, name in enumerate(full, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(name, 16)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions


def load_plan() -> dict:
    data = json.loads(PLAN.read_text(encoding='utf-8'))
    return {d['slug']: d for d in data['districts']}


def split_rows(rows: list[dict], polygons) -> tuple[list[dict], list[dict], int]:
    inside: list[dict] = []
    outside: list[dict] = []
    no_coords = 0
    for row in rows:
        try:
            lat = float(row['latitude'])
            lon = float(row['longitude'])
        except (KeyError, TypeError, ValueError):
            no_coords += 1
            inside.append(row)  # cannot test it, so do not throw it away
            continue
        if any(poly.contains(lat, lon) for poly in polygons):
            inside.append(row)
        else:
            outside.append(row)
    return inside, outside, no_coords


def clean_one(slug: str, apply: bool, out_override: Path | None) -> None:
    plan = load_plan()
    if slug not in plan:
        print(f'! unknown district: {slug}')
        return
    record = plan[slug]
    src = BASE / f'{slug}_mobile_all.xlsx'
    polygon_file = ROOT / 'maharashtra' / record['polygon']
    if not src.exists():
        print(f'! missing workbook: {src}')
        return
    if not polygon_file.exists():
        print(f'! missing polygon: {polygon_file}')
        return

    polygons = load_polygon_geojson(polygon_file)
    shop_header, shops = read_table(src, 'shops')
    filt_header, filtered = read_table(src, 'filtered_out')
    if not shop_header:
        print(f'! no shops sheet in {src.name}')
        return
    if not filt_header:
        filt_header = shop_header

    shops_in, shops_out, no_coords = split_rows(shops, polygons)
    filt_in, filt_out, _ = split_rows(filtered, polygons)

    for row in shops_out:
        row['sourceSheet'] = 'shops'

    ka = sum(1 for r in shops_out
             if str(r.get('countryCode') or '').upper() == 'KA')
    ma = sum(1 for r in shops_out
             if str(r.get('countryCode') or '').upper() == 'MA')

    print(f'\n{record["name"]}:')
    print(f'  shops total       : {len(shops)}')
    print(f'  inside district   : {len(shops_in)}')
    print(f'  outside district  : {len(shops_out)}  (KA {ka} / MA {ma})')
    print(f'  no coordinates    : {no_coords} (kept)')
    print(f'  filtered total    : {len(filtered)}  '
          f'(inside {len(filt_in)} / outside {len(filt_out)})')

    out = out_override or (src if apply else src.with_name(f'{src.stem}_clean.xlsx'))
    if apply and not out_override:
        backup = src.with_name(f'{src.stem}_raw.xlsx')
        if not backup.exists():
            shutil.copy2(src, backup)
            print(f'  backup            : {backup.name}')

    wb = Workbook()
    ws = wb.active
    ws.title = 'shops'
    write_table(ws, shop_header, shops_in)

    if shops_out:
        write_table(wb.create_sheet('outside_district'), shop_header,
                    shops_out, extra=('sourceSheet',))

    ws3 = wb.create_sheet('filtered_out')
    write_table(ws3, filt_header, filt_in)

    summary = wb.create_sheet('summary')
    summary.append(['metric', 'value'])
    for cell in summary[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color=HEADER_FILL)
    rows = [
        ('district', record['name']),
        ('rows in shops sheet (original)', len(shops)),
        ('shops inside district', len(shops_in)),
        ('shops outside district', len(shops_out)),
        ('outside with KA state code', ka),
        ('outside with MA state code', ma),
        ('shops with no coordinates (kept)', no_coords),
        ('filtered_out rows (original)', len(filtered)),
        ('filtered_out inside district', len(filt_in)),
        ('filtered_out outside district (dropped)', len(filt_out)),
        ('shops with phone', sum(
            1 for r in shops_in if str(r.get('phone') or '').strip())),
        ('shops without phone', sum(
            1 for r in shops_in if not str(r.get('phone') or '').strip())),
    ]
    for label, value in rows:
        summary.append([label, value])
    summary.column_dimensions['A'].width = 40
    summary.column_dimensions['B'].width = 50

    wb.save(out)
    print(f'  saved             : {out}')


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('districts', nargs='+', help='district slugs, e.g. kolhapur')
    ap.add_argument('--apply', action='store_true',
                    help='overwrite the workbook (keeps a _raw backup)')
    ap.add_argument('--out', type=Path, default=None,
                    help='write one cleaned file here (single district only)')
    args = ap.parse_args()

    if args.out and len(args.districts) > 1:
        print('--out only works with a single district.')
        return 1

    for slug in args.districts:
        clean_one(slug.strip().lower().replace(' ', '_'), args.apply, args.out)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
