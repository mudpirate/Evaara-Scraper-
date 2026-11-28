"""Merge two half-district accessory workbooks into one, deduplicated.

Used when a district is split across two parallel workers. Rows are matched on
the Google place ID in the Maps URL (same key as merge_maharashtra_districts.py),
with a title/street/city fallback for rows that have no parseable ID.

Usage:
    python merge_accessories_halves.py <base.xlsx> <extra.xlsx> <out.xlsx>
"""
from __future__ import annotations

import re
import sys
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

PID_HEX = re.compile(r'!1s(0x[0-9a-fA-F]+:0x[0-9a-fA-F]+)')
PID_B64 = re.compile(r'!19s(ChIJ[A-Za-z0-9_\-]+)')
CID = re.compile(r'[?&]cid=(\d+)')


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


def row_key(row: dict) -> str:
    pid = place_id(row)
    if pid:
        return pid
    return 'text:' + '|'.join(
        norm(row.get(k)) for k in ('title', 'street', 'city')
    )


def sheet_rows(ws) -> tuple[list[str], list[tuple]]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(h) if h is not None else '' for h in rows[0]]
    return header, rows[1:]


def main() -> None:
    if len(sys.argv) != 4:
        sys.exit('usage: python merge_accessories_halves.py <base.xlsx> <extra.xlsx> <out.xlsx>')

    base_path = Path(sys.argv[1])
    extra_path = Path(sys.argv[2])
    out_path = Path(sys.argv[3])

    for p in (base_path, extra_path):
        if not p.exists():
            sys.exit(f'missing input: {p}')

    base = load_workbook(base_path)
    extra = load_workbook(extra_path)

    total_added = 0
    for sheet in extra.sheetnames:
        xws = extra[sheet]
        xheader, xrows = sheet_rows(xws)
        if not xheader:
            continue

        if sheet not in base.sheetnames:
            bws = base.create_sheet(sheet)
            bws.append(xheader)
            for col_idx in range(1, len(xheader) + 1):
                try:
                    bws.cell(row=1, column=col_idx)._style = copy(
                        xws.cell(row=1, column=col_idx)._style)
                except Exception:
                    pass
            bheader = xheader
            existing: set[str] = set()
        else:
            bws = base[sheet]
            bheader, brows = sheet_rows(bws)
            existing = {
                row_key(dict(zip(bheader, r))) for r in brows
                if any(v not in (None, '') for v in r)
            }

        added = 0
        for src_row, r in enumerate(xrows, start=2):
            if not any(v not in (None, '') for v in r):
                continue
            row = dict(zip(xheader, r))
            key = row_key(row)
            if key in existing:
                continue
            existing.add(key)
            bws.append([row.get(h, '') for h in bheader])
            target_row = bws.max_row
            for col_idx in range(1, len(bheader) + 1):
                try:
                    src = xws.cell(row=src_row, column=col_idx)
                    dst = bws.cell(row=target_row, column=col_idx)
                    if src.has_style:
                        dst._style = copy(src._style)
                except Exception:
                    pass
            added += 1

        if added:
            total_added += added
            print(f'{sheet}: +{added} rows')

    base.save(out_path)
    print(f'wrote {out_path} (added {total_added} unique rows)')


if __name__ == '__main__':
    main()
