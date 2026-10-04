# -*- coding: utf-8 -*-
"""Inspect both files to merge."""
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import openpyxl
from pathlib import Path

FILES = [
    r'C:\Users\Nomesh\Desktop\evaara\google_scraper\mobileaccessoriesvijapura_cleaned.xlsx',
    r'D:\dsa exam\mobile_vijaypur_all_CLEANED (1).xlsx',
]

for f in FILES:
    p = Path(f)
    print('=' * 80)
    print('FILE:', f, '| exists:', p.exists(), '| size:', p.stat().st_size if p.exists() else '-')
    if not p.exists():
        continue
    wb = openpyxl.load_workbook(p, read_only=True)
    print('  sheets:', wb.sheetnames)
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            print(f'  [{sn}] {len(rows)-1} data rows')
            print('    cols:', rows[0][:30])
            print('    row1:', rows[1][:30] if len(rows) > 1 else '-')
            print('    row2:', rows[2][:30] if len(rows) > 2 else '-')
        else:
            print(f'  [{sn}] (empty)')
    wb.close()
print('=' * 80)