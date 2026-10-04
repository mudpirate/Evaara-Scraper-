# -*- coding: utf-8 -*-
"""Analyze mobileaccessoriesvijapura.xlsx borderline rows."""
import sys
import openpyxl
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SRC = 'mobileaccessoriesvijapura.xlsx'
GREEN = 'C6EFCE'

wb = openpyxl.load_workbook(SRC, read_only=True)
ws = wb['shops']
rows = list(ws.iter_rows(values_only=True))
cols = [str(c).lower() for c in rows[0]]
wb.close()

# identify green rows using normal (non read-only) load
wb2 = openpyxl.load_workbook(SRC)
ws2 = wb2['shops']
green_titles = set()
for r in ws2.iter_rows(min_row=2):
    vals = [c.value for c in r]
    title = str(vals[1] or '')
    fills = [(c.fill.patternType or '', getattr(c.fill.fgColor, 'rgb', None)) for c in r[:10]]
    is_green = any(pt == 'solid' and rgb and str(rgb)[-6:].upper() == GREEN for pt, rgb in fills)
    if is_green:
        green_titles.add((title, str(vals[5] or ''), str(vals[-1] or '')))
wb2.close()

print('=== GREEN (chain) rows:', len(green_titles), '===')
for t, c, cat in sorted(green_titles):
    print(f'  {t} | {c} | {cat}')

# Borderline categories worth seeing
border_cats = {'Electronics repair shop', 'Phone repair service', 'Telecommunications service provider',
               'Electronics Company', 'Store', 'Repair service', 'Car stereo store', 'Appliance store',
               'Computer store', 'Computer accessories store', 'Internet shop', 'Watch store',
               'none', 'None', ''}
print('\n=== BORDERLINE CATEGORY ROWS ===')
for r in rows[1:]:
    cat = str(r[-1] or '').strip()
    if cat in border_cats and (str(r[1]), str(r[5]), cat) not in green_titles:
        print(f'  {str(r[1])[:60]:60s} | {str(r[5])[:30]:30s} | {cat}')