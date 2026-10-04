# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import openpyxl
from collections import Counter

p = r'C:\Users\Nomesh\Desktop\evaara\google_scraper\_mobile_vijaypur_merged_all.xlsx'
wb = openpyxl.load_workbook(p, read_only=True)
print('sheets:', wb.sheetnames)
ws = wb['shops']
rows = list(ws.iter_rows(values_only=True))
print('header (first 12):', rows[0][:12])
print('total data rows:', len(rows) - 1)
cats = Counter(str(r[-1] or '') for r in rows[1:])
print('categoryName counts:')
for k, v in cats.most_common():
    print(f'   {k!r} = {v}')

# how many have empty citySearched (from File B)
nosearch = sum(1 for r in rows[1:] if not str(r[0] or '').strip())
hassearch = sum(1 for r in rows[1:] if str(r[0] or '').strip())
print(f'\nrows WITH citySearched: {hassearch}  (from file A)')
print(f'rows WITHOUT citySearched: {nosearch}  (from file B)')
wb.close()