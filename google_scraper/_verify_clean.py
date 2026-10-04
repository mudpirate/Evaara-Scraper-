# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import openpyxl
from collections import Counter

wb = openpyxl.load_workbook('mobileaccessoriesvijapura_cleaned.xlsx')
ws = wb['shops']
print('kept rows:', ws.max_row - 1)
cats = Counter(str(r[-1].value or '').strip() for r in ws.iter_rows(min_row=2))
print('categories:')
for k, v in cats.most_common():
    print(f'   {k!r} = {v}')

# show any row whose category is not in the keep set (should be none)
KEEP = {'cell phone store', 'cell phone accessory store',
        'mobile phone repair shop', 'electronics store',
        'electronics repair shop', 'phone repair service',
        'electronics company'}
print('\nRows kept with NON-keep categories:')
n = 0
for r in ws.iter_rows(min_row=2):
    cat = str(r[-1].value or '').strip().lower()
    if cat not in KEEP:
        print(f'   [{cat}] {r[1].value}')
        n += 1
print('count:', n)

ws2 = wb['removed']
print('\nremoved rows:', ws2.max_row - 1)
rems = Counter(str(r[-1].value) for r in ws2.iter_rows(min_row=2))
for k, v in rems.most_common(12):
    print(f'   {k} = {v}')