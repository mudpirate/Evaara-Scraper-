# -*- coding: utf-8 -*-
# Restores chain-store rows (scrape.py CHAIN_NAMES) from the 'removed'
# sheet back into 'shops', then applies the original scrape.py highlights:
#   - chain-store row  -> green fill across the row (C6EFCE)
#   - missing phone    -> yellow fill on the phone cell (FFEB9C)
# Usage: python _restore_chains.py  (processes every *_CLEANED.xlsx)
import os, re, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'districts excel')

# Same list as scrape.py
CHAIN_NAMES = [
    'poorvika', 'sangeetha', 'pai mobiles', 'ss mobiles', 'mi store',
    'i central', 'i centre', 'icentre', 'icentral', 'invent tribe',
    'icare', 'imagine', 'iplanet', 'maple', 'unicorn', 'invent',
    'ivenus', 'idestiny', 'samsung cafe', 'oppo', 'vivo', 'phonorange',
    'phono range',
    'my jio store', 'ntel mobiles', 'sara mobiles', 'pai international',
    'samsung service center', 'reliance digital', 'xiaomi',
]

CHAIN_FILL = 'C6EFCE'   # light green highlight for chain-store rows
NO_PHONE_FILL = 'FFEB9C'  # light yellow highlight for missing phone numbers


def is_chain_match(title: str) -> bool:
    if not title:
        return False
    t = str(title).lower()
    return any(
        re.search(rf'(?<![a-z]){re.escape(k)}(?![a-z])', t)
        for k in CHAIN_NAMES
    )


def process(path):
    shops = pd.read_excel(path, sheet_name='shops')
    removed = pd.read_excel(path, sheet_name='removed')

    was_removed = len(shops)
    chain_mask = removed['title'].apply(is_chain_match)
    chains = removed[chain_mask].copy()
    rest = removed[~chain_mask].copy()

    shops = pd.concat([shops, chains], ignore_index=True)

    summary = pd.DataFrame([
        ['kept_shops', len(shops)],
        ['restored_chains', len(chains)],
        ['still_removed_irrelevant', len(rest)],
        ['chain_highlight', f'green {CHAIN_FILL} across row'],
        ['missing_phone_highlight', f'yellow {NO_PHONE_FILL} on phone cell'],
        ['chain_list', '; '.join(CHAIN_NAMES)],
    ])

    with pd.ExcelWriter(path, engine='openpyxl') as w:
        shops.to_excel(w, sheet_name='shops', index=False)
        rest.to_excel(w, sheet_name='removed', index=False)
        summary.to_excel(w, sheet_name='summary', index=False, header=False)

    # apply highlights
    wb = load_workbook(path)
    ws = wb['shops']
    headers = [c.value for c in ws[1]]
    phone_col = headers.index('phone') + 1
    title_col = headers.index('title') + 1
    green = PatternFill('solid', start_color=CHAIN_FILL)
    yellow = PatternFill('solid', start_color=NO_PHONE_FILL)
    n_chain = n_nophone = 0
    for row in range(2, ws.max_row + 1):
        title = ws.cell(row=row, column=title_col).value
        chain = is_chain_match(title)
        phone = str(ws.cell(row=row, column=phone_col).value or '').strip()
        if chain:
            n_chain += 1
            for col in range(1, len(headers) + 1):
                if col != phone_col:
                    ws.cell(row=row, column=col).fill = green
            if phone:
                ws.cell(row=row, column=phone_col).fill = green
            else:
                ws.cell(row=row, column=phone_col).fill = yellow
                n_nophone += 1
        elif not phone:
            ws.cell(row=row, column=phone_col).fill = yellow
            n_nophone += 1
    wb.save(path)
    print(f'{os.path.basename(path)}: shops {was_removed} -> {len(shops)} '
          f'(+{len(chains)} chains restored) | chain rows: {n_chain} | '
          f'missing-phone cells: {n_nophone}')


if __name__ == '__main__':
    for fn in sorted(os.listdir(FOLDER)):
        if fn.endswith('_CLEANED.xlsx'):
            process(os.path.join(FOLDER, fn))
