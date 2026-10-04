# -*- coding: utf-8 -*-
"""
Merge two cleaned Vijayapura Excel files and remove duplicates across both.

Files:
  A = mobileaccessoriesvijapura_cleaned.xlsx  (google_scraper)
  B = mobile_vijaypur_all_CLEANED (1).xlsx    (D:/dsa exam)

File A has a leading 'citySearched' column; File B does not, so File B rows get
a blank citySearched.  Rows are deduped by normalized (title, street, city,
phone).  Output:  mobile_vijaypur_merged_all.xlsx
"""
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SRC_A = Path(r'C:\Users\Nomesh\Desktop\evaara\google_scraper\mobileaccessoriesvijapura_cleaned.xlsx')
SRC_B = Path(r'D:\dsa exam\mobile_vijaypur_all_CLEANED (1).xlsx')
DST = Path(r'C:\Users\Nomesh\Desktop\evaara\google_scraper\_mobile_vijaypur_merged_all.xlsx')

# Canonical column order — File A shape (with citySearched) extended to any width.
BASE = ['title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
        'countryCode', 'website', 'phone']
COLS = ['citySearched'] + BASE + [f'categories/{i}' for i in range(11)] + ['url', 'categoryName']
COLS_SET = set(COLS)


def norm(v):
    return '' if v is None else str(v).strip().lower()


def row_to_dict(row_vals, col_names):
    """Map a row of values + its header names into the canonical COLS order."""
    d = dict(zip(col_names, row_vals))
    return [d.get(c, '') for c in COLS]


def load_shops(path):
    """Return (col_names, list_of_rows).  Reads the 'shops' sheet."""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb['shops']
    it = ws.iter_rows(values_only=True)
    try:
        header = next(it)
    except StopIteration:
        wb.close()
        return [], []
    rows = [list(r) for r in it]
    wb.close()
    return list(header), rows


def dedup_key(vals):
    return (norm(vals[1]),   # title
            norm(vals[4]),   # street
            norm(vals[5]),   # city
            norm(vals[9]))   # phone


def main():
    total_from, seen, merged = {'A': 0, 'B': 0}, set(), []

    for label, path in (('A', SRC_A), ('B', SRC_B)):
        if not path.exists():
            print(f'!! {label} missing: {path}', flush=True)
            sys.exit(1)
        header, rows = load_shops(path)
        added = 0
        for r in rows:
            vals = row_to_dict(r, header)
            if len(vals) != len(COLS):
                vals = (vals + [''] * len(COLS))[:len(COLS)]
            total_from[label] += 1
            key = dedup_key(vals)
            if key in seen:
                continue
            if label == 'B':
                vals[0] = ''  # File B has no citySearched column
            seen.add(key)
            merged.append(vals)
            added += 1
        print(f'{label}: total {total_from[label]}  | kept {added} '
              f'| dupes {total_from[label] - added}', flush=True)

    # ---- write output ----
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'shops'
    ws.append(COLS)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', start_color='DDDDDD')
        c.alignment = Alignment(horizontal='center', vertical='center')

    phone_col = COLS.index('phone')
    thin = Side(style='thin', color='CCCCCC')
    for vals in merged:
        ws.append(vals)
        rn = ws.max_row
        for ci, val in enumerate(vals, 1):
            ws.cell(row=rn, column=ci).border = Border(
                left=thin, right=thin, top=thin, bottom=thin)
        if not norm(vals[phone_col]):
            ws.cell(row=rn, column=phone_col + 1).fill = \
                PatternFill('solid', start_color='FFEB9C')

    widths = {'A': 20, 'B': 40, 'C': 10, 'D': 12, 'E': 36, 'F': 20,
              'G': 14, 'H': 10, 'I': 32, 'J': 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'

    # dupes sheet
    w2 = wb.create_sheet('duplicates_removed')
    w2.append(['dup_key', 'title'])
    for r in merged:
        pass
    # We track dup counts inside loop; print summary instead.
    wb.remove(w2)

    w3 = wb.create_sheet('summary')
    w3.append(['metric', 'value'])
    w3.append(['fileA', str(SRC_A)])
    w3.append(['fileA_rows', total_from['A']])
    w3.append(['fileB', str(SRC_B)])
    w3.append(['fileB_rows', total_from['B']])
    w3.append(['merged_unique_rows', len(merged)])
    w3.append(['duplicates_removed', (total_from['A'] + total_from['B']) - len(merged)])

    wb.save(DST)
    print(f'\nMERGED: {len(merged)} unique rows '
          f'({total_from["A"]} + {total_from["B"]} -> removed '
          f'{(total_from["A"] + total_from["B"]) - len(merged)} dupes) '
          f'-> {DST}', flush=True)

    # quick preview of any duplicates that overlapped between A and B
    print('\nSample unique rows:', flush=True)
    for r in merged[:8]:
        print('  ', r[1], '|', r[9])


if __name__ == '__main__':
    main()