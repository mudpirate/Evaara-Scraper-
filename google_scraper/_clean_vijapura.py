# -*- coding: utf-8 -*-
"""
Clean mobileaccessoriesvijapura.xlsx:
  - Keep ONLY: cell/mobile accessory store, mobile shop (cell phone store),
    electronic store, mobile repair shop.
  - REMOVE: green (chain-store) rows + all other irrelevant rows.
Writes mobileaccessoriesvijapura_cleaned.xlsx  (shops / removed / summary)
Original file stays untouched.
"""
import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SRC = Path('mobileaccessoriesvijapura.xlsx')
DST = Path('mobileaccessoriesvijapura_cleaned.xlsx')
GREEN = 'C6EFCE'
YELLOW = 'FFEB9C'

KEEP_CATEGORIES = {
    'cell phone store', 'cell phone accessory store',
    'mobile phone repair shop', 'electronics store',
    'electronics repair shop', 'phone repair service', 'electronics company',
}

# Title keywords used when the category is missing/borderline.
# Only words BOUNDARED to these stem prefixes count, so 'automobiles'
# won't accidentally match 'mobil'.
TITLE_PREFIXES = ('mobil', 'phone', 'fone', 'cellphone', 'cell')
ELEC_PREFIXES = ('elec',)
COMM_WORDS = ('communication', 'communications', 'telicom',
              'telecommunication', 'telecom')

# Whole-word auto/vehicle signals that make a title IRRELEVANT even if it
# also contains 'mobile' (e.g. "AUTO MOBILES & SEAT COVERS").
AUTO_WORDS = {'auto', 'automobile', 'automobiles', 'automotive', 'car', 'cars',
              'bike', 'bikes', 'scooter', 'scooters', 'vehicle', 'vehicles',
              'two', 'wheeler', 'motorcycle', 'motorcycles', 'atv', 'truck',
              'trucks', 'tractor', 'tractors', 'helmet', 'helmets', 'ev',
              'evs', 'seat', 'seats', 'dealer', 'dealers', 'showroom'}


def norm(v):
    return str(v or '').strip().lower()


def title_signal(title: str) -> bool:
    """Decide relevance from the business NAME alone.

    Keep if the name clearly says mobile/phone/cell shop or is an electronics
    store.  Reject auto/vehicle names (auto mobiles, EV dealers, helmets...).
    """
    t = norm(title)
    words = [w for w in t.split() if w]
    has_auto = any(w in AUTO_WORDS for w in words)
    has_comm = any(w in COMM_WORDS for w in words)
    has_elec = any(w.startswith(p) for w in words for p in ELEC_PREFIXES)
    has_dev = any(w.startswith(p) for w in words for p in TITLE_PREFIXES)

    # "AUTO MOBILES", "EV WORLD", "CAR SHOWROOM" -> auto => reject
    if has_auto and not (has_comm or has_elec):
        return False
    # pure "Automobiles" word (star automobiles) -> reject
    if any(w.startswith('automobil') for w in words) and not has_dev \
            and not has_elec and not has_comm:
        return False
    # electronic-store signal from name (electronics, electronic)
    if has_elec and not has_auto:
        return True
    if has_comm:
        return True
    return has_dev


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb['shops']
    rows = list(ws.iter_rows())
    header = [c.value for c in rows[0]]
    data = rows[1:]
    print(f'Source: {SRC.name}  total rows: {len(data)}')

    kept, removed_green, removed_irr = [], [], []
    for r in data:
        vals = [c.value for c in r]
        title = norm(vals[1])
        cat = norm(vals[-1])
        # green detection — chain-store row filled across non-phone cells
        is_green = any(
            c.fill.patternType == 'solid' and c.fill.fgColor
            and str(c.fill.fgColor.rgb)[-6:].upper() == GREEN
            for c in r[:10]
        )
        phone = norm(vals[9])
        if is_green:
            removed_green.append(vals)
            continue
        # category says relevant -> keep (unless title is clearly NOT mobile
        # e.g. computer/auto mislabeled; but SHRIRAM MOBILE'S under computer
        # store should stay, so title-first for ambiguous categories)
        if cat in KEEP_CATEGORIES:
            kept.append((vals, cat))
        elif title_signal(title):
            kept.append((vals, title))
        else:
            cat_show = cat if cat and cat != 'none' else '(no category)'
            removed_irr.append((vals, cat_show))

    n_green = len(removed_green)
    n_irr = len(removed_irr)
    print(f'KEEP: {len(kept)} | REMOVE green/chains: {n_green} | '
          f'REMOVE irrelevant: {n_irr}')

    # ---- write cleaned workbook ----
    out = openpyxl.Workbook()
    w = out.active
    w.title = 'shops'
    w.append(header)
    for c in w[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', start_color='DDDDDD')
        c.alignment = Alignment(horizontal='center')

    from scrape import is_chain_match  # for yellow/no-phone styling only
    for vals, src_sig in kept:
        w.append(vals)
        rn = w.max_row
        phone_val = norm(vals[9])
        if not phone_val:
            w.cell(row=rn, column=10).fill = PatternFill('solid', start_color=YELLOW)

    widths = {'A': 20, 'B': 36, 'C': 10, 'D': 12, 'E': 32, 'F': 18,
              'G': 18, 'H': 12, 'I': 32, 'J': 18}
    for col, wd in widths.items():
        w.column_dimensions[col].width = wd
    w.freeze_panes = 'A2'

    # removed sheet
    w2 = out.create_sheet('removed')
    w2.append(header + ['reason'])
    for c in w2[1]:
        c.font = Font(bold=True)
    for vals in removed_green:
        w2.append(vals + ['chain store (was green)'])
    for vals, reason in removed_irr:
        w2.append(vals + [f'irrelevant ({reason})'])
    w2.column_dimensions['A'].width = 20
    w2.column_dimensions['B'].width = 36

    # summary sheet
    w3 = out.create_sheet('summary')
    w3.append(['metric', 'value'])
    w3.append(['source', SRC.name])
    w3.append(['total_rows', len(data)])
    w3.append(['kept_rows', len(kept)])
    w3.append(['removed_chain_rows', n_green])
    w3.append(['removed_irrelevant_rows', n_irr])
    w3.append(['kept_categories',
               ', '.join(sorted(KEEP_CATEGORIES))])

    out.save(DST)
    print(f'Saved: {DST.name}  (kept {len(kept)} rows)')

    # preview what was removed (top few)
    print('\n--- Sample removed (chain/green) ---')
    for v in removed_green[:8]:
        print('   ', v[1], '|', v[-1])
    print('\n--- Sample removed (irrelevant) ---')
    for v, r in removed_irr[:15]:
        print('   ', v[1], '|', r)
    print('\n--- Sample KEPT (first 15) ---')
    for v, s in kept[:15]:
        print('   ', v[1], '|', s)


if __name__ == '__main__':
    main()