# -*- coding: utf-8 -*-
# Cleans a district excel: keeps only mobile/cell phone related shops.
# Usage: python _clean_districts.py file1.xlsx file2.xlsx ...
# Writes <name>_CLEANED.xlsx with sheets: shops / removed / summary
import sys, os, unicodedata
import pandas as pd
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'districts excel')

PHONE_BRANDS = [
    'realme', 'vivo', 'oppo', 'oneplus', 'xiaomi', 'redmi', 'samsung',
    'nokia', 'motorola', 'infinix', 'tecno', 'lava', 'iqoo',
    'nothing phone', 'pixel', 'iphone',
]

RULE = ('KEEP if category has cell phone/mobile phone/phone repair/telecommunication '
        'OR title has mobil(+kannada)/phone/cell/telecom/telcom/communication/jio/airtel/'
        'mi store/recharge/phone-brands(realme,vivo,oppo,etc); '
        'auto+mobile without phone/cell excluded')


def classify(row, cat_cols):
    for c in cat_cols:
        v = str(row[c]).lower() if pd.notna(row[c]) else ''
        if ('cell phone' in v or 'mobile phone' in v or 'phone repair' in v
                or 'telecommunication' in v):
            return True, 'category:' + str(row[c])
    t = unicodedata.normalize('NFKC', str(row['title'])).lower() if pd.notna(row['title']) else ''
    has_mobil = ('mobil' in t or 'moble' in t or 'moblie' in t or 'mobail' in t
                 or 'ಮೊಬೈಲ' in t or 'ಮೋಬೈಲ' in t or 'mobihub' in t)
    has_phone = 'phone' in t or 'phono' in t or 'fone' in t
    has_cell = 'cell' in t
    has_tele = ('telecom' in t or 'telcom' in t or 'telicom' in t or 'communicat' in t
                or 'jio' in t or 'airtel' in t or 'vodafone' in t
                or ' bsnl' in t or 'mi store' in t or 'mi exclusive' in t
                or 'mi show' in t or 'mi service' in t or 'mi authorized' in t
                or 'sangeetha' in t or 'recharge' in t)
    has_brand = any(b in t for b in PHONE_BRANDS)
    has_auto = 'auto' in t
    if has_mobil and has_auto and not (has_phone or has_cell or has_tele):
        return False, 'auto-vehicle exclusion'
    if has_mobil or has_phone or has_cell or has_tele or has_brand:
        return True, 'title-keyword'
    return False, 'no-mobile-signal'


def clean_one(src_fn):
    src = os.path.join(FOLDER, src_fn)
    dst = os.path.join(FOLDER, src_fn.replace('.xlsx', '_CLEANED.xlsx'))
    df = pd.read_excel(src, sheet_name='shops')
    if len(df) == 0:
        print(src_fn, '=> EMPTY (0 rows), nothing to clean')
        return
    cat_cols = [c for c in df.columns if c.startswith('categories/')]
    flags = df.apply(lambda r: classify(r, cat_cols), axis=1)
    df['_keep'] = [x[0] for x in flags]
    df['_reason'] = [x[1] for x in flags]
    keep = df[df['_keep']].copy()
    rem = df[~df['_keep']].copy()
    summary = pd.DataFrame([
        ['source_file', src_fn],
        ['total_rows', len(df)],
        ['kept_rows', len(keep)],
        ['removed_rows', len(rem)],
        ['rule', RULE],
    ])
    for k, v in keep['_reason'].value_counts().items():
        summary.loc[len(summary)] = [k, v]
    with pd.ExcelWriter(dst, engine='openpyxl') as w:
        keep.drop(columns=['_keep', '_reason']).to_excel(w, sheet_name='shops', index=False)
        rem.drop(columns=['_keep']).to_excel(w, sheet_name='removed', index=False)
        summary.to_excel(w, sheet_name='summary', index=False, header=False)
    print(src_fn, '=> TOTAL:', len(df), '| KEEP:', len(keep), '| REMOVE:', len(rem))
    print(keep['_reason'].value_counts().to_string())
    print('--- REMOVED ---')
    print(rem[['title', 'categories/0']].to_string())


if __name__ == '__main__':
    for fn in sys.argv[1:]:
        clean_one(fn)
