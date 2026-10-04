# -*- coding: utf-8 -*-
"""Merge each district _*_parts folder into one universal FINAL excel, deduped.

- Stacks all *_parts/*.xlsx shops sheets per district
- Dedupes by robust Google place key from URL (fallback: title+phone+street+city)
- Applies mobile-clean filter (same as _clean_districts.py)
- Restores chain stores (same as _restore_chains.py)
- Writes districts excel/universal_<name>_district_FINAL.xlsx + combined master
"""
import os, re, sys, unicodedata
from pathlib import Path
from collections import Counter

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

HERE = Path(__file__).parent
OUTDIR = HERE / "districts excel"
OUTDIR.mkdir(exist_ok=True)

DISTRICTS = {
    "_chikmangaluru_parts": "universal_chikmangaluru_district_FINAL.xlsx",
    "_chitradurga_parts": "universal_chitradurga_district_FINAL.xlsx",
    "_dakshinkannada_parts": "universal_dakshin_kannada_district_FINAL.xlsx",
    "_davanagere_parts": "universal_davanagere_district_FINAL.xlsx",
    "_hassan_parts": "universal_hassan_district_FINAL.xlsx",
    "_haveri_parts": "universal_haveri_district_FINAL.xlsx",
    "_shivmoga_parts": "universal_shivmoga_district_FINAL.xlsx",
    "_udupi_parts": "universal_udupi_district_FINAL.xlsx",
    "_uttarkannada_parts": "universal_uttar_kannada_district_FINAL.xlsx",
    "_kodagu_parts": "universal_kodagu_district_FINAL.xlsx",
}

PHONE_BRANDS = ['realme','vivo','oppo','oneplus','xiaomi','redmi','samsung',
    'nokia','motorola','infinix','tecno','lava','iqoo','nothing phone','pixel','iphone']

CHAIN_NAMES = ['poorvika','sangeetha','pai mobiles','ss mobiles','mi store',
    'i central','i centre','icentre','icentral','invent tribe','icare','imagine',
    'iplanet','maple','unicorn','invent','ivenus','idestiny','samsung cafe','oppo',
    'vivo','phonorange','phono range','my jio store','ntel mobiles','sara mobiles',
    'pai international','samsung service center','reliance digital','xiaomi']
CHAIN_FILL = 'C6EFCE'
NO_PHONE_FILL = 'FFEB9C'

def place_key(url):
    if url is None or (isinstance(url, float) and pd.isna(url)):
        return None
    u = str(url).strip()
    if not u or u.lower() == 'nan':
        return None
    m = re.search(r'16s%2Fg%2F([^!&]+)', u)
    if m:
        return 'g/' + m.group(1)
    m = re.search(r'19s(ChIJ[^?&]+)', u)
    if m:
        return m.group(1)
    return u.strip().lower()[:200]

def fallback_key(row):
    t = str(row.get('title','') or '').strip().lower()
    p = re.sub(r'[^0-9]','', str(row.get('phone','') or ''))[-10:]
    s = str(row.get('street','') or '').strip().lower()
    c = str(row.get('city','') or '').strip().lower()
    return f"fb|{t}|{p}|{s}|{c}"

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

def is_chain_match(title: str) -> bool:
    if not title:
        return False
    t = str(title).lower()
    return any(re.search(rf'(?<![a-z]){re.escape(k)}(?![a-z])', t) for k in CHAIN_NAMES)

def merge_one(parts_folder: Path):
    files = sorted(parts_folder.glob('*.xlsx'))
    dfs = []
    per_file_counts = {}
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name='shops')
        except Exception as e:
            print(f"  !! skip {f.name}: {e}")
            continue
        per_file_counts[f.name] = len(df)
        df['_source_file'] = f.name
        dfs.append(df)
    if not dfs:
        return None, {'files': 0, 'stacked': 0, 'per_file': {}}
    stacked = pd.concat(dfs, ignore_index=True)
    total_stacked = len(stacked)
    # dedupe by place_key, fallback if missing
    keys = []
    for _, r in stacked.iterrows():
        k = place_key(r.get('url'))
        if k is None:
            k = fallback_key(r)
        keys.append(k)
    stacked['_dedupe_key'] = keys
    before = len(stacked)
    deduped = stacked.drop_duplicates(subset=['_dedupe_key'], keep='first').copy()
    dupes_removed = before - len(deduped)
    # drop helper cols for output (keep _source_file? drop to match universal schema)
    deduped = deduped.drop(columns=['_dedupe_key'])
    info = {'files': len(files), 'stacked': total_stacked,
            'unique': len(deduped), 'dupes_removed': dupes_removed,
            'per_file': per_file_counts}
    return deduped, info

def clean_and_restore(deduped: pd.DataFrame):
    cat_cols = [c for c in deduped.columns if c.startswith('categories/')]
    flags = deduped.apply(lambda r: classify(r, cat_cols), axis=1)
    deduped['_keep'] = [x[0] for x in flags]
    deduped['_reason'] = [x[1] for x in flags]
    keep = deduped[deduped['_keep']].copy()
    rem = deduped[~deduped['_keep']].copy()
    # restore chains from removed
    chain_mask = rem['title'].apply(is_chain_match) if len(rem) else pd.Series([], dtype=bool)
    chains = rem[chain_mask].copy() if len(rem) else rem.copy()
    rest = rem[~chain_mask].copy() if len(rem) else rem.copy()
    if len(chains):
        chains['_reason'] = chains['_reason'] + ' + chain-restored'
    shops = pd.concat([keep, chains], ignore_index=True) if len(chains) else keep.copy()
    # drop helper _keep, keep _reason + _source_file for traceability? match existing schema:
    # existing universals have _reason but no _source_file. Keep _source_file as extra col for usefulness.
    return shops, rest, chains

def write_final(shops, rest, info, parts_folder_name, out_path: Path, per_file):
    # order columns: canonical 22 + _reason + _source_file
    base_cols = ['title','totalScore','reviewsCount','street','city','state','countryCode',
                 'website','phone'] + [f'categories/{i}' for i in range(11)] + ['url','categoryName']
    # ensure all base cols exist
    for c in base_cols:
        if c not in shops.columns:
            shops[c] = None
    # final shops col order
    shop_cols = base_cols + ['_reason','_source_file'] if '_source_file' in shops.columns else base_cols + ['_reason']
    shops_out = shops[shop_cols] if len(shops) else pd.DataFrame(columns=shop_cols)
    if '_source_file' in rest.columns:
        rest_cols = base_cols + ['_reason','_source_file']
        rest_cols = [c for c in rest_cols if c in rest.columns]
        rest_out = rest[rest_cols]
    else:
        rest_out = rest
    summary_rows = [
        ['district_parts_folder', parts_folder_name],
        ['part_files_merged', info['files']],
        ['total_rows_stacked', info['stacked']],
        ['unique_after_dedupe', info['unique']],
        ['duplicates_removed', info['dupes_removed']],
        ['kept_shops_final', len(shops_out)],
        ['restored_chains', len(rest[rest['_reason'].astype(str).str.contains('chain-restored', na=False)]) if len(rest) else 0],
        ['still_removed_irrelevant', len(rest_out)],
        ['dedupe_key', 'google-place-key from url (fallback title+phone+street+city)'],
        ['chain_highlight', f'green {CHAIN_FILL} across row'],
        ['missing_phone_highlight', f'yellow {NO_PHONE_FILL} on phone cell'],
    ]
    for fn, cnt in sorted(per_file.items()):
        summary_rows.append([f'part:{fn}', cnt])
    summary = pd.DataFrame(summary_rows)
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        shops_out.to_excel(w, sheet_name='shops', index=False)
        rest_out.to_excel(w, sheet_name='removed', index=False)
        summary.to_excel(w, sheet_name='summary', index=False, header=False)
    # highlights
    wb = load_workbook(out_path)
    ws = wb['shops']
    headers = [c.value for c in ws[1]]
    try:
        phone_col = headers.index('phone') + 1
        title_col = headers.index('title') + 1
    except ValueError:
        wb.save(out_path)
        return
    green = PatternFill('solid', start_color=CHAIN_FILL)
    yellow = PatternFill('solid', start_color=NO_PHONE_FILL)
    for row in range(2, ws.max_row + 1):
        title = ws.cell(row=row, column=title_col).value
        chain = is_chain_match(title)
        phone = str(ws.cell(row=row, column=phone_col).value or '').strip()
        if chain:
            for col in range(1, len(headers) + 1):
                if col != phone_col:
                    ws.cell(row=row, column=col).fill = green
            ws.cell(row=row, column=phone_col).fill = green if phone and phone.lower()!='nan' else yellow
        elif not phone or phone.lower()=='nan':
            ws.cell(row=row, column=phone_col).fill = yellow
    wb.save(out_path)

def main():
    results = {}
    all_shops = []
    for folder_name, out_name in DISTRICTS.items():
        pf = HERE / folder_name
        print(f"\n=== {folder_name} -> {out_name} ===")
        if not pf.exists() or not list(pf.glob('*.xlsx')):
            print("  SKIP: empty / no part files (kodagu not scraped yet)")
            results[folder_name] = None
            continue
        deduped, info = merge_one(pf)
        print(f"  files={info['files']} stacked={info['stacked']} unique={info['unique']} dupes_removed={info['dupes_removed']}")
        shops, rest, chains = clean_and_restore(deduped)
        print(f"  shops_final={len(shops)} (chains_restored={len(chains)}) removed={len(rest)}")
        out_path = OUTDIR / out_name
        write_final(shops, rest, info, folder_name, out_path, info['per_file'])
        print(f"  WROTE -> {out_path}")
        # verify no dupes in output
        keys = [place_key(u) for u in shops['url']] if len(shops) else []
        s = pd.Series(keys)
        dup_in_output = int(s.duplicated(keep=False).sum()) if len(s) else 0
        print(f"  verify: unique_keys={s.nunique() if len(s) else 0} dup_rows_in_output={dup_in_output}")
        results[folder_name] = {'info': info, 'shops': len(shops), 'path': str(out_path)}
        tmp = shops.copy()
        tmp['_district'] = folder_name
        all_shops.append(tmp)
    # combined master across districts (global dedupe)
    if all_shops:
        master = pd.concat(all_shops, ignore_index=True)
        mkeys = []
        for _, r in master.iterrows():
            k = place_key(r.get('url'))
            if k is None:
                k = fallback_key(r)
            mkeys.append(k)
        master['_mkey'] = mkeys
        before = len(master)
        master_nodup = master.drop_duplicates(subset=['_mkey'], keep='first').drop(columns=['_mkey'])
        cross_dupes = before - len(master_nodup)
        print(f"\n=== COMBINED MASTER: stacked_districts={before} cross_district_dupes={cross_dupes} final={len(master_nodup)} ===")
        mout = OUTDIR / "ALL_DISTRICTS_UNIVERSAL_MERGED_FINAL.xlsx"
        base_cols = ['title','totalScore','reviewsCount','street','city','state','countryCode',
                     'website','phone'] + [f'categories/{i}' for i in range(11)] + ['url','categoryName']
        cols = base_cols + [c for c in ['_reason','_source_file','_district'] if c in master_nodup.columns]
        with pd.ExcelWriter(mout, engine='openpyxl') as w:
            master_nodup[cols].to_excel(w, sheet_name='shops', index=False)
            pd.DataFrame([['districts_merged', len(all_shops)],['stacked', before],
                          ['cross_district_dupes_removed', cross_dupes],['final_unique', len(master_nodup)]]).to_excel(w, sheet_name='summary', index=False, header=False)
        print(f"WROTE MASTER -> {mout}")
    print("\nDONE")

if __name__ == '__main__':
    main()
