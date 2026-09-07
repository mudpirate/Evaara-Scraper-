# -*- coding: utf-8 -*-
# Checks for duplicate shop rows: within each file and across files.
import os, re, sys
import pandas as pd
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'districts excel')


def place_key(url):
    if not isinstance(url, str) or not url.strip():
        return None
    m = re.search(r'16s%2Fg%2F([^!&]+)', url)
    if m:
        return 'g/' + m.group(1)
    m = re.search(r'19s(ChIJ[^?&]+)', url)
    if m:
        return m.group(1)
    return url.strip().lower()[:120]


files = sorted(f for f in os.listdir(FOLDER) if f.endswith('_CLEANED.xlsx'))
print('files:', len(files))
per_file = {}
for fn in files:
    df = pd.read_excel(os.path.join(FOLDER, fn), sheet_name='shops')
    keys = [place_key(u) for u in df['url']]
    df['_key'] = keys
    dup_in = int((df['_key'].notna() & df['_key'].duplicated(keep=False)).sum())
    per_file[fn] = set(k for k in keys if k)
    print(f'{fn}: rows={len(df)} unique_keys={df["_key"].nunique()} '
          f'rows_in_dupes_within_file={dup_in}')

print()
print('=== CROSS-FILE DUPLICATES (same shop in multiple files) ===')
all_keys = {}
for fn, keys in per_file.items():
    for k in keys:
        all_keys.setdefault(k, []).append(fn)
cross = {k: v for k, v in all_keys.items() if len(v) > 1}
print('total unique shops overall:', len(all_keys))
print('shops appearing in multiple files:', len(cross))
pair_counts = Counter()
for k, fns in cross.items():
    fs = tuple(sorted(fns))
    for i in range(len(fs)):
        for j in range(i + 1, len(fs)):
            pair_counts[(fs[i], fs[j])] += 1
for (a, b), c in pair_counts.most_common(40):
    print(f'{c:4d}  {a}  <->  {b}')
