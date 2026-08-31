# -*- coding: utf-8 -*-
"""Quick status viewer for run_districts.py progress. Usage: python _status.py"""
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

TXT = 'districts.txt'
total_places = 0
done_places = 0
rows = []


def slugify(t):
    return re.sub(r'[-\s]+', '_', re.sub(r'[^\w\s-]', '', t.lower())).strip('_')[:60]


text = open(TXT, encoding='utf-8').read()
groups = [g for g in re.split(r'[;\n]', text) if g.strip() and not g.strip().startswith('#')]

print()
print(f"{'#':>2}  {'DISTRICT':<20} {'PLACES':>6}  {'FILE':<38} {'STATUS'}")
print('-' * 100)
for i, g in enumerate(groups, 1):
    district, _, cities = g.partition(':')
    district = district.strip()
    n = len([c for c in cities.split(',') if c.strip()])
    total_places += n
    fname = f"mobile_{slugify(district)}_all.xlsx"
    # Files may land in this folder OR in Documents (CWD of the running process)
    alt = os.path.expanduser(f"~\\Documents\\{fname}")
    path = None
    if os.path.exists(fname):
        path = fname
    elif os.path.exists(alt):
        path = alt
    if path:
        import datetime
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(path))
        run_start = datetime.datetime.fromtimestamp(os.path.getmtime('run_districts.pid')) \
            if os.path.exists('run_districts.pid') else datetime.datetime.fromtimestamp(os.path.getmtime(TXT))
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            ws = wb['shops']
            r = max(ws.max_row - 1, 0)
            wb.close()
        except Exception:
            r = 0
        if mtime > run_start:
            done_places += n
            loc = 'this folder' if path == fname else 'Documents'
            status = f"DONE - {r} rows ({mtime:%d-%b %H:%M}, {loc})"
        else:
            status = f"OLD file from earlier run ({mtime:%d-%b}) - will be overwritten"
        rows.append((i, district, n, fname, status))
    else:
        rows.append((i, district, n, fname, 'pending'))

for i, district, n, fname, status in rows:
    print(f"{i:>2}  {district:<20} {n:>6}  {fname:<38} {status}")

print('-' * 100)
pct = 100 * done_places / total_places if total_places else 0
print(f"Progress: {done_places}/{total_places} places ({pct:.0f}%)")
print(f"District files: {sum(1 for r in rows if 'DONE' in r[4])}/{len(rows)}")
print()
