# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl"]
# ///
"""Merge several scraper workbooks into one deduplicated master workbook.

Usage:
    python merge_results.py master.xlsx a.xlsx b.xlsx
    python merge_results.py master.xlsx a.xlsx b.xlsx \
        --label a.xlsx="cell phone store" \
        --label b.xlsx="smartphone accessories; mobile phone repair shop"

A duplicate is the same normalized title + street + city. Rows with no
street fall back to title + city. Same phone alone is NOT enough: a chain
can share one number across different branches, so phone is never used as
the dedupe key. The richest copy of each shop is kept and missing fields
(street, phone, website, ...) are filled from the other copies.

Output sheets:
    shops         - one row per shop, plus a foundIn column
    filtered_out  - removed rows from all inputs, also deduplicated
"""

import re
import sys
from pathlib import Path

from openpyxl import load_workbook

from scrape import write_xlsx


def read_sheet(path: Path, sheet: str = 'shops') -> tuple[list[dict], list[str]]:
    """Read a sheet into row dictionaries plus its header column list."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            return [], []
        ws = wb[sheet]
        rows = ws.iter_rows(values_only=True)
        try:
            header = next(rows)
        except StopIteration:
            return [], []
        cols = [str(c or '').strip() for c in header]
        if 'title' not in cols:
            raise ValueError(
                f"{path.name}: no 'title' column in '{sheet}' - "
                "this does not look like a scraper output file."
            )
        out: list[dict] = []
        for values in rows:
            if not values or all(v is None for v in values):
                continue
            r = {c: values[i] if i < len(values) else None
                 for i, c in enumerate(cols) if c}
            r['categoryName'] = r.get('categoryName') or r.get('categories/0') or ''
            out.append(r)
        return out, cols
    finally:
        wb.close()


def norm(value) -> str:
    """Lowercase, drop punctuation, collapse whitespace (Unicode-safe)."""
    return re.sub(r'[^\w]+', ' ', str(value or '').strip().lower()).strip()


def dedupe_key(row: dict) -> tuple:
    """Identity of a shop: GPS first, then title + street + city.

    Grid workbooks carry latitude/longitude, which is the strongest identity
    signal. Older workbooks fall back to normalized address text.
    """
    title = norm(row.get('title'))
    lat, lng = row.get('latitude'), row.get('longitude')
    if lat not in (None, '') and lng not in (None, ''):
        try:
            return ('geo', title, round(float(lat), 5), round(float(lng), 5))
        except (TypeError, ValueError):
            pass
    city = norm(row.get('city'))
    street = norm(row.get('street'))
    return (title, street, city) if street else (title, city)


def richness(row: dict) -> int:
    """How much useful contact data a row carries (used to pick the keeper)."""
    fields = ('phone', 'website', 'street', 'url', 'totalScore',
              'reviewsCount', 'categoryName')
    return sum(1 for f in fields if str(row.get(f) or '').strip())


def merge(rows: list[dict]) -> list[dict]:
    """Deduplicate rows, keeping the richest copy and merging blank fields."""
    groups: dict[tuple, list[dict]] = {}
    order: list[tuple] = []
    for row in rows:
        key = dedupe_key(row)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    merged: list[dict] = []
    for key in order:
        group = groups[key]
        group.sort(key=richness, reverse=True)  # stable: ties keep input order
        base = dict(group[0])
        seen_files: list[str] = []
        for row in group:
            label = str(row.get('foundIn') or '').strip()
            if label and label not in seen_files:
                seen_files.append(label)
            for field, value in row.items():
                if field in ('foundIn', 'title'):
                    continue
                if str(base.get(field) or '').strip() == '' and \
                        str(value or '').strip() != '':
                    base[field] = value
        base['foundIn'] = '; '.join(seen_files)
        merged.append(base)
    return merged


def parse_args(args: list[str]) -> tuple[Path, list[Path], dict[str, str]]:
    labels: dict[str, str] = {}
    while '--label' in args:
        i = args.index('--label')
        if i + 1 >= len(args):
            print('--label needs FILE=LABEL')
            sys.exit(1)
        spec = args[i + 1]
        if '=' not in spec:
            print(f"--label expects FILE=LABEL, got '{spec}'")
            sys.exit(1)
        name, label = spec.split('=', 1)
        labels[Path(name.strip()).name.lower()] = label.strip()
        del args[i:i + 2]

    if len(args) < 3:
        print(
            'Usage:\n'
            '  python merge_results.py master.xlsx a.xlsx b.xlsx [c.xlsx ...]\n'
            '  python merge_results.py master.xlsx a.xlsx b.xlsx '
            '--label a.xlsx="cell phone store"'
        )
        sys.exit(1)
    return Path(args[0]), [Path(a) for a in args[1:]], labels


def main() -> None:
    out, files, labels = parse_args(list(sys.argv[1:]))
    missing = [str(p) for p in files if not p.exists()]
    if missing:
        print('File(s) not found:\n  ' + '\n  '.join(missing))
        sys.exit(1)

    shops: list[dict] = []
    dropped: list[dict] = []
    for path in files:
        label = labels.get(path.name.lower(), path.stem)
        rows, _cols = read_sheet(path, 'shops')
        bad, _ = read_sheet(path, 'filtered_out')
        print(f"+ {path.name}: {len(rows)} shops, {len(bad)} filtered out")
        for row in rows + bad:
            row['foundIn'] = label
        shops.extend(rows)
        dropped.extend(bad)

    shops = merge(shops)
    dropped = merge(dropped)
    dropped_keys = {dedupe_key(r) for r in shops}
    dropped = [r for r in dropped if dedupe_key(r) not in dropped_keys]

    extra_cols = ['foundIn']
    for r in shops + dropped:
        for col in ('latitude', 'longitude'):
            if col not in extra_cols and str(r.get(col) or '').strip():
                extra_cols.append(col)

    saved = write_xlsx(shops, dropped, out, extra_cols=extra_cols)
    print(
        f"\n+ merged {len(files)} files -> {len(shops)} unique shops, "
        f"{len(dropped)} filtered rows -> {saved}"
    )


if __name__ == '__main__':
    main()
