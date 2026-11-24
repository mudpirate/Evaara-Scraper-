# /// script
# requires-python = ">=3.11"
# dependencies = ["patchright", "rich", "openpyxl"]
# ///
"""Apply the scraper's accuracy filter to an EXISTING Excel file.

Use this when you already scraped before the filter was added and do not
want to re-scrape. It keeps only phone/electronics categories, removes the
junk categories/names (insurance, waterworks, restaurant, bank, auto parts,
mobile caterer, ...), and writes a new workbook with:
  shops         - cleaned rows
  filtered_out  - rows that were removed (for inspection)

Usage:
    python clean_existing.py "cell phone store" old_results.xlsx
    python clean_existing.py "cell phone store" a.xlsx b.xlsx c.xlsx

By default the output is <name>_cleaned.xlsx next to each source file.
Use -o to choose one output file (single input only):
    python clean_existing.py "cell phone store" old_results.xlsx -o clean.xlsx

The keyword is only used to choose the category rules. For mobile phone
data, "cell phone store" is the right choice.
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

from scrape import filter_by_category, write_xlsx


def read_rows(path: Path) -> tuple[list[dict], list[str]]:
    """Read the 'shops' sheet (or the active sheet) into row dictionaries.

    Returns (rows, header_columns) so extra columns such as citySearched can
    be preserved in the cleaned workbook.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb['shops'] if 'shops' in wb.sheetnames else wb.active
        rows = ws.iter_rows(values_only=True)
        try:
            header = next(rows)
        except StopIteration:
            return [], []
        cols = [str(c or '').strip() for c in header]
        if 'title' not in cols:
            raise ValueError(
                f"{path.name}: no 'title' column found. "
                "This does not look like a scraper output file."
            )
        if 'categoryName' not in cols and 'categories/0' not in cols:
            raise ValueError(
                f"{path.name}: no 'categoryName' or 'categories/0' column found. "
                "Re-scrape with the current scraper, or add a category column."
            )
        out: list[dict] = []
        for values in rows:
            if not values or all(v is None for v in values):
                continue
            r = {c: values[i] if i < len(values) else None for i, c in enumerate(cols) if c}
            r['categoryName'] = r.get('categoryName') or r.get('categories/0') or ''
            out.append(r)
        return out, cols
    finally:
        wb.close()


def clean_file(src: Path, keyword: str, dst: Path | None = None) -> Path:
    rows, cols = read_rows(src)
    kept, dropped, intent = filter_by_category(rows, keyword)
    if intent is None:
        print(
            f"! {src.name}: keyword '{keyword}' matched no category rule; "
            "no rows were filtered."
        )
    out = dst or src.with_name(f"{src.stem}_cleaned.xlsx")
    saved = write_xlsx(kept, dropped, out, extra_cols=cols)
    print(
        f"+ {src.name}: {len(rows)} rows -> {len(kept)} kept, "
        f"{len(dropped)} filtered -> {saved}"
    )
    return saved


def main() -> None:
    args = list(sys.argv[1:])
    if len(args) < 2:
        print(
            "Usage:\n"
            "  python clean_existing.py \"cell phone store\" file1.xlsx [file2.xlsx ...]\n"
            "  python clean_existing.py \"cell phone store\" file.xlsx -o cleaned.xlsx"
        )
        sys.exit(1)

    keyword = args.pop(0)
    out: Path | None = None
    if '-o' in args:
        i = args.index('-o')
        if i + 1 >= len(args):
            print("-o needs an output filename")
            sys.exit(1)
        out = Path(args[i + 1])
        del args[i:i + 2]
        if len(args) != 1:
            print("-o can only be used with one input file")
            sys.exit(1)

    files = [Path(a) for a in args]
    missing = [str(p) for p in files if not p.exists()]
    if missing:
        print("File(s) not found:\n  " + "\n  ".join(missing))
        sys.exit(1)

    for i, src in enumerate(files):
        try:
            clean_file(src, keyword, out if i == 0 else None)
        except Exception as exc:
            print(f"! {src.name}: {type(exc).__name__}: {exc}")
            sys.exit(1)


if __name__ == '__main__':
    main()
