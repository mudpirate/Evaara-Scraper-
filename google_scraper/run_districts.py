# /// script
# requires-python = ">=3.11"
# dependencies = ["rich", "openpyxl"]
# ///
"""
Google Maps Multi-District Batch Scraper.

Give it a districts.txt (or an inline string) with districts + cities and it
runs the scraper DISTRICT BY DISTRICT, saving ONE Excel FILE PER DISTRICT.

FORMAT (one district per line or ';'-separated):
    district1: city1, city2, city3
    district2: cityA, cityB

Every city becomes the query:  "city, district, karnataka"
Output per district:  mobile_<district>_all.xlsx

WAYS TO RUN:

  1. From a text file (recommended — one district per line, # = comment):
     python run_districts.py districts.txt

     districts.txt example:
         mandya: mandya, maddur, malavalli, nagamangala
         kolar: kolar, bangarpet, malur
         chamarajanagar: chamarajanagar, kollegal, gundlupet

  2. Inline string (use ';' between district groups):
     python run_districts.py "mandya: mandya, maddur; kolar: kolar, malur"

  3. Custom output name base (per-district files get _<district> appended):
     python run_districts.py districts.txt my_shops.xlsx

  4. Merge EVERYTHING into one single Excel instead (old behaviour):
     python run_districts.py --merged districts.txt

  5. Preview without scraping (prints the queries and exits):
     python run_districts.py --dry-run districts.txt

The state (karnataka) is appended automatically. KEYWORDS below holds
the 3 fixed queries (cell phone store / smartphone accessories / repair).
These specific phrases avoid Google's broad "mobile" matches on addresses
and reviews.
"""

import importlib.util
import os
import re
import subprocess
import sys
from pathlib import Path

# Force UTF-8 output on Windows consoles
if sys.platform == 'win32':
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    os.environ.setdefault('PYTHONUTF8', '1')
    os.environ.setdefault('PYTHONIOENCODING', 'utf-8')

# Fixed keywords — only these 3 run per district city.
# Specific phrases are intentional: plain "mobile" also matches addresses
# ("Mobile Avenue") and reviews ("they insured my mobile phone").
KEYWORDS = [
    "cell phone store",
    "smartphone accessories",
    "mobile phone repair shop",
]

# State is hardcoded so every query stays within Karnataka.
STATE = "karnataka"


def _bootstrap():
    needed = []
    for pkg in ['rich', 'openpyxl']:
        if importlib.util.find_spec(pkg) is None:
            needed.append(pkg)
    if needed:
        print(f"[first-run] installing {', '.join(needed)}...", file=sys.stderr, flush=True)
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", *needed],
            check=True,
        )


_bootstrap()

from rich.console import Console
from rich.panel import Panel
from rich.table import Table

console = Console()


def slugify(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:60]


def parse_groups(text: str) -> list[tuple[str, list[str]]]:
    """Parse 'district: city, city; district: city, city' into groups.

    Newlines act as group separators too (one district per line).
    Lines starting with '#' are ignored (comments).
    Entries without ':' are treated as raw place phrases
    (e.g. 'bengaluru, karnataka').
    """
    text = text.lstrip('\ufeff')  # strip UTF-8 BOM if present
    text = text.replace('\r\n', ';').replace('\n', ';')
    groups: list[tuple[str, list[str]]] = []
    for chunk in text.split(';'):
        chunk = chunk.strip()
        if not chunk or chunk.startswith('#'):
            continue
        if ':' in chunk:
            district, _, cities = chunk.partition(':')
            district = district.strip()
            city_list = [c.strip() for c in cities.split(',') if c.strip()]
            if district and city_list:
                groups.append((district, city_list))
                continue
        # No ':' — treat as a raw place phrase
        groups.append(('', [chunk]))
    return groups


def build_places(groups: list[tuple[str, list[str]]]) -> list[str]:
    places: list[str] = []
    for district, cities in groups:
        if not district:
            # Raw phrase — append state if missing
            for c in cities:
                places.append(c if STATE in c.lower() else f"{c}, {STATE}")
        else:
            for c in cities:
                places.append(f"{c}, {district}, {STATE}")
    return places

# --- PART2 ---

def run_one(places: list[str], out: str) -> str | None:
    """Run scrape.py once for a place-set. Returns the saved file or None."""
    env = os.environ.copy()
    # Keep the accuracy filter ON. Set KEEP_ALL=1 yourself only for raw data.
    env.setdefault('KEEP_ALL', '0')
    env.setdefault('HEADLESS', '1')  # HEADLESS=0 to watch the browser
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUNBUFFERED'] = '1'

    script = Path(__file__).parent / 'scrape.py'
    cmd = [sys.executable, str(script), ", ".join(KEYWORDS), "; ".join(places), out]
    console.print(f"[dim]{' '.join(cmd)}[/]\n")

    saved_file = None
    try:
        proc = subprocess.Popen(
            cmd, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', bufsize=1,
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            print(line, end='')
            if line.startswith('FILE_PATH:'):
                saved_file = line.split('FILE_PATH:', 1)[1].strip()
        proc.wait()
    except KeyboardInterrupt:
        console.print("\n[yellow]Interrupted — partial data saved by scrape.py[/]")
        sys.exit(130)

    if proc.returncode != 0:
        console.print(f"[red]X scrape.py exited with code {proc.returncode}[/]")
        return None

    if not saved_file or not Path(saved_file).exists():
        return out
    return saved_file


def summarize(saved_file: str | None) -> tuple[int, int]:
    """Print row/phone counts for a saved workbook. Returns (rows, phones)."""
    if not saved_file or not Path(saved_file).exists():
        return 0, 0
    try:
        from openpyxl import load_workbook
        wb = load_workbook(saved_file)
        ws = wb['shops']
        rows = max(ws.max_row - 1, 0)
        phones = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[8])
        console.print(
            f"[bold green]✓[/] [bold]{rows}[/] unique businesses "
            f"({phones} with phone) → [yellow]{Path(saved_file).resolve()}[/]"
        )
        wb.close()
        return rows, phones
    except Exception as e:
        console.print(f"[yellow]Saved to {saved_file} (couldn't verify: {e})[/]")
        return 0, 0


def main() -> None:
    args = list(sys.argv[1:])
    dry_run = '--dry-run' in args
    merged = '--merged' in args
    for flag in ('--dry-run', '--merged'):
        while flag in args:
            args.remove(flag)

    if not args:
        console.print(Panel.fit(
            "[bold cyan]Multi-District Google Maps Scraper[/]\n"
            "[dim]Format:[/] district1: city1, city2; district2: cityA, cityB\n"
            "[dim]Or a text file:[/] python run_districts.py districts.txt\n"
            "[dim]Preview first:[/] python run_districts.py --dry-run ...",
            border_style="cyan",
        ))
        sys.exit(1)

    source = args[0]
    out = args[1] if len(args) > 1 else None

    text = source
    src_path = Path(source)
    if src_path.suffix.lower() == '.txt' or src_path.exists():
        if src_path.exists():
            text = src_path.read_text(encoding='utf-8')
        else:
            console.print(f"[red]File not found: {src_path}[/]")
            sys.exit(1)

    groups = parse_groups(text)
    if not groups:
        console.print("[red]No places parsed.[/]")
        sys.exit(1)

    places = build_places(groups)

    console.print(Panel.fit(
        "[bold cyan]Google Maps Multi-District Scraper[/]\n"
        f"[dim]Keyword:[/] {'  |  '.join(KEYWORDS)}    [dim]State:[/] [yellow]{STATE}[/]\n"
        "[dim]Mode:[/] " + ("merged — ONE Excel for everything" if merged else "one Excel PER DISTRICT"),
        border_style="cyan",
    ))
    console.print("\n[bold]District groups:[/]")
    for district, cities in groups:
        if district:
            console.print(f"  [cyan]•[/] [bold]{district}[/]: {', '.join(cities)}")
        else:
            console.print(f"  [cyan]•[/] [dim](raw)[/] {', '.join(cities)}")

    n_queries = len(KEYWORDS) * len(places)
    console.print(f"\n[bold]Queries to run[/] [dim]({n_queries} total):[/]")
    for kw in KEYWORDS:
        for p in places:
            console.print(f"  [cyan]•[/] {kw} in {p}")

    if merged:
        districts = [d for d, _ in groups if d]
        if not out:
            name_parts = '_'.join(slugify(d) for d in districts[:3]) or 'mixed'
            out = f"mobile_{name_parts}_all.xlsx"
        console.print(f"[dim]\n→ Saving to:[/] [yellow]{out}[/]\n")
    else:
        n_files = len({d or 'mixed' for d, _ in groups})
        console.print(f"[dim]\n→ Saving to:[/] [yellow]{n_files} Excel files, one per district[/]")

    if dry_run:
        console.print("[yellow]--dry-run: nothing was scraped.[/]")
        sys.exit(0)

    if merged:
        saved = run_one(places, out)
        summarize(saved)
        sys.exit(0)

    # One Excel per district — run district by district.
    results: list[tuple[str, str | None, int, int]] = []
    total_rows = 0
    for i, (district, cities) in enumerate(groups, 1):
        dplaces = build_places([(district, cities)])
        dslug = slugify(district) if district else 'mixed'
        if out:
            base = Path(out).stem
            out_name = f"{base}_{dslug}.xlsx" if district else out
        else:
            out_name = f"mobile_{dslug}_all.xlsx"
        console.rule(f"[bold cyan]{i}/{len(groups)} · {district or 'mixed'} → {out_name}")
        saved = run_one(dplaces, out_name)
        rows, phones = summarize(saved)
        total_rows += rows
        results.append((district or 'mixed', saved, rows, phones))

    # Final summary table
    console.print("\n[bold]Summary:[/]")
    table = Table(show_header=True, header_style="bold")
    table.add_column("District")
    table.add_column("Businesses", justify="right")
    table.add_column("With phone", justify="right")
    table.add_column("File", style="yellow", overflow="fold")
    for district, saved, rows, phones in results:
        table.add_row(district, str(rows), str(phones), saved or '—')
    table.add_row("[bold]TOTAL[/]", f"[bold]{total_rows}[/]", "", "")
    console.print(table)


if __name__ == "__main__":
    main()
