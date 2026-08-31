# /// script
# requires-python = ">=3.11"
# dependencies = ["rich", "openpyxl"]
# ///
"""
Google Maps Batch Scraper — fixed keywords, any city/district.

The keywords below stay the same every run. You only change the
city/district/places. Every keyword runs for every place, all results
are merged into ONE Excel file with duplicates removed across all queries.

Run it:
  python run_batch.py              (interactive — enter city/district)
  python run_batch.py "bidar, humnabad, basavakalyan"   (one-shot)

Config: edit KEYWORDS below if you want different fixed keywords.
"""

import importlib.util
import os
import re
import subprocess
import sys
from pathlib import Path

# Force UTF-8 output on Windows consoles (cp1252 can't print unicode arrows)
if sys.platform == 'win32':
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    os.environ.setdefault('PYTHONUTF8', '1')
    os.environ.setdefault('PYTHONIOENCODING', 'utf-8')

# Fixed keywords — same every run. Change these if you want other categories.
KEYWORDS = ["mobile phones"]

# State is hardcoded so every query stays within Karnataka.
# Change this if you ever need another state.
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
from rich.prompt import Prompt

console = Console()


def slugify(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:40]


def main() -> None:
    console.print(Panel.fit(
        "[bold cyan]Google Maps Batch Scraper[/]\n"
        "[dim]Fixed keywords:  " + "  |  ".join(KEYWORDS) + "\n"
        f"State hardcoded: [yellow]{STATE}[/] — query becomes "
        "'city, district, " + STATE + "'",
        border_style="cyan",
    ))

    if len(sys.argv) >= 2:
        places_input = sys.argv[1]
    else:
        places_input = Prompt.ask(
            "\n[bold]City, district[/] [dim](e.g. belura, bidar district)\n"
            "Multiple cities: separate with ;  (e.g. belura, bidar district; humnabad, bidar district)[/]"
        ).strip()

    # Each entry is "city, district" (no state needed — it's hardcoded).
    entries = [e.strip() for e in places_input.split(';') if e.strip()]
    if not entries:
        console.print("[red]No places entered.[/]")
        sys.exit(1)

    # Build full place phrases: "city, district, karnataka" (+ state)
    places = []
    for e in entries:
        if STATE in e.lower():
            places.append(e)
        else:
            places.append(f"{e}, {STATE}")

    first = places[0]
    kw_slug = slugify(KEYWORDS[0].split()[0])
    city_slug = slugify(entries[0].split(',')[0].strip())
    out = f"{kw_slug}_{city_slug}_all.xlsx"

    console.print(f"\n[bold]Queries to run[/] [dim]({len(KEYWORDS) * len(places)} total):[/]")
    for kw in KEYWORDS:
        for p in places:
            console.print(f"  [cyan]\u2022[/] {kw} in {p}")
    console.print(f"[dim]\n\u2192 Saving to:[/] [yellow]{out}[/]\n")

    env = os.environ.copy()
    env['KEEP_ALL'] = '1'          # no category filtering — keep maximum data
    env.setdefault('HEADLESS', '1')  # HEADLESS=0 to watch the browser
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUNBUFFERED'] = '1'

    script = Path(__file__).parent / 'scrape.py'
    cmd = [sys.executable, str(script), ", ".join(KEYWORDS), "; ".join(places), out]
    console.print(f"[dim]{' '.join(cmd)}[/]\n")

    try:
        proc = subprocess.Popen(
            cmd, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', bufsize=1,
        )
        saved_file = None
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
        sys.exit(proc.returncode)

    if not saved_file or not Path(saved_file).exists():
        # Fallback: newest matching xlsx in this folder
        matches = sorted(Path(__file__).parent.glob('*.xlsx'), key=lambda p: p.stat().st_mtime)
        saved_file = str(matches[-1]) if matches else out

    try:
        from openpyxl import load_workbook
        wb = load_workbook(saved_file)
        ws = wb['shops']
        phones = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r[8])
        console.print(
            f"\n[bold green]Done:[/] {ws.max_row - 1} unique businesses "
            f"({phones} with phone) \u2192 [yellow]{Path(saved_file).resolve()}[/]"
        )
        wb.close()
    except Exception as e:
        console.print(f"[yellow]Saved to {saved_file} (couldn't verify: {e})[/]")


if __name__ == "__main__":
    main()