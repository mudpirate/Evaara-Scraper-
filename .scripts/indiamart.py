# /// script
# requires-python = ">=3.11"
# dependencies = ["patchright", "rich", "openpyxl"]
# ///
"""
IndiaMART scraper — extracts supplier/business listings via __NEXT_DATA__ JSON.

Zero setup: needs only Python 3.11+. Missing deps and Chromium are
auto-installed on first run. uv is optional.

Usage:
  python indiamart.py "mobile accessories" "bangalore"
  python indiamart.py "mobile accessories, phone chargers, usb cables" "bangalore"
  python indiamart.py "mobile accessories, phone chargers" "delhi, mumbai, bangalore"
  python indiamart.py "usb cables" "hubli" output.xlsx

Environment variables:
  HEADLESS=1       run without visible browser (default: visible)
  DELAY_MIN_MS     min jitter between page loads (default 600)
  DELAY_MAX_MS     max jitter between page loads (default 2000)
  MAX_PAGES        max pages per query (default 20)
  PROFILE_DIR      persistent Chrome profile dir

Note on phone numbers:
  IndiaMART gates real phone numbers behind login/OTP. The `pns` field
  (Preferred Number Service) is a virtual relay number — calling it
  connects you to the seller through IndiaMART's system. Good enough
  for first contact in B2B outreach.
"""

# --- stdlib ---------------------------------------------------------------
import asyncio
import importlib.util
import os
import random
import re
import subprocess
import sys
from pathlib import Path
from urllib.parse import quote_plus


# --- bootstrap ------------------------------------------------------------
def _bootstrap_deps() -> None:
    required = ['patchright', 'rich', 'openpyxl']
    missing = [p for p in required if importlib.util.find_spec(p) is None]
    if not missing:
        return
    print(f"[first-run] installing: {', '.join(missing)}",
          file=sys.stderr, flush=True)
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", *missing],
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"pip install failed ({e}).\n"
              f"  Install manually:  pip install {' '.join(missing)}\n"
              f"  Or use uv:         uv run indiamart.py ...",
              file=sys.stderr, flush=True)
        sys.exit(1)
    print("[first-run] ✓ deps installed", file=sys.stderr, flush=True)


_bootstrap_deps()


# --- third-party ----------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from patchright.async_api import Page, async_playwright
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt

console = Console()

# --- config ---------------------------------------------------------------
HEADLESS = os.getenv("HEADLESS", "0") == "1"
DELAY_MIN_MS = int(os.getenv("DELAY_MIN_MS", "600"))
DELAY_MAX_MS = int(os.getenv("DELAY_MAX_MS", "2000"))
MAX_PAGES = int(os.getenv("MAX_PAGES", "20"))
PROFILE_DIR = Path(os.getenv(
    "PROFILE_DIR",
    str(Path.home() / ".cache" / "indiamart-scraper" / "profile"),
))
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
)

STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
Object.defineProperty(navigator, 'plugins', { get: () => [
    { name: 'Chrome PDF Plugin' },
    { name: 'Chrome PDF Viewer' },
    { name: 'Native Client' },
]});
window.chrome = { runtime: {}, loadTimes: () => {}, csi: () => {} };
const _query = window.navigator.permissions && window.navigator.permissions.query;
if (_query) {
    window.navigator.permissions.query = (p) =>
        p.name === 'notifications'
            ? Promise.resolve({ state: 'default' })
            : _query(p);
}
"""


# --- helpers --------------------------------------------------------------
def ensure_chromium() -> None:
    if os.name == 'nt':
        local = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData' / 'Local')))
    elif sys.platform == 'darwin':
        local = Path.home() / 'Library' / 'Caches'
    else:
        local = Path(os.environ.get('XDG_CACHE_HOME', str(Path.home() / '.cache')))
    for dirname in ('ms-playwright', 'ms-patchright'):
        base = local / dirname
        if base.exists() and any(d.name.startswith('chromium') for d in base.iterdir()):
            return
    console.print(Panel.fit(
        "[bold yellow]First-time setup[/]\n"
        "[dim]Downloading Chromium (~150 MB, one-time)...[/]",
        border_style="yellow",
    ))
    for mod in ('patchright', 'playwright'):
        try:
            subprocess.run([sys.executable, "-m", mod, "install", "chromium"], check=True)
            console.print("[bold green]✓ Chromium ready[/]\n")
            return
        except (subprocess.CalledProcessError, FileNotFoundError):
            continue
    console.print("[red]✗ Chromium install failed.[/]")
    sys.exit(1)


def build_url(keyword: str, city: str, page_num: int = 1) -> str:
    """Build IndiaMART search URL."""
    url = f"https://dir.indiamart.com/search.mp?ss={quote_plus(keyword)}"
    if city:
        url += f"&city={quote_plus(city)}"
    if page_num > 1:
        url += f"&page={page_num}"
    return url


def slugify(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:60]


# --- anti-bot -------------------------------------------------------------
async def jitter(page: Page) -> None:
    await page.wait_for_timeout(random.randint(DELAY_MIN_MS, DELAY_MAX_MS))


async def is_blocked(page: Page) -> bool:
    url = page.url or ''
    if any(x in url.lower() for x in ['captcha', 'blocked', '/sorry/', 'challenge']):
        return True
    for sel in ['iframe[src*="captcha"]', '#challenge-form', '#px-captcha',
                'form[action*="challenge"]']:
        try:
            if await page.locator(sel).count() > 0:
                return True
        except Exception:
            pass
    try:
        title = await page.title()
        if any(x in title.lower() for x in ('access denied', 'blocked', 'robot')):
            return True
    except Exception:
        pass
    return False


async def handle_block(page: Page) -> None:
    if not await is_blocked(page):
        return

    console.print(Panel.fit(
        "[bold red]⚠ Blocked / CAPTCHA detected[/]\n\n"
        f"[yellow]URL:[/]  {page.url}\n"
        f"[yellow]Mode:[/] {'HEADLESS' if HEADLESS else 'HEADED'}",
        border_style="red",
    ))
    if HEADLESS:
        try:
            await page.screenshot(path="im_blocked.png", full_page=True)
            console.print("[yellow]Screenshot:[/] im_blocked.png")
        except Exception:
            pass
        console.print(
            "[red]Re-run without HEADLESS=1 to solve the challenge. "
            "Profile is persistent — afterwards headless will work.[/]"
        )
        sys.exit(2)
    else:
        is_tty = sys.stdin.isatty() if hasattr(sys.stdin, 'isatty') else False
        if is_tty:
            console.print(
                "[bold]Solve in the visible browser "
                "[dim](ENTER optional — auto-detects when cleared)[/][/]"
            )

            async def _wait_for_enter():
                loop = asyncio.get_running_loop()
                await loop.run_in_executor(None, sys.stdin.readline)

            enter_task = asyncio.create_task(_wait_for_enter())
        else:
            console.print(
                "[bold]Solve in the visible browser. "
                "[dim](auto-resuming when cleared, up to 10 min)[/][/]"
            )
            enter_task = None

        max_wait_s, check_every_s = 600, 5
        elapsed = 0
        while elapsed < max_wait_s:
            if not await is_blocked(page):
                if enter_task and not enter_task.done():
                    enter_task.cancel()
                console.print(f"[green]✓ Resuming[/] [dim](after {elapsed}s)[/]\n")
                return
            await asyncio.sleep(check_every_s)
            elapsed += check_every_s
            if elapsed and elapsed % 30 == 0:
                console.print(
                    f"[dim]still waiting... ({elapsed}s / {max_wait_s}s)[/]"
                )
        if enter_task and not enter_task.done():
            enter_task.cancel()
        console.print("[red]Still blocked after 10 min. Aborting.[/]")
        sys.exit(2)


# --- data extraction via __NEXT_DATA__ ------------------------------------
def parse_next_data(raw: dict) -> tuple[list[dict], bool]:
    """Extract listings from __NEXT_DATA__ JSON.

    IndiaMART stores results at:
      props.pageProps.searchResponse.results[].fields

    Returns (listings, has_next_page).
    """
    try:
        page_props = raw['props']['pageProps']
        search_resp = page_props.get('searchResponse') or {}
        results = search_resp.get('results', []) if isinstance(search_resp, dict) else []
        has_next = bool(search_resp.get('nextPage', False)) if isinstance(search_resp, dict) else False
    except (KeyError, TypeError, AttributeError):
        return [], False

    listings = []
    for item in results:
        fields = item.get('fields') if isinstance(item, dict) else None
        if not fields:
            continue
        listings.append(fields)

    return listings, has_next


def extract_fields(f: dict) -> dict:
    """Map IndiaMART's raw fields to clean output dict."""
    # Rating
    rating = f.get('supplier_rating') or ''
    if rating:
        try:
            rating = float(rating)
        except (ValueError, TypeError):
            rating = ''

    rating_count = f.get('rating_count') or ''
    if rating_count:
        try:
            rating_count = int(str(rating_count).replace(',', ''))
        except (ValueError, TypeError):
            rating_count = ''

    # Price
    price = str(f.get('price_f', '') or f.get('itemprice', '') or '').strip()
    moq = str(f.get('moq', '') or '').strip()
    moq_type = str(f.get('moq_type', '') or '').strip()
    if moq and moq_type:
        moq = f"{moq} {moq_type}"

    # GST
    gst = str(f.get('gstNumber', '') or '').strip()
    gst_verified = bool(f.get('gstVerifiedFlag'))

    # PNS (virtual relay number)
    pns = str(f.get('pns', '') or '').strip()

    # Company URL
    catalog_url = str(f.get('catalog_url', '') or '').strip()
    if catalog_url and not catalog_url.startswith('http'):
        catalog_url = f"https://{catalog_url}"

    # Product detail URL
    detail_url = str(f.get('desktop_title_url', '') or f.get('title_url', '') or '').strip()
    if detail_url and not detail_url.startswith('http'):
        detail_url = f"https://dir.indiamart.com{detail_url}"

    # Specs
    specs_raw = str(f.get('isq', '') or '')
    specs = specs_raw.replace('==', ': ').replace('||', ', ') if specs_raw else ''

    return {
        'companyName': str(f.get('companyname', '') or '').strip(),
        'product': str(f.get('title', '') or '').strip(),
        'price': price,
        'moq': moq,
        'city': str(f.get('city', '') or '').strip(),
        'state': str(f.get('state', '') or '').strip(),
        'district': str(f.get('district', '') or '').strip(),
        'pincode': str(f.get('zipcode', '') or '').strip(),
        'pns': pns,
        'gst': gst,
        'gstVerified': 'Yes' if gst_verified else '',
        'rating': rating,
        'ratingCount': rating_count,
        'memberSince': str(f.get('memberSinceDisplay', '') or '').strip(),
        'annualTurnover': str(f.get('annualTurnover', '') or '').strip(),
        'category': str(f.get('mcatname', '') or '').strip(),
        'bizType': str(f.get('biztype', '') or '').strip(),
        'specs': specs,
        'companyUrl': catalog_url,
        'productUrl': detail_url,
        '_glsrid': str(f.get('glusrid', '') or '').strip(),
        '_displayid': str(f.get('displayid', '') or '').strip(),
    }


# --- scraping -------------------------------------------------------------
async def scrape_query(page: Page, keyword: str, city: str) -> list[dict]:
    """Scrape all pages for one keyword+city combo."""
    import time as _time
    label = f"{keyword} in {city}" if city else keyword
    console.print(f"[bold cyan]→[/] [yellow]{label}[/]")
    _start = _time.monotonic()

    all_listings: list[dict] = []
    seen_ids: set[str] = set()
    page_num = 1
    dry_streak = 0

    while page_num <= MAX_PAGES:
        url = build_url(keyword, city, page_num)
        if page_num == 1:
            console.print(f"[dim]{url}[/]")

        for attempt in range(2):  # 1 retry
            try:
                await page.goto(url, wait_until="domcontentloaded")
                await handle_block(page)
                await jitter(page)
                break
            except Exception:
                if attempt == 0:
                    await page.wait_for_timeout(2000)
                    continue
                raise

        try:
            raw = await page.evaluate("""
                () => {
                    const el = document.querySelector('#__NEXT_DATA__');
                    return el ? JSON.parse(el.textContent) : null;
                }
            """)
        except Exception:
            raw = None

        if not raw:
            if page_num == 1:
                console.print("[red]  ✗ no __NEXT_DATA__ found[/]")
            break

        listings, has_next = parse_next_data(raw)

        if not listings:
            if page_num == 1:
                console.print("[red]  ✗ no results[/]")
            break

        new_count = 0
        for entry in listings:
            fields = extract_fields(entry)
            uid = fields.get('_displayid') or fields.get('_glsrid', '')
            if uid and uid in seen_ids:
                continue
            if uid:
                seen_ids.add(uid)
            all_listings.append(fields)
            new_count += 1

        console.print(
            f"  [dim]page {page_num}: +{new_count} new"
            f"  ({len(all_listings)} unique)[/]"
        )

        if new_count == 0:
            dry_streak += 1
            if dry_streak >= 3:
                console.print(f"  [dim]3 pages with no new results — stopping early[/]")
                break
        else:
            dry_streak = 0

        if not has_next:
            break

        if len(listings) < 5:
            break

        page_num += 1

    elapsed = _time.monotonic() - _start
    console.print(
        f"[bold green]✓[/] {len(all_listings)} suppliers "
        f"({page_num} page{'s' if page_num > 1 else ''}, {elapsed:.0f}s)"
    )
    return all_listings


def dedupe(rows: list[dict]) -> list[dict]:
    """Dedupe by displayid or company name + city."""
    seen: set[str] = set()
    out = []
    for r in rows:
        key = r.get('_displayid', '')
        if not key:
            key = f"{r.get('companyName', '').lower().strip()}|{r.get('city', '').lower()}|{r.get('product', '').lower().strip()}"
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


# --- output ---------------------------------------------------------------
def write_xlsx(rows: list[dict], out: Path) -> None:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        'companyName', 'product', 'price', 'moq', 'city', 'state',
        'district', 'pincode', 'pns', 'gst', 'gstVerified', 'rating',
        'ratingCount', 'memberSince', 'annualTurnover', 'category',
        'bizType', 'specs', 'companyUrl', 'productUrl',
    ]
    headers = [
        'Company', 'Product', 'Price', 'MOQ', 'City', 'State',
        'District', 'Pincode', 'Phone (PNS)', 'GST', 'GST Verified', 'Rating',
        'Ratings Count', 'Member Since', 'Annual Turnover', 'Category',
        'Business Type', 'Specifications', 'Company URL', 'Product URL',
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "suppliers"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color='DDDDDD')
        cell.alignment = Alignment(horizontal='center')

    for r in rows:
        ws.append([r.get(c, '') for c in cols])

    widths = {
        'A': 32, 'B': 30, 'C': 14, 'D': 12, 'E': 14, 'F': 14,
        'G': 14, 'H': 10, 'I': 16, 'J': 18, 'K': 12, 'L': 8,
        'M': 12, 'N': 14, 'O': 18, 'P': 24, 'Q': 16, 'R': 30,
        'S': 36, 'T': 50,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    wb.save(out)


# --- CLI ------------------------------------------------------------------
def get_inputs() -> tuple[list[str], list[str], Path]:
    """Interactive mode."""
    console.print(Panel.fit(
        "[bold cyan]IndiaMART Scraper[/]\n"
        "[dim]Extract B2B supplier listings via __NEXT_DATA__ JSON[/]",
        border_style="cyan",
    ))

    kw_input = Prompt.ask(
        "\n[bold]Product/keyword(s)[/] "
        "[dim](comma-separated, e.g. mobile accessories, phone chargers, usb cables)[/]"
    ).strip()
    keywords = [k.strip() for k in kw_input.split(',') if k.strip()]

    city_input = Prompt.ask(
        "[bold]City(s)[/] "
        "[dim](comma-separated, e.g. bangalore, delhi, mumbai)[/]"
    ).strip()
    cities = [c.strip() for c in city_input.split(',') if c.strip()]

    default_out = f"im_{slugify(keywords[0])}_{slugify(cities[0])}.xlsx"
    out_str = Prompt.ask("[bold]Output file[/]", default=default_out).strip()
    out = Path(out_str)
    if out.suffix.lower() != '.xlsx':
        out = out.with_suffix('.xlsx')

    total = len(keywords) * len(cities)
    console.print(f"\n[dim]{len(keywords)} keyword(s) × {len(cities)} city(s) = {total} search(es):[/]")
    for kw in keywords:
        for c in cities:
            console.print(f"  [cyan]•[/] {kw} in {c}")
    console.print(f"[dim]→ saving to[/] [yellow]{out}[/]\n")

    return keywords, cities, out


def parse_args() -> tuple[list[str], list[str], Path]:
    """One-shot: indiamart.py "kw1, kw2" "city1, city2" [out.xlsx].

    Keywords and cities are both comma-separated.
    """
    kw_input = sys.argv[1]
    keywords = [k.strip() for k in kw_input.split(',') if k.strip()]
    city_input = sys.argv[2]
    cities = [c.strip() for c in city_input.split(',') if c.strip()]

    if len(sys.argv) > 3:
        out = Path(sys.argv[3])
    else:
        out = Path(f"im_{slugify(keywords[0])}_{slugify(cities[0])}.xlsx")
    if out.suffix.lower() != '.xlsx':
        out = out.with_suffix('.xlsx')
    return keywords, cities, out


async def main():
    if len(sys.argv) == 1:
        keywords, cities, out = get_inputs()
    elif len(sys.argv) >= 3:
        keywords, cities, out = parse_args()
    else:
        console.print("[bold]Usage:[/]")
        console.print(
            "  python indiamart.py                                        "
            "[dim](interactive)[/]"
        )
        console.print('  python indiamart.py "keyword1, keyword2" "city1, city2"')
        console.print(
            '  python indiamart.py "mobile accessories, phone chargers" "bangalore" '
            "[out.xlsx]"
        )
        console.print("  [dim](or: uv run indiamart.py ...)[/]")
        sys.exit(1)

    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    console.print(f"[dim]profile: {PROFILE_DIR}[/]")

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=HEADLESS,
            locale="en-US",
            viewport={"width": 1400, "height": 900},
            user_agent=USER_AGENT,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        await ctx.add_init_script(STEALTH_JS)
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        global _partial_rows, _partial_out
        _partial_out = out
        all_rows: list[dict] = []
        for kw in keywords:
            for city in cities:
                try:
                    rows = await scrape_query(page, kw, city)
                    all_rows.extend(rows)
                    _partial_rows = list(all_rows)
                except SystemExit:
                    raise
                except Exception as e:
                    console.print(f"[red]✗[/] {kw} in {city}: {e}")

        await ctx.close()

    before = len(all_rows)
    deduped = dedupe(all_rows)
    after = len(deduped)
    console.print(
        f"\n[bold]Total:[/] {before} scraped → "
        f"[green]{after}[/] unique ([yellow]{before - after}[/] dupes removed)"
    )

    write_xlsx(deduped, out)
    console.print(f"[bold green]✓[/] saved [yellow]{after}[/] → [yellow]{out}[/]")


_partial_rows: list[dict] = []
_partial_out: Path | None = None


if __name__ == "__main__":
    ensure_chromium()
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        console.print("\n[yellow]Interrupted[/]")
        if _partial_rows and _partial_out:
            partial_file = _partial_out.with_stem(_partial_out.stem + '_partial')
            try:
                write_xlsx(_partial_rows, partial_file)
                console.print(
                    f"[yellow]Partial save:[/] {len(_partial_rows)} rows → "
                    f"[yellow]{partial_file}[/]"
                )
            except Exception:
                pass
        os._exit(130)
    except EOFError:
        console.print(
            "\n[red]No interactive input available.[/] "
            'Use one-shot: [cyan]python indiamart.py "keyword" "city"[/]'
        )
        os._exit(1)
    os._exit(0)
