# /// script
# requires-python = ">=3.11"
# dependencies = ["patchright", "rich", "openpyxl"]
# ///
"""
Google Maps scraper — free Apify replacement.

Zero setup: needs only Python 3.11+. Missing Python deps and Chromium are
auto-installed on first run. `uv` is optional.

Interactive (prompts for input):
  python scrape.py           # or: uv run scrape.py

One-shot:
  python scrape.py "mobile shops" "bidar, humnabad, basavakalyan"
  python scrape.py "restaurants" "koramangala bangalore" out.xlsx

Environment variables:
  HEADLESS=1     run Chromium without a window
  KEEP_ALL=1     disable category filter
  DELAY_MIN_MS   min jitter (ms) between extractions (default 600)
  DELAY_MAX_MS   max jitter (ms) between extractions (default 1800)
  PROFILE_DIR    persistent Chrome profile dir (default ~/.cache/maps-scraper/profile)
"""
# --- stdlib only (safe to import before bootstrap) ----------------------
import asyncio
import importlib.util
import os
import random
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


# --- First-run bootstrap ------------------------------------------------
# Makes this script zero-setup on any machine with Python 3.11+.
# - If you run with `uv run`, uv resolves deps from the `# /// script` header
#   above and this block is a no-op.
# - If you run with plain `python scrape.py`, this block pip-installs any
#   missing third-party packages on first invocation and then continues.
def _bootstrap_deps() -> None:
    required = ['patchright', 'rich', 'openpyxl']
    missing = [p for p in required if importlib.util.find_spec(p) is None]
    if not missing:
        return
    print(f"[first-run] installing Python deps: {', '.join(missing)}",
          file=sys.stderr, flush=True)
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", *missing],
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(
            f"[first-run] pip install failed ({e}).\n"
            f"  Install manually:  pip install {' '.join(missing)}\n"
            f"  Or use uv:         uv run scrape.py ...",
            file=sys.stderr, flush=True,
        )
        sys.exit(1)
    print("[first-run] ✓ Python deps installed", file=sys.stderr, flush=True)


_bootstrap_deps()


# --- third-party (guaranteed available after bootstrap) -----------------
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from patchright.async_api import Page, async_playwright
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.prompt import Prompt

console = Console()

HEADLESS = os.getenv("HEADLESS", "0") == "1"  # HEADLESS=1 to run without window
KEEP_ALL = os.getenv("KEEP_ALL", "0") == "1"  # KEEP_ALL=1 to disable category filter
DELAY_MIN_MS = int(os.getenv("DELAY_MIN_MS", "600"))  # min jitter between extracts
DELAY_MAX_MS = int(os.getenv("DELAY_MAX_MS", "1800"))  # max jitter
PARALLEL_TABS = int(os.getenv("PARALLEL_TABS", "3"))  # concurrent extraction tabs
PROFILE_DIR = Path(os.getenv("PROFILE_DIR", str(Path.home() / ".cache" / "maps-scraper" / "profile")))
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")

NETWORK_ERROR_HINTS = (
    "ERR_NAME_NOT_RESOLVED",
    "ERR_INTERNET_DISCONNECTED",
    "ERR_CONNECTION_TIMED_OUT",
    "ERR_PROXY_CONNECTION_FAILED",
    "ERR_TUNNEL_CONNECTION_FAILED",
    "ERR_CONNECTION_RESET",
)


def is_network_error(exc: Exception) -> bool:
    msg = str(exc)
    return any(hint in msg for hint in NETWORK_ERROR_HINTS)


def print_browser_network_error(exc: Exception) -> None:
    console.print(Panel.fit(
        "[bold red]Browser network/DNS failure[/]\n"
        "Chromium could not resolve or reach Google Maps. This is not a keyword/filter issue.\n\n"
        "Try these on Windows:\n"
        "1. Open https://www.google.com/maps in Chrome/Edge.\n"
        "2. Run: ipconfig /flushdns\n"
        "3. Run: nslookup www.google.com\n"
        "4. Disable VPN/proxy/firewall filtering temporarily, then retry.\n"
        "5. Use Visible mode once to confirm the browser itself can load Google Maps.\n\n"
        f"[dim]{type(exc).__name__}: {exc}[/]",
        border_style="red",
    ))


async def check_browser_network(page: Page) -> None:
    try:
        await page.goto("https://www.google.com/", wait_until="commit", timeout=30000)
    except Exception as exc:
        if is_network_error(exc):
            print_browser_network_error(exc)
            raise SystemExit(1)
        raise

# Stealth patches — hide common automation fingerprints from Google's detector.
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
CARD_WAIT_MS = 700
SCROLL_WAIT_MS = 1800

COUNTRY_CODES = {
    'india': 'IN', 'united states': 'US', 'usa': 'US',
    'united kingdom': 'GB', 'uk': 'GB', 'canada': 'CA', 'australia': 'AU',
    'germany': 'DE', 'france': 'FR', 'japan': 'JP', 'china': 'CN',
    'pakistan': 'PK', 'bangladesh': 'BD', 'sri lanka': 'LK', 'nepal': 'NP',
    'singapore': 'SG', 'malaysia': 'MY', 'thailand': 'TH', 'indonesia': 'ID',
    'philippines': 'PH', 'vietnam': 'VN', 'south korea': 'KR', 'taiwan': 'TW',
    'uae': 'AE', 'united arab emirates': 'AE', 'saudi arabia': 'SA',
    'qatar': 'QA', 'kuwait': 'KW', 'oman': 'OM', 'bahrain': 'BH',
}

# Intent → terms matched against Google's structured `categoryName` field
# (short, clean strings like "Cell phone store"), NOT against free-form titles.
# Single-word terms use word-boundary match (so 'mobile' won't match 'snowmobile'
# or 'automobile'). Multi-word terms use substring match. KEEP_ALL=1 disables.
CATEGORY_TERMS = {
    'mobile':     ['phone', 'cell', 'telecom', 'electronics', 'computer', 'smartphone', 'gadget'],
    'phone':      ['phone', 'cell', 'telecom', 'electronics', 'computer', 'smartphone'],
    'laptop':     ['computer', 'laptop', 'electronics'],
    'restaurant': ['restaurant', 'cafe', 'food', 'dining', 'pizza', 'biryani', 'dhaba', 'eatery', 'bistro', 'coffee'],
    'cafe':       ['cafe', 'coffee', 'restaurant', 'bistro', 'tea'],
    'pharmacy':   ['pharmacy', 'chemist', 'drug store', 'medical store'],
    'medical':    ['hospital', 'clinic', 'medical', 'pharmacy', 'chemist'],
    'salon':      ['salon', 'spa', 'parlor', 'parlour', 'barber', 'beauty'],
    'beauty':     ['salon', 'spa', 'parlor', 'parlour', 'beauty', 'cosmetic'],
    'hospital':   ['hospital', 'clinic', 'medical', 'maternity', 'nursing'],
    'clinic':     ['clinic', 'hospital', 'medical', 'doctor'],
    'school':     ['school', 'kindergarten', 'pre-school'],
    'college':    ['college', 'university', 'institute'],
    'tuition':    ['tutor', 'tuition', 'coaching', 'institute', 'academy'],
    'gym':        ['gym', 'fitness', 'health club', 'crossfit', 'yoga'],
    'hotel':      ['hotel', 'lodge', 'resort', 'inn', 'guest house', 'motel'],
    'gas':        ['gas station', 'petrol', 'fuel'],
    'grocery':    ['grocery', 'supermarket', 'hypermarket', 'general store', 'mart'],
    'clothing':   ['clothing', 'apparel', 'boutique', 'fashion', 'tailor', 'garment'],
    'jewelry':    ['jewelry', 'jewellery', 'goldsmith', 'gold', 'silver'],
    'bank':       ['bank', 'atm', 'credit union'],
    'car':        ['car dealer', 'auto repair', 'automobile', 'vehicle'],
    'bike':       ['motorcycle', 'two-wheeler', 'bicycle'],
}


def ensure_chromium() -> None:
    """First-run setup: download Chromium for Patchright/Playwright."""
    if os.name == 'nt':
        local = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / 'AppData' / 'Local')))
    elif sys.platform == 'darwin':
        local = Path.home() / 'Library' / 'Caches'
    else:
        local = Path(os.environ.get('XDG_CACHE_HOME', str(Path.home() / '.cache')))

    # Check both patchright and playwright browser caches
    for dirname in ('ms-playwright', 'ms-patchright'):
        base = local / dirname
        if base.exists() and any(d.name.startswith('chromium') for d in base.iterdir()):
            return

    console.print(Panel.fit(
        "[bold yellow]First-time setup[/]\n"
        "[dim]Downloading Chromium browser (~150 MB, one-time)...[/]",
        border_style="yellow",
    ))
    # Try patchright first, fall back to playwright
    for mod in ('patchright', 'playwright'):
        try:
            subprocess.run(
                [sys.executable, "-m", mod, "install", "chromium"],
                check=True,
            )
            console.print("[bold green]✓ Chromium ready[/]\n")
            return
        except (subprocess.CalledProcessError, FileNotFoundError):
            continue
    console.print("[red]✗ Chromium install failed.[/]")
    console.print("[yellow]Install manually:[/]  python -m patchright install chromium")
    sys.exit(1)


def match_intent(keyword: str) -> tuple[str | None, list[str] | None]:
    """Match the search keyword to an intent + its category terms."""
    kw = keyword.lower()
    for key in sorted(CATEGORY_TERMS, key=len, reverse=True):
        if key in kw:
            return key, CATEGORY_TERMS[key]
    return None, None


def _category_matches(category: str, terms: list[str]) -> bool:
    """Prefix-on-word match for single terms (so 'telecom' catches
    'telecommunications' and 'phone' catches 'phones'), substring for
    multi-word terms. Avoids false positives like 'snowmobile' for 'mobile'."""
    text = category.lower()
    words = re.findall(r"[a-z]+", text)
    for t in terms:
        t = t.strip().lower()
        if not t:
            continue
        if ' ' in t or '-' in t:
            if t in text:
                return True
        else:
            if any(w.startswith(t) for w in words):
                return True
    return False


def filter_by_category(rows: list[dict], keyword: str) -> tuple[list[dict], list[dict], str | None]:
    """Split rows into (kept, dropped, intent_used)."""
    intent, terms = match_intent(keyword)
    if not terms:
        return rows, [], None
    kept, dropped = [], []
    for r in rows:
        if _category_matches(r.get('categoryName') or '', terms):
            kept.append(r)
        else:
            dropped.append(r)
    return kept, dropped, intent

async def jitter(page) -> None:
    await page.wait_for_timeout(random.randint(DELAY_MIN_MS, DELAY_MAX_MS))


async def is_captcha(page) -> bool:
    """Detect Google's anti-bot challenges."""
    url = page.url or ''
    if '/sorry/' in url or 'captcha' in url:
        return True
    selectors = (
        'iframe[src*="recaptcha"]',
        'form[action*="/sorry/"]',
        'div#recaptcha',
        '#captcha-form',
    )
    for sel in selectors:
        try:
            if await page.locator(sel).count() > 0:
                return True
        except Exception:
            pass
    return False


async def handle_captcha(page) -> None:
    """Pause if Google blocks us. Headed = wait for human, headless = abort."""
    if not await is_captcha(page):
        return
    console.print()
    console.print(Panel.fit(
        "[bold red]⚠ CAPTCHA detected[/]\n\n"
        f"[yellow]URL:[/]  {page.url}\n"
        f"[yellow]Mode:[/] {'HEADLESS (no window)' if HEADLESS else 'HEADED'}",
        border_style="red",
    ))
    if HEADLESS:
        screenshot = Path("captcha_screenshot.png").absolute()
        try:
            await page.screenshot(path=str(screenshot), full_page=True)
            console.print(f"[yellow]Screenshot:[/] {screenshot}")
        except Exception:
            pass
        console.print(
            "[red]Re-run WITHOUT HEADLESS=1 once to solve the captcha. "
            "Your profile is persistent — afterwards headless will work.[/]"
        )
        sys.exit(2)
    else:
        is_tty = sys.stdin.isatty() if hasattr(sys.stdin, 'isatty') else False
        if is_tty:
            console.print(
                "[bold]Solve the captcha in the visible browser, "
                "then press [cyan]ENTER[/] here (or just solve it — "
                "the script auto-detects either way).[/]"
            )

            async def _wait_for_enter():
                loop = asyncio.get_running_loop()
                await loop.run_in_executor(None, sys.stdin.readline)

            enter_task = asyncio.create_task(_wait_for_enter())
        else:
            console.print(
                "[bold]Solve the captcha in the visible browser. "
                "[dim](no TTY — auto-resuming as soon as it clears, up to 10 min)[/][/]"
            )
            enter_task = None

        # Poll until the captcha clears, up to ~10 minutes. Works in both
        # foreground (TTY) and background (no-TTY) runs — ENTER is optional.
        max_wait_s = 600
        check_every_s = 5
        elapsed = 0
        while elapsed < max_wait_s:
            if not await is_captcha(page):
                if enter_task and not enter_task.done():
                    enter_task.cancel()
                console.print(f"[green]✓ Resuming[/] [dim](after {elapsed}s)[/]\n")
                return
            await asyncio.sleep(check_every_s)
            elapsed += check_every_s
            if elapsed and elapsed % 30 == 0:
                console.print(
                    f"[dim]still waiting for captcha solve... "
                    f"({elapsed}s / {max_wait_s}s)[/]"
                )
        if enter_task and not enter_task.done():
            enter_task.cancel()
        console.print("[red]Captcha still present after 10 min. Aborting.[/]")
        sys.exit(2)


def parse_address(full: str) -> dict:
    if not full:
        return {}
    parts = [p.strip() for p in full.split(',')]
    out = {}
    if len(parts) >= 2:
        country_raw = parts[-1]
        out['countryCode'] = COUNTRY_CODES.get(country_raw.lower(), country_raw[:2].upper())
    if len(parts) >= 3:
        state_raw = parts[-2]
        out['state'] = re.sub(r'\s*\d{4,7}.*$', '', state_raw).strip()
        out['city'] = parts[-3]
        if len(parts) > 3:
            out['street'] = ', '.join(parts[:-3])
    elif len(parts) == 2:
        out['city'] = parts[0]
    return out


async def scroll_all(page: Page) -> int:
    feed = page.locator('div[role="feed"]')
    try:
        await feed.wait_for(state="visible", timeout=10000)
    except Exception:
        return 0
    end_marker = page.get_by_text("You've reached the end of the list")
    prev, stale = 0, 0
    with console.status("[cyan]Scrolling...") as s:
        while True:
            await feed.evaluate("el => el.scrollTo(0, el.scrollHeight)")
            await page.wait_for_timeout(SCROLL_WAIT_MS)
            count = await page.locator('div[role="feed"] a.hfpxzc').count()
            s.update(f"[cyan]Loaded {count} results")
            if await end_marker.count() > 0:
                break
            if count == prev:
                stale += 1
                if stale >= 4:
                    break
            else:
                stale = 0
            prev = count
    return await page.locator('div[role="feed"] a.hfpxzc').count()


async def extract(page: Page, href: str) -> dict:
    # 60s timeout handles slow networks; Google Maps place pages can be heavy
    await page.goto(href, wait_until="domcontentloaded", timeout=60000)
    await handle_captcha(page)
    await page.wait_for_selector('h1.DUwDvf', timeout=30000)
    await jitter(page)

    d: dict = {}
    d['title'] = (await page.locator('h1.DUwDvf').first.inner_text()).strip()

    rb = page.locator('div.F7nice')
    if await rb.count() > 0:
        m = re.search(r'([\d.]+)\s*\(([\d,]+)\)', await rb.first.inner_text())
        if m:
            d['totalScore'] = float(m.group(1))
            d['reviewsCount'] = int(m.group(2).replace(',', ''))

    for sel in ['button.DkEaL', 'button[jsaction*="pane.wfvdle"]', 'button[jsaction*="category"]']:
        cat = page.locator(sel).first
        if await cat.count() > 0:
            txt = (await cat.inner_text()).strip()
            if txt and len(txt) < 80:
                d['categoryName'] = txt
                break

    addr_el = page.locator('button[data-item-id="address"]').first
    if await addr_el.count() > 0:
        full = ((await addr_el.get_attribute('aria-label')) or '').replace('Address: ', '').strip()
        d['_full_address'] = full
        d.update(parse_address(full))

    web = page.locator('a[data-item-id="authority"]').first
    if await web.count() > 0:
        d['website'] = await web.get_attribute('href')

    phone = page.locator('button[data-item-id^="phone"]').first
    if await phone.count() > 0:
        d['phone'] = ((await phone.get_attribute('aria-label')) or '').replace('Phone: ', '').strip()

    # Real place coords: check the resolved page.url first (has !3d!4d after
    # Google resolves the place ID), fall back to the original href.
    d['url'] = page.url
    m = re.search(r'!3d(-?[\d.]+)!4d(-?[\d.]+)', page.url)
    if not m:
        m = re.search(r'!3d(-?[\d.]+)!4d(-?[\d.]+)', href)
    if not m:
        m = re.search(r'/@(-?[\d.]+),(-?[\d.]+)', page.url)
    if m:
        d['_lat'], d['_lng'] = float(m.group(1)), float(m.group(2))
    return d


async def _extract_worker(tab: Page, hrefs: list[str], results: list[dict],
                          lock: asyncio.Lock, prog, task_id) -> None:
    """Worker that extracts places from its chunk of hrefs using its own tab."""
    for href in hrefs:
        for attempt in range(2):
            try:
                data = await extract(tab, href)
                async with lock:
                    results.append(data)
                    prog.update(task_id, advance=1,
                                description=f"[cyan]{data.get('title','?')[:38]}")
                break
            except Exception as e:
                if attempt == 0:
                    await tab.wait_for_timeout(2000)
                    continue
                async with lock:
                    console.print(f"[red]![/] {type(e).__name__}: {e}")
                    prog.update(task_id, advance=1)


async def scrape_query(page: Page, query: str, ctx=None) -> list[dict]:
    console.print(f"[bold cyan]→[/] [yellow]{query}[/]")
    url = f"https://www.google.com/maps/search/{query.replace(' ', '+')}"
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as exc:
        if is_network_error(exc):
            print_browser_network_error(exc)
            raise SystemExit(1)
        raise
    await handle_captcha(page)

    total = await scroll_all(page)
    if total == 0:
        console.print("[red]✗[/] no results")
        return []

    # Harvest every card's canonical href upfront — stable place URLs.
    raw_hrefs = await page.locator('div[role="feed"] a.hfpxzc').evaluate_all(
        "els => els.map(e => e.href).filter(Boolean)"
    )
    seen_ids: set[str] = set()
    hrefs: list[str] = []
    for h in raw_hrefs:
        m = re.search(r'!1s(0x[0-9a-f]+:0x[0-9a-f]+)', h)
        pid = m.group(1) if m else h
        if pid not in seen_ids:
            seen_ids.add(pid)
            hrefs.append(h)

    console.print(f"[bold green]✓[/] {len(hrefs)} unique places (from {total} cards)")

    n_tabs = min(PARALLEL_TABS, len(hrefs)) if ctx else 1
    if n_tabs < 2:
        # Fallback: single-tab sequential extraction (original behavior)
        ctx = None

    results: list[dict] = []
    lock = asyncio.Lock()

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
    ) as prog:
        task_id = prog.add_task(
            f"[cyan]Extracting ({n_tabs} tab{'s' if n_tabs > 1 else ''})",
            total=len(hrefs),
        )

        if ctx and n_tabs > 1:
            # Split hrefs into N chunks, one per tab
            chunks: list[list[str]] = [[] for _ in range(n_tabs)]
            for i, href in enumerate(hrefs):
                chunks[i % n_tabs].append(href)

            # Create extra tabs (reuse the main page as tab 0)
            tabs = [page]
            for _ in range(n_tabs - 1):
                tabs.append(await ctx.new_page())

            try:
                workers = [
                    _extract_worker(tabs[i], chunks[i], results, lock, prog, task_id)
                    for i in range(n_tabs)
                ]
                await asyncio.gather(*workers)
            finally:
                # Close extra tabs (keep the original page)
                for tab in tabs[1:]:
                    await tab.close()
        else:
            # Single-tab sequential
            for i, href in enumerate(hrefs):
                for attempt in range(2):
                    try:
                        data = await extract(page, href)
                        results.append(data)
                        prog.update(task_id, advance=1,
                                    description=f"[cyan]{data.get('title','?')[:38]}")
                        break
                    except Exception as e:
                        if attempt == 0:
                            await page.wait_for_timeout(2000)
                            continue
                        console.print(f"[red]![/] #{i}: {type(e).__name__}: {e}")
                        prog.update(task_id, advance=1)

    return results


def dedupe(rows: list[dict]) -> list[dict]:
    seen: dict = {}
    for r in rows:
        key = (
            r.get('title', '').lower().strip(),
            round(r.get('_lat', 0), 5),
            round(r.get('_lng', 0), 5),
        )
        if key not in seen:
            seen[key] = r
    return list(seen.values())


# Chain stores / brand names to highlight in the output Excel.
# Any result whose title contains one of these (word-boundary match,
# case-insensitive) gets highlighted. Add/remove names here as needed.
CHAIN_NAMES = [
    'poorvika', 'sangeetha', 'pai mobiles', 'ss mobiles', 'mi store',
    'i central', 'i centre', 'icentre', 'icentral', 'invent tribe',
    'icare', 'imagine', 'iplanet', 'maple', 'unicorn', 'invent',
    'ivenus', 'idestiny', 'samsung cafe', 'oppo', 'vivo', 'phonorange',
    'phono range',
    'my jio store', 'ntel mobiles', 'sara mobiles', 'pai international',
    'samsung service center', 'reliance digital', 'xiaomi',
]

CHAIN_FILL = 'C6EFCE'  # light green highlight for chain-store rows
NO_PHONE_FILL = 'FFEB9C'  # light yellow highlight for missing phone numbers


def is_chain_match(title: str) -> bool:
    """True if the business title matches any highlighted chain/brand."""
    if not title:
        return False
    t = title.lower()
    return any(
        re.search(rf'(?<![a-z]){re.escape(k)}(?![a-z])', t)
        for k in CHAIN_NAMES
    )


def _apply_highlights(ws, cols: list, values: list, chain_match: bool) -> None:
    """Apply per-row highlighting:
    - Chain store rows get a green fill across the row.
    - Any phone cell with NO number gets a yellow fill.
    """
    row_num = ws.max_row
    phone_idx = cols.index('phone')
    phone_val = str(values[phone_idx] or '').strip()
    green = PatternFill('solid', start_color=CHAIN_FILL)
    yellow = PatternFill('solid', start_color=NO_PHONE_FILL)
    if chain_match:
        for idx in range(len(values)):
            if idx == phone_idx:
                continue
            ws.cell(row=row_num, column=idx + 1).fill = green
    if phone_val:
        if chain_match:
            ws.cell(row=row_num, column=phone_idx + 1).fill = green
    else:
        ws.cell(row=row_num, column=phone_idx + 1).fill = yellow


def write_xlsx(rows: list[dict], dropped: list[dict], out: Path) -> Path:
    """Write results to Excel with chain highlights + missing-phone
    highlights. Returns the actual file saved (auto-fallback if locked)."""
    # Defensive: ensure parent directory exists before saving
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        'title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
        'countryCode', 'website', 'phone',
    ]
    cat_cols = [f'categories/{i}' for i in range(11)]
    cols += cat_cols + ['url', 'categoryName']

    def _row(r: dict) -> list:
        primary = r.get('categoryName', '') or ''
        cats = [primary] + [''] * 10 if primary else [''] * 11
        return [
            r.get('title', ''),
            r.get('totalScore', ''),
            r.get('reviewsCount', ''),
            r.get('street', ''),
            r.get('city', ''),
            r.get('state', ''),
            r.get('countryCode', ''),
            r.get('website', ''),
            r.get('phone', ''),
            *cats,
            r.get('url', ''),
            primary,
        ]

    def _format_sheet(ws):
        ws.append(cols)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill('solid', start_color='DDDDDD')
            cell.alignment = Alignment(horizontal='center')

    wb = Workbook()
    ws = wb.active
    ws.title = "shops"
    _format_sheet(ws)
    for r in rows:
        values = _row(r)
        ws.append(values)
        _apply_highlights(ws, cols, values, is_chain_match(r.get('title', '')))

    widths = {
        'A': 36, 'B': 10, 'C': 12, 'D': 32, 'E': 18, 'F': 18,
        'G': 12, 'H': 32, 'I': 18,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'

    if dropped:
        ws2 = wb.create_sheet("filtered_out")
        _format_sheet(ws2)
        for r in dropped:
            values = _row(r)
            ws2.append(values)
            _apply_highlights(ws2, cols, values, is_chain_match(r.get('title', '')))
        for col, w in widths.items():
            ws2.column_dimensions[col].width = w
        ws2.freeze_panes = 'A2'

    return save_workbook(wb, out)


def save_workbook(wb, out: Path) -> Path:
    """Save the workbook, auto-fallback to name_2.xlsx, name_3.xlsx, ...
    if the requested file is locked (e.g. open in Excel)."""
    try:
        wb.save(out)
        return out
    except PermissionError:
        pass
    stem = out.stem
    for i in range(2, 100):
        candidate = out.with_name(f"{stem}_{i}.xlsx")
        try:
            wb.save(candidate)
            console.print(f"[yellow]'{out.name}' is locked (open in Excel?) — "
                          f"saved to [bold]{candidate.name}[/] instead[/]")
            return candidate
        except PermissionError:
            continue
    console.print(f"[red]Could not save '{out.name}' — it's locked by another "
                  f"program. Close it and re-run.[/]")
    raise


def slugify(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:60]


def get_inputs() -> tuple[str, list[str], Path]:
    """Interactive mode — prompt the user for inputs."""
    console.print(Panel.fit(
        "[bold cyan]Google Maps Scraper[/]\n"
        "[dim]Free alternative to Apify — works with any keyword + place(s)[/]",
        border_style="cyan",
    ))

    keyword = Prompt.ask(
        "\n[bold]What to search for[/] [dim](comma-separated for multiple, "
        "e.g. mobile phones, mobile accessories, mobile shop)[/]"
    ).strip()

    place_input = Prompt.ask(
        "[bold]Place(s)[/] [dim](e.g. belura, bidar district, karnataka. "
        "Multiple places: separate with ; like 'belura, bidar; humnabad, bidar')[/]"
    ).strip()

    # Each place phrase can contain commas (city, district, state), so
    # multiple places are separated by ';' instead.
    places = [p.strip() for p in place_input.split(';') if p.strip()]
    keywords = [k.strip() for k in keyword.split(',') if k.strip()]
    queries = [f"{kw} in {p}" for kw in keywords for p in places]

    default_out = f"{slugify(keywords[0])}_{slugify(places[0])}.xlsx"
    out_str = Prompt.ask(
        "[bold]Output file[/]",
        default=default_out,
    ).strip()
    out = Path(out_str)

    console.print()
    console.print(f"[dim]Will run {len(queries)} search(es):[/]")
    for q in queries:
        console.print(f"  [cyan]•[/] {q}")
    console.print(f"[dim]→ saving to[/] [yellow]{out}[/]\n")

    return keyword, queries, out


def parse_args() -> tuple[str, list[str], Path]:
    """One-shot mode — `scrape.py "kw1, kw2" "place1; place2" [out.xlsx]`.

    Multiple keywords run as a cartesian product (every keyword x every place)
    in one browser session; results are deduped across ALL queries so a
    business that shows up under multiple keywords appears only once.

    Place phrases can contain commas (e.g. "belura, bidar district,
    karnataka"); multiple places are separated by ';'.
    """
    keyword = sys.argv[1]
    place_input = sys.argv[2]
    places = [p.strip() for p in place_input.split(';') if p.strip()]
    keywords = [k.strip() for k in keyword.split(',') if k.strip()]
    queries = [f"{kw} in {p}" for kw in keywords for p in places]
    if len(sys.argv) > 3:
        out = Path(sys.argv[3])
    else:
        first_kw = slugify(keywords[0]) if keywords else 'search'
        out = Path(f"{first_kw}_{slugify(places[0])}.xlsx")
    return keyword, queries, out


async def main():
    if len(sys.argv) == 1:
        keyword, queries, out = get_inputs()
    elif len(sys.argv) >= 3:
        keyword, queries, out = parse_args()
    else:
        console.print("[bold]Usage:[/]")
        console.print("  python scrape.py                                      [dim](interactive)[/]")
        console.print("  python scrape.py \"keyword\" \"place1, place2\" [out.xlsx]")
        console.print("  [dim](or: uv run scrape.py ...)[/]")
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
        await check_browser_network(page)

        global _partial_rows, _partial_out
        _partial_out = out
        all_rows = []
        query_errors = 0
        for q in queries:
            try:
                rows = await scrape_query(page, q, ctx=ctx)
                all_rows.extend(rows)
                _partial_rows = list(all_rows)
            except SystemExit:
                raise
            except Exception as e:
                query_errors += 1
                console.print(f"[red]✗[/] {q}: {e}")

        await ctx.close()

    if query_errors and not all_rows:
        console.print("[bold red]All queries failed before any rows were scraped; not writing an empty workbook.[/]")
        raise SystemExit(1)

    before = len(all_rows)
    deduped = dedupe(all_rows)
    after = len(deduped)
    console.print(
        f"\n[bold]Total:[/] {before} scraped → "
        f"[green]{after}[/] unique ([yellow]{before - after}[/] dupes removed)"
    )

    if KEEP_ALL:
        kept, dropped, intent = deduped, [], None
        console.print("[dim]KEEP_ALL=1 → category filter disabled[/]")
    else:
        kept, dropped, intent = filter_by_category(deduped, keyword)

    if intent:
        console.print(
            f"[bold]Filter:[/] intent='[cyan]{intent}[/]' → "
            f"[green]{len(kept)} kept[/], [red]{len(dropped)} filtered[/] "
            f"[dim](set KEEP_ALL=1 to disable)[/]"
        )
        if dropped:
            from collections import Counter
            top = Counter(r.get('categoryName') or '(none)' for r in dropped).most_common(8)
            console.print("[dim]top filtered categories:[/]")
            for c, n in top:
                console.print(f"  [red]✗[/] {c}: {n}")
    elif not KEEP_ALL:
        console.print(f"[dim]No category filter for keyword '{keyword}' (no intent matched)[/]")

    saved = write_xlsx(kept, dropped, out)
    sheets = "shops" + (" + filtered_out" if dropped else "")
    console.print(f"[bold green]✓[/] saved [yellow]{len(kept)}[/] → [yellow]{saved}[/] ({sheets})")
    console.print(f"[bold magenta]FILE_PATH:[/] {saved}")

    n_chain = sum(1 for r in kept if is_chain_match(r.get('title', '')))
    n_no_phone = sum(1 for r in kept if not str(r.get('phone') or '').strip())
    if n_chain or n_no_phone:
        console.print(
            f"[green]\u2728[/] {n_chain} chain-store rows highlighted green "
            f"| [yellow]{n_no_phone} rows with NO phone highlighted yellow[/]"
        )


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
                write_xlsx(_partial_rows, [], partial_file)
                console.print(
                    f"[yellow]Partial save:[/] {len(_partial_rows)} rows → "
                    f"[yellow]{partial_file}[/]"
                )
            except Exception:
                pass
        os._exit(130)
    except EOFError:
        console.print("\n[red]No interactive input available.[/] "
                      "Use one-shot mode: [cyan]python scrape.py \"keyword\" \"place\"[/]")
        os._exit(1)
    # Hard exit to avoid playwright's 300s thread-join timeout on Windows
    os._exit(0)
