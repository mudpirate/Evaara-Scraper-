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
  python scrape.py "cell phone store" "bidar, humnabad, basavakalyan"
  python scrape.py "restaurants" "koramangala bangalore" out.xlsx

GPS grid mode (full coverage of a city/taluka/district):
  python scrape.py "cell phone store" --area "barshi, maharashtra" --step 0.02
  python scrape.py "cell phone store" --bbox 18.07,75.52,18.39,75.84 --step 0.05
  python scrape.py "cell phone store" --area "barshi" --step 0.02 --dry-run

  Grid mode searches every cell of a checkerboard over the area, so small
  village shops are not lost when Google's 120-result cap cuts a broad
  search short. Dense cells that still hit 120 results are split
  automatically into four smaller cells.

Environment variables:
  HEADLESS=1      run Chromium without a window
  KEEP_ALL=1      disable category + junk-name filter (raw results)
  STRICT_QUERY=1  wrap the phrase in quotes (experimental — Maps often
                  treats a quoted phrase as a place and returns only the city)
  DELAY_MIN_MS    min jitter (ms) between extractions (default 600)
  DELAY_MAX_MS    max jitter (ms) between extractions (default 1800)
  PROFILE_DIR     persistent Chrome profile dir (default ~/.cache/maps-scraper/profile)

Accuracy notes:
  Google Maps matches query words against addresses and reviews too, so a
  search for "mobile" can return an insurance agency on "Mobile Avenue" or
  a waterworks company with a review mentioning mobile phones. This script
  filters on the structured Google category (not the address), and searches
  with specific terms like "cell phone store" / "mobile phone repair shop".
"""
# --- stdlib only (safe to import before bootstrap) ----------------------
import asyncio
import importlib.util
import math
import os
import random
import re
import subprocess
import sys
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote_plus

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
from rich.prompt import Confirm, Prompt

from grid import (GridCell, bbox_size_km, grid_shape, lookup_area_bbox,
                  make_cells, parse_bbox)

console = Console()

HEADLESS = os.getenv("HEADLESS", "1") == "1"  # HEADLESS=0 to watch the browser
KEEP_ALL = os.getenv("KEEP_ALL", "0") == "1"  # KEEP_ALL=1 to disable category + junk filter
STRICT_QUERY = os.getenv("STRICT_QUERY", "0") == "1"  # optional phrase quoting (usually hurts Maps)
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
# (short, clean strings like "Cell phone store"), NOT against free-form titles
# or addresses. Single-word terms use word-prefix match; multi-word terms use
# substring match. KEEP_ALL=1 disables the whole filter.
#
# Mobile/phone uses an inclusion list on purpose: broad terms like "computer"
# or plain "mobile" let in computer shops, mobile caterers and Oldsmobile
# dealers. Only categories that actually sell/repair phones are allowed.
MOBILE_CATEGORY_TERMS = [
    'cell phone',
    'cellphone',
    'mobile phone',
    'mobile store',
    'mobile shop',
    'mobile accessory',
    'mobile accessories',
    'phone store',
    'phone shop',
    'phone accessory',
    'phone accessories',
    'phone repair',
    'smartphone',
    'telecommunication service',
    'telecommunications service',
    'telecom service',
    'telephone company',
    'electronic store',
    'electronics store',
    'electronic shop',
    'electronics shop',
    'electronic repair',
    'electronics repair',
    'electronic company',
    'electronics company',
    'electronic wholesaler',
    'electronics wholesaler',
    'electronic accessor',
    'electronics accessor',
]

ELECTRONICS_CATEGORY_TERMS = MOBILE_CATEGORY_TERMS + [
    'appliance store',
    'home appliance store',
    'computer store',
    'computer accessor',
    'computer repair',
]

CATEGORY_TERMS = {
    'mobile':      MOBILE_CATEGORY_TERMS,
    'phone':       MOBILE_CATEGORY_TERMS,
    'telecom':     MOBILE_CATEGORY_TERMS,
    'electronic':  ELECTRONICS_CATEGORY_TERMS,
    'electronics': ELECTRONICS_CATEGORY_TERMS,
    'laptop':      ['computer', 'laptop', 'electronics'],
    'restaurant':  ['restaurant', 'cafe', 'food', 'dining', 'pizza', 'biryani', 'dhaba', 'eatery', 'bistro', 'coffee'],
    'cafe':        ['cafe', 'coffee', 'restaurant', 'bistro', 'tea'],
    'pharmacy':    ['pharmacy', 'chemist', 'drug store', 'medical store'],
    'medical':     ['hospital', 'clinic', 'medical', 'pharmacy', 'chemist'],
    'salon':       ['salon', 'spa', 'parlor', 'parlour', 'barber', 'beauty'],
    'beauty':      ['salon', 'spa', 'parlor', 'parlour', 'beauty', 'cosmetic'],
    'hospital':    ['hospital', 'clinic', 'medical', 'maternity', 'nursing'],
    'clinic':      ['clinic', 'hospital', 'medical', 'doctor'],
    'school':      ['school', 'kindergarten', 'pre-school'],
    'college':     ['college', 'university', 'institute'],
    'tuition':     ['tutor', 'tuition', 'coaching', 'institute', 'academy'],
    'gym':         ['gym', 'fitness', 'health club', 'crossfit', 'yoga'],
    'hotel':       ['hotel', 'lodge', 'resort', 'inn', 'guest house', 'motel'],
    'gas':         ['gas station', 'petrol', 'fuel'],
    'grocery':     ['grocery', 'supermarket', 'hypermarket', 'general store', 'mart'],
    'clothing':    ['clothing', 'apparel', 'boutique', 'fashion', 'tailor', 'garment'],
    'jewelry':     ['jewelry', 'jewellery', 'goldsmith', 'gold', 'silver'],
    'bank':        ['bank', 'atm', 'credit union'],
    'car':         ['car dealer', 'auto repair', 'automobile', 'vehicle'],
    'bike':        ['motorcycle', 'two-wheeler', 'bicycle'],
}

# Structured categories that must never count as a phone/electronics shop.
# These are checked against `categoryName` only. Google also matches search
# words against addresses, so an insurance office on "Mobile Avenue" can
# otherwise slip through.
JUNK_CATEGORY_TERMS = [
    'insur',
    'waterwork', 'water work', 'water utility', 'water supply',
    'plumb',
    'restaurant', 'cafe', 'cafeteria', 'food court', 'caterer', 'food truck',
    'bank', 'atm', 'financ', 'loan',
    'real estate', 'property dealer', 'travel agency',
    'mobile home', 'oldsmobile', 'snowmobile',
    'auto parts', 'car dealer', 'automobile dealer',
    'motorcycle dealer', 'atv dealer', 'atv repair',
    'chauffeur', 'taxi service',
    'notary', 'attorney', 'lawyer',
    'doctor', 'physician', 'hospital', 'clinic', 'pharmac',
    'school', 'college', 'tuition',
    'pet grooming', 'veterinar',
]

# Junk words checked against the business NAME. Kept narrow on purpose:
# a repair shop can legitimately be called "Mobile Hospital" or "Mobile Care",
# and many real shops sit "opposite SBI Bank", so addresses are never used.
JUNK_NAME_TERMS = [
    'insurance', 'waterworks', 'water works', 'plumber', 'plumbing',
    'restaurant', 'cafe', 'bank', 'mobile home', 'mobile caterer',
    'oldsmobile', 'snowmobile', 'auto parts', 'car dealer',
]

# If the name itself clearly says phone shop, don't reject it just because it
# also mentions a locality such as "Bank More" or "Cafe Street".
PHONE_NAME_SIGNALS = (
    'mobile', 'cell', 'phone', 'telecom', 'smartphone', 'electronic', 'recharge',
)

# Only these searches get the junk exclusion. A restaurant search must not
# drop restaurants, and a bank search must not drop banks.
MOBILE_INTENTS = {'mobile', 'phone', 'telecom', 'electronic', 'electronics'}


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


def _name_matches(name: str, terms: list[str]) -> bool:
    """Whole-word match for business names.

    Single words only match complete words, so 'bank' rejects "X Insurance &
    Bank" but not "Banke Mobiles". Multi-word terms match as a substring.
    Addresses are deliberately never checked: real shops often sit
    "opposite SBI Bank" or on "Mobile Avenue".
    """
    text = (name or '').lower()
    words = set(re.findall(r"[a-z]+", text))
    for t in terms:
        t = t.strip().lower()
        if not t:
            continue
        if ' ' in t or '-' in t:
            if t in text:
                return True
        elif t in words:
            return True
    return False


def _has_phone_signal(name: str) -> bool:
    """True when the business name itself clearly signals a phone shop."""
    words = re.findall(r"[a-z]+", (name or '').lower())
    return any(w.startswith(sig) for w in words for sig in PHONE_NAME_SIGNALS)


def filter_by_category(rows: list[dict], keyword: str) -> tuple[list[dict], list[dict], str | None]:
    """Split rows into (kept, dropped, intent_used).

    Positive rule: the structured category must be on the intent's allowlist.
    Negative rule (mobile, telecom and electronics searches only): drop known
    other trades such as insurance, waterworks, restaurants and banks, even if
    Google put them in the results because of an address or review match.
    """
    intent, terms = match_intent(keyword)
    if not terms:
        return rows, [], None
    check_junk = intent in MOBILE_INTENTS
    kept, dropped = [], []
    for r in rows:
        category = str(r.get('categoryName') or r.get('categories/0') or '')
        name = str(r.get('title') or '')
        junk_name = (
            _name_matches(name, JUNK_NAME_TERMS)
            and not _has_phone_signal(name)
        )
        if check_junk and (
            _category_matches(category, JUNK_CATEGORY_TERMS)
            or junk_name
        ):
            dropped.append(r)
            continue
        if _category_matches(category, terms):
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
                          lock: asyncio.Lock, prog, task_id,
                          seen_ids: set[str] | None = None) -> None:
    """Worker that extracts places from its chunk of hrefs using its own tab."""
    for href in hrefs:
        for attempt in range(2):
            try:
                data = await extract(tab, href)
                async with lock:
                    results.append(data)
                    if seen_ids is not None:
                        pid = place_id_from_href(href)
                        if pid:
                            seen_ids.add(pid)
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


def place_id_from_href(href: str) -> str | None:
    """Google place ID from a Maps URL (`!1s0x...:0x...`)."""
    m = re.search(r'!1s(0x[0-9a-fA-F]+:0x[0-9a-fA-F]+)', href or '')
    return m.group(1) if m else None


def coords_from_href(href: str) -> tuple[float, float] | None:
    """Place coordinates from a Maps URL (`!3d<lat>!4d<lon>`)."""
    m = re.search(r'!3d(-?[\d.]+)!4d(-?[\d.]+)', href or '')
    if not m:
        return None
    try:
        return float(m.group(1)), float(m.group(2))
    except ValueError:
        return None


def in_bounds(lat: float, lng: float,
              bounds: tuple[float, float, float, float]) -> bool:
    south, west, north, east = bounds
    return south <= lat <= north and west <= lng <= east


async def scrape_query(page: Page, query: str, ctx=None,
                       center: tuple[float, float] | None = None,
                       zoom: int | None = None,
                       seen_ids: set[str] | None = None,
                       stats: dict | None = None,
                       bounds: tuple[float, float, float, float] | None = None) -> list[dict]:
    if stats is not None:
        stats.clear()
        stats.update(cards=0, unique=0, fresh=0, skipped=0,
                     extracted=0, outside=0, capped=False)

    where = ""
    if center:
        where = (f" [dim]@ {center[0]:.4f},{center[1]:.4f} "
                 f"z{zoom or 15}[/]")
    console.print(f"[bold cyan]→[/] [yellow]{query}[/]{where}")

    base = f"https://www.google.com/maps/search/{quote_plus(query)}"
    url = (f"{base}/@{center[0]:.6f},{center[1]:.6f},{zoom or 15}z"
           if center else base)
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as exc:
        if is_network_error(exc):
            print_browser_network_error(exc)
            raise SystemExit(1)
        raise
    await handle_captcha(page)

    total = await scroll_all(page)
    if stats is not None:
        stats['cards'] = total
        stats['capped'] = total >= 120
    if total == 0:
        console.print("[red]✗[/] no results")
        return []

    # Harvest every card's canonical href upfront — stable place URLs.
    raw_hrefs = await page.locator('div[role="feed"] a.hfpxzc').evaluate_all(
        "els => els.map(e => e.href).filter(Boolean)"
    )
    seen_local: set[str] = set()
    hrefs: list[str] = []
    outside = 0
    for h in raw_hrefs:
        if bounds:
            coords = coords_from_href(h)
            if coords and not in_bounds(coords[0], coords[1], bounds):
                outside += 1
                continue
        pid = place_id_from_href(h) or h
        if pid not in seen_local:
            seen_local.add(pid)
            hrefs.append(h)
    if stats is not None:
        stats['unique'] = len(hrefs)
        stats['outside'] = outside
    if outside:
        console.print(
            f"[yellow]  ↳ {outside} result(s) outside the grid area skipped[/]"
        )
    if bounds and not hrefs:
        console.print("[dim]  no places inside the grid area[/]")
        return []
    if stats is not None and stats['capped']:
        console.print(
            "[yellow]![/] 120-result cap reached; small shops may be missing "
            "in this area (grid mode splits dense cells automatically)."
        )

    if seen_ids is not None:
        fresh: list[str] = []
        for h in hrefs:
            pid = place_id_from_href(h) or h
            if pid not in seen_ids:
                fresh.append(h)
        skipped = len(hrefs) - len(fresh)
        if stats is not None:
            stats['skipped'] = skipped
            stats['fresh'] = len(fresh)
        suffix = f" [dim]· {skipped} already collected[/]" if skipped else ""
        console.print(
            f"[bold green]✓[/] {len(hrefs)} unique places "
            f"(from {total} cards){suffix}"
        )
        if not fresh:
            console.print("[dim]  no new places in this cell[/]")
            return []
        hrefs = fresh
    else:
        console.print(
            f"[bold green]✓[/] {len(hrefs)} unique places (from {total} cards)"
        )

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
                    _extract_worker(tabs[i], chunks[i], results, lock, prog,
                                    task_id, seen_ids)
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
                        if seen_ids is not None:
                            pid = place_id_from_href(href)
                            if pid:
                                seen_ids.add(pid)
                        prog.update(task_id, advance=1,
                                    description=f"[cyan]{data.get('title','?')[:38]}")
                        break
                    except Exception as e:
                        if attempt == 0:
                            await page.wait_for_timeout(2000)
                            continue
                        console.print(f"[red]![/] #{i}: {type(e).__name__}: {e}")
                        prog.update(task_id, advance=1)

    if bounds:
        inside: list[dict] = []
        dropped = 0
        for r in results:
            lat, lng = r.get('_lat'), r.get('_lng')
            if lat is None or lng is None:
                # No coordinates: keep it rather than risk losing a real shop.
                inside.append(r)
            elif in_bounds(float(lat), float(lng), bounds):
                inside.append(r)
            else:
                dropped += 1
        if dropped:
            console.print(
                f"[yellow]  ↳ {dropped} extracted result(s) outside the "
                "grid area dropped[/]"
            )
        results = inside
        if stats is not None:
            stats['outside'] = stats.get('outside', 0) + dropped

    if stats is not None:
        stats['extracted'] = len(results)
    return results


def dedupe(rows: list[dict]) -> list[dict]:
    """Drop the same place seen twice (grid cells overlap).

    Coordinates are the strongest signal; fall back to the Google place ID,
    then to title + address for rows where coordinates could not be read.
    """
    seen: dict = {}
    for r in rows:
        title = str(r.get('title') or '').strip().lower()
        lat, lng = r.get('_lat'), r.get('_lng')
        if lat is not None and lng is not None:
            key = ('geo', title, round(float(lat), 5), round(float(lng), 5))
        else:
            pid = place_id_from_href(r.get('url') or '')
            if pid:
                key = ('pid', pid)
            else:
                key = ('addr', title,
                       str(r.get('street') or '').strip().lower(),
                       str(r.get('city') or '').strip().lower())
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


def write_xlsx(rows: list[dict], dropped: list[dict], out: Path,
               extra_cols: list[str] | None = None) -> Path:
    """Write results to Excel with chain highlights + missing-phone
    highlights. Returns the actual file saved (auto-fallback if locked).

    `extra_cols` lets clean_existing.py preserve columns from an existing
    workbook (for example citySearched). They are written after the standard
    scraper columns. scrape.py itself leaves this unset.
    """
    # Defensive: ensure parent directory exists before saving
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        'title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
        'countryCode', 'website', 'phone',
    ]
    cat_cols = [f'categories/{i}' for i in range(11)]
    cols += cat_cols + ['url', 'categoryName']
    extra: list[str] = []
    for c in (extra_cols or []):
        if c and c not in cols and c not in extra:
            extra.append(c)
    cols += extra

    def _row(r: dict) -> list:
        primary = r.get('categoryName', '') or ''
        cats = [primary] + [''] * 10 if primary else [''] * 11
        values = [
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
        values += [r.get(c, '') for c in extra]
        return values

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


# Vague keywords cast a very wide net on Google Maps. Rewrite the common
# generic phone searches to specific store/repair phrases before searching.
QUERY_REWRITES = {
    'mobile': 'mobile phone store',
    'mobile shop': 'mobile phone store',
    'mobile shops': 'mobile phone store',
    'mobile store': 'mobile phone store',
    'mobile stores': 'mobile phone store',
    'mobile phone': 'mobile phone store',
    'mobile phones': 'mobile phone store',
    'mobile accessory': 'smartphone accessories',
    'mobile accessories': 'smartphone accessories',
    'mobile accesories': 'smartphone accessories',
    'mobile repair': 'mobile phone repair shop',
    'mobile repairs': 'mobile phone repair shop',
    'mobile shop and accessories': 'mobile phone store',
    'mobile phones and accessories': 'mobile phone store',
    'mobile accessories and repair': 'smartphone accessories',
    'electronics and appliances': 'electronics store',
    'electronics & appliances': 'electronics store',
}


def build_term(keyword: str) -> str:
    """Rewrite a vague keyword to a specific business phrase."""
    term = keyword.strip()
    term = QUERY_REWRITES.get(term.lower(), term)
    if STRICT_QUERY and '"' not in term:
        term = f'"{term}"'
    return term


def build_query(keyword: str, place: str) -> str:
    """Build a Maps query from a specific business phrase + place.

    Google Maps does not treat quoted phrases like Google Search does: a
    quoted query such as "cell phone store" in Bidar often resolves to the
    city itself and returns no shops. Quoting is therefore opt-in only
    (STRICT_QUERY=1); the real accuracy gain comes from the specific phrase
    rewrites above and the category filter.
    """
    return f"{build_term(keyword)} in {place}"


def get_inputs() -> tuple[str, list[str], Path]:
    """Interactive mode — prompt the user for inputs."""
    console.print(Panel.fit(
        "[bold cyan]Google Maps Scraper[/]\n"
        "[dim]Free alternative to Apify — works with any keyword + place(s)[/]",
        border_style="cyan",
    ))

    keyword = Prompt.ask(
        "\n[bold]What to search for[/] [dim](comma-separated for multiple, "
        "e.g. cell phone store, smartphone accessories, mobile phone repair shop)[/]"
    ).strip()

    place_input = Prompt.ask(
        "[bold]Place(s)[/] [dim](e.g. belura, bidar district, karnataka. "
        "Multiple places: separate with ; like 'belura, bidar; humnabad, bidar')[/]"
    ).strip()

    # Each place phrase can contain commas (city, district, state), so
    # multiple places are separated by ';' instead.
    places = [p.strip() for p in place_input.split(';') if p.strip()]
    keywords = [k.strip() for k in keyword.split(',') if k.strip()]
    queries = [build_query(kw, p) for kw in keywords for p in places]

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
    queries = [build_query(kw, p) for kw in keywords for p in places]
    if len(sys.argv) > 3:
        out = Path(sys.argv[3])
    else:
        first_kw = slugify(keywords[0]) if keywords else 'search'
        out = Path(f"{first_kw}_{slugify(places[0])}.xlsx")
    return keyword, queries, out


@dataclass
class GridPlan:
    """Everything grid mode needs before the browser starts."""

    cells: list[GridCell]
    terms: list[str]
    bbox: tuple[float, float, float, float]
    step: float
    min_step: float
    area_name: str | None
    zoom: int | None
    yes: bool
    dry_run: bool

    @property
    def bounds(self) -> tuple[float, float, float, float]:
        """Overall area plus a margin for shops just outside the edge.

        Google Maps sometimes ignores the coordinate and falls back to the
        browser's default city when a cell has few local shops. That fallback
        can be hundreds of kilometres away, so every result is checked
        against this box before it is kept.
        """
        margin = max(0.02, self.step)
        lat_min, lon_min, lat_max, lon_max = self.bbox
        return (lat_min - margin, lon_min - margin,
                lat_max + margin, lon_max + margin)


def is_grid_mode(args: list[str]) -> bool:
    return any(a in ('--bbox', '--area')
               or a.startswith(('--bbox=', '--area=')) for a in args)


def parse_grid_args(args: list[str]) -> tuple[str, Path, GridPlan]:
    """Parse `scrape.py "keywords" --bbox ... | --area ... [options]`."""
    def fail(message: str):
        console.print(f"[red]{message}[/]")
        console.print(
            "[dim]Examples:[/]\n"
            "  python scrape.py \"cell phone store\" --area \"barshi, "
            "maharashtra\" --step 0.02\n"
            "  python scrape.py \"cell phone store\" --bbox "
            "18.07,75.52,18.39,75.84 --step 0.05\n"
            "  python scrape.py \"cell phone store\" --area \"barshi\" "
            "--step 0.02 --dry-run"
        )
        sys.exit(1)

    if not args or args[0].startswith('-'):
        fail('Grid mode needs the keyword first, e.g. "cell phone store".')

    keyword = args[0]
    terms = [k.strip() for k in keyword.split(',') if k.strip()]
    if not terms:
        fail('No keyword given.')
    terms = [build_term(t) for t in terms]

    bbox = area = out = None
    step, min_step = 0.02, 0.01
    min_step_given = False
    zoom = None
    yes = dry_run = False

    def take_value(i: int, name: str) -> tuple[str, int]:
        arg = args[i]
        if '=' in arg:
            return arg.split('=', 1)[1], i + 1
        if i + 1 >= len(args):
            fail(f'{name} needs a value.')
        return args[i + 1], i + 2

    i = 1
    while i < len(args):
        arg = args[i]
        name = arg.split('=', 1)[0]
        try:
            if name == '--bbox':
                value, i = take_value(i, '--bbox')
                bbox = parse_bbox(value)
            elif name == '--area':
                value, i = take_value(i, '--area')
                area = value.strip()
                if not area:
                    fail('--area needs a place name.')
            elif name == '--step':
                value, i = take_value(i, '--step')
                step = float(value)
            elif name == '--min-step':
                value, i = take_value(i, '--min-step')
                min_step = float(value)
                min_step_given = True
            elif name == '--zoom':
                value, i = take_value(i, '--zoom')
                zoom = int(value)
            elif name == '--out':
                value, i = take_value(i, '--out')
                out = Path(value)
            elif arg in ('--yes', '-y'):
                yes = True
                i += 1
            elif arg == '--dry-run':
                dry_run = True
                i += 1
            else:
                fail(f'Unknown grid option: {arg}')
        except ValueError as exc:
            fail(str(exc))

    if bbox and area:
        fail('Use either --bbox or --area, not both.')
    if not bbox and not area:
        fail('Grid mode needs --bbox lat1,lon1,lat2,lon2 or --area "place name".')
    if not math.isfinite(step) or step <= 0:
        fail('--step must be a finite number greater than 0.')
    if not min_step_given:
        min_step = min(min_step, step)
    if not math.isfinite(min_step) or min_step <= 0:
        fail('--min-step must be a finite number greater than 0.')
    if min_step > step:
        fail('--min-step cannot be larger than --step.')
    if zoom is not None and not (3 <= zoom <= 20):
        fail('--zoom must be between 3 and 20.')

    area_name = None
    if area:
        console.print(f"[dim]Looking up bounding box for[/] [yellow]{area}[/] ...")
        try:
            bbox, area_name = lookup_area_bbox(area)
        except Exception as exc:
            console.print(f"[red]Could not look up '{area}':[/] {exc}")
            console.print(
                "[dim]Pass --bbox lat1,lon1,lat2,lon2 instead — right-click "
                "two opposite corners in Google Maps and copy the "
                "coordinates.[/]"
            )
            sys.exit(1)

    cells = make_cells(bbox, step, zoom=zoom)
    plans = GridPlan(
        cells=cells,
        terms=terms,
        bbox=bbox,
        step=step,
        min_step=min_step,
        area_name=area_name,
        zoom=zoom,
        yes=yes,
        dry_run=dry_run,
    )

    lat_min, lon_min, lat_max, lon_max = bbox
    rows, cols = grid_shape(bbox, step)
    height_km, width_km = bbox_size_km(bbox)
    zooms = sorted({c.zoom for c in cells})
    zoom_label = (f"auto → {zooms[0]}z" if len(zooms) == 1
                  else f"auto → {zooms[0]}-{zooms[-1]}z")
    if zoom is not None:
        zoom_label = f"{zoom}z (fixed)"
    searches = len(cells) * len(terms)
    cell_word = "cell" if len(cells) == 1 else "cells"
    keyword_word = "keyword" if len(terms) == 1 else "keywords"

    console.print(Panel.fit(
        "[bold cyan]GPS grid search[/]\n"
        f"[bold]Area:[/] {area_name or 'custom bounding box'}\n"
        f"[bold]BBox:[/] S {lat_min:.5f}  W {lon_min:.5f}  "
        f"N {lat_max:.5f}  E {lon_max:.5f}\n"
        f"[dim]({height_km:.1f} km tall x {width_km:.1f} km wide)[/]\n"
        f"[bold]Cells:[/] {rows} x {cols} = [cyan]{len(cells)}[/] "
        f"(step {step:g}° ≈ {step * 111.32:.1f} km)\n"
        f"[bold]Zoom:[/] {zoom_label}\n"
        f"[bold]Searches:[/] {len(cells)} {cell_word} x {len(terms)} "
        f"{keyword_word} = [cyan]{searches}[/]",
        border_style="cyan",
    ))
    console.print(
        "[dim]Cells that hit Google's 120-result cap are split into 4 "
        "smaller cells automatically (down to --min-step "
        f"{min_step:g}°).[/]"
    )

    if out is None:
        if area_name:
            label = area_name.split(',')[0]
        else:
            label = f"bbox_{lat_min:.2f}_{lon_min:.2f}"
        out = Path(f"{slugify(terms[0])}_{slugify(label)}_grid.xlsx")

    return keyword, out, plans


async def main():
    grid: GridPlan | None = None
    if is_grid_mode(sys.argv[1:]):
        keyword, out, grid = parse_grid_args(sys.argv[1:])
        queries = grid.terms
    elif len(sys.argv) == 1:
        keyword, queries, out = get_inputs()
    elif len(sys.argv) >= 3:
        keyword, queries, out = parse_args()
    else:
        console.print("[bold]Usage:[/]")
        console.print("  python scrape.py                                      [dim](interactive)[/]")
        console.print("  python scrape.py \"keyword\" \"place1, place2\" [out.xlsx]")
        console.print("  python scrape.py \"keyword\" --area \"city, state\" --step 0.02")
        console.print("  python scrape.py \"keyword\" --bbox lat1,lon1,lat2,lon2 --step 0.05")
        console.print("  [dim](or: uv run scrape.py ...)[/]")
        sys.exit(1)

    if grid:
        searches = len(grid.cells) * len(grid.terms)
        if grid.dry_run:
            console.print("[yellow]--dry-run:[/] stopping before browser launch; "
                          "no file written.")
            return
        if searches > 50 and not grid.yes:
            lo = searches * 60 / 3600
            hi = searches * 180 / 3600
            question = (
                f"[yellow]This grid needs {searches} searches[/] "
                f"(roughly {lo:.1f}-{hi:.1f} hours). Continue?"
            )
            if sys.stdin.isatty():
                if not Confirm.ask(question, default=False):
                    console.print("[dim]Cancelled.[/]")
                    return
            else:
                console.print(
                    f"[red]{searches} searches with no terminal to confirm. "
                    "Re-run with --yes to start it.[/]"
                )
                sys.exit(1)

    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    console.print(f"[dim]profile: {PROFILE_DIR}[/]")
    if STRICT_QUERY:
        console.print(
            "[yellow]STRICT_QUERY=1: quoted phrases often make Google Maps "
            "return only the city, not shops. Unset it for normal use.[/]"
        )

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
        all_rows: list[dict] = []
        query_errors = 0
        seen_ids: set[str] = set()
        splits = 0

        if grid:
            for term in grid.terms:
                queue: deque[GridCell] = deque(grid.cells)
                cell_no = 0
                while queue:
                    cell = queue.popleft()
                    cell_no += 1
                    stats: dict = {}
                    try:
                        rows = await scrape_query(
                            page, term, ctx=ctx, center=cell.center,
                            zoom=cell.zoom, seen_ids=seen_ids, stats=stats,
                            bounds=grid.bounds,
                        )
                    except SystemExit:
                        raise
                    except Exception as e:
                        query_errors += 1
                        console.print(f"[red]✗[/] cell {cell_no}: {e}")
                        continue
                    if rows:
                        all_rows.extend(rows)
                        _partial_rows = list(all_rows)
                        _save_partial(verbose=False)
                    console.print(
                        f"[dim]  cell {cell_no}: +{stats.get('extracted', 0)} "
                        f"new | total {len(all_rows)} | {len(queue)} queued[/]"
                    )
                    if stats.get('capped') and stats.get('outside', 0) == 0:
                        if cell.can_split(grid.min_step):
                            children = cell.split()
                            splits += 1
                            console.print(
                                f"[yellow]  ↳ 120-result cap: splitting into "
                                f"{len(children)} sub-cells at zoom "
                                f"{children[0].zoom}[/]"
                            )
                            queue.extendleft(reversed(children))
                        else:
                            console.print(
                                "[red]  ↳ 120-result cap at minimum step: some "
                                "places in this cell may be missing.[/]"
                            )
                    elif stats.get('capped'):
                        console.print(
                            "[dim]  ↳ cap includes out-of-area fallback results; "
                            "not splitting this cell[/]"
                        )
        else:
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
    if grid:
        console.print(
            f"[dim]Grid: {len(grid.cells)} initial cells, step "
            f"{grid.step:g}°, {splits} cap split(s)[/]"
        )

    extra_cols = None
    if grid:
        for r in deduped:
            if r.get('_lat') is not None and r.get('_lng') is not None:
                r['latitude'] = round(float(r['_lat']), 6)
                r['longitude'] = round(float(r['_lng']), 6)
        extra_cols = ['latitude', 'longitude']

    if KEEP_ALL:
        kept, dropped, intent = deduped, [], None
        console.print("[dim]KEEP_ALL=1 → category + junk filter disabled[/]")
    else:
        kept, dropped, intent = filter_by_category(deduped, keyword)

    if intent:
        console.print(
            f"[bold]Filter:[/] intent='[cyan]{intent}[/]' → "
            f"[green]{len(kept)} kept[/], [red]{len(dropped)} filtered[/] "
            f"[dim](category allowlist + junk exclusion; KEEP_ALL=1 to disable)[/]"
        )
        if dropped:
            from collections import Counter
            top = Counter(r.get('categoryName') or '(none)' for r in dropped).most_common(8)
            console.print("[dim]top filtered categories:[/]")
            for c, n in top:
                console.print(f"  [red]✗[/] {c}: {n}")
    elif not KEEP_ALL:
        console.print(f"[dim]No category filter for keyword '{keyword}' (no intent matched)[/]")

    saved = write_xlsx(kept, dropped, out, extra_cols=extra_cols)
    # A completed run supersedes its checkpoint.
    try:
        out.with_stem(out.stem + '_partial').unlink(missing_ok=True)
    except Exception:
        pass
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


def _save_partial(verbose: bool = True) -> Path | None:
    """Write a checkpoint workbook so a crash never loses scraped rows."""
    if not (_partial_rows and _partial_out):
        return None
    partial_file = _partial_out.with_stem(_partial_out.stem + '_partial')
    rows = _partial_rows
    extra_cols = None
    if any(r.get('_lat') is not None and r.get('_lng') is not None
           for r in rows):
        rows = [
            dict(r,
                 latitude=round(float(r['_lat']), 6),
                 longitude=round(float(r['_lng']), 6))
            if r.get('_lat') is not None and r.get('_lng') is not None
            else r
            for r in rows
        ]
        extra_cols = ['latitude', 'longitude']
    try:
        write_xlsx(rows, [], partial_file, extra_cols=extra_cols)
        if verbose:
            console.print(
                f"[yellow]Partial save:[/] {len(rows)} rows → "
                f"[yellow]{partial_file}[/]"
            )
        return partial_file
    except Exception:
        return None


if __name__ == "__main__":
    ensure_chromium()
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        console.print("\n[yellow]Interrupted[/]")
        _save_partial()
        os._exit(130)
    except SystemExit as exc:
        _save_partial()
        code = exc.code if isinstance(exc.code, int) else 1
        os._exit(code)
    except EOFError:
        console.print("\n[red]No interactive input available.[/] "
                      "Use one-shot mode: [cyan]python scrape.py \"keyword\" \"place\"[/]")
        os._exit(1)
    # Hard exit to avoid playwright's 300s thread-join timeout on Windows
    os._exit(0)
