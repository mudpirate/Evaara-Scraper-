# /// script
# requires-python = ">=3.11"
# dependencies = ["playwright", "rich", "openpyxl"]
# ///
"""
JustDial scraper — extracts hydrated result cards and verifies listing fields.

Zero setup: needs only Python 3.11+. Missing deps and Chromium are
auto-installed on first run. uv is optional.

Usage:
  python jd.py "mobile shops" "bangalore"
  python jd.py "mobile shops" "koramangala bangalore, jayanagar bangalore"
  python jd.py "restaurants" "mumbai" food.xlsx

Environment variables:
  HEADLESS=1       run without visible browser
  DELAY_MIN_MS     min jitter between page loads (default 800)
  DELAY_MAX_MS     max jitter between page loads (default 2500)
  MAX_PAGES        max pages per query (default 15, JD caps at ~15)
  PROFILE_DIR      persistent Chrome profile dir
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


# --- bootstrap ------------------------------------------------------------
def _bootstrap_deps() -> None:
    required = ['playwright', 'rich', 'openpyxl']
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
              f"  Or use uv:         uv run jd.py ...",
              file=sys.stderr, flush=True)
        sys.exit(1)
    print("[first-run] ✓ deps installed", file=sys.stderr, flush=True)


_bootstrap_deps()


# --- third-party ----------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
try:
    from playwright.async_api import Page, async_playwright
except ImportError:
    from patchright.async_api import Page, async_playwright
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt

console = Console()

# --- config ---------------------------------------------------------------
HEADLESS = os.getenv("HEADLESS", "0") == "1"
DELAY_MIN_MS = int(os.getenv("DELAY_MIN_MS", "800"))
DELAY_MAX_MS = int(os.getenv("DELAY_MAX_MS", "2500"))
MAX_PAGES = int(os.getenv("MAX_PAGES", "15"))
PROFILE_DIR = Path(os.getenv(
    "PROFILE_DIR",
    str(Path.home() / ".cache" / "justdial-scraper" / "profile"),
))
# Do not pin an old Chrome version. Akamai compares the user agent with the
# actual browser build and returns an empty HTML shell when they disagree.
USER_AGENT = os.getenv('JD_USER_AGENT', '')
BROWSER_CHANNEL = os.getenv('JD_BROWSER_CHANNEL', 'chrome')

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


_location_cache: dict[str, tuple[str, str | None]] = {}

_LOCATION_STOPWORDS = {
    'district', 'state', 'india', 'karnataka', 'maharashtra', 'delhi',
    'kerala', 'tamil', 'nadu', 'telangana', 'andhra', 'pradesh',
}
# Landmark/POI hints — JD's autocomplete returns hotels, hospitals, roads,
# etc. alongside the plain city. These are never valid search "areas".
_POI_HINTS = (
    'hotel', 'hospital', 'school', 'college', 'university', 'temple',
    'church', 'masjid', 'mosque', 'resort', 'mall', 'airport', 'railway',
    'station', 'road', 'street', 'circle', 'square', 'garden', 'park',
    'lodge', 'guest house', 'clinic', 'cinema', 'theatre', 'club', 'tower',
    'palace', 'fort', 'lake', 'dam', 'stadium', 'market', 'bridge',
    'marriott', 'taj ', 'leela', 'hospital', 'choultry', 'petrol',
)
_INDIAN_STATES = {
    'andhra pradesh', 'arunachal pradesh', 'assam', 'bihar', 'chhattisgarh',
    'goa', 'gujarat', 'haryana', 'himachal pradesh', 'jharkhand',
    'karnataka', 'kerala', 'madhya pradesh', 'maharashtra', 'manipur',
    'meghalaya', 'mizoram', 'nagaland', 'odisha', 'punjab', 'rajasthan',
    'sikkim', 'tamil nadu', 'telangana', 'tripura', 'uttar pradesh',
    'uttarakhand', 'west bengal', 'delhi',
}


def fallback_location(place: str) -> tuple[str, str | None]:
    """Infer city/area without treating a state or district as the city."""
    words = [w for w in re.findall(r"[A-Za-z]+", place) if w.lower() not in _LOCATION_STOPWORDS]
    if not words:
        return place.strip().title(), None
    city = words[-1].title()
    area = '-'.join(w.title() for w in words[:-1]) or None
    return city, area


async def resolve_location(page, place: str) -> tuple[str, str | None]:
    """Resolve user input to JD's (city, area) using their autocomplete API.

    Requires a browser page (playwright) to call the API with proper cookies.
    Returns (city, area) or falls back to manual parsing if API fails.
    """
    key = place.strip().lower()
    if key in _location_cache:
        return _location_cache[key]

    # Fallback: manual parsing
    fallback_city, fallback_area = fallback_location(place)

    try:
        # Navigate to JD homepage and use the autocomplete search box
        current = page.url or ''
        if 'justdial.com' not in current:
            await page.goto('https://www.justdial.com', wait_until='domcontentloaded')
            await page.wait_for_timeout(1500)

        city_input = page.locator('#city-auto-sug')
        if await city_input.count() > 0:
            await city_input.click()
            await city_input.fill('')
            await city_input.type(key, delay=80)
            await page.wait_for_timeout(2500)

            # Read autocomplete suggestions
            suggestions = await page.locator(
                '.city-sug-list li, .locationSug li, '
                '[class*="suggest"] li, [class*="autosug"] li'
            ).all()

            candidates = []
            query_tokens = [
                token for token in re.findall(r"[a-z]+", key)
                if token not in _LOCATION_STOPWORDS
            ]
            for index, s in enumerate(suggestions):
                text = (await s.inner_text()).strip()
                if not text or 'detect' in text.lower():
                    continue
                lower_text = text.lower()
                score = sum(3 for token in query_tokens if token in lower_text)
                if query_tokens and all(token in lower_text for token in query_tokens):
                    score += 100
                # Penalize POI-style suggestions (hotels, hospitals, roads, …)
                # so the plain city ("Bengaluru", "Bidar") wins the tie.
                if any(hint in lower_text for hint in _POI_HINTS):
                    score -= 50
                # Prefer the shortest matching suggestion — the canonical city
                # name is much shorter than "Bengaluru Marriott Hotel Whitefield".
                if score > 0:
                    score -= len(text) // 40
                candidates.append((score, -index, s, text))

            if not candidates or max(candidates)[0] <= 0:
                raise ValueError(f"no location suggestion matches '{place}'")

            _, _, selected, _ = max(candidates, key=lambda item: (item[0], item[1]))
            await selected.click()
            await page.wait_for_timeout(1000)
            new_val = (await city_input.input_value()).strip()
            if new_val:
                # JD commonly returns "Bidari, Bagalkot". For a state-level
                # result such as "Bidar, Karnataka", the left side is the city.
                if ',' in new_val:
                    left, right = [part.strip() for part in new_val.split(',', 1)]
                    if right.lower() in _INDIAN_STATES:
                        city, area = left, None
                    else:
                        city = right
                        area_text = left
                        # If the auto-filled "area" is really a landmark
                        # (hotel/hospital/road, …), JD suggested a POI, not a
                        # locality — drop the area and search the whole city.
                        if any(hint in area_text.lower() for hint in _POI_HINTS):
                            area = None
                        else:
                            area = '-'.join(w.title() for w in area_text.split())
                else:
                    city = new_val
                    area = None
                console.print(f"[dim]  JD location: {place} -> {city}"
                              f"{f' (area: {area})' if area else ''}[/]")
                _location_cache[key] = (city, area)
                return city, area
    except Exception as e:
        console.print(f"[dim]  JD location resolve failed for '{place}': {e}[/]")

    console.print(f"[dim]  JD location: using '{fallback_city}' as-is[/]")
    _location_cache[key] = (fallback_city, fallback_area)
    return fallback_city, fallback_area


def build_url(keyword: str, city: str, area: str | None = None, page_num: int = 1) -> str:
    """Build JustDial search URL from resolved city + area."""
    kw_slug = '-'.join(w.title() for w in keyword.strip().split())
    path = f"{kw_slug}-in-{area}" if area else kw_slug
    url = f"https://www.justdial.com/{city}/{path}"
    if page_num > 1:
        url += f"/page-{page_num}"
    return url


def place_label(keyword: str, place: str) -> str:
    parts = place.strip().split()
    if len(parts) == 1:
        return f"{keyword} in {parts[0].title()}"
    city = parts[-1].title()
    area = ' '.join(w.title() for w in parts[:-1])
    return f"{keyword} in {area}, {city}"


def slugify(text: str) -> str:
    s = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '_', s).strip('_')[:60]


# --- anti-bot -------------------------------------------------------------
async def jitter(page: Page) -> None:
    await page.wait_for_timeout(random.randint(DELAY_MIN_MS, DELAY_MAX_MS))


async def is_blocked(page: Page) -> bool:
    url = page.url or ''
    if any(x in url.lower() for x in ['/sorry/', 'blocked', 'captcha', 'challenge']):
        return True
    for sel in ['iframe[src*="captcha"]', '#challenge-form',
                'form[action*="challenge"]', '#px-captcha']:
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
            await page.screenshot(path="jd_blocked.png", full_page=True)
            console.print("[yellow]Screenshot:[/] jd_blocked.png")
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


# --- data extraction -------------------------------------------------------

# The current JustDial results page renders cards client-side and no longer
# puts listing data in `listData`. Keep the selectors in one browser-side
# expression so the Python parser only receives small, serializable records.
DOM_CARDS_JS = r"""
() => {
    const clean = (value) => (value || '').replace(/\s+/g, ' ').trim();
    const absolute = (href) => {
        try { return href ? new URL(href, location.origin).href : ''; }
        catch (_) { return ''; }
    };
    const phoneFrom = (value) => {
        const matches = clean(value).match(/(?:\+91[\s-]?)?0?[6-9]\d{9}/g) || [];
        return matches.length ? matches[0] : '';
    };
    const cards = Array.from(document.querySelectorAll(
        '#jdresult_box .resultbox, .results_listing_container .resultbox'
    ));

    return cards.map((card) => {
        const title = card.querySelector(
            '.resultbox_title_anchor, .resultbox_title_anchorbox, h2'
        );
        const address = card.querySelector(
            'address .locatcity, .resultbox_address .locatcity, ' +
            '.resultbox_address address'
        );
        const rating = card.querySelector('[aria-label^="Ratings"], .resultbox_totalrate');
        const reviewText = card.querySelector('.resultbox_countrate');
        const href = title && title.getAttribute('href') ||
            (card.querySelector('a[href*="_BZDET"]') || {}).href || '';
        const phoneLink = card.querySelector(
            'a[href^="tel:"], [data-phone], [data-number]'
        );
        const phoneValue = phoneLink && (
            phoneLink.getAttribute('href') ||
            phoneLink.getAttribute('data-phone') ||
            phoneLink.getAttribute('data-number') ||
            phoneLink.innerText
        ) || '';
        const amenities = Array.from(card.querySelectorAll('.amenities_tabs'))
            .map((el) => clean(el.innerText)).filter(Boolean);
        const addressText = clean(address && address.innerText);
        const text = clean(card.innerText);
        const pincode = (addressText.match(/\b[1-9]\d{5}\b/) || [''])[0];
        const cityParts = addressText.split(',').map(clean).filter(Boolean);

        return {
            name: clean(title && title.innerText),
            rating: clean(rating && (rating.getAttribute('aria-label') || rating.innerText))
                .replace(/^Ratings\s*:\s*/i, ''),
            reviewsText: clean(reviewText && reviewText.innerText),
            address: addressText,
            area: cityParts.length > 1 ? cityParts[cityParts.length - 2] : '',
            city: cityParts.length ? cityParts[cityParts.length - 1] : '',
            pincode,
            phone: phoneFrom(phoneValue) || phoneFrom(text),
            whatsapp: '',
            categoryName: amenities[0] || '',
            url: absolute(href),
            docid: card.id || card.getAttribute('keyid') || '',
        };
    }).filter((row) => row.name && row.address);
}
"""


def normalize_phone(value: object) -> str:
    """Return a phone only when it looks like a valid Indian mobile number."""
    raw = str(value or '').strip()
    digits = re.sub(r'\D', '', raw)
    if digits.startswith('91') and len(digits) == 12:
        digits = digits[2:]
    elif digits.startswith('0') and len(digits) == 11:
        digits = digits[1:]
    if len(digits) != 10 or digits[0] not in '6789':
        return ''
    return f'+91 {digits}'


def parse_dom_cards(raw_cards: list[dict]) -> list[dict]:
    """Normalize browser DOM card records and discard non-business cards."""
    rows = []
    for raw in raw_cards or []:
        if not isinstance(raw, dict):
            continue
        name = str(raw.get('name', '') or '').strip()
        address = str(raw.get('address', '') or '').strip()
        url = str(raw.get('url', '') or '').strip()
        if not name or not address or '/guides/' in url.lower():
            continue
        reviews_raw = str(raw.get('reviewsText', '') or '')
        review_match = re.search(r'([\d,]+)', reviews_raw)
        rating_raw = str(raw.get('rating', '') or '')
        rating_match = re.search(r'\d+(?:\.\d+)?', rating_raw)
        rows.append({
            'name': name,
            'rating': rating_match.group(0) if rating_match else '',
            'reviewsCount': int(review_match.group(1).replace(',', ''))
                if review_match else '',
            'address': address,
            'area': str(raw.get('area', '') or '').strip(),
            'city': str(raw.get('city', '') or '').strip(),
            'pincode': str(raw.get('pincode', '') or '').strip(),
            'phone': normalize_phone(raw.get('phone', '')),
            'whatsapp': normalize_phone(raw.get('whatsapp', '')),
            'categories': [str(raw.get('categoryName', '')).strip()]
                if raw.get('categoryName') else [],
            'categoryName': str(raw.get('categoryName', '') or '').strip(),
            'lat': '',
            'lon': '',
            'url': url,
            '_docid': str(raw.get('docid', '') or '').strip(),
        })
    return rows


def is_valid_listing(row: dict) -> bool:
    """Reject SEO, guide, and header rows returned by stale page payloads."""
    name = str(row.get('name', '') or '').strip()
    address = str(row.get('address', '') or '').strip()
    url = str(row.get('url', '') or '').lower()
    if not name or not address or '/guides/' in url:
        return False
    if name.lower().startswith(('mobile phone dealers in ', 'popular mobile phone dealers')):
        return False
    return bool('/' in url and ('_bzdet' in url or '/nct-' not in url))
def parse_next_data(raw: dict) -> tuple[list[dict], int | None]:
    """Extract listings from __NEXT_DATA__ JSON.

    JustDial (Next.js) stores results in columnar format:
      results.columns = ["docid", "name", "distance", ...]  (88 fields)
      results.data    = {"0": [...], "1": [...], ...}        (up to 10 per page)

    Returns (listings, total_count).
    """
    try:
        list_data = raw['props']['pageProps']['listData']
        if not isinstance(list_data, dict):
            return [], None
        results = list_data.get('results', {})
        columns = results.get('columns', [])
        data = results.get('data', {})
    except (KeyError, TypeError):
        return [], None

    total = None
    try:
        total = int(list_data.get('totalNumberofResults', 0)) or None
    except (TypeError, ValueError):
        pass

    if not columns or not data:
        return [], total

    # data can be a list of lists (common) or a dict of {"0": [...], ...}
    if isinstance(data, dict):
        items = [data[k] for k in sorted(data, key=lambda k: int(k) if str(k).isdigit() else 999)]
    elif isinstance(data, list):
        items = data
    else:
        return [], total

    listings = []
    for row in items:
        if not isinstance(row, list):
            continue
        entry = {}
        for i, col in enumerate(columns):
            if i < len(row):
                entry[col] = row[i]
        listings.append(entry)

    return listings, total


def extract_fields(entry: dict) -> dict:
    """Map JustDial's columnar fields to clean output dict."""
    # Reviews — "369 Ratings" or similar
    reviews_raw = str(entry.get('totalReviews', '') or '')
    m = re.search(r'([\d,]+)', reviews_raw)
    reviews = int(m.group(1).replace(',', '')) if m else None

    # Categories — comma-separated in `type`
    cats_raw = str(entry.get('type', '') or '')
    categories = [c.strip() for c in cats_raw.split(',') if c.strip()]

    # Phone (JD virtual number, plaintext in VNumber when supplied)
    phone = normalize_phone(
        entry.get('VNumber') or entry.get('phone') or entry.get('mobile')
    )

    # WhatsApp
    wp = entry.get('wpnumber')
    whatsapp = ''
    if isinstance(wp, list) and wp:
        whatsapp = str(wp[0])
    elif isinstance(wp, str) and wp:
        whatsapp = wp

    # Detail URL
    weburl = str(entry.get('weburl', '') or '').strip()
    if weburl and not weburl.startswith('/'):
        weburl = '/' + weburl
    url = f"https://www.justdial.com{weburl}" if weburl else ''

    return {
        'name': str(entry.get('name', '') or '').strip(),
        'rating': entry.get('compRating') or '',
        'reviewsCount': reviews or '',
        'address': str(
            entry.get('NewAddress') or entry.get('address') or
            entry.get('fullAddress') or ''
        ).strip(),
        'area': str(entry.get('area', '') or '').strip(),
        'city': str(entry.get('city', '') or '').strip(),
        'pincode': str(entry.get('pincode', '') or '').strip(),
        'phone': phone,
        'whatsapp': whatsapp,
        'categories': categories,
        'categoryName': categories[0] if categories else '',
        'lat': entry.get('lat') or '',
        'lon': entry.get('lon') or '',
        'url': url,
        '_docid': str(entry.get('docid', '') or '').strip(),
    }


async def wait_for_dom_cards(page: Page) -> bool:
    """Wait for hydrated result cards without treating an empty shell as data."""
    try:
        await page.wait_for_function(
            """() => document.querySelectorAll(
                '#jdresult_box .resultbox, .results_listing_container .resultbox'
            ).length > 0""",
            timeout=45000,
        )
        return True
    except Exception:
        return False


async def reveal_phone(page: Page, listing: dict) -> str:
    """Best-effort phone reveal; never use a number from another business card."""
    try:
        buttons = page.get_by_text('Show Number', exact=True)
        if await buttons.count():
            await buttons.first.click(timeout=8000)
            await page.wait_for_timeout(1200)
        name = str(listing.get('name', '') or '').lower()
        return await page.evaluate(
            """(businessName) => {
                const valid = (value) => {
                    const digits = (value || '').replace(/\\D/g, '');
                    const d = digits.startsWith('91') && digits.length === 12
                        ? digits.slice(2) : digits.startsWith('0') && digits.length === 11
                        ? digits.slice(1) : digits;
                    return /^[6-9]\\d{9}$/.test(d) ? '+91 ' + d : '';
                };
                const links = Array.from(document.querySelectorAll('a[href^="tel:"]'));
                for (const link of links) {
                    let node = link;
                    for (let depth = 0; node && depth < 7; depth++, node = node.parentElement) {
                        if ((node.innerText || '').toLowerCase().includes(businessName)) {
                            const phone = valid(link.getAttribute('href') || link.innerText);
                            if (phone) return phone;
                        }
                    }
                }
                return '';
            }""",
            name,
        )
    except Exception:
        return ''


async def enrich_phone(page: Page, listing: dict) -> str:
    """Open a listing only when its card did not expose a phone number."""
    url = str(listing.get('url', '') or '').strip()
    if listing.get('phone') or not url or '_BZDET' not in url.upper():
        return ''
    try:
        await page.goto(url, wait_until='domcontentloaded', timeout=45000)
        await jitter(page)
        phone = await reveal_phone(page, listing)
        if phone:
            listing['phone'] = phone
        return phone
    except Exception:
        return ''


# --- scraping -------------------------------------------------------------
async def scrape_query(page: Page, keyword: str, place: str) -> list[dict]:
    """Scrape real hydrated listing cards, with JSON as a legacy fallback."""
    label = place_label(keyword, place)
    console.print(f"[bold cyan]→[/] [yellow]{label}[/]")

    # Resolve location via JD's autocomplete API
    city, area = await resolve_location(page, place)

    all_listings: list[dict] = []
    seen_docids: set[str] = set()
    total_estimated: int | None = None
    page_num = 1
    dry_streak = 0

    while page_num <= MAX_PAGES:
        url = build_url(keyword, city, area, page_num)
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

        # Detect silent city redirect — JD sends unknown cities to Mumbai
        if page_num == 1:
            # Compare against the RESOLVED city (not raw user input)
            expected = city.lower().replace('-', '')
            final_url = page.url.lower()
            url_parts = final_url.replace('https://www.justdial.com/', '').split('/')
            actual_city = url_parts[0].replace('-', '') if url_parts else ''
            if actual_city and expected not in actual_city and actual_city not in expected:
                console.print(
                    f"[red]  ! JustDial redirected '{city}' -> '{actual_city}'[/]\n"
                    f"[yellow]  City not recognized by JustDial. Skipping.[/]"
                )
                return []

        listings: list[dict] = []
        if await wait_for_dom_cards(page):
            try:
                raw_cards = await page.evaluate(DOM_CARDS_JS)
                listings = parse_dom_cards(raw_cards)
            except Exception:
                listings = []

        # Older pages may still expose the columnar payload. It is only used
        # after DOM extraction and every row is validated before export.
        if not listings:
            try:
                raw = await page.evaluate("""
                    () => {
                        const el = document.querySelector('#__NEXT_DATA__');
                        return el ? JSON.parse(el.textContent) : null;
                    }
                """)
            except Exception:
                raw = None
            if raw:
                raw_listings, total = parse_next_data(raw)
                if page_num == 1 and total is not None:
                    total_estimated = total
                listings = [
                    fields for entry in raw_listings
                    if is_valid_listing(fields := extract_fields(entry))
                ]

        if not listings:
            if page_num == 1:
                console.print("[red]  ✗ no rendered business listings[/]")
            break

        new_count = 0
        for fields in listings:
            docid = fields.get('_docid', '')
            if docid and docid in seen_docids:
                continue
            if docid:
                seen_docids.add(docid)
            all_listings.append(fields)
            new_count += 1

        console.print(
            f"  [dim]page {page_num}: +{new_count} new"
            f"  ({len(all_listings)}"
            f"{f'/{total_estimated}' if total_estimated else ''}"
            f" unique)[/]"
        )

        if new_count == 0:
            dry_streak += 1
            if dry_streak >= 3:
                console.print(f"  [dim]3 pages with no new results — stopping early[/]")
                break
        else:
            dry_streak = 0

        if len(listings) < 10:
            break

        page_num += 1

    # JustDial normally hides phone numbers behind Show Number. Enrich only a
    # bounded number of rows and accept a number only when the detail page
    # associates it with the same business name.
    if os.getenv('ENRICH_PHONES', '1') == '1':
        try:
            limit = max(0, int(os.getenv('MAX_PHONE_ENRICH', '15')))
        except ValueError:
            limit = 15
        enriched = 0
        for listing in all_listings[:limit]:
            if not listing.get('phone') and await enrich_phone(page, listing):
                enriched += 1
        if enriched:
            console.print(f"  [dim]revealed {enriched} verified mobile number(s)[/]")

    console.print(
        f"[bold green]✓[/] {len(all_listings)} listings "
        f"({page_num} page{'s' if page_num > 1 else ''})"
    )

    return all_listings


def dedupe(rows: list[dict]) -> list[dict]:
    """Dedupe by docid (JustDial's unique listing identifier)."""
    seen: set[str] = set()
    out = []
    for r in rows:
        key = r.get('_docid', '')
        if not key:
            key = f"{r.get('name', '').lower().strip()}|{r.get('lat', '')}|{r.get('lon', '')}"
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


# --- output ---------------------------------------------------------------
def write_xlsx(rows: list[dict], out: Path) -> None:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        'name', 'rating', 'reviewsCount', 'address', 'area', 'city',
        'pincode', 'phone', 'whatsapp', 'categoryName', 'url',
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "listings"
    ws.append(cols)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', start_color='DDDDDD')
        cell.alignment = Alignment(horizontal='center')

    for r in rows:
        ws.append([r.get(c, '') for c in cols])

    widths = {
        'A': 36, 'B': 8, 'C': 12, 'D': 42, 'E': 22, 'F': 14,
        'G': 10, 'H': 16, 'I': 16, 'J': 32, 'K': 50,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    wb.save(out)


# --- CLI ------------------------------------------------------------------
def get_inputs() -> tuple[str, list[str], Path]:
    """Interactive mode."""
    console.print(Panel.fit(
        "[bold cyan]JustDial Scraper[/]\n"
        "[dim]Extract real business cards with names, addresses, and phones[/]",
        border_style="cyan",
    ))

    keyword = Prompt.ask(
        "\n[bold]Category[/] "
        "[dim](e.g. mobile shops, restaurants, salons)[/]"
    ).strip()

    place_input = Prompt.ask(
        "[bold]City or areas[/] "
        "[dim](e.g. bangalore, or koramangala bangalore, jayanagar bangalore)[/]"
    ).strip()

    places = [p.strip() for p in place_input.split(',') if p.strip()]

    default_out = f"jd_{slugify(keyword)}_{slugify(places[0])}.xlsx"
    out_str = Prompt.ask("[bold]Output file[/]", default=default_out).strip()

    console.print()
    for p in places:
        console.print(f"  [cyan]•[/] {place_label(keyword, p)}")
    console.print(f"[dim]→ saving to[/] [yellow]{out_str}[/]\n")

    out = Path(out_str)
    if out.suffix.lower() != '.xlsx':
        out = out.with_suffix('.xlsx')
    return keyword, places, out


def parse_args() -> tuple[str, list[str], Path]:
    """One-shot: jd.py "keyword" "place1, place2" [out.xlsx]."""
    keyword = sys.argv[1]
    place_input = sys.argv[2]
    places = [p.strip() for p in place_input.split(',') if p.strip()]

    if len(sys.argv) > 3:
        out = Path(sys.argv[3])
    else:
        out = Path(f"jd_{slugify(keyword)}_{slugify(places[0])}.xlsx")
    if out.suffix.lower() != '.xlsx':
        out = out.with_suffix('.xlsx')
    return keyword, places, out


async def main():
    if len(sys.argv) == 1:
        keyword, places, out = get_inputs()
    elif len(sys.argv) >= 3:
        keyword, places, out = parse_args()
    else:
        console.print("[bold]Usage:[/]")
        console.print(
            "  python jd.py                                       "
            "[dim](interactive)[/]"
        )
        console.print('  python jd.py "keyword" "city"')
        console.print(
            '  python jd.py "keyword" "area1 city, area2 city" '
            "[out.xlsx]"
        )
        console.print("  [dim](or: uv run jd.py ...)[/]")
        sys.exit(1)

    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    console.print(f"[dim]profile: {PROFILE_DIR}[/]")

    async with async_playwright() as p:
        launch_options = {
            'user_data_dir': str(PROFILE_DIR),
            'headless': HEADLESS,
            'locale': 'en-US',
            'viewport': {'width': 1400, 'height': 900},
            'args': ['--disable-blink-features=AutomationControlled'],
            'ignore_default_args': ['--enable-automation'],
        }
        if USER_AGENT:
            launch_options['user_agent'] = USER_AGENT
        if BROWSER_CHANNEL:
            launch_options['channel'] = BROWSER_CHANNEL
        try:
            ctx = await p.chromium.launch_persistent_context(**launch_options)
        except Exception:
            if not BROWSER_CHANNEL:
                raise
            console.print(
                f"[dim]Chrome channel unavailable; falling back to bundled Chromium[/]"
            )
            launch_options.pop('channel', None)
            ctx = await p.chromium.launch_persistent_context(**launch_options)
        await ctx.add_init_script(STEALTH_JS)
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        global _partial_rows, _partial_out
        _partial_out = out
        all_rows: list[dict] = []
        for place in places:
            try:
                rows = await scrape_query(page, keyword, place)
                all_rows.extend(rows)
                _partial_rows = list(all_rows)
            except SystemExit:
                raise
            except Exception as e:
                console.print(f"[red]✗[/] {place}: {e}")

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
            'Use one-shot: [cyan]python jd.py "keyword" "city"[/]'
        )
        os._exit(1)
    os._exit(0)
