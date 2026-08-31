"""LIVE check: what does IndiaMART serve the scraper right now (bengaluru)?"""
import asyncio
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PROFILE = r"C:\Users\Nomesh\.cache\indiamart-scraper\profile"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")

URLS = [
    ("search-city", "https://dir.indiamart.com/search.mp?ss=mobile+accessories&city=bengaluru"),
    ("search-nocity", "https://dir.indiamart.com/search.mp?ss=mobile+accessories"),
    ("impcat", "https://dir.indiamart.com/impcat/mobile-phone-accessories.html"),
]

SELECTORS = [
    ".card", ".srp", ".lng__cont", ".listing", "[class*='supplier']",
    "[class*='product']", "[class*='result']", ".card--list",
]


async def main():
    from patchright.async_api import async_playwright

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=PROFILE, headless=True, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        for name, url in URLS:
            print(f"\n## {name}: {url}")
            try:
                resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            except Exception as e:
                print(f"  goto failed: {e}")
                continue
            await page.wait_for_timeout(3000)
            print(f"  status={resp.status if resp else 'n/a'} final={page.url}")

            info = await page.evaluate("""() => {
                const out = { next: null, ppKeys: null, srType: 'n/a', blen: 0, head: '' };
                const el = document.querySelector('#__NEXT_DATA__');
                if (el) {
                    out.next = true;
                    try {
                        const data = JSON.parse(el.textContent);
                        const pp = data.props && data.props.pageProps;
                        out.ppKeys = pp ? Object.keys(pp) : [];
                        const sr = pp && pp.searchResponse;
                        out.srType = sr === null ? 'None' : Array.isArray(sr) ? 'list' : typeof sr;
                        if (sr && typeof sr === 'object' && !Array.isArray(sr)) {
                            out.srResults = Array.isArray(sr.results) ? sr.results.length : 'n/a';
                            out.srKeys = Object.keys(sr).slice(0, 25);
                        }
                    } catch (e) { out.next = 'parse-error'; }
                }
                out.blen = document.body ? document.body.innerHTML.length : 0;
                out.head = (document.body ? document.body.innerText : '').slice(0, 300);
                return out;
            }""")
            for k, v in info.items():
                print(f"  {k}: {v}")

            counts = {}
            for sel in SELECTORS:
                try:
                    counts[sel] = await page.locator(sel).count()
                except Exception:
                    counts[sel] = 'err'
            print(f"  dom counts: {counts}")

            # visible links that look like supplier/product detail pages
            links = await page.evaluate("""() => {
                const set = new Set();
                document.querySelectorAll('a[href]').forEach(a => {
                    const h = a.href || '';
                    if (/indiamart\.com\/(proddetail|supplier)/.test(h)) set.add(h.slice(0, 120));
                });
                return Array.from(set).slice(0, 10);
            }""")
            print(f"  supplier/detail links: {links}")

        await ctx.close()


asyncio.run(main())
print("\nDONE")
