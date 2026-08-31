"""Inspect current JustDial + IndiaMART page structure for building DOM/API fallbacks."""
import asyncio
import json
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")

STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
Object.defineProperty(navigator, 'plugins', { get: () => [{}] });
window.chrome = { runtime: {} };
"""


def say(s):
    print(s, flush=True)


async def goto_retry(page, url, tries=3, wait_until="domcontentloaded", timeout=90000):
    last = None
    for i in range(tries):
        try:
            resp = await page.goto(url, wait_until=wait_until, timeout=timeout)
            return resp, None
        except Exception as e:
            last = e
            say(f"  (goto attempt {i + 1} failed: {type(e).__name__}: {str(e)[:80]})")
            await page.wait_for_timeout(2000)
    return None, last


def json_maybe(text):
    try:
        return json.loads(text)
    except Exception:
        return None


async def probe_jd(page):
    say("\n########## JUSTDIAL ##########")
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    captured = []

    def on_resp(r):
        u = r.url
        if 'justdial' in u and not any(x in u for x in
            ['.css', '.js?', '.png', '.jpeg', '.jpg', '.gif', '.svg', '.woff', '.webp', 'analytics', 'bazaarvoice']):
            captured.append((r.status, r.url[:220]))
    page.on("response", on_resp)

    resp, err = await goto_retry(page, url)
    if err:
        say(f"GOTO ERR: {err}")
        page.remove_listener("response", on_resp)
        return
    say(f"status={resp.status if resp else 'n/a'} final_url={page.url[:150]}")
    await page.wait_for_timeout(8000)
    # scroll to trigger lazy loading
    for _ in range(4):
        await page.mouse.wheel(0, 2000)
        await page.wait_for_timeout(700)

    say(f"title: {(await page.title())[:110]!r}")
    nd = await page.evaluate("""() => { const el = document.querySelector('#__NEXT_DATA__');
                return el ? JSON.parse(el.textContent) : null; }""")
    say(f"__NEXT_DATA__: {'present' if nd else 'MISSING'}")

    # structure: find anchors to business profile pages
    out = await page.evaluate("""
      () => {
        const out = {};
        const as = Array.from(document.querySelectorAll('a[href*="justdial.com/"]'))
            .map(a => a.getAttribute('href') || '')
            .filter(h => /justdial\\.com\\/[^/]+\\/[^/]+/.test(h) && !/page-\\d/.test(h) && !/nct-/.test(h));
        out.listing_hrefs = [...new Set(as.slice(0, 8))];
        // pick the first business anchor and climb to a plausible card
        const first = document.querySelector('a[href*="justdial.com/"]');
        let card = first, depth = 0, best;
        while (card && depth < 8) {
            const cls = (card.className || '').toString();
            if (cls && !/jsx/.test(cls) && (cls.includes('card') || cls.includes('result') || cls.includes('store') || cls.includes('jcard'))) {
                best = card;
            }
            card = card.parentElement; depth++;
        }
        let card_html = '', card_class = '';
        if (best) { card_class = best.className; card_html = best.outerHTML; }
        // fallback: element containing the first restaurant-ish family
        return { out, card_class, card_html: card_html.slice(0, 6000), body_tail: (document.body.innerText || '').slice(0, 700) };
      }
    """)
    say("sample listing hrefs: " + "\n  ".join(out["listing_hrefs"]))
    say(f"CARD class: {out['card_class'][:150]}")
    say("---- CARD HTML ----")
    say(out["card_html"])
    say("---- BODY TAIL ----")
    say(out["body_tail"])
    say("---- captured (justdial) ----")
    for s, u in captured[:15]:
        say(f"  {s} {u}")
    page.remove_listener("response", on_resp)


async def probe_im(page):
    say("\n########## INDIAMART ##########")
    urls = [
        ("search-mobile", "https://dir.indiamart.com/search.mp?ss=mobile+accessories&city=bidar"),
        ("search-mobile-v4", "https://dir.indiamart.com/search.mp?ss=mobile+accessories&city=bidar&v=4"),
        ("impcat", "https://dir.indiamart.com/impcat/mobile-phones.html"),
        ("search-nocity", "https://dir.indiamart.com/search.mp?ss=mobile+phones"),
    ]
    json_bodies = []

    async def on_resp(r):
        u = r.url
        if 'indiamart' not in u:
            return
        if any(x in u for x in ['analytics', 'google', 'doubleclick', 'facebook', '.gif', '.png', '.woff', '.svg', '.css', '.js?', 'components']):
            return
        ct = r.headers.get('content-type', '')
        if 'json' in ct or '.json' in u or 'api' in u.lower():
            try:
                txt = await r.text()
                if len(txt) > 80 and any(k in txt for k in ('companyname', 'displayid', 'searchResponse', 'supplier')):
                    json_bodies.append(u[:200] + "|||" + txt[:1500])
            except Exception:
                pass
    page.on("response", on_resp)

    for name, url in urls:
        say(f"\n--- {name} : {url}")
        resp, err = await goto_retry(page, url)
        if err:
            say(f"  GOTO ERR: {err}")
            continue
        say(f"  status={resp.status if resp else 'n/a'} final_url={page.url[:150]}")
        for _ in range(3):
            await page.mouse.wheel(0, 1500)
            await page.wait_for_timeout(900)
        nd = await page.evaluate("""() => { const el = document.querySelector('#__NEXT_DATA__');
                    return el ? JSON.parse(el.textContent) : null; }""")
        pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
        sr = pp.get('searchResponse')
        say(f"  searchResponse: {type(sr).__name__}")
        if isinstance(sr, dict):
            results = sr.get('results')
            if isinstance(results, list):
                say(f"    results={len(results)}")
                for it in results[:3]:
                    f = it.get('fields', {}) if isinstance(it, dict) else {}
                    say(f"      {str(f.get('companyname'))[:42]!r} | {str(f.get('title'))[:36]!r} | {f.get('city')}")
        info = await page.evaluate("""() => {
            const body = document.body ? document.body.innerText : '';
            const supp = document.querySelectorAll('[class*="supplier"],[class*="listing"]').length;
            const rs = (body.match(/Rs/g) || []).length;
            return {blen: body.length, supp, rs, head: body.slice(0, 180)};
        }""")
        say(f"  body_len={info['blen']} supplier_els={info['supp']} 'Rs'={info['rs']}")
        say(f"  head: {info['head']!r}")

    page.remove_listener("response", on_resp)
    say("\n-- JSON bodies captured --")
    for b in json_bodies[:6]:
        say(b)


async def main():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_jd",
            headless=False, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        await ctx.add_init_script(STEALTH_JS)
        try:
            await asyncio.wait_for(probe_jd(page), timeout=180)
        except Exception as e:
            say(f"JD prob err: {type(e).__name__}: {e}")
        try:
            await asyncio.wait_for(probe_im(page), timeout=240)
        except Exception as e:
            say(f"IM prob err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())