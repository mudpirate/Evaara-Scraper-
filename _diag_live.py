"""LIVE check: what do JustDial + IndiaMART actually serve the scraper today?

Mirrors the real scrapers' browser setup (persistent profile, same UA/stealth)
so this reflects a real run. Headless by default; pass `headed` to force.
"""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

HEADED = len(sys.argv) > 1 and sys.argv[1] == 'headed'

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


async def nd_eval(page):
    return await page.evaluate(
        """() => { const el = document.querySelector('#__NEXT_DATA__');
                  return el ? JSON.parse(el.textContent) : null; }"""
    )


async def goto_retry(page, url, tries=3):
    last = None
    for i in range(tries):
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            return resp, None
        except Exception as e:
            last = e
            say(f"  (goto attempt {i + 1} failed: {type(e).__name__}: {str(e)[:90]})")
            await page.wait_for_timeout(2500)
    return None, last


async def probe_jd(page):
    say("\n########## JUSTDIAL ##########")
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    resp, err = await goto_retry(page, url)
    if err:
        say(f"GOTO ERR (all attempts): {type(err).__name__}: {err}")
        return
    say(f"status: {resp.status if resp else 'n/a'}  url: {page.url[:120]}")
    await page.wait_for_timeout(12000)  # let client-side render settle

    say(f"title: {repr((await page.title()))[:110]}")
    nd = await nd_eval(page)
    if nd is None:
        say("__NEXT_DATA__: MISSING")
    else:
        pp = nd.get('props', {}).get('pageProps', {})
        say(f"__NEXT_DATA__: present, pageProps keys={len(pp)}")
        say(f"  keys: {list(pp.keys())[:25]}")
        ld = pp.get('listData')
        say(f"  listData type: {type(ld).__name__}")
        if isinstance(ld, dict):
            res = ld.get('results', {})
            if isinstance(res, dict):
                data = res.get('data', {})
                say(f"  results.columns={len(res.get('columns', []))} "
                    f"results.data items={len(data)} total={ld.get('totalNumberofResults')}")
            elif isinstance(res, list):
                say(f"  results is list, len={len(res)}")

    info = await page.evaluate("""() => {
        const body = document.body ? document.body.innerText : '';
        const links = Array.from(document.querySelectorAll('a[href*="justdial.com/"]'))
            .map(a => a.getAttribute('href')).filter(h => h);
        const names = links.map(h => decodeURIComponent((h||'').split('/').pop() || ''))
            .filter(n => n && !/[?#]/.test(n) && !/^nct-/.test(n));
        const cards = document.querySelectorAll('[class*="resultbox"],[class*="jcard"]').length;
        const showNum = (body.match(/Show Number/g) || []).length;
        const whatsapp = document.querySelectorAll('a[href*="wa.me"], a[href*="whatsapp.com"]').length;
        return {blen: body.length, cards, showNum, whatsapp, names: [...new Set(names)].slice(0, 12)};
    }""")
    say(f"body_len={info['blen']} resultbox/cards_els={info['cards']} "
        f"showNumber={info['showNum']} whatsapp_links={info['whatsapp']}")
    say("listing-ish links: " + ", ".join(info['names']))
    head = await page.evaluate("() => (document.body ? document.body.innerText : '').slice(0, 300)")
    say("body head: " + repr(head))


async def probe_im(page):
    say("")
    say("########## INDIAMART ##########")
    url = "https://dir.indiamart.com/search.mp?ss=mobile+phones&city=bidar"
    captured = []

    def on_resp(r):
        u = r.url
        if 'indiamart.com' in u and not any(x in u for x in
            ['analytics', 'google', 'doubleclick', 'facebook', '.gif', '.png', '.woff', 'svg']):
            captured.append((r.status, u[:200]))
    page.on("response", on_resp)
    resp, err = await goto_retry(page, url)
    if err:
        say(f"GOTO ERR (all attempts): {type(err).__name__}: {err}")
        page.remove_listener("response", on_resp)
        return
    say(f"status: {resp.status if resp else 'n/a'}  url: {page.url[:130]}")
    await page.wait_for_timeout(10000)
    page.remove_listener("response", on_resp)

    nd = await nd_eval(page)
    pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
    say(f"__NEXT_DATA__: present={nd is not None}  pageProps keys={len(pp)}")
    if pp:
        sr = pp.get('searchResponse')
        say(f"searchResponse: {type(sr).__name__}")
        if isinstance(sr, dict):
            results = sr.get('results')
            if isinstance(results, list):
                say(f"  results len={len(results)}")
                for it in results[:4]:
                    f = it.get('fields', {}) if isinstance(it, dict) else {}
                    say(f"   - {str(f.get('companyname'))[:40]!r} | {str(f.get('title'))[:40]!r} | {f.get('city')}")
        say(f"isqResponseData: {type(pp.get('isqResponseData')).__name__}")

    info = await page.evaluate("""() => {
        const body = document.body ? document.body.innerText : '';
        const links = Array.from(document.querySelectorAll('a[href*="indiamart.com"]'))
            .map(a => a.href).filter(h => /[A-Za-z0-9-]+$/.test(h));
        const cards = document.querySelectorAll('[class*="listing-card"], [class*="supplier"], [class*="result"]').length;
        return {blen: body.length, cards, links: [...new Set(links)].slice(0, 12), first: body.slice(0, 300)};
    }""")
    say(f"body_len={info['blen']} listing/supplier_els={info['cards']}")
    say("supplier-ish links: " + ", ".join(info['links']))
    say("body head: " + repr(info['first']))
    say("captured responses (first 10):")
    for s, u in captured[:10]:
        say(f"  {s} {u}")


async def main():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_jd",
            headless=not HEADED, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled", "--disable-http2"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        await ctx.add_init_script(STEALTH_JS)
        try:
            await asyncio.wait_for(probe_jd(page), timeout=120)
        except Exception as e:
            say(f"JD prob err: {type(e).__name__}: {e}")
        try:
            await asyncio.wait_for(probe_im(page), timeout=120)
        except Exception as e:
            say(f"IM prob err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())