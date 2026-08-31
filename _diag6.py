"""Diagnostic part 6: DOM structure of JD cards + IM page for building fallback extractors."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")


def say(s):
    print(s, flush=True)


async def probe_jd(page):
    say("\n########## JD ##########")
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    try:
        resp = await page.goto(url, wait_until="load", timeout=90000)
        say(f"status {resp.status if resp else 'n/a'}")
    except Exception as e:
        say(f"GOTO ERR: {type(e).__name__}: {e}")
        return
    await page.wait_for_timeout(9000)
    say(f"title: {(await page.title())[:100]!r}")
    info = await page.evaluate("""() => {
        const sels = ['section.jcard', '.jcard', '[class*="resultbox"]',
                      '[class*="store-details"]', '[class*="result-card"]',
                      '[class*="list_result"]', '[class*="card"]'];
        const counts = {};
        for (const s of sels) counts[s] = document.querySelectorAll(s).length;
        const body = document.body ? document.body.innerText : '';
        const tel = document.querySelectorAll('a[href*="tel:"]').length;
        const wa = document.querySelectorAll('a[href*="wa.me"], [href*="api.whatsapp.com"]').length;
        const showNum = document.querySelectorAll('[class*="show-number"], [class*="ShowNumber"]').length;
        return {counts, blen: body.length, tel, wa, showNum, first: body.slice(0, 500)};
    }""")
    say(f"selectors: {info['counts']}")
    say(f"body_len={info['blen']} tel_links={info['tel']} whatsapp_links={info['wa']} showNumber_els={info['showNum']}")
    say("body head: " + repr(info['first']))
    # best selector -> dump first card html
    html = await page.evaluate("""() => {
        const sels = ['section.jcard', '.jcard', '[class*="resultbox"]',
                      '[class*="store-details"]', '[class*="result-card"]', '[class*="list_result"]'];
        let best = null, n = 0;
        for (const s of sels) {
            const els = document.querySelectorAll(s);
            if (els.length > n) { n = els.length; best = s; }
        }
        if (!best) return 'NO_CARD_SELECTOR';
        const el = document.querySelector(best);
        return 'SELECTOR=' + best + ' count=' + n + '\\n' + el.outerHTML.slice(0, 2800);
    }""")
    say(html)


async def probe_im(page):
    say("\n########## IndiaMART ##########")
    url = "https://dir.indiamart.com/search.mp?ss=mobile%20phones"
    captured = []

    def on_resp(resp):
        u = resp.url
        if ('indiamart' in u and ('search' in u.lower() or '.json' in u.lower() or 'api' in u.lower())):
            captured.append((resp.status, u[:160]))
    page.on('response', on_resp)
    try:
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        say(f"status {resp.status if resp else 'n/a'}")
    except Exception as e:
        say(f"GOTO ERR: {type(e).__name__}: {e}")
        page.remove_listener('response', on_resp)
        return
    await page.wait_for_timeout(8000)
    page.remove_listener('response', on_resp)
    nd = await page.evaluate(
        """() => { const el = document.querySelector('#__NEXT_DATA__');
                  return el ? JSON.parse(el.textContent) : null; }""")
    if nd is None:
        say("__NEXT_DATA__ MISSING")
    else:
        pp = nd.get('props', {}).get('pageProps', {})
        say(f"pageProps ALL keys ({len(pp)}): {list(pp.keys())}")
        sr = pp.get('searchResponse')
        say(f"searchResponse: {type(sr).__name__}")
        if isinstance(sr, dict):
            say(f"  sr keys: {list(sr.keys())[:20]}")
    info = await page.evaluate("""() => {
        const body = document.body ? document.body.innerText : '';
        const imLink = document.querySelectorAll('a[href*="indiamart.com"]').length;
        const imp = document.querySelectorAll('a[href*="/impcat/"]').length;
        const rs = document.body ? (document.body.innerText.match(/Rs/g) || []).length : 0;
        return {blen: body.length, imLink, imp, rs, first: body.slice(0, 400)};
    }""")
    say(f"body_len={info['blen']} im_links={info['imLink']} impcat_links={info['imp']} 'Rs' count={info['rs']}")
    say("body head: " + repr(info['first']))
    say("captured responses:")
    for s, u in captured[-12:]:
        say(f"  {s} {u}")


async def main():
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir="C:\\Users\\Nomesh\\Desktop\\evaara\\_diag_profile6",
            headless=False, locale="en-US", viewport={"width": 1400, "height": 900},
            user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        try:
            await asyncio.wait_for(probe_jd(page), timeout=180)
        except Exception as e:
            say(f"JD wrapper err: {type(e).__name__}: {e}")
        try:
            await asyncio.wait_for(probe_im(page), timeout=120)
        except Exception as e:
            say(f"IM wrapper err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())