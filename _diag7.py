"""Diagnostic part 7: JD exact card markup + IM alternate endpoints."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from playwright.async_api import async_playwright

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")


def say(s):
    print(s, flush=True)


async def probe_jd_card(page):
    say("\n########## JD card markup ##########")
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    try:
        resp = await page.goto(url, wait_until="load", timeout=90000)
        say(f"status {resp.status if resp else 'n/a'}")
    except Exception as e:
        say(f"GOTO ERR: {type(e).__name__}: {e}")
        return
    await page.wait_for_timeout(9000)
    html = await page.evaluate("""() => {
        // find deepest element whose trimmed text starts with a known shop name
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
        let n = walker.nextNode();
        const hits = [];
        while (n) {
            const t = (n.textContent || '').trim();
            if (hits.length < 5 && /Sohel Mobile/.test(t) && n.children.length === 0) {
                // climb to a plausible card ancestor
                let a = n, depth = 0;
                while (a && depth < 6 && !/resultbox|jcard|card/i.test(a.className || '')) {
                    a = a.parentElement; depth++;
                }
                if (a && hits.indexOf(a) === -1) {
                    hits.push(a);
                    const cls = (a.className || '').toString();
                    return {
                        card_class: cls.slice(0, 300),
                        card_html: a.outerHTML.slice(0, 3000),
                        card_children: a.children.length,
                        tag: a.tagName,
                    };
                }
            }
            n = walker.nextNode();
        }
        return {hits: hits.length};
    }""")
    say("card_class: " + str(html.get('card_class')))
    say("tag=" + str(html.get('tag')) + " children=" + str(html.get('card_children')))
    say("--- card html ---")
    say(str(html.get('card_html'))[:3000])


async def probe_im2(page):
    say("\n########## IndiaMART alternate ##########")
    urls = [
        ("IM-cat", "https://dir.indiamart.com/impcat/mobile-phones.html"),
        ("IM-search-xhr", "https://dir.indiamart.com/search.mp?ss=mobile%20phones&v=4"),
    ]
    for name, url in urls:
        say(f"\n--- {name} : {url}")
        xhr = []
        handlers = []

        def make_handler():
            def h(r):
                if 'indiamart' in r.url:
                    xhr.append((r.status, r.url[:170]))
            return h
        h = make_handler()
        handlers.append(h)
        page.on('response', h)
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
            say(f"status {resp.status if resp else 'n/a'}")
        except Exception as e:
            say(f"GOTO ERR: {type(e).__name__}: {e}")
            for hh in handlers:
                page.remove_listener('response', hh)
            continue
        for _ in range(3):
            await page.mouse.wheel(0, 1200)
            await page.wait_for_timeout(1500)
        await page.wait_for_timeout(3000)
        nd = await page.evaluate(
            """() => { const el = document.querySelector('#__NEXT_DATA__');
                      return el ? JSON.parse(el.textContent) : null; }""")
        pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
        sr = pp.get('searchResponse')
        isq = pp.get('isqResponseData')
        say(f"title: {(await page.title())[:70]!r}")
        say(f"searchResponse: {type(sr).__name__}   isqResponseData: {type(isq).__name__}")
        if isinstance(sr, dict):
            say(f"  sr keys: {list(sr.keys())[:12]}")
        info = await page.evaluate("""() => {
            const body = document.body ? document.body.innerText : '';
            const cards = document.querySelectorAll('[class*="supplier"], [class*="listing"], [class*="result"]').length;
            const rs = (body.match(/Rs/g) || []).length;
            return {blen: body.length, cards, rs, first: body.slice(0, 300)};
        }""")
        say(f"body_len={info['blen']} result_els={info['cards']} 'Rs'={info['rs']}")
        say("body head: " + repr(info['first']))
        say("im responses:")
        for s, u in xhr[-6:]:
            say(f"  {s} {u}")
        for hh in handlers:
            page.remove_listener('response', hh)


async def main():
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir="C:\\Users\\Nomesh\\Desktop\\evaara\\_diag_profile7",
            headless=False, locale="en-US", viewport={"width": 1400, "height": 900},
            user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        try:
            await asyncio.wait_for(probe_jd_card(page), timeout=180)
        except Exception as e:
            say(f"JD wrapper err: {type(e).__name__}: {e}")
        try:
            await asyncio.wait_for(probe_im2(page), timeout=180)
        except Exception as e:
            say(f"IM wrapper err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())