"""Deeper structure probe for JD cards + IM category/API paths. Saves HTML to files."""
import asyncio
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


async def goto_retry(page, url, tries=2, wait_until="domcontentloaded", timeout=90000):
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


async def wait_for_text(page, needle, timeout_ms=25000):
    try:
        await page.wait_for_function(
            "t => (document.body ? document.body.innerText.includes(t) : false)",
            arg=needle, timeout=timeout_ms)
        return True
    except Exception:
        return False


async def probe_jd(page):
    say("\n########## JD: card markup ##########")
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    resp, err = await goto_retry(page, url)
    if err:
        say(f"GOTO ERR: {err}")
        return
    say(f"status={resp.status if resp else 'n/a'} url={page.url[:130]}")
    ok = await wait_for_text(page, "Ratings", 30000)
    say(f"waited for 'Ratings' in body: {ok}")
    await page.wait_for_timeout(5000)
    for _ in range(5):
        await page.mouse.wheel(0, 1800)
        await page.wait_for_timeout(600)

    say(f"title: {(await page.title())[:100]!r}")
    has_nd = await page.evaluate("() => !!document.querySelector('#__NEXT_DATA__')")
    say(f"__NEXT_DATA__ present: {has_nd}")

    result = await page.evaluate("""
      () => {
        const anchors = Array.from(document.querySelectorAll('a'));
        const biz = anchors
          .map(a => a.getAttribute('href')||'')
          .filter(h => /justdial\\.com\\/\\w+\\/[\\w.-]+(\\?|$)/.test(h) || /^\\/[A-Z][a-z]+\\/[\\w.-]+(\\?|$)/.test(h))
          .filter(h => !/nct-/.test(h));
        const out = { biz_hrefs: [...new Set(biz)].slice(0, 10) };
        const els = Array.from(document.querySelectorAll('a,h2,h3,h4,div,[class*="resultbox"]'))
          .filter(e => e.children.length < 6 && /Ratings/.test(e.textContent||''));
        let card = null;
        if (els.length) {
          let el = els[0]; card = el;
          for (let i=0;i<10 && card;i++){ if((card.className||'').toString().includes('card')) break; card=card.parentElement; }
        }
        out.card_html = (card ? card.outerHTML : 'NO_CARD').slice(0, 9000);
        out.card_class = card ? ((card.className||'').toString().slice(0,200)) : '';
        out.body_head = (document.body?document.body.innerText:'').slice(0,500);
        return out;
      }
    """)
    say("biz hrefs: " + " | ".join(result.get('biz_hrefs', [])))
    say(f"card_class: {result['card_class']}")
    with open("_jd_card.html", "w", encoding="utf-8") as f:
        f.write(result["card_html"])
        f.write("\n\n=== BODY HEAD ===\n")
        f.write(result["body_head"])
    say("wrote _jd_card.html len=" + str(len(result["card_html"])))


async def probe_im(page):
    say("\n== IM: homepage-then-search + impcat ##########")
    r, _ = await goto_retry(page, "https://www.indiamart.com", tries=1)
    say(f"homepage status={r.status if r else 'n/a'}")
    await page.wait_for_timeout(2500)
    resp, err = await goto_retry(page, "https://dir.indiamart.com/search.mp?ss=mobile+phones&city=bidar")
    if not err:
        say(f"search-after-home status={resp.status if resp else 'n/a'}")
        await page.wait_for_timeout(3000)
        nd = await page.evaluate("""() => { const el=document.querySelector('#__NEXT_DATA__');
            return el ? JSON.parse(el.textContent) : null; }""")
        pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
        sr = pp.get('searchResponse')
        if isinstance(sr, dict):
            say(f"  searchResponse dict, results={len(sr.get('results') or [])}")
        else:
            say(f"  searchResponse: {type(sr).__name__}  (not dict after homepage)")
        info = await page.evaluate("""() => {
            const b = document.body?document.body.innerText:'';
            return {blen:b.length, rs:(b.match(/Rs/g)||[]).length, head:b.slice(0,160)};
        }""")
        say(f"  body_len={info['blen']} Rs={info['rs']} head={info['head']!r}")
    resp, err = await goto_retry(page, "https://dir.indiamart.com/impcat/mobile-phones.html")
    if not err:
        say(f"impcat status={resp.status if resp else 'n/a'}")
        await page.wait_for_timeout(4000)
        nd = await page.evaluate("""() => { const el=document.querySelector('#__NEXT_DATA__');
            return el ? JSON.parse(el.textContent) : null; }""")
        pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
        sr = pp.get('searchResponse')
        say(f"impcat searchResponse: {type(sr).__name__}")
        if isinstance(sr, dict):
            res = sr.get('results')
            if isinstance(res, list) and res:
                f0 = res[0].get('fields', {})
                say(f"impcat sr results={len(res)}; sample: {str(f0)[:700]}")
        html = await page.evaluate("""
          () => {
            const els = Array.from(document.querySelectorAll('[class*="supplier"],[class*="listing"],[class*="lcard"],[class*="result"]'))
              .filter(e => e.children.length > 2);
            if (!els.length) return 'NO_SUPPLIER_CARDS';
            return els[0].outerHTML.slice(0, 5000);
          }
        """)
        with open("_im_card.html", "w", encoding="utf-8") as f:
            f.write(html)
        say("wrote _im_card.html len=" + str(len(html)))
        bh = await page.evaluate("() => (document.body?document.body.innerText:'').slice(0,400)")
        say("impcat body head: " + repr(bh))
    else:
        say(f"impcat GOTO ERR: {err}")


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
            await asyncio.wait_for(probe_jd(page), timeout=150)
        except Exception as e:
            say(f"JD err: {type(e).__name__}: {e}")
        try:
            await asyncio.wait_for(probe_im(page), timeout=150)
        except Exception as e:
            say(f"IM err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())