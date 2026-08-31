"""Test fresh profile (no cookies/history) against JD + IM."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
Object.defineProperty(navigator, 'plugins', { get: () => [{}] });
window.chrome = { runtime: {}, loadTimes: () => {}, csi: () => {} };
"""


def say(s):
    print(s, flush=True)


async def goto_retry(page, url, tries=2, timeout=90000):
    for i in range(tries):
        try:
            return await page.goto(url, wait_until="domcontentloaded", timeout=timeout), None
        except Exception as e:
            if i == tries - 1:
                return None, e
            await page.wait_for_timeout(2000)


async def probe_jd(page):
    say("\n## JD ##")
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    resp, err = await goto_retry(page, url)
    if err:
        say(f"GOTO ERR: {err}")
        return
    say(f"status={resp.status if resp else 'n/a'} url={page.url[:110]}")
    for _ in range(10):
        await page.wait_for_timeout(3000)
        try:
            ok = await page.evaluate("() => (document.body?document.body.innerText.includes('Ratings'):false)")
        except Exception:
            ok = False
        if ok:
            break
    info = await page.evaluate("""() => {
        const b = document.body ? document.body.innerText : '';
        const biz = Array.from(document.querySelectorAll('a'))
            .map(a => a.getAttribute('href')||'')
            .filter(h => /justdial\\.com\\/[A-Za-z]+\\/[\\w.-]+(\\?|$)/.test(h))
            .filter(h => !/nct-/.test(h));
        return {blen: b.length, title: document.title.slice(0,90), biz: [...new Set(biz)].slice(0,8),
                head: b.split('\\n').filter(l => l.trim()).slice(0, 22)};
    }""")
    say(f"ratings present: {ok}")
    say(f"title={info['title']!r} body_len={info['blen']}")
    say("biz: " + " | ".join(info['biz']))
    say("lines: " + " // ".join(info['head']))


async def probe_im(page):
    say("\n## IM ##")
    url = "https://dir.indiamart.com/search.mp?ss=mobile+phones&city=bidar"
    resp, err = await goto_retry(page, url)
    if err:
        say(f"GOTO ERR: {err}")
        return
    say(f"status={resp.status if resp else 'n/a'}")
    await page.wait_for_timeout(6000)
    nd = await page.evaluate("""() => { const el=document.querySelector('#__NEXT_DATA__');
        return el ? JSON.parse(el.textContent) : null; }""")
    pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
    sr = pp.get('searchResponse')
    if isinstance(sr, dict):
        say(f"searchResponse dict, results={len(sr.get('results') or [])}")
    else:
        say(f"searchResponse: {type(sr).__name__}")
    info = await page.evaluate("""() => {
        const b = document.body ? document.body.innerText : '';
        return {blen: b.length, rs: (b.match(/Rs/g)||[]).length,
                supp: document.querySelectorAll('[class*=\"supplier\"],[class*=\"listing\"]').length,
                lines: b.split('\\n').filter(l => l.trim()).slice(0, 18)};
    }""")
    say(f"body_len={info['blen']} Rs={info['rs']} supplier_els={info['supp']}")
    say("lines: " + " // ".join(info['lines']))


async def main():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_fresh",
            headless=False, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        await ctx.add_init_script(STEALTH_JS)
        try:
            await asyncio.wait_for(probe_jd(page), timeout=120)
        except Exception as e:
            say(f"JD err: {type(e).__name__}: {e}")
        try:
            await asyncio.wait_for(probe_im(page), timeout=90)
        except Exception as e:
            say(f"IM err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())