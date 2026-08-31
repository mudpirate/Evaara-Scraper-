"""Test IndiaMART URL variants for any that return actual searchResponse/results."""
import asyncio, shutil, sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
PROFILE = r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_im_var"


def say(s):
    print(s, flush=True)


async def test(page, name, url):
    say(f"\n-- {name}\n   {url}")
    try:
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        say(f"   GOTO ERR: {type(e).__name__}: {e}")
        return
    say(f"   status={resp.status if resp else 'n/a'}")
    await page.wait_for_timeout(5000)
    try:
        nd = await page.evaluate("""() => { const el=document.querySelector('#__NEXT_DATA__');
            return el ? JSON.parse(el.textContent) : null; }""")
    except Exception:
        nd = None
    pp = nd.get("props", {}).get("pageProps", {}) if nd else {}
    sr = pp.get("searchResponse")
    n = 0
    if isinstance(sr, dict):
        res = sr.get("results")
        if isinstance(res, list):
            n = len(res)
            for it in res[:2]:
                f = it.get("fields", {})
                say(f"     SR: {str(f.get('companyname'))[:38]!r} | {str(f.get('title'))[:34]!r} | {f.get('city')}")
    info = await page.evaluate("""() => {
        const b = document.body ? document.body.innerText : '';
        return {blen: b.length, rs: (b.match(/Rs/g)||[]).length,
                cards: document.querySelectorAll('[class*=\"lcard\"],[class*=\"supplier\"][class]').length,
                names: b.split('\\n').filter(l=>l.trim()&&l.trim().length>6).slice(0,8)};
    }""")
    say(f"   SR results={n} body_len={info['blen']} Rs={info['rs']} card-el-hits={info['cards']}")
    say("   head: " + " // ".join(info['names']))


async def main():
    from patchright.async_api import async_playwright
    shutil.rmtree(PROFILE, ignore_errors=True)
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=PROFILE, headless=False, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        urls = [
            ("broad+city",  "https://dir.indiamart.com/search.mp?ss=mobile+phones&city=bidar"),
            ("specific",    "https://dir.indiamart.com/search.mp?ss=usb+cables+manufacturer&city=bidar"),
            ("nocity",      "https://dir.indiamart.com/search.mp?ss=usb+cables"),
            ("v2",          "https://dir.indiamart.com/search.mp?ss=mobile+phones&v=2"),
            ("hyphen",      "https://dir.indiamart.com/search.mp?ss=mobile-phones"),
        ]
        for name, url in urls:
            try:
                await asyncio.wait_for(test(page, name, url), timeout=70)
            except Exception as e:
                say(f"   wrapped err: {type(e).__name__}: {e}")
        await ctx.close()
    say("\nDONE")


asyncio.run(main())