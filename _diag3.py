"""Diagnostic part 3: JustDial rendered DOM inspect."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from playwright.async_api import async_playwright

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")


async def probe(page):
    url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    print(f"GET {url}")
    try:
        resp = await page.goto(url, wait_until="load", timeout=90000)
        print(f"status {resp.status if resp else 'n/a'}")
    except Exception as e:
        print(f"GOTO ERR: {type(e).__name__}: {e}")
        return
    await page.wait_for_timeout(8000)
    print(f"title: {(await page.title())[:120]!r}")
    print(f"url: {page.url[:160]}")
    info = await page.evaluate("""() => {
        const body = document.body ? document.body.innerText : '';
        const els = {
          links_justdial: document.querySelectorAll('a[href*="justdial.com"]').length,
          store_blocks: document.querySelectorAll('[class*="card"], [class*="store"], [class*="resultbox"], .jcard').length,
          phone_els: document.querySelectorAll('[class*="phone"], [class*="tel"]').length,
        };
        return {len: body.length, first: body.slice(0, 600), els: els};
    }""")
    print(f"  body_len={info['len']}  els={info['els']}")
    print("  body head:", repr(info['first']))


async def main():
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir="C:\\Users\\Nomesh\\Desktop\\evaara\\_diag_profile_jd",
            headless=False, locale="en-US", viewport={"width": 1400, "height": 900},
            user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        try:
            await asyncio.wait_for(probe(page), timeout=150)
        except Exception as e:
            print(f"wrapper err: {type(e).__name__}: {e}")
        await ctx.close()


asyncio.run(main())