"""Diagnostic part 2: IndiaMART query variants."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from patchright.async_api import async_playwright

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")


async def probe(page, name: str, url: str) -> None:
    print(f"\n=== {name} : {url}")
    try:
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        print(f"  status: {resp.status if resp else 'n/a'}")
    except Exception as e:
        print(f"  GOTO ERR: {type(e).__name__}: {e}")
        return
    await page.wait_for_timeout(2000)
    try:
        nd = await page.evaluate(
            """() => { const el = document.querySelector('#__NEXT_DATA__');
                    return el ? JSON.parse(el.textContent) : null; }""")
    except Exception as e:
        print(f"  NEXT_DATA EVAL ERR: {e}")
        return
    pp = nd.get('props', {}).get('pageProps', {}) if nd else {}
    sr = pp.get('searchResponse') if nd else None
    print(f"  title: {(await page.title())[:60]}  searchResponse: {type(sr).__name__}")
    if isinstance(sr, dict):
        results = sr.get('results')
        if isinstance(results, list):
            print(f"    len(results)={len(results)}")
            for i, item in enumerate(results[:3]):
                f = item.get('fields', {}) if isinstance(item, dict) else {}
                print(f"      [{i}] company={f.get('companyname','?')[:45]!r} city={f.get('city','?')} title={str(f.get('title','?'))[:40]!r}")


async def main():
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir="C:\\Users\\Nomesh\\Desktop\\evaara\\_diag_profile",
            headless=False, locale="en-US", viewport={"width": 1400, "height": 900},
            user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for name, url in [
            ("IM-nocity", "https://dir.indiamart.com/search.mp?ss=mobile%20phones"),
            ("IM-dealers", "https://dir.indiamart.com/search.mp?ss=mobile+phone+dealers"),
            ("IM-access", "https://dir.indiamart.com/search.mp?ss=mobile+accessories"),
        ]:
            try:
                await asyncio.wait_for(probe(page, name, url), timeout=120)
            except Exception as e:
                print(f"  wrapper err: {type(e).__name__}: {e}")
        await ctx.close()


asyncio.run(main())