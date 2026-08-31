"""Test IndiaMART with a FRESH profile (was always None in warm profiles)."""
import asyncio, shutil, sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
PROFILE = r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_im_fresh"


def say(s):
    print(s, flush=True)


async def goto_retry(page, url, tries=3, timeout=90000):
    for i in range(tries):
        try:
            return await page.goto(url, wait_until="domcontentloaded", timeout=timeout), None
        except Exception as e:
            if i == tries - 1:
                return None, e
            await page.wait_for_timeout(2000)


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
        url = "https://dir.indiamart.com/search.mp?ss=mobile+phones&city=bidar"
        raw_html = {}

        async def on_resp(r):
            if 'search.mp' in r.url and r.request.is_navigation_request():
                try:
                    raw_html["t"] = await r.text()
                except Exception:
                    pass
        page.on("response", on_resp)
        resp, err = await goto_retry(page, url)
        if err:
            say(f"GOTO ERR: {err}")
        else:
            say(f"status={resp.status if resp else 'n/a'}")
        await page.wait_for_timeout(8000)
        raw = raw_html.get("t", "")
        say(f"raw html len={len(raw)}; has 'searchResponse': {bool(raw) and 'searchResponse' in raw}")
        if raw:
            import re
            m = re.search(r'<script id="__NEXT_DATA__" type="application/json"[^>]*>(.*?)</script>', raw, re.S)
            say(f"__NEXT_DATA__ in raw: {bool(m)}")
            if m:
                import json
                try:
                    nd = json.loads(m.group(1))
                    pp = nd.get("props", {}).get("pageProps", {})
                    sr = pp.get("searchResponse")
                    say(f"  searchResponse: {type(sr).__name__}")
                    if isinstance(sr, dict):
                        results = sr.get("results")
                        if isinstance(results, list):
                            say(f"    results={len(results)}")
                            for it in results[:3]:
                                f = it.get("fields", {})
                                say(f"      {str(f.get('companyname'))[:40]!r} | {str(f.get('title'))[:36]!r}")
                    keys = [k for k in pp.keys()]
                    say(f"  pageProps keys: {keys[:20]}")
                except Exception as e:
                    say(f"  json err: {e}")
        info = await page.evaluate("""() => {
            const b = document.body ? document.body.innerText : '';
            return {blen: b.length, rs: (b.match(/Rs/g)||[]).length,
                    supp: document.querySelectorAll('[class*=\"supplier\"],[class*=\"listing\"]').length,
                    head: b.split('\\n').filter(l=>l.trim()).slice(0,16)};
        }""")
        say(f"body_len={info['blen']} Rs={info['rs']} supplier_els={info['supp']}")
        say("lines: " + " // ".join(info['lines']))
        await ctx.close()
    say("\nDONE")


asyncio.run(main())