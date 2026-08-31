"""Check what JD puts in raw __NEXT_DATA__ at domcontentloaded vs after hydration."""
import asyncio
import json
import re
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

PROFILE = r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_deep2"


def say(s):
    print(s, flush=True)


async def main():
    from playwright.async_api import async_playwright
    shutil.rmtree(PROFILE, ignore_errors=True)
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=PROFILE, headless=False, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
        raw_holder = {}

        async def on_resp(r):
            if r.url.startswith("https://www.justdial.com/Bidar") and r.request.is_navigation_request():
                try:
                    txt = await r.text()
                    raw_holder["html"] = txt
                except Exception:
                    pass
        page.on("response", on_resp)

        resp = await page.goto(url, wait_until="domcontentloaded", timeout=120000)
        say(f"status={resp.status if resp else 'n/a'}")
        # read __NEXT_DATA__ from RAW response html immediately
        raw_html = raw_holder.get("html", "")
        m = re.search(r'<script id="__NEXT_DATA__" type="application/json"[^>]*>(.*?)</script>', raw_html, re.S)
        say(f"raw html len={len(raw_html)}; __NEXT_DATA__ in response: {bool(m)}")
        if m:
            try:
                nd = json.loads(m.group(1))
                pp = nd.get("props", {}).get("pageProps", {})
                say(f"  pageProps keys: {list(pp.keys())[:30]}")
                ld = pp.get("listData")
                say(f"  listData type: {type(ld).__name__}")
                if isinstance(ld, dict):
                    res = ld.get("results", {})
                    if isinstance(res, dict):
                        say(f"    results columns={len(res.get('columns', []))} "
                            f"data items={len(res.get('data', {}))} total={ld.get('totalNumberofResults')}")
                    else:
                        say(f"    results type: {type(res).__name__}")
            except Exception as e:
                say(f"  json parse err: {e}")

        await page.wait_for_timeout(12000)
        info = await page.evaluate("""() => {
            const b = document.body ? document.body.innerText : '';
            const nd = document.querySelector('#__NEXT_DATA__');
            const cards = document.querySelectorAll('[class*="resultbox"]');
            return {blen: b.length, next_data_gone: nd === null,
                    ratings: b.includes('Ratings'),
                    first_names: (b.match(/(?:^|\\n)([A-Z][A-Za-z0-9 &.,'-]{8,60})(?:\\n)/g)||[]).slice(0, 8),
                    head: b.slice(0, 260)};
        }""")
        say(f"after hydration: body_len={info['blen']} __NEXT_DATA__ gone={info['next_data_gone']} "
            f"ratings={info['ratings']}")
        say("candidate names: " + " | ".join(x.strip() for x in info['first_names']))
        say("head: " + repr(info['head']))
        await ctx.close()
    say("\nDONE")


asyncio.run(main())