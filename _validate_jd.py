"""Validate JD DOM fallback extraction against the live site with retries."""
import asyncio
import importlib.util
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

spec = importlib.util.spec_from_file_location("jd_mod", r".scripts\jd.py")
jd = importlib.util.module_from_spec(spec)
spec.loader.exec_module(jd)


def say(s):
    print(s, flush=True)


async def one_attempt(page, url):
    try:
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        return {"err": f"{type(e).__name__}: {e}"}
    ok = False
    import time as _t
    deadline = _t.monotonic() + 40
    while _t.monotonic() < deadline:
        try:
            ok = await page.evaluate("() => (document.body && document.body.innerText || '').includes('Ratings')")
        except Exception:
            ok = False
        if ok:
            break
        try:
            await page.mouse.wheel(0, 1500)
        except Exception:
            pass
        await page.wait_for_timeout(1500)
    if not ok:
        bl = 0
        try:
            bl = await page.evaluate("() => (document.body ? document.body.innerText : '').length")
        except Exception:
            pass
        return {"empty": True, "body_len": bl}
    js_cards = await page.evaluate(jd.DOM_CARDS_JS)
    rows = jd.parse_dom_cards(js_cards)
    return {"rows": rows, "raw_count": len(js_cards)}


async def main():
    from playwright.async_api import async_playwright
    url = "https://www.justdial.com/Delhi/Mobile-Phone-Dealers"
    for attempt in range(5):
        say(f"\n===== attempt {attempt + 1} =====")
        prof = r"C:\Users\Nomesh\Desktop\evaara\_diag_valprof"
        shutil.rmtree(prof, ignore_errors=True)
        async with async_playwright() as p:
            ctx = await p.chromium.launch_persistent_context(
                user_data_dir=prof, headless=False, locale="en-US",
                viewport={"width": 1400, "height": 900}, user_agent=UA,
                args=["--disable-blink-features=AutomationControlled"],
                ignore_default_args=["--enable-automation"],
            )
            page = ctx.pages[0] if ctx.pages else await ctx.new_page()
            # try up to 2 navigations in the same session
            result = None
            for nav in range(2):
                result = await one_attempt(page, url)
                if result.get("empty"):
                    say(f"  nav{nav + 1}: empty shell (body_len={result['body_len']}) — retrying nav")
                    await page.wait_for_timeout(3000)
                    continue
                if "err" in result:
                    say(f"  nav{nav + 1}: err {result['err']}")
                    break
                say(f"  nav{nav + 1}: RENDERED, raw js_cards={result['raw_count']} rows={len(result['rows'])}")
                for r in result["rows"][:6]:
                    say(f"   - {r['name']!r} | {r['rating']} | {r['reviewsCount']} | "
                        f"{r['address'][:45]} | {r['categoryName']} | phone={r['phone']}")
                break
            await ctx.close()
            if result and "rows" in result and result["rows"]:
                say("\nVALIDATION PASSED")
                return
    say("\nVALIDATION FAILED — no rendered session in this window (JD intermittently walls)")


asyncio.run(main())