"""Dump exactly what JD serves (raw HTML + console errors) on a brand-new profile."""
import asyncio
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

PROFILE = r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_deep"


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

        console_msgs, page_errors = [], []
        page.on("console", lambda m: console_msgs.append(f"{m.type}:{m.text[:200]}"))
        page.on("pageerror", lambda e: page_errors.append(str(e)[:300]))

        url = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
        try:
            resp = await page.goto(url, wait_until="load", timeout=120000)
            say(f"status={resp.status if resp else 'n/a'} final_url={page.url[:130]}")
            raw = await resp.text()
            say(f"raw html len={len(raw)}")
            say("RAW HEAD: " + repr(raw[:1200]))
            say("RAW has 'Sohel': " + str('Sohel' in raw))
            nd_in_raw = "__NEXT_DATA__" in raw
            say("RAW has __NEXT_DATA__ tag: " + str(nd_in_raw))
        except Exception as e:
            say(f"GOTO ERR: {type(e).__name__}: {e}")
            await ctx.close()
            return

        await page.wait_for_timeout(12000)
        info = await page.evaluate("""() => {
            const b = document.body ? document.body.innerText : '';
            const scripts = Array.from(document.scripts).map(s => (s.src||'inline').split('/').pop().slice(0,60));
            return {ready: document.readyState, html_len: document.documentElement.outerHTML.length,
                    blen: b.length, scripts: scripts.slice(0, 15), head: b.slice(0, 200)};
        }""")
        say(f"readyState={info['ready']} html_len={info['html_len']} body_len={info['blen']}")
        say("scripts: " + ", ".join(info['scripts']))
        say("body head: " + repr(info['head']))
        say("console msgs: " + " | ".join(console_msgs[-8:]))
        say("pageerrors: " + " | ".join(page_errors[-6:]))
        await ctx.close()
    say("\nDONE")


asyncio.run(main())