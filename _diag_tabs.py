"""Test JD card rendering: (A) direct goto, (B) same URL in a NEW tab, (C) homepage->target."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")


def say(s):
    print(s, flush=True)


async def goto_retry(page, url, tries=2, timeout=90000):
    for i in range(tries):
        try:
            return await page.goto(url, wait_until="domcontentloaded", timeout=timeout), None
        except Exception as e:
            if i == tries - 1:
                return None, e
            await page.wait_for_timeout(1500)


async def snap(page, label):
    await page.wait_for_timeout(8000)
    try:
        info = await page.evaluate("""() => {
            const b = document.body ? document.body.innerText : '';
            const biz = Array.from(document.querySelectorAll('a'))
                .map(a => a.getAttribute('href')||'')
                .filter(h => /justdial\\.com\\/[A-Za-z]+\\/[\\w.-]+(\\?|$)/.test(h))
                .filter(h => !/nct-/.test(h));
            return {blen: b.length, title: document.title.slice(0,70),
                    ratings: b.includes('Ratings'), biz: [...new Set(biz)].length,
                    head: b.split('\\n').filter(l=>l.trim()).slice(0,8)};
        }""")
    except Exception as e:
        say(f"  [{label}] eval err {e}")
        return
    say(f"[{label}] ratings={info['ratings']} body_len={info['blen']} " +
        f"title={info['title']!r} biz_links={info['biz']}")
    if info['head']:
        say(f"      lines: {' // '.join(info['head'])}")


async def main():
    from playwright.async_api import async_playwright
    target = "https://www.justdial.com/Bidar/Mobile-Phone-Dealers"
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=r"C:\Users\Nomesh\Desktop\evaara\_diag_profile_jd",
            headless=False, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        say("== A: direct goto on the existing page ==")
        await goto_retry(page, target)
        await snap(page, "A")

        say("\n== B: same URL in a brand-new tab ==")
        tab = await ctx.new_page()
        await goto_retry(tab, target)
        await snap(tab, "B")
        await tab.close()

        say("\n== C: homepage first, then new tab target ==")
        await goto_retry(page, "https://www.justdial.com/Bidar", tries=1)
        await page.wait_for_timeout(4000)
        tab2 = await ctx.new_page()
        await goto_retry(tab2, target)
        await snap(tab2, "C")
        await tab2.close()

        await ctx.close()
    say("\nDONE")


asyncio.run(main())