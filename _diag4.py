"""Diagnostic part 4: JustDial card DOM structure."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from playwright.async_api import async_playwright

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")


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
        await page.goto("https://www.justdial.com/Bidar/Mobile-Phone-Dealers",
                        wait_until="load", timeout=90000)
        await page.wait_for_timeout(6000)
        html = await page.evaluate("""() => {
            const cards = Array.from(document.querySelectorAll('[class*="jcard"], [class*="resultbox"], [class*="card "]'));
            if (!cards.length) {
                // fallback: dump a section of body
                const main = document.querySelector('main') || document.body;
                return 'NO CARDS; main html len=' + main.innerHTML.length;
            }
            const c = cards[0];
            return c.outerHTML.slice(0, 3000);
        }""")
        print(html[:3500])
        await ctx.close()


asyncio.run(main())