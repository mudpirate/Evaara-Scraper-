"""Diagnostic part 5: JustDial DOM structure (raw)."""
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
        info = await page.evaluate("""() => {
            const out = {};
            out.doctags = document.documentElement.outerHTML.slice(0, 200);
            // find any element containing 'Sohel Mobile'
            const nodes = [];
            document.querySelectorAll('*').forEach((el) => {
                if (el.children.length === 0 && /Sohel Mobile/.test(el.textContent || '')) {
                    nodes.push(el.tagName + '.' + (el.className || ''));
                }
            });
            out.matches = nodes.slice(0, 6);
            // find parent chain
            const it = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
            let n = it.nextNode(), parent = null;
            while (n) {
                if ((n.textContent || '').trim().startsWith('Sohel Mobile')) { parent = n; break; }
                n = it.nextNode();
            }
            out.parent_html = parent ? parent.outerHTML.slice(0, 2600) : 'NOT FOUND';
            return out;
        }""")
        print("doctags:", info["doctags"])
        print("matches:", info["matches"])
        print("--- parent outerHTML ---")
        print(info["parent_html"])
        await ctx.close()


asyncio.run(main())