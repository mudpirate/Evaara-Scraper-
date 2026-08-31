"""Dump one full IndiaMART impcat product card + all cards' structure."""
import asyncio
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

PROFILE = r"C:\Users\Nomesh\.cache\indiamart-scraper\profile"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")


async def main():
    from patchright.async_api import async_playwright

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=PROFILE, headless=True, locale="en-US",
            viewport={"width": 1400, "height": 900}, user_agent=UA,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        url = "https://dir.indiamart.com/impcat/mobile-phone-accessories.html"
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(3000)

        # Full innerText of the first 3 cards
        cards_text = await page.evaluate("""() => {
            const cards = document.querySelectorAll('article.template7-product-card, [class*="template7-product-card"]');
            return Array.from(cards).slice(0, 3).map(c => (c.innerText || '').trim().slice(0, 500));
        }""")
        for i, t in enumerate(cards_text):
            print(f"\n--- card {i} innerText ---\n{t}")

        # Classes + short text inside first card (to locate company/city/rating)
        sub = await page.evaluate("""() => {
            const card = document.querySelector('article.template7-product-card, [class*="template7-product-card"]');
            if (!card) return [];
            const out = [];
            card.querySelectorAll('[class]').forEach(el => {
                const cls = (el.getAttribute('class') || '').trim();
                if (cls) out.push({ cls, text: (el.innerText || '').trim().slice(0, 80) });
            });
            return out;
        }""")
        print("\n--- first card classes/text ---")
        for s in sub[:40]:
            print(f"  .{s['cls']}  =>  {s['text']!r}")

        await ctx.close()


asyncio.run(main())
print("\nDONE")
