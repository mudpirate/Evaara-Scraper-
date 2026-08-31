"""Dump IndiaMART impcat category-page DOM structure for the scraper fallback."""
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
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(3000)
        print(f"status={resp.status if resp else 'n/a'} final={page.url}")

        # Unique classes containing 'supplier' or 'product'
        info = await page.evaluate("""() => {
            const classes = new Set();
            document.querySelectorAll('[class]').forEach(el => {
                const cls = el.getAttribute('class') || '';
                for (const c of cls.split(/\\s+/)) {
                    const lc = c.toLowerCase();
                    if (lc.includes('supplier') || lc.includes('product') ||
                        lc.includes('listing') || lc.includes('card')) classes.add(c);
                }
            });
            return Array.from(classes).slice(0, 60);
        }""")
        print("classes:", info)

        # Find anchors to proddetail and dump their ancestry
        detail = await page.evaluate("""() => {
            const anchors = Array.from(document.querySelectorAll('a[href*="/proddetail/"]'));
            if (!anchors.length) return { count: 0 };
            const a = anchors[0];
            let chain = [];
            let node = a;
            for (let d = 0; node && d < 6; d++, node = node.parentElement) {
                chain.push(node.tagName + '.' + (node.className || '').toString().trim().slice(0, 60));
            }
            return {
                count: anchors.length,
                href: a.href.slice(0, 140),
                text: (a.innerText || '').trim().slice(0, 200),
                chain,
                html: a.closest('[class]').outerHTML.slice(0, 1200),
            };
        }""")
        print("\n--- first proddetail anchor ---")
        for k, v in detail.items():
            print(f"{k}: {v}")

        await ctx.close()


asyncio.run(main())
print("\nDONE")
