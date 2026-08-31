"""Find how to build an impcat category URL for a given keyword."""
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

        # 1) search shell: what category info does it expose?
        await page.goto("https://dir.indiamart.com/search.mp?ss=mobile+accessories&city=bengaluru",
                        wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(2500)
        meta = await page.evaluate("""() => {
            const el = document.querySelector('#__NEXT_DATA__');
            if (!el) return { next: false };
            try {
                const pp = JSON.parse(el.textContent).props.pageProps;
                return {
                    keyword: pp.keyword,
                    contextMcatId: pp.contextMcatId,
                    contextCatId: pp.contextCatId,
                    mcatid: pp.mcatid,
                    responseCity: pp.responseCity,
                };
            } catch (e) { return { next: 'err' }; }
        }""")
        print("search-shell meta:", meta)

        # sidebar category links on the shell (mcat / impcat links)
        links = await page.evaluate("""() => {
            const out = [];
            document.querySelectorAll('a[href*="/impcat/"], a[href*="mcatid"]').forEach(a => {
                out.push({ text: (a.innerText || '').trim().slice(0, 60), href: (a.href || '').slice(0, 160) });
            });
            return out.slice(0, 15);
        }""")
        print("category links on shell:", links)

        # 2) candidate impcat URLs
        candidates = [
            "https://dir.indiamart.com/impcat/mobile-accessories.html",
            "https://dir.indiamart.com/impcat/mobile-phone-accessories.html",
            "https://dir.indiamart.com/impcat/mobile-phones.html",
            "https://dir.indiamart.com/impcat/mobile-accessories-135.html",
        ]
        for c in candidates:
            try:
                r = await page.goto(c, wait_until="domcontentloaded", timeout=45000)
                await page.wait_for_timeout(2000)
                n_cards = await page.locator('article.template7-product-card, [class*="template7-product-card"]').count()
                n_prod = await page.locator('[class*="template7-product-name"]').count()
                print(f"{c}\n    status={r.status if r else 'n/a'} cards={n_cards} prod_names={n_prod}")
            except Exception as e:
                print(f"{c}\n    ERROR {e}")

        await ctx.close()


asyncio.run(main())
print("\nDONE")
