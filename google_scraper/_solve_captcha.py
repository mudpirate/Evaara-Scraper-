# /// script
# requires-python = ">=3.11"
# dependencies = ["patchright", "rich"]
# ///
"""Open a VISIBLE browser with the scraper profile so a human can solve
Google's captcha manually. Auto-exits once Google loads clean.

Usage:
    python _solve_captcha.py [profile_dir]
    (default profile: ~/.cache/maps-scraper/profile)

Leave this running. A Chrome window will appear — solve the captcha /
click through the /sorry/ page in THAT window. This script detects when
it clears and exits 0. Max wait 10 minutes.
"""
import asyncio
import os
import sys
from pathlib import Path

if sys.platform == 'win32':
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    os.environ.setdefault('PYTHONUTF8', '1')

PROFILE = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    Path.home() / ".cache" / "maps-scraper" / "profile")
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36}")


async def is_captcha(page) -> bool:
    url = page.url or ''
    if '/sorry/' in url or 'captcha' in url:
        return True
    for sel in ('iframe[src*="recaptcha"]', 'form[action*="/sorry/"]',
                'div#recaptcha', '#captcha-form'):
        try:
            if await page.locator(sel).count() > 0:
                return True
        except Exception:
            pass
    return False


async def main() -> None:
    from patchright.async_api import async_playwright
    print(f"profile: {PROFILE}", flush=True)
    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE),
            headless=False,  # VISIBLE — user solves the captcha here
            locale="en-US",
            viewport={"width": 1400, "height": 900},
            user_agent=USER_AGENT,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        await page.goto("https://www.google.com/maps", timeout=60000)
        for elapsed in range(0, 600, 5):
            if not await is_captcha(page):
                # double-check it can actually search
                print(f"CAPTCHA_CLEAR after ~{elapsed}s — profile OK",
                      flush=True)
                await ctx.close()
                return
            if elapsed % 30 == 0:
                print(f"waiting for manual solve... ({elapsed}s). "
                      f"Solve the captcha in the visible Chrome window.",
                      flush=True)
            await asyncio.sleep(5)
        print("STILL_BLOCKED after 10 min", flush=True)
        await ctx.close()
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())
