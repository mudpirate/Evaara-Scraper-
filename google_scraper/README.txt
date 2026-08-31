GOOGLE MAPS SCRAPER
====================

Scrapes business listings from Google Maps (shops, phone numbers,
websites, ratings, addresses) into an Excel file.

Only needs Python 3.11+. Everything else (patchright, rich, openpyxl,
Chromium browser) is installed automatically on first run.


HOW TO RUN
----------

Option 1 - double-click:
    Double-click run.bat -> enters the city/district -> scrapes all fixed
    keywords -> saves one Excel file with duplicates removed.

Option 2 - batch (recommended for regular use):
    The keyword "mobile phones" and the STATE (karnataka) are FIXED.
    You only enter the city and district:

    python run_batch.py
    Then enter: belura, bidar district
    Or pass it directly:
    python run_batch.py "belura, bidar district"

    This runs:  mobile phones in belura, bidar district, karnataka

    The state is appended automatically, so results stay within Karnataka
    (e.g. "belura" won't pick up Belur in Howrah, West Bengal).

    Multiple cities in one run — separate with ; (semicolon):
    python run_batch.py "belura, bidar district; humnabad, bidar district; basavakalyan, bidar district"

    Results are merged into ONE Excel file with duplicates removed.

    To change the fixed keyword or state, edit KEYWORDS / STATE in run_batch.py.

Option 3 - direct (single keyword, one-shot):
    python scrape.py "mobile phones" "belura, bidar district, karnataka"
    python scrape.py "mobile phones" "belura, bidar district, karnataka; humnabad, bidar district, karnataka"

Option 4 - interactive:
    python scrape.py
    Then enter: keyword(s), place(s) (use ; between multiple places), output file name.


EXAMPLES
--------
    python run_batch.py "belura, bidar district"
    python run_batch.py "belura, bidar district; humnabad, bidar district; basavakalyan, bidar district"
    python run_batch.py "mysuru, mysuru district"
    python scrape.py "restaurants" "koramangala, bengaluru urban district, karnataka" out.xlsx

The state (karnataka) is added automatically. Multiple places are
separated with ; and each keyword runs for each place.


USEFUL SETTINGS (environment variables)
---------------------------------------
    HEADLESS=1        run without a visible browser window
    KEEP_ALL=1        disable the category filter (keep every result)
    DELAY_MIN_MS      min delay between extractions (ms, default 600)
    DELAY_MAX_MS      max delay between extractions (ms, default 1800)

Example:
    set HEADLESS=1 && python scrape.py "mobile shops" "bidar"


OUTPUT
------
    The Excel file has two sheets:
      "shops"        - businesses matching your search
      "filtered_out" - unrelated results Google returned (for transparency)

    Duplicates are removed across ALL queries in a run. A business is
    considered a duplicate if it has the same name AND same location
    (coordinates), so the same shop found under "mobile phones" and
    "mobile accessories" is saved only once.

    CHAIN-STORE HIGHLIGHTING:
    Known chains/brands (Poorvika, Sangeetha, Pai Mobiles, SS Mobiles,
    Mi Store, My Jio Store, NTel Mobiles, Sara Mobiles, Pai International,
    Samsung Service Center, Reliance Digital, Xiaomi, Apple resellers,
    Samsung Cafe, Oppo, Vivo, PhonoRange, ...)
    are highlighted in GREEN when they appear in the results.

    MISSING PHONE HIGHLIGHTING:
    Any business with NO phone number gets its phone cell highlighted
    in YELLOW, so you can instantly see which stores are missing numbers
    (even ones that aren't chain stores).

    To add/remove highlighted names, edit the CHAIN_NAMES list at the
    top of scrape.py.

    If the output file is open in Excel, the scraper automatically saves
    to a new file (name_2.xlsx, name_3.xlsx, ...) instead of failing.

    Typical results: 70-120 shops per locality.

    If you press Ctrl+C mid-run, what was collected so far is saved to
    a "<name>_partial.xlsx" file - nothing is lost.


TROUBLESHOOTING
---------------
    Captcha appears (headless):  re-run WITHOUT HEADLESS=1 once, solve the
        captcha in the visible browser. The browser profile remembers, so
        headless works again afterwards.

    First run downloads Chromium (~150 MB, one-time). Needs internet.

    If the browser was blocked, wait ~30 minutes and re-run slower.
