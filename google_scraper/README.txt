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
    The specific keyword "cell phone store" and the STATE (karnataka) are
    FIXED. You only enter the city and district:

    python run_batch.py
    Then enter: belura, bidar district
    Or pass it directly:
    python run_batch.py "belura, bidar district"

    This runs:  cell phone store in belura, bidar district, karnataka

    The state is appended automatically, so results stay within Karnataka
    (e.g. "belura" won't pick up Belur in Howrah, West Bengal).

    Multiple cities in one run — separate with ; (semicolon):
    python run_batch.py "belura, bidar district; humnabad, bidar district; basavakalyan, bidar district"

    Results are merged into ONE Excel file with duplicates removed.

    To change the fixed keyword or state, edit KEYWORDS / STATE in run_batch.py.

Option 3 - direct (single keyword, one-shot):
    python scrape.py "cell phone store" "belura, bidar district, karnataka"
    python scrape.py "cell phone store" "belura, bidar district, karnataka; humnabad, bidar district, karnataka"

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


GPS GRID MODE (FULL AREA COVERAGE)
----------------------------------
A broad search like "cell phone store in <district>" stops at Google's
~120-result cap, so small village shops never appear. Grid mode avoids that
by covering the area with a checkerboard of cells and searching once per
cell with a coordinate-centered Maps URL:

    https://www.google.com/maps/search/cell+phone+store/@18.234,75.688,14z

Every cell returns its own local set of shops, so nothing depends on one
oversized result list.

1. AUTOMATIC AREA BOUNDARY (needs internet for the lookup)
   Give a place name; the bounding box is fetched from OpenStreetMap
   Nominatim (no API key), then split into cells:

       python scrape.py "cell phone store" --area "barshi, maharashtra" --step 0.02

2. MANUAL BOUNDING BOX
   In Google Maps, right-click two opposite corners of the area and copy
   their coordinates (latitude,longitude). Order does not matter:

       python scrape.py "cell phone store" --bbox 18.07,75.52,18.39,75.84 --step 0.05

3. PREVIEW WITHOUT SCRAPING

       python scrape.py "cell phone store" --area "barshi, maharashtra" --step 0.02 --dry-run

   Prints the bbox, cell count, zoom and number of searches, then exits.

4. ALL THREE KEYWORDS (large grid — --yes skips the confirmation)

       python scrape.py "cell phone store, smartphone accessories, mobile phone repair shop" --area "barshi, maharashtra" --step 0.05 --yes

   This is 49 cells x 3 keywords = 147 searches; expect several hours.

OPTIONS
    --area "place, state"   look up the bounding box automatically
    --bbox lat1,lon1,lat2,lon2   use a manual box (exclusive with --area)
    --step DEGREES          cell size; default 0.02 (~2.2 km)
    --min-step DEGREES      smallest split size; default 0.01 (~1.1 km)
    --zoom N                force a zoom; default is picked per cell so the
                            viewport always covers the cell
    --out FILE.xlsx         output path
    --yes                   skip the confirmation prompt for large grids
    --dry-run               print the plan and exit

CHOOSING A STEP
    0.02  (~2.2 km)  towns and cities: most complete, slowest
    0.05  (~5.6 km)  talukas / rural districts: good balance
    0.10  (~11 km)   first pass over a very large district

    A run costs roughly 60-180 seconds (1-3 minutes) per search, and each
    keyword is searched at every cell. Example: 49 cells x 3 keywords =
    147 searches. Grids over 50 searches ask for confirmation unless --yes
    is passed.

120-RESULT CAP INSIDE A CELL
    If a dense cell still returns exactly 120 in-area results, the scraper
    splits it into four smaller cells (and searches each at a closer zoom),
    repeating until the cap is relieved or --min-step is reached. Splits are
    automatic; the run log shows them. A cap that is only full because Google
    added out-of-area fallback results is not split, since the local results
    are already exhausted. The default floor is 0.01° (~1.1 km); lower
    --min-step to 0.005 for an extremely dense market.

CHECKPOINTS (LONG RUNS)
    Grid runs save <output>_partial.xlsx after every cell. If the browser
    crashes, the network drops, or you press Ctrl+C, the shops collected so
    far are still on disk (with latitude/longitude). The checkpoint is
    removed automatically when the final workbook is written.

DEDUPLICATION
    Neighbouring cells overlap, so the same shop is often found several
    times. Results are deduplicated by name + GPS coordinates (falling back
    to Google's place ID, then name + address). Shops already collected are
    skipped before extraction, which also saves time.

    Grid workbooks add "latitude" and "longitude" columns after the normal
    scraper columns.

AREA SAFETY CHECK
    Google Maps sometimes ignores a coordinate when a cell has no local
    shops and falls back to the browser's default city, which can be
    hundreds of kilometres away. Every result is checked against the grid
    bounding box (plus a margin of one --step) before extraction, so those
    far-away fallback results are skipped. The run log shows how many were
    skipped. A cell with no in-area results is not split.

TIP: the automatic --area lookup uses the full administrative boundary
(e.g. the whole Barshi taluka, not just the town). Use a manual --bbox for a
tighter area.


ACCURATE RESULTS (FALSE-POSITIVE FILTER)
----------------------------------------
Google Maps does not only match business names/categories. It also matches
query words against street addresses and customer reviews. That is why a
search for "mobile" can return:
    - an insurance company with a review saying "they insured my phone"
    - a waterworks company located on "Mobile Avenue"
    - a mobile caterer, Oldsmobile dealer, or auto-parts shop

This scraper now fixes that in two ways:

1. CATEGORY ALLOWLIST + JUNK EXCLUSION
   A result is kept only when Google's structured category (categoryName)
   is a phone/electronics category, such as:
       Cell phone store
       Cell phone accessory store
       Mobile phone repair shop
       Phone repair service
       Electronics store
       Telecommunications service provider
   Results are also rejected when the category or business NAME is a known
   other trade: insurance, waterworks, plumber, restaurant, bank, mobile home,
   Oldsmobile/snowmobile, auto parts, etc.
   Addresses are never used for filtering, because a legitimate phone shop
   often sits "opposite SBI Bank" or on "Mobile Avenue".

2. SPECIFIC SEARCH PHRASES
   Generic keywords are rewritten to specific business phrases:
       mobile shops        ->  mobile phone store
       mobile phones       ->  mobile phone store
       mobile accessories  ->  smartphone accessories
       electronics+appl.   ->  electronics store
   This avoids the very broad standalone word "mobile".

   NOTE ON QUOTES: wrapping the phrase in quotes (as you would on Google
   Search) was tested against Google Maps. A query like
   "cell phone store" in Bidar returns only the city "Bidar" and no shops,
   while the unquoted phrase returns the actual stores. Quoting is therefore
   OFF by default. Set STRICT_QUERY=1 only if you want to experiment.


USEFUL SETTINGS (environment variables)
---------------------------------------
    HEADLESS=1        run without a visible browser window
    KEEP_ALL=1        disable category + junk filter (raw results)
    STRICT_QUERY=1    wrap phrases in quotes (experimental; usually breaks Maps)
    DELAY_MIN_MS      min delay between extractions (ms, default 600)
    DELAY_MAX_MS      max delay between extractions (ms, default 1800)

Example:
    set HEADLESS=1 && python scrape.py "cell phone store" "bidar"


OUTPUT
------
    The Excel file has two sheets:
      "shops"        - businesses that passed the category allowlist + junk filter
      "filtered_out" - unrelated results Google returned (insurance, waterworks,
                       restaurants, banks, auto shops, etc., for transparency)

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

    Typical results: 40-120 shops per locality, depending on city size.
    Strict phrases deliberately trade a little coverage for accuracy.

    If you press Ctrl+C mid-run, what was collected so far is saved to
    a "<name>_partial.xlsx" file - nothing is lost.


CLEANING OLD FILES (NO RE-SCRAPE NEEDED)
---------------------------------------
If you already have Excel files scraped before this fix, clean them with the
same category allowlist + junk filter:

    python clean_existing.py "cell phone store" old_results.xlsx

This writes old_results_cleaned.xlsx with two sheets:
    shops         - cleaned rows
    filtered_out  - removed rows (insurance, waterworks, auto parts, etc.)

Clean several files at once:

    python clean_existing.py "cell phone store" a.xlsx b.xlsx c.xlsx

Choose a custom output name (single file only):

    python clean_existing.py "cell phone store" old_results.xlsx -o clean.xlsx


MERGING SEVERAL RESULTS FILES
-----------------------------
If you ran one query per file (for example "cell phone store" and
"smartphone accessories"), merge them into one master workbook:

    python merge_results.py north_solapur_mobile_master.xlsx north_solapur_mobile.xlsx north_solapur_mobile_other.xlsx

Shops are matched on title + street + city, so the same shop found by two
queries is written once. A shared phone number alone is NOT treated as a
duplicate - chains often use one number for several branches. The richest
copy is kept and blank fields (phone, website, street) are filled from the
other copy.

Add readable query labels with --label:

    python merge_results.py master.xlsx a.xlsx b.xlsx \
        --label a.xlsx="cell phone store" \
        --label b.xlsx="smartphone accessories; mobile phone repair shop"

The output has the usual two sheets:
    shops         - one row per unique shop
    filtered_out  - removed rows from all input files

An extra "foundIn" column records which query file(s) each shop came from.


TROUBLESHOOTING
---------------
    Captcha appears (headless):  re-run WITHOUT HEADLESS=1 once, solve the
        captcha in the visible browser. The browser profile remembers, so
        headless works again afterwards.

    First run downloads Chromium (~150 MB, one-time). Needs internet.

    If the browser was blocked, wait ~30 minutes and re-run slower.
