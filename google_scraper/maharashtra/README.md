# Maharashtra mobile-shop run — operating guide

2 keywords, one workbook per district, polygon-clipped grids.

## What is already set up

| File | Purpose |
|---|---|
| `maharashtra/districts.json` | 36 districts: bbox, polygon path, area, cell counts |
| `maharashtra/polygons/*.geojson` | Cached district boundaries from OSM/Nominatim |
| `fetch_maharashtra_districts.py` | Rebuilds the plan (safe to re-run; cached) |
| `run_maharashtra_districts.ps1` | District queue with retries, checkpoint resume, captcha stop |
| `merge_maharashtra_districts.py` | Statewide dedupe + chain marking + master workbook |
| `clean_district_polygon.py` | Splits old workbooks into in-district shops and border spill-over |
| `../../scrape.py --polygon` | Clips grid cells **and results** to the district boundary |

Measured plan: **13,284 clipped cells** (from 22,083 rectangle cells) ×
2 keywords = **26,568 searches**.

Solapur is already finished. Its taluka master was cleaned against the district
polygon: **1,171 in-district shops** (including 68 rows without coordinates),
with 278 out-of-district rows moved to `outside_district` (25 Karnataka, 252
neighbouring Maharashtra districts). Kolhapur is also finished: **1,026
in-district shops**, 661 out-of-district rows quarantined (295 Karnataka, 366
Maharashtra). The remaining 34 districts are ~12,296 cells and ~24,592 searches.

## Run it

```powershell
cd C:\Users\Nom\Evaara-Scraper-\google_scraper

# pilot: one mid-size district
.\run_maharashtra_districts.ps1 -Only kolhapur

# a few districts
.\run_maharashtra_districts.ps1 -Limit 3

# everything, largest first (default)
.\run_maharashtra_districts.ps1

# optional third keyword everywhere (not recommended)
.\run_maharashtra_districts.ps1 -IncludeAccessories

# third keyword only in metro districts (fresh run: all 3 keywords)
.\run_maharashtra_districts.ps1 -Only mumbai_city,mumbai_suburban,thane,pune,nagpur,nashik -IncludeAccessories

# add the third keyword to districts already finished, without redoing 1-2
.\run_maharashtra_districts.ps1 -Only pune,nashik -AccessoriesOnly
# -> maharashtra_districts/<slug>_accessories.xlsx; merged automatically

# dry run: print the commands without launching Chrome
.\run_maharashtra_districts.ps1 -DryRun
```

The runner skips districts whose workbook already exists, so simply re-running
the same command continues where it stopped. A district that crashes is retried
up to 3 times and resumes from its `_partial.xlsx` checkpoint.

## Resume behaviour

Two checkpoints per district:

1. `maharashtra_districts/<slug>_mobile_all_partial.xlsx` — scraped rows.
2. `maharashtra_districts/<slug>_mobile_all.progress.jsonl` — one line per
   cell already searched (including cap-split children).

On a restart the scraper reloads the rows and **skips searched cells**, so a
crash at hour 6 of an 8-hour district loses minutes, not the whole night. Both
files are deleted automatically after a successful finish.

## District boundary filtering

`countryCode` in the output is Google's **state** code, not a country: `MA` =
Maharashtra, `KA` = Karnataka, `GA` = Goa. The first Kolhapur run contained 295
`KA` rows because a Maps search centred on a border cell returns shops from the
other side of the line (Belagavi, Nipani, Sankeshwar), and the old filter only
checked the rectangular bbox.

Two fixes are in place:

1. `scrape.py` tests every result's coordinates against the district polygon
   before extraction (`outside_area` in the stats/log). Out-of-district results
   are skipped and never written to the workbook.
2. `clean_district_polygon.py` re-checks older workbooks. Rows outside the
   polygon move to an `outside_district` sheet instead of being deleted, so
   border mistakes can be reviewed.

```powershell
python clean_district_polygon.py kolhapur solapur          # dry run -> _clean.xlsx
python clean_district_polygon.py kolhapur solapur --apply  # in place + _raw backup
```

The `shops` sheet is now strictly in-district. `outside_district` is kept for
audit and is **not** merged into the statewide master. Rows without coordinates
cannot be tested and stay in `shops`. Resumed checkpoints are re-checked before
they are loaded, and the final rows are checked again before the workbook is
written, so an old partial cannot reintroduce border shops.

## Speed tuning

The scraper now:

- polls for new cards while scrolling instead of always sleeping a fixed
  interval, and stops after 3 stale checks instead of 4 (`SCROLL_WAIT_MS`,
  `SCROLL_POLL_MS`, `SCROLL_STALE_LIMIT`);
- splits a capped cell only when at least `CAP_SPLIT_MIN` (default 90) results
  are actually inside the district, instead of splitting whenever the 120-cap
  contains any out-of-area padding;
- drops out-of-district results before opening place pages, so it never spends
  phone/website extraction time on border-state shops.

Do **not** raise `--step` to save time without a coverage test. A 0.07° test
cell found 17 shops where a finer grid found 48 in the same box: Google returns
a relevance-limited set, so coarser cells lose shops even when they do not hit
the 120 cap. `0.05` is the validated default.

## Stop conditions

| Exit | Meaning | Action |
|---|---|---|
| 0 | district finished | none |
| 2 | CAPTCHA | solve once with `HEADLESS=0`, then rerun; queue resumes |
| 1 | 2 districts failed in a row | check network, rerun |

## Monitor

```powershell
Get-Content .\maharashtra_districts\_logs\status.txt -Tail 30 -Wait
Get-Content .\maharashtra_districts\_logs\runs.csv -Tail 10
```

Per-district raw output is in `maharashtra_districts/_logs/<slug>.log`.

## Merge

```powershell
python merge_maharashtra_districts.py
# -> maharashtra_districts/maharashtra_merged_master.xlsx
```

Sheets: `shops` (deduped), `duplicates`, `filtered_out`, `summary`.
Dedupe is by Google place ID; the richest copy wins and blanks are filled from
the other copies. Chains are green, missing-phone rows yellow. Each district's
`outside_district` sheet is not merged; its row count is reported in `summary`.

If `<slug>_accessories.xlsx` exists (from `-AccessoriesOnly`), the merge reads
it as an extra source for that district and dedupes it against the main file.

Partial merge while the run is going:

```powershell
python merge_maharashtra_districts.py --only kolhapur,pune,nashik
```

## Measured expectations

| Metric | Value |
|---|---|
| Cells | 13,284 (45% fewer than rectangles) |
| Searches @2 keywords | 26,568 |
| Time per search (Kolhapur) | 21.3 s |
| Time per search (Solapur) | 18.9 s |
| Continuous runtime | ~6–6.5 days |
| 10 h/night | ~15–16 nights |
| 8 h/night | ~18–20 nights |
| Expected in-district shops | ~30,000–35,000 |

Raw hits are expected around 40,000–45,000; 15–40% of them sit across a
district or state border and are removed by the polygon test. Cap splits in
Mumbai/Thane/Pune add time; rural districts run close to the cell count.

Scroll polling and the in-area cap rule were added after the Kolhapur pilot.
They should save a few seconds on capped cells; re-measure after the first
district and adjust the runtime estimate if needed.

## Recommended order

1. ~~Kolhapur pilot~~ — done; 1,026 in-district shops. Border filter and speed
   tuning are in place.
2. **Metro districts** while you can check results: Mumbai City, Mumbai
   Suburban, Thane, Pune, Nagpur, Nashik.
3. **Remaining districts** unattended, largest first (the default order).
4. Optional: third keyword only in the metro districts listed above.
5. Merge, then run the `filtered_out` recovery pass for phone-named rows.

## Notes

- `countryCode` is Google's state code, not a country (`MA` = Maharashtra,
  `KA` = Karnataka, `GA` = Goa). It is kept under that name for compatibility.
- One browser process at a time. Do not run two districts in parallel on the
  same IP; captcha risk outweighs the speed gain.
- Windows sleep is disabled on AC; the monitor turning off is harmless.
- The `--step 0.05 --min-step 0.01` default is deliberate. A 0.07° test lost
  shops (17 vs 48 in the same box), so change the step only with a coverage test.
- Speed knobs: `CAP_SPLIT_MIN=90`, `SCROLL_WAIT_MS=2500`,
  `SCROLL_POLL_MS=250`, `SCROLL_STALE_LIMIT=3`. Lowering
  `DELAY_MIN_MS`/`DELAY_MAX_MS` speeds extraction but raises captcha risk.
- To re-fetch a boundary: `python fetch_maharashtra_districts.py --only pune --force`.
