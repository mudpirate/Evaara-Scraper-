# Running the Google Scraper — Maharashtra Districts

Everything needed is inside:

```text
C:\Users\Nom\Evaara-Scraper-\google_scraper
```

Output workbooks go to:

```text
C:\Users\Nom\Evaara-Scraper-\google_scraper\maharashtra_districts
```

Current state:

- **36 districts** in the plan.
- **Kolhapur** and **Solapur** are finished and are skipped automatically.
- **34 districts remain**: 12,296 grid cells / 24,592 base searches.
- 2 keywords: `cell phone store` + `mobile phone repair shop`.
- Base grid `0.05°`; capped cells split at 40 in-district results, down to **z15 (~2.8 km)** everywhere, **z16 (~1.4 km)** for the six metros only, or ~175 m — whichever comes first.
- The metro list is `mumbai_city`, `mumbai_suburban`, `thane`, `pune`, `nagpur`, `nashik`. Each district in `districts.json` carries its own `max_zoom` and `metro` flag.
- Splits are breadth-first: every base cell is searched before any sub-cell, so one dense market cannot starve the rest of the district.
- Estimated remaining time: **9–11 days continuous**, or **22–27 nights at 10 h/night**.

Renamed districts (the old names still work with `-Only`):

| Old name | Official name |
|---|---|
| Ahmednagar | **Ahilyanagar** |
| Aurangabad | **Chhatrapati Sambhajinagar** |
| Osmanabad | **Dharashiv** |

> Run only **one** scraper at a time. Two instances on the same IP increase captcha risk.

---

## 1. Open the terminal

Open **PowerShell** and go to the scraper folder:

```powershell
cd C:\Users\Nom\Evaara-Scraper-\google_scraper

# Allow the local runner in this PowerShell window only.
# Needed if you see "running scripts is disabled on this system".
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

The execution-policy command affects only the current window; it does not change Windows globally.

Check Python:

```powershell
python --version
```

Python 3.11 or newer is required. This machine has 3.13.

Check the plan:

```powershell
python -c "import json; d=json.load(open(r'maharashtra\districts.json', encoding='utf-8')); print('step:', d['step'], '| split floor:', d['min_step'], '| max zoom:', d['max_zoom']); print('cells:', d['totals']['cells_clipped'], '| base searches:', d['totals']['queries_2_keywords'])"
```

Expected output:

```text
step: 0.05 | split floor: 0.00125 | max zoom: 15 (metros 16)
cells: 13284 | base searches: 26568
```

Optional dry run (prints commands, does not launch Chrome):

```powershell
.\run_maharashtra_districts.ps1 -Limit 1 -DryRun
```

---

## 2. Start the full run

```powershell
.\run_maharashtra_districts.ps1
```

What the runner does:

- Runs districts **largest first**.
- Uses **z15** for every district and **z16** automatically for the six metros. `-MaxZoom N` overrides both for one run.
- Applies the **zero-yield rule**: if the last 50 split cells produce 0–1 new shops, it saves the workbook, exits with code 3, and the runner moves to the next district.
- Skips any district whose `*_mobile_all.xlsx` already exists.
- Retries a failed district up to **3 times**.
- Resumes a crashed district from its `_partial.xlsx` checkpoint.
- Waits 60 seconds between districts.
- Stops the queue on a captcha (exit code 2).
- Stops after 2 districts fail in a row (network problem).

Keep the terminal open while it runs. If the terminal closes, the scraper stops — just run the same command again and it resumes.

### Metro split (keyword 3, z16)

Run the optional accessories keyword only in the six metros, where dense wholesale markets make it worthwhile. This writes separate `<slug>_accessories.xlsx` files and leaves the main workbooks untouched.

```powershell
.\run_maharashtra_districts.ps1 -MetrosOnly -AccessoriesOnly
```

`-MetrosOnly` picks `mumbai_city`, `mumbai_suburban`, `thane`, `pune`, `nagpur`, `nashik`. Those districts get z16 from the plan; everything else stays z15. Confirm with a dry run first:

```powershell
.\run_maharashtra_districts.ps1 -MetrosOnly -AccessoriesOnly -DryRun
```

Kolhapur is already done at z16 for accessories, so it is not in this pass.

### Zero-yield stop rule

Base grid cells are **always** searched. The rule only looks at refinement cells (z16 in metros, z16/z15 in a split run):

| Setting | Default | Meaning |
|---|---:|---|
| `-ZeroYieldWindow` | 50 | number of recent split cells to check |
| `-ZeroYieldMax` | 1 | if that window yields at most this many new shops, stop |

When it triggers, the scraper writes the workbook and exits with code **3**. The runner treats 3 as a completed district and continues to the next one — it is not a failure and does not trigger a retry.

Evidence from Kolhapur: 87 split cells produced **1** new shop, so the rule would have stopped after 50 cells and saved ~2–3 hours.

To disable for a one-off test:

```powershell
.\run_maharashtra_districts.ps1 -ZeroYieldWindow 9999
# or per scrape.py call: --no-zero-yield
```

### Start it minimized in the background

Use this if you do not want the PowerShell window in the way:

```powershell
Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','C:\Users\Nom\Evaara-Scraper-\google_scraper\run_maharashtra_districts.ps1' -WindowStyle Minimized
```

Progress still goes to `maharashtra_districts\_logs\status.txt`.

### Resume after stopping

Run the same command again:

```powershell
.\run_maharashtra_districts.ps1
```

Finished districts are skipped, and the interrupted district continues from its checkpoint.

### Run only specific districts

```powershell
# one district
.\run_maharashtra_districts.ps1 -Only pune

# several districts
.\run_maharashtra_districts.ps1 -Only mumbai_city,mumbai_suburban,thane

# first 3 remaining districts
.\run_maharashtra_districts.ps1 -Limit 3

# alphabetical instead of largest-first
.\run_maharashtra_districts.ps1 -Order name
```

### Optional accessories pass

The plan's default third keyword is `smartphone accessories`. To add it to districts that already have their main 2-keyword workbook, use `-AccessoriesOnly`:

```powershell
.\run_maharashtra_districts.ps1 -Only mumbai_city,mumbai_suburban,thane,pune,nagpur,nashik -AccessoriesOnly
```

This writes `<slug>_accessories.xlsx` and does **not** redo the first two keywords. The merge reads it automatically.

To use a different accessories phrase for one run, add `-AccessoriesKeyword`:

```powershell
.\run_maharashtra_districts.ps1 -Only kolhapur -AccessoriesOnly -AccessoriesKeyword "mobile accessories"
```

The override applies only to that run; the plan default stays `smartphone accessories`.

Note: the scraper rewrites `mobile accessories` to `smartphone accessories` internally because that phrase performs better on Google Maps. The output is still the accessories pass, saved as `<slug>_accessories.xlsx`.

### Run with 2 workers (split one district)

For a large district, split the bbox at a grid-aligned latitude and run two `scrape.py` processes side by side. Each worker needs its own `PROFILE_DIR`, because Chrome cannot share one profile.

Example: Kolhapur split into south and north halves at `16.443546` (14 rows from the south edge):

```powershell
# South half: resumes from the existing checkpoint and partial workbook
$env:PROFILE_DIR = "$HOME\.cache\maps-scraper\profile"
python .\scrape.py "mobile accessories" `
  --bbox 15.743546,73.687277,16.443546,74.700079 `
  --polygon .\maharashtra\polygons\kolhapur.geojson `
  --step 0.05 --min-step 0.00125 --max-zoom 16 --yes `
  --out .\maharashtra_districts\kolhapur_accessories.xlsx `
  --seed .\maharashtra_districts\kolhapur_accessories_partial.xlsx

# North half: separate profile, fresh output
$env:PROFILE_DIR = "$HOME\.cache\maps-scraper\profile_kolhapur_b"
python .\scrape.py "mobile accessories" `
  --bbox 16.443546,73.687277,17.181344,74.700079 `
  --polygon .\maharashtra\polygons\kolhapur.geojson `
  --step 0.05 --min-step 0.00125 --max-zoom 16 --yes `
  --out .\maharashtra_districts\kolhapur_accessories_north.xlsx
```

After both finish, merge the halves. Duplicates are removed by Google place ID:

```powershell
python .\merge_accessories_halves.py `
  .\maharashtra_districts\kolhapur_accessories.xlsx `
  .\maharashtra_districts\kolhapur_accessories_north.xlsx `
  .\maharashtra_districts\kolhapur_accessories.xlsx
```

Rules for a split run:

- Split at a latitude that is a multiple of `0.05°` from the district's south edge, so the two grids line up and leave no gap.
- Give each worker its own `PROFILE_DIR`; copy the main profile once, then remove Chrome's `SingletonLock`/`lockfile` from the copy.
- Two workers on one IP roughly double captcha risk. Keep them on the **same** district, and stop both if either gets a captcha.
- This is a manual procedure. The queue runner still runs one district at a time.

Watch each half:

```powershell
(Get-Content .\maharashtra_districts\kolhapur_accessories.progress.jsonl | Measure-Object -Line).Lines
(Get-Content .\maharashtra_districts\kolhapur_accessories_north.progress.jsonl | Measure-Object -Line).Lines
```

### Redo a district from scratch

This overwrites the existing workbook:

```powershell
# if a checkpoint exists, remove it first for a true fresh start
Remove-Item .\maharashtra_districts\pune_mobile_all_partial.xlsx -ErrorAction SilentlyContinue
Remove-Item .\maharashtra_districts\pune_mobile_all.progress.jsonl -ErrorAction SilentlyContinue

.\run_maharashtra_districts.ps1 -Force -Only pune
```

To redo Kolhapur and Solapur with the new 40 / 175 m settings:

```powershell
# optional: back up the cleaned workbooks first
Copy-Item .\maharashtra_districts\kolhapur_mobile_all.xlsx .\maharashtra_districts\kolhapur_mobile_all.backup.xlsx
Copy-Item .\maharashtra_districts\solapur_mobile_all.xlsx .\maharashtra_districts\solapur_mobile_all.backup.xlsx

.\run_maharashtra_districts.ps1 -Force -Only kolhapur,solapur
```

---

## 3. Monitor progress

### Live log (best while watching)

```powershell
Get-Content .\maharashtra_districts\_logs\status.txt -Tail 20 -Wait
```

This updates as the runner works. Press **Ctrl+C** to stop watching — the scraper keeps running.

### One-shot snapshot

```powershell
Get-Content .\maharashtra_districts\_logs\status.txt -Tail 20
```

### Refresh every 60 seconds

```powershell
while ($true) {
    Clear-Host
    Get-Content .\maharashtra_districts\_logs\status.txt -Tail 15
    Start-Sleep -Seconds 60
}
```

Press **Ctrl+C** to stop the monitor only.

### Progress dashboard

Copy and paste this whole block:

```powershell
$plan = Get-Content .\maharashtra\districts.json -Raw | ConvertFrom-Json
$out  = '.\maharashtra_districts'

$done = @($plan.districts | Where-Object {
    Test-Path (Join-Path $out ($_.slug + '_mobile_all.xlsx'))
})

$totalCells = ($plan.districts | Measure-Object cells_clipped -Sum).Sum
$doneCells  = ($done | Measure-Object cells_clipped -Sum).Sum

"Districts done: $($done.Count)/$($plan.districts.Count)"
"Grid cells:     {0:N0}/{1:N0} ({2:P1})" -f $doneCells, $totalCells, ($doneCells / $totalCells)

"`nPartial workbooks:"
$parts = @(Get-ChildItem $out -Filter '*_partial.xlsx' -ErrorAction SilentlyContinue)
if ($parts.Count -eq 0) {
    '  none'
} else {
    $parts | Select-Object Name,
        @{n='MB';e={[math]::Round($_.Length/1MB,2)}},
        LastWriteTime | Format-Table -AutoSize
}

"Active search progress:"
$prog = @(Get-ChildItem $out -Filter '*.progress.jsonl' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending)
if ($prog.Count -eq 0) {
    '  none'
} else {
    foreach ($p in $prog) {
        "  {0}: {1:N0} searches done" -f $p.Name,
            ((Get-Content $p.FullName | Measure-Object -Line).Lines)
    }
}
```

### What is running right now?

```powershell
Get-Content .\maharashtra_districts\_logs\status.txt -Tail 6
$latest = Get-ChildItem .\maharashtra_districts\_logs\*.log |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
$latest | Select-Object Name, LastWriteTime
Get-Content $latest.FullName -Tail 25
```

### Completed workbooks

```powershell
Get-ChildItem .\maharashtra_districts\*_mobile_all.xlsx |
    Sort-Object LastWriteTime |
    Select-Object Name,
        @{n='MB';e={[math]::Round($_.Length/1MB,2)}},
        LastWriteTime |
    Format-Table -AutoSize
```

### District run history

```powershell
Import-Csv .\maharashtra_districts\_logs\runs.csv |
    Format-Table timestamp, district, exit_code, elapsed_seconds, size_bytes -AutoSize
```

`elapsed_seconds` is the time that district took. A finished district has `exit_code = 0`.

### Per-district raw log

```powershell
Get-Content .\maharashtra_districts\_logs\pune.log -Tail 40
```

Replace `pune` with the district slug.

---

## 4. Captcha or crash

### Captcha

The runner stops with:

```text
CAPTCHA <District> - queue stopped
```

Solve it once in a visible browser, then rerun only that district:

```powershell
$env:HEADLESS = '0'
.\run_maharashtra_districts.ps1 -Only pune
```

Replace `pune` with the district named in the `CAPTCHA` line. Solve the puzzle in the Chrome window; the scraper detects it and continues. When that district finishes, turn headless mode back on and resume the full queue:

```powershell
Remove-Item Env:HEADLESS
.\run_maharashtra_districts.ps1
```

If it stopped in the middle of a district, the rerun resumes from the checkpoint.

### Crash, power cut, or terminal closed

Just rerun:

```powershell
.\run_maharashtra_districts.ps1
```

The district resumes from:

- `maharashtra_districts\<slug>_mobile_all_partial.xlsx` — rows already scraped
- `maharashtra_districts\<slug>_mobile_all.progress.jsonl` — cells already searched

Both are deleted automatically after the district finishes.

### Network failure

If 2 districts fail in a row, the queue stops. Check the internet, then rerun the same command.

---

## 5. Merge the districts

After all districts are done (or for a partial merge at any time):

```powershell
python .\merge_maharashtra_districts.py
```

Output:

```text
maharashtra_districts\maharashtra_merged_master.xlsx
```

Sheets:

- `shops` — deduplicated, strictly in-district shops
- `duplicates` — duplicate place IDs that were merged
- `filtered_out` — rejected categories / junk names
- `summary` — counts, including rows found outside district polygons

Merge only selected districts:

```powershell
python .\merge_maharashtra_districts.py --only kolhapur,solapur,pune
```

The merge deduplicates by Google place ID. Chains are green; rows without a phone are yellow. `outside_district` audit sheets are never merged.

---

## 6. Important files

| Path | Purpose |
|---|---|
| `run_maharashtra_districts.ps1` | District queue, retries, resume, captcha stop |
| `scrape.py` | Google Maps scraper, polygon filtering, cap splitting |
| `fetch_maharashtra_districts.py` | Rebuilds `maharashtra\districts.json` |
| `merge_maharashtra_districts.py` | Statewide dedupe + master workbook |
| `merge_accessories_halves.py` | Merges two half-district accessory workbooks |
| `maharashtra\districts.json` | 36-district plan: bbox, polygon, cells |
| `maharashtra\polygons\*.geojson` | District boundaries |
| `maharashtra\README.md` | Detailed technical notes |
| `maharashtra_districts\*_mobile_all.xlsx` | One finished workbook per district |
| `maharashtra_districts\_logs\status.txt` | Runner log / current status |
| `maharashtra_districts\_logs\runs.csv` | One row per district attempt |
| `maharashtra_districts\_logs\<slug>.log` | Full raw output for one district |

Do **not** delete `_partial.xlsx` or `.progress.jsonl` while a district is running — that is the resume checkpoint.

---

## 7. Timing notes

Measured on Kolhapur: about 19–21 seconds per search.

| Work | Estimate |
|---|---|
| 24,592 base searches | ~5.5–6 days continuous |
| With 40 / z15 cap splits (z16 metros) | ~9–12 days continuous |
| 10 hours per night | ~22–29 nights |
| 8 hours per night | ~27–36 nights |

Most extra time is in Mumbai, Pune, Thane, Nagpur and Nashik. Rural districts stay close to the base cell count.

### Why the split cap is z15 (z16 in the metros)

Measured on the dense north corner of Kolhapur:

| Zoom | Cell size | New shops | Searches |
|---|---|---:|---:|
| z14 | ~5.6 km | 45 | 3 |
| z15 | ~2.8 km | 5 | 3 |
| z16 | ~1.4 km | 1 | 10 |
| z17–z19 | 700–175 m | 4 | 397 |

z15 captures 91% of the extra leads, z16 captures 93% for ~10 more searches. Everything deeper is mostly the same top 120 repeated from neighbouring tiny cells. So `scrape.py` defaults to **z15**, and the six metros are raised to **z16** through each district's `max_zoom` in `districts.json`. Use `-MaxZoom N` / `--max-zoom N` only for a deliberate one-off override.

Keep the laptop plugged in and prevent Windows from sleeping. Turning the monitor off is fine.

---

## 8. Do not

- Do **not** raise the grid step. `0.05°` is the validated value; a `0.07°` test lost shops.
- Do **not** run two scraper instances on **different districts** at the same time. The only supported parallel setup is the 2-worker split of one district described above.
- Do **not** delete `_partial.xlsx` or `.progress.jsonl` while a district is running.
- Do **not** close the terminal if you want the run to continue — use the minimized launch command above, or simply rerun the queue later.

---

## 9. Quick reference

```powershell
# go to the scraper
cd C:\Users\Nom\Evaara-Scraper-\google_scraper

# allow local scripts in this window (needed on some systems)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# start / resume everything
.\run_maharashtra_districts.ps1

# watch live
Get-Content .\maharashtra_districts\_logs\status.txt -Tail 20 -Wait

# one snapshot
Get-Content .\maharashtra_districts\_logs\status.txt -Tail 20

# see completed workbooks
Get-ChildItem .\maharashtra_districts\*_mobile_all.xlsx |
    Sort-Object LastWriteTime | Select-Object Name, LastWriteTime

# merge everything
python .\merge_maharashtra_districts.py
```

If something looks wrong, check the last 20 lines of `status.txt` and the newest `_logs\<district>.log` first.
