# Maharashtra district queue - 2 keywords, polygon-clipped grids.
#
#   .\run_maharashtra_districts.ps1                      # all 36, biggest first
#   .\run_maharashtra_districts.ps1 -Only kolhapur        # pilot one district
#   .\run_maharashtra_districts.ps1 -Limit 3              # first 3 districts
#   .\run_maharashtra_districts.ps1 -DryRun               # print commands only
#   .\run_maharashtra_districts.ps1 -IncludeAccessories   # optional 3rd keyword
#   .\run_maharashtra_districts.ps1 -Only pune -AccessoriesOnly
#                                                        # 3rd keyword only,
#                                                        # into a separate file
#   .\run_maharashtra_districts.ps1 -MetrosOnly -AccessoriesOnly
#                                                        # 6 metros, 3rd keyword
#
# Zoom: z15 everywhere by default; the 6 metro districts in the plan are
# automatically raised to z16. -MaxZoom N overrides both for one run.
#
# Restart-safe: skips districts whose final workbook exists and resumes a
# crashed district from its _partial.xlsx checkpoint.
#
# Stop conditions:
#   exit code 2  -> CAPTCHA, queue stops (solve in a visible browser, rerun)
#   exit code 3  -> zero-yield rule: district saved, queue continues
#   2 districts fail in a row -> network is down, queue stops
#
# Zero-yield rule: after 50 refinement (split) cells produce 0 or 1 new
# shop, the scraper saves what it has and exits 3 so the runner moves on.
# Base grid cells are always completed; only splits are cut short.

param(
    [string]$Only = '',
    [int]$Limit = 0,
    [switch]$IncludeAccessories,
    [switch]$AccessoriesOnly,
    [string]$AccessoriesKeyword = '',
    [switch]$DryRun,
    [switch]$Force,
    [int]$PauseSeconds = 60,
    [ValidateSet('cells', 'name')]
    [string]$Order = 'cells',
    [switch]$MetrosOnly,
    [int]$MaxZoom = 0,
    [int]$ZeroYieldWindow = 50,
    [int]$ZeroYieldMax = 1
)

$ErrorActionPreference = 'Continue'

$root   = $PSScriptRoot
$python = "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"
if (-not (Test-Path $python)) { $python = 'python' }

$planPath = Join-Path $root 'maharashtra\districts.json'
$outDir   = Join-Path $root 'maharashtra_districts'
$logDir   = Join-Path $outDir '_logs'
New-Item -ItemType Directory -Force -Path $outDir, $logDir | Out-Null
$statusFile = Join-Path $logDir 'status.txt'
$runsFile   = Join-Path $logDir 'runs.csv'

function Log($message) {
    $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $message
    Add-Content -Path $statusFile -Value $line -Encoding UTF8
    Write-Output $line
}

if (-not (Test-Path $planPath)) {
    Write-Error "Missing $planPath - run: python fetch_maharashtra_districts.py"
    exit 1
}

$plan = Get-Content $planPath -Raw | ConvertFrom-Json
$districts = @($plan.districts)

if ($Only) {
    # Old names still work after the 2023 renames.
    $aliases = @{
        'ahmednagar' = 'ahilyanagar'
        'aurangabad' = 'chhatrapati_sambhajinagar'
        'osmanabad'  = 'dharashiv'
    }
    $wanted = $Only.Split(',') | ForEach-Object {
        $key = $_.Trim().ToLower().Replace(' ', '_')
        if ($aliases.ContainsKey($key)) { $aliases[$key] } else { $key }
    }
    $districts = $districts | Where-Object {
        $wanted -contains $_.slug -or
        $wanted -contains $_.name.ToLower().Replace(' ', '_')
    }
}
if ($MetrosOnly) {
    $districts = $districts | Where-Object { $_.metro -eq $true -or $_.slug -in @('mumbai_city','mumbai_suburban','thane','pune','nagpur','nashik') }
}
if (-not $districts -or $districts.Count -eq 0) {
    Write-Error "No districts matched -Only '$Only' / -MetrosOnly."
    exit 1
}

if ($Order -eq 'name') {
    $districts = @($districts | Sort-Object name)
} else {
    $districts = @($districts | Sort-Object cells_clipped -Descending)
}
if ($Limit -gt 0) {
    $districts = @($districts | Select-Object -First $Limit)
}

if ($AccessoriesOnly -and $IncludeAccessories) {
    Write-Error "-AccessoriesOnly and -IncludeAccessories cannot be combined."
    exit 1
}

$outSuffix = ''
if ($AccessoriesOnly) {
    # Post-pass for districts already scraped with the two main keywords:
    # writes <slug>_accessories.xlsx so nothing has to be redone.
    # -AccessoriesKeyword overrides the plan's optional keyword for one run.
    $keywords = if ($AccessoriesKeyword) {
        $AccessoriesKeyword
    } else {
        [string]$plan.optional_keyword
    }
    $outSuffix = '_accessories'
} else {
    if ($AccessoriesKeyword) {
        Write-Error "-AccessoriesKeyword requires -AccessoriesOnly."
        exit 1
    }
    $keywords = ($plan.keywords -join ', ')
    if ($IncludeAccessories) {
        $keywords = "$keywords, $($plan.optional_keyword)"
    }
}
$keywordCount = $keywords.Split(',').Count
$step = [string]$plan.step
$minStep = [string]$plan.min_step

$totalCells = ($districts | Measure-Object -Property cells_clipped -Sum).Sum
$totalSearches = $totalCells * $keywordCount

# Keep a machine-readable log for the later merge/report step.
if (-not (Test-Path $runsFile)) {
    'timestamp,district,attempt,exit_code,elapsed_seconds,file,size_bytes' |
        Out-File -FilePath $runsFile -Encoding UTF8
}

$zoomPolicy = if ($MaxZoom -gt 0) { "override z$MaxZoom" } else { 'plan (z15, metros z16)' }
Log "RUNNER START pid=$PID districts=$($districts.Count) cells=$totalCells base_searches=$totalSearches keywords=[$keywords] zoom=$zoomPolicy"
Log "Output: $outDir"

$index = 0
$doneCount = 0
$failed = @()
$consecutiveFail = 0

foreach ($d in $districts) {
    $index++
    if ($outSuffix) { $stem = $d.slug + $outSuffix }
    else { $stem = $d.slug + '_mobile_all' }
    $out = Join-Path $outDir ($stem + '.xlsx')
    $polygon = Join-Path $root ('maharashtra\' + ($d.polygon -replace '/', '\'))
    $bbox = ($d.bbox -join ',')

    if (-not $Force -and (Test-Path $out) -and ((Get-Item $out).Length -gt 5000)) {
        Log "[$index/$($districts.Count)] SKIP $($d.name) (output exists)"
        continue
    }
    if (-not (Test-Path $polygon)) {
        Log "[$index/$($districts.Count)] FAIL $($d.name) - polygon missing: $polygon"
        $failed += $d.name
        continue
    }

    $scrape = Join-Path $root 'scrape.py'
    $logFile = Join-Path $logDir ($d.slug + '.log')
    $districtDone = $false
    # Plan-driven zoom: z15 everywhere, z16 for the six metros.
    # -MaxZoom N overrides the plan for the whole run.
    if ($MaxZoom -gt 0) { $zoom = $MaxZoom }
    elseif ($d.max_zoom) { $zoom = [int]$d.max_zoom }
    else { $zoom = [int]$plan.max_zoom }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $seedPath = Join-Path $outDir ($stem + '_partial.xlsx')
        $seedArgs = @()
        if (Test-Path $seedPath) {
            $seedArgs = @('--seed', $seedPath)
            Log "[$index/$($districts.Count)] RESUME $($d.name) attempt $attempt from checkpoint"
        } else {
            Log "[$index/$($districts.Count)] START $($d.name) attempt $attempt cells=$($d.cells_clipped) area=$($d.area_km2)km2 zoom=$zoom"
        }

        $cmdLine = "$python `"$scrape`" `"$keywords`" --bbox $bbox --polygon `"$polygon`" --step $step --min-step $minStep --max-zoom $zoom --zero-yield-window $ZeroYieldWindow --zero-yield-max $ZeroYieldMax --yes --out `"$out`""
        if ($seedArgs.Count -gt 0) { $cmdLine += " --seed `"$seedPath`"" }

        if ($DryRun) {
            Write-Output $cmdLine
            $districtDone = $true
            break
        }

        Log "  $cmdLine"
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        & $python $scrape $keywords --bbox $bbox --polygon $polygon `
            --step $step --min-step $minStep --max-zoom $zoom `
            --zero-yield-window $ZeroYieldWindow --zero-yield-max $ZeroYieldMax `
            --yes --out $out @seedArgs 2>&1 |
            Tee-Object -FilePath $logFile -Append
        $code = $LASTEXITCODE
        $sw.Stop()
        $elapsed = [int]$sw.Elapsed.TotalSeconds

        $size = 0
        if (Test-Path $out) { $size = (Get-Item $out).Length }
        "{0},{1},{2},{3},{4},{5},{6}" -f `
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $d.slug, $attempt, $code, $elapsed, $out, $size |
            Add-Content -Path $runsFile -Encoding UTF8

        if ($code -eq 2) {
            Log "[$index/$($districts.Count)] CAPTCHA $($d.name) - queue stopped. Solve in a visible browser, then rerun."
            Log "RUNNER FINISHED (captcha) done=$doneCount failed=$($failed -join ', ')"
            exit 2
        }

        if ($code -eq 3 -and (Test-Path $out) -and ((Get-Item $out).Length -gt 5000)) {
            Log "[$index/$($districts.Count)] STOPPED EARLY (zero yield) $($d.name) in $([TimeSpan]::FromSeconds($elapsed).ToString('hh\:mm\:ss')) -> $out"
            $districtDone = $true
            $doneCount++
            $consecutiveFail = 0
            break
        }

        if ($code -eq 0 -and (Test-Path $out) -and ((Get-Item $out).Length -gt 5000)) {
            Log "[$index/$($districts.Count)] DONE $($d.name) in $([TimeSpan]::FromSeconds($elapsed).ToString('hh\:mm\:ss')) -> $out"
            $districtDone = $true
            $doneCount++
            $consecutiveFail = 0
            break
        }

        Log "[$index/$($districts.Count)] ATTEMPT FAILED $($d.name) attempt $attempt code=$code after $([TimeSpan]::FromSeconds($elapsed).ToString('hh\:mm\:ss'))"
        if ($attempt -lt 3) {
            $wait = 120 * $attempt
            Log "  waiting $wait s before retrying $($d.name)"
            Start-Sleep -Seconds $wait
        }
    }

    if (-not $districtDone) {
        $failed += $d.name
        $consecutiveFail++
        Log "[$index/$($districts.Count)] FAILED $($d.name) after 3 attempts"
        if ($consecutiveFail -ge 2) {
            Log "2 districts in a row failed - network looks down. Stopping queue."
            Log "RUNNER FINISHED (network) done=$doneCount failed=$($failed -join ', ')"
            exit 1
        }
    }

    if ($PauseSeconds -gt 0 -and -not $DryRun) {
        Start-Sleep -Seconds $PauseSeconds
    }
}

$failedList = if ($failed.Count) { $failed -join ', ' } else { 'none' }
Log "RUNNER FINISHED done=$doneCount failed=$failedList"
