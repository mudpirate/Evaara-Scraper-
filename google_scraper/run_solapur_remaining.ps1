# Solapur remaining talukas - sequential Google Maps grid runs
# Fresh full run: all 3 keywords per taluka, one taluka after another.
# Restart-safe: skips talukas whose final file already exists.
# Network-resilient: a failed taluka is retried up to 3 times (seeded from its
# checkpoint on retries) before moving on.

$ErrorActionPreference = 'Continue'

$root   = 'C:\Users\Nom\Evaara-Scraper-\google_scraper'
$python = "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"
if (-not (Test-Path $python)) { $python = 'python' }

$outDir = Join-Path $root 'solapur_talukas'
$logDir = Join-Path $outDir '_logs'
New-Item -ItemType Directory -Force -Path $outDir, $logDir | Out-Null
$statusFile = Join-Path $logDir 'status.txt'

function Log($message) {
    $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $message
    Add-Content -Path $statusFile -Value $line -Encoding UTF8
    Write-Output $line
}

$env:HEADLESS          = '1'
$env:KEEP_ALL          = '0'
$env:PYTHONUTF8        = '1'
$env:PYTHONIOENCODING  = 'utf-8'
$env:PYTHONUNBUFFERED  = '1'

$keywords = 'cell phone store, smartphone accessories, mobile phone repair shop'
$step     = '0.05'
$minStep  = '0.01'

# bboxes from OpenStreetMap administrative boundaries (south,west,north,east)
$talukas = @(
    @{ Name = 'Pandharpur';  Slug = 'pandharpur';  Bbox = '17.5200631,75.0953292,17.9341523,75.5724620' },
    @{ Name = 'Malshiras';   Slug = 'malshiras';   Bbox = '17.5880669,74.6082682,18.0308931,75.2136468';
       Keywords = 'smartphone accessories, mobile phone repair shop';
       Seed = 'malshiras_mobile_all_partial.xlsx' },
    @{ Name = 'Sangola';     Slug = 'sangola';     Bbox = '17.1463676,74.8485017,17.6757141,75.3375434' },
    @{ Name = 'Akkalkot';    Slug = 'akkalkot';    Bbox = '17.3027059,75.8905485,17.7276394,76.4277000' },
    @{ Name = 'Madha';       Slug = 'madha';       Bbox = '17.8793089,75.0204634,18.2632616,75.7041263' },
    @{ Name = 'Mohol';       Slug = 'mohol';       Bbox = '17.5299461,75.3573703,18.0112227,75.8230376' },
    @{ Name = 'Karmala';     Slug = 'karmala';     Bbox = '18.0878108,74.7958660,18.5537936,75.4035178' },
    @{ Name = 'Mangalwedha'; Slug = 'mangalwedha'; Bbox = '17.1865101,75.2798652,17.6169466,75.6571769' }
)

$total = $talukas.Count
Log "RUNNER START pid=$PID talukas=$total step=$step (full 3-keyword runs)"

$index = 0
$consecutiveFail = 0
$failedTalukas = @()

foreach ($t in $talukas) {
    $index++
    $out = Join-Path $outDir ($t.Slug + '_mobile_all.xlsx')
    if ((Test-Path $out) -and ((Get-Item $out).Length -gt 5000)) {
        Log "[$index/$total] SKIP $($t.Name) (output exists)"
        continue
    }

    $logFile = Join-Path $logDir ($t.Slug + '.log')
    $done = $false

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $kw = if ($t.ContainsKey('Keywords')) { $t.Keywords } else { $keywords }
        $seedArgs = @()
        $seedPath = Join-Path $outDir ($t.Slug + '_mobile_all_partial.xlsx')
        if ($t.ContainsKey('Seed')) {
            $namedSeed = Join-Path $outDir $t.Seed
            if (Test-Path $namedSeed) { $seedPath = $namedSeed }
        }
        if (Test-Path $seedPath) {
            $seedArgs = @('--seed', $seedPath)
            Log "[$index/$total] RESUMING $($t.Name) from checkpoint (attempt $attempt)"
        } else {
            Log "[$index/$total] START $($t.Name) attempt $attempt bbox=$($t.Bbox)"
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $scrape = Join-Path $root 'scrape.py'
        & $python $scrape $kw --bbox $t.Bbox --step $step --min-step $minStep --yes --out $out @seedArgs 2>&1 | Tee-Object -FilePath $logFile -Append
        $code = $LASTEXITCODE
        $sw.Stop()
        $elapsed = $sw.Elapsed.ToString('hh\:mm\:ss')

        if ($code -eq 2) {
            Log "[$index/$total] CAPTCHA $($t.Name) - queue stopped. Solve in a visible browser, then rerun."
            Log "RUNNER FINISHED (captcha)"
            return
        }

        if ($code -eq 0 -and (Test-Path $out)) {
            Log "[$index/$total] DONE $($t.Name) in $elapsed -> $out"
            $done = $true
            $consecutiveFail = 0
            break
        }

        Log "[$index/$total] ATTEMPT FAILED $($t.Name) attempt $attempt code=$code after $elapsed"
        if ($attempt -lt 3) {
            $wait = 120 * $attempt
            Log "[$index/$total] waiting $wait s before retrying $($t.Name)"
            Start-Sleep -Seconds $wait
        }
    }

    if (-not $done) {
        $consecutiveFail++
        $failedTalukas += $t.Name
        Log "[$index/$total] FAILED $($t.Name) after 3 attempts"
        if ($consecutiveFail -ge 2) {
            Log "2 talukas in a row failed - network looks down. Stopping queue."
            Log "RUNNER FINISHED (network)"
            return
        }
    }
}

$failedList = if ($failedTalukas.Count) { $failedTalukas -join ', ' } else { 'none' }
Log "RUNNER FINISHED - failed talukas: $failedList"
