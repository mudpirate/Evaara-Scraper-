# Solapur finish - Mangalwedha only.
# Karmala was skipped by user request; its checkpoint
# (karmala_mobile_all_partial.xlsx) is kept for a later resume.

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

$talukas = @(
    @{ Name = 'Mangalwedha'; Slug = 'mangalwedha'; Bbox = '17.1865101,75.2798652,17.6169466,75.6571769' }
)

$total = $talukas.Count
Log "RUNNER START pid=$PID talukas=$total step=$step (Mangalwedha only; Karmala skipped by user, checkpoint kept)"

$index = 0
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
            Log "[$index/$total] CAPTCHA $($t.Name) - queue stopped."
            Log "RUNNER FINISHED (captcha)"
            return
        }

        if ($code -eq 0 -and (Test-Path $out)) {
            Log "[$index/$total] DONE $($t.Name) in $elapsed -> $out"
            $done = $true
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
        $failedTalukas += $t.Name
        Log "[$index/$total] FAILED $($t.Name) after 3 attempts"
    }
}

$failedList = if ($failedTalukas.Count) { $failedTalukas -join ', ' } else { 'none' }
Log "RUNNER FINISHED - failed talukas: $failedList"
