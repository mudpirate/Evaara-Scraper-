$ErrorActionPreference = 'Continue'
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:HEADLESS = '1'
$env:PARALLEL_TABS = '1'
$env:DELAY_MIN_MS = '400'
$env:DELAY_MAX_MS = '1400'
$dir = 'C:\Users\Nomesh\Desktop\evaara\output\2026-08-15_mobile_phones_bidar'
$scr = 'C:\Users\Nomesh\Desktop\evaara\.scripts'
$p = Start-Process -FilePath 'python' -ArgumentList "`"$scr\scrape.py`" `"mobile phone shops`" `"bidar, humnabad, basavakalyan`" `"$dir\maps_mobile_phones_bidar.xlsx`"" -RedirectStandardOutput "$dir\maps.log" -RedirectStandardError "$dir\maps.err.log" -WindowStyle Hidden -PassThru
$p.Id | Set-Content "$dir\maps.pid"
Write-Host "maps launched pid=$($p.Id)"