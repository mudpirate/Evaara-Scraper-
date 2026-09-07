# Waits for both Vijayapura workers, then merges to the final ONE Excel.
# Launched in background; logs to _vijaypur_merge.log
$ErrorActionPreference = "SilentlyContinue"
Set-Location "C:\Users\Nomesh\Desktop\evaara\google_scraper"
$w1 = Get-Content "_vijaypur_w1.pid"
$w2 = Get-Content "_vijaypur_w2.pid"
"waiting for workers $w1 and $w2 ..." | Out-File "_vijaypur_merge.log" -Encoding utf8
try { Wait-Process -Id $w1 -ErrorAction SilentlyContinue } catch {}
try { Wait-Process -Id $w2 -ErrorAction SilentlyContinue } catch {}
"both workers exited, merging ..." | Out-File "_vijaypur_merge.log" -Append -Encoding utf8
& python _merge_vijaypur_final.py >> "_vijaypur_merge.log" 2>&1
"watcher done" | Out-File "_vijaypur_merge.log" -Append -Encoding utf8
