# Runs the Google Maps scraper for 10 Tumkuru-district cities one by one
# -> single Excel file `mobile_tumkuru_all.xlsx`.
# HEADLESS=1 (background). Set to "0" temporarily if a CAPTCHA needs
# manual solving in the visible browser window.
$env:HEADLESS = "1"
$env:KEEP_ALL = "1"
$env:PYTHONUNBUFFERED = "1"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
Set-Location "C:\Users\Nomesh\Desktop\evaara\google_scraper"
& python scrape.py "mobile phones" "tumkuru city, tumkuru district, karnataka; sira, tumkuru district, karnataka; madhugiri, tumkuru district, karnataka; koratagere, tumkuru district, karnataka; chiknayakanhalli, tumkuru district, karnataka; gubbi, tumkuru district, karnataka; turuvekere, tumkuru district, karnataka; tiptur, tumkuru district, karnataka; kunigal, tumkuru district, karnataka; nelamangala, tumkuru district, karnataka" "mobile_tumkuru_all.xlsx" 2>&1 | Tee-Object -FilePath "run_tumkuru.log"
exit $LASTEXITCODE