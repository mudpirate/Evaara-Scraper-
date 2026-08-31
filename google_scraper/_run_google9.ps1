# Runs the Google Maps scraper for 9 cities one by one -> single Excel file.
# HEADLESS=0 (visible browser) so that if Google shows a CAPTCHA it can be
# solved right in the window; the scraper auto-resumes once it clears.
# The persistent browser profile (~/.cache/maps-scraper/profile) keeps working
# for future headless runs afterwards.
$env:HEADLESS = "0"
$env:KEEP_ALL = "1"
$env:PYTHONUNBUFFERED = "1"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
Set-Location "C:\Users\Nomesh\Desktop\evaara\google_scraper"
& python scrape.py "mobile phones" "raichur, raichur district, karnataka; lingsugur, raichur district, karnataka; manvi, raichur district, karnataka; sindhanur, raichur district, karnataka; mudgal, raichur district, karnataka; siruguppa, ballari district, karnataka; karatagi, koppal district, karnataka; kanakagiri, koppal district, karnataka; gangavathi, koppal district, karnataka" "mobile_raichur_all.xlsx"