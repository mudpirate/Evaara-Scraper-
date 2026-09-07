# Runs the Google Maps scraper for 9 cities one by one -> single Excel file.
# HEADLESS=1 (background). Set to "0" temporarily if a CAPTCHA needs
# manual solving in the visible browser window.
$env:HEADLESS = "1"
$env:KEEP_ALL = "1"
$env:PYTHONUNBUFFERED = "1"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
Set-Location "C:\Users\Nomesh\Desktop\evaara\google_scraper"
& python scrape.py "mobile phones" "raichur, raichur district, karnataka; lingsugur, raichur district, karnataka; manvi, raichur district, karnataka; sindhanur, raichur district, karnataka; mudgal, raichur district, karnataka; siruguppa, ballari district, karnataka; karatagi, koppal district, karnataka; kanakagiri, koppal district, karnataka; gangavathi, koppal district, karnataka" "mobile_raichur_all.xlsx"