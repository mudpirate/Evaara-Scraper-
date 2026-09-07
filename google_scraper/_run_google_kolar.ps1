# Runs the Google Maps scraper for 10 Kolar-district cities one by one
# -> single Excel file `mobile_kolar_all.xlsx`.
# HEADLESS=1 (background). Set to "0" temporarily if a CAPTCHA needs
# manual solving in the visible browser window.
$env:HEADLESS = "1"
$env:KEEP_ALL = "1"
$env:PYTHONUNBUFFERED = "1"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
Set-Location "C:\Users\Nomesh\Desktop\evaara\google_scraper"
& python scrape.py "mobile phones" "kolar city, kolar district, karnataka; gudibanda, kolar district, karnataka; ballapur, kolar district, karnataka; gauribindapur, kolar district, karnataka; sidlaghatta, kolar district, karnataka; mangarpet, kolar district, karnataka; bangarpete, kolar district, karnataka; chintamani, kolar district, karnataka; srinivaspura, kolar district, karnataka; mulhar, kolar district, karnataka" "mobile_kolar_all.xlsx" 2>&1 | Tee-Object -FilePath "run_kolar.log"
exit $LASTEXITCODE