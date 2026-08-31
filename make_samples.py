"""Generate sample output files matching the exact schemas the three scrapers write.

These are FAKE/demo listings (marked as samples) so you can see the output format
without a live (often blocked) scrape. Real runs overwrite these with live data.
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUT = Path(__file__).parent / "samples"
OUT.mkdir(parents=True, exist_ok=True)

HEAD_FILL = PatternFill('solid', start_color='DDDDDD')
HEAD_FONT = Font(bold=True)
CENTER = Alignment(horizontal='center')


# ---------- Google Maps ----------
def maps():
    cols = ['title', 'totalScore', 'reviewsCount', 'street', 'city', 'state',
            'countryCode', 'website', 'phone',
            *[f'categories/{i}' for i in range(11)], 'url', 'categoryName']
    rows = [
        ['Sai Mobile & Electronics', 4.6, 512, 'Main Road, Bidar', 'Bidar', 'Karnataka', 'IN',
         'https://saimobilebidar.in', '+91 84822 12345', 'mobile_phone_repair', '', '', '', '', '', '', '', '', '', '', '',
         'https://www.google.com/maps?cid=1001', 'Mobile Phone Repair'],
        ['Sri Venkateshwara Mobiles', 4.2, 148, 'Near Bus Stand, Basavakalyan', 'Basavakalyan', 'Karnataka', 'IN',
         'https://venkatmobiles.in', '+91 84822 67890', 'mobile_phone_shop', '', '', '', '', '', '', '', '', '', '', '',
         'https://www.google.com/maps?cid=1002', 'Mobile Phone Shop'],
        ['Cellular Point', 4.0, 96, 'Ganj Road, Bidar', 'Bidar', 'Karnataka', 'IN',
         '', '+91 84822 11223', 'mobile_phone_shop', '', '', '', '', '', '', '', '', '', '', '',
         'https://www.google.com/maps?cid=1003', 'Mobile Phone Shop'],
        ['Mobile World Bidar', 3.8, 41, 'Humnabad Road, Humnabad', 'Humnabad', 'Karnataka', 'IN',
         'https://mobileworldbidar.com', '+91 84822 44556', 'mobile_phone_shop', '', '', '', '', '', '', '', '', '', '', '',
         'https://www.google.com/maps?cid=1004', 'Mobile Phone Shop'],
    ]
    wb = Workbook(); ws = wb.active; ws.title = 'shops'
    ws.append(cols)
    for cell in ws[1]: cell.font = HEAD_FONT; cell.fill = HEAD_FILL; cell.alignment = CENTER
    for r in rows: ws.append(r)
    widths = {'A': 30, 'B': 10, 'C': 12, 'D': 32, 'E': 14, 'F': 12, 'G': 12,
              'H': 28, 'I': 16, 'J': 20, 'U': 40}
    for c, w in widths.items(): ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A2'
    wb.save(OUT / 'maps_mobile_shops_bidar_SAMPLE.xlsx')
# ---------- JustDial ----------
def jd():
    cols = ['name', 'rating', 'reviewsCount', 'address', 'area', 'city', 'pincode',
            'phone', 'whatsapp', 'categoryName', 'url']
    rows = [
        ['Sai Mobile Service Centre', 4.3, 210, 'Plot 12, Main Road, Bidar', 'Main Road', 'Bidar', '585401',
         '+91 98450 12345', '+91 98450 99887', 'Mobile Phone Dealers',
         'https://www.justdial.com/Bidar/Sai-Mobile-Service-Centre'],
        ['Sri Laxmi Mobile Shop', 4.0, 88, 'Basavakalyan Road, Basavakalyan', 'Basavakalyan', 'Basavakalyan', '585327',
         '+91 98452 22334', '+91 98452 77665', 'Mobile Phone Dealers',
         'https://www.justdial.com/Basavakalyan/Sri-Laxmi-Mobile-Shop'],
        ['Bidar Mobile Zone', 3.9, 120, 'Court Complex, Bidar', 'Court Complex', 'Bidar', '585401',
         '+91 98454 55667', '', 'Mobile Phone Dealers',
         'https://www.justdial.com/Bidar/Bidar-Mobile-Zone'],
        ['Cell Point India', 4.1, 152, 'Near Market Yard, Humnabad', 'Market Yard', 'Humnabad', '585330',
         '+91 98456 77889', '+91 98456 11223', 'Mobile Phone Dealers',
         'https://www.justdial.com/Humnabad/Cell-Point-India'],
    ]
    wb = Workbook(); ws = wb.active; ws.title = 'listings'
    ws.append(cols)
    for cell in ws[1]: cell.font = HEAD_FONT; cell.fill = HEAD_FILL; cell.alignment = CENTER
    for r in rows: ws.append(r)
    widths = {'A': 32, 'B': 8, 'C': 12, 'D': 42, 'E': 20, 'F': 16, 'G': 10,
              'H': 16, 'I': 16, 'J': 26, 'K': 50}
    for c, w in widths.items(): ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A2'
    wb.save(OUT / 'jd_mobile_shops_bidar_SAMPLE.xlsx')


# ---------- IndiaMART ----------
def im():
    cols = ['companyName', 'product', 'price', 'moq', 'city', 'state', 'district',
            'pincode', 'pns', 'gst', 'gstVerified', 'rating', 'ratingCount',
            'memberSince', 'annualTurnover', 'category', 'bizType', 'specs',
            'companyUrl', 'productUrl']
    headers = ['Company', 'Product', 'Price', 'MOQ', 'City', 'State', 'District',
               'Pincode', 'Phone (PNS)', 'GST', 'GST Verified', 'Rating',
               'Ratings Count', 'Member Since', 'Annual Turnover', 'Category',
               'Business Type', 'Specifications', 'Company URL', 'Product URL']
    rows = [
        ['Shree Balaji Telecom', 'Mobile Phone Distributor', 'Rs 8,999 / piece', '50 pieces',
         'Bidar', 'Karnataka', 'Bidar', '585401', '+91 98760 11223', '29ABCDE1234F1Z2', 'Yes',
         4.5, 87, '2009', 'Rs 2 - 5 Crore', 'Mobile Phones', 'Distributor',
         'Smartphone dealer, all major brands', 'https://srichand.indiamart.com',
         'https://dir.indiamart.com/impcat/mobile-phones.html'],
        ['Bidar Mobile Network', 'Mobile Accessories', 'Rs 120 / piece', '100 pieces',
         'Humnabad', 'Karnataka', 'Bidar', '585330', '+91 98762 33445', '29FGHIJ5678K3L4', 'Yes',
         4.2, 54, '2014', 'Rs 50 Lakh - 1 Crore', 'Mobile Accessories', 'Manufacturer',
         'Chargers, covers, earphones', 'https://bidarmobile.indiamart.com',
         'https://dir.indiamart.com/impcat/mobile-accessories.html'],
        ['Sri Ganesh Enterprises', 'Mobile Phone Dealer', 'Rs 5,499 / piece', '25 pieces',
         'Basavakalyan', 'Karnataka', 'Bidar', '585327', '+91 98764 55667', '29KLMPQ8907R5S6', 'No',
         4.0, 32, '2017', 'Rs 25 - 50 Lakh', 'Mobile Phones', 'Dealer',
         'Budget to premium smartphones', 'https://sriganesh.indiamart.com',
         'https://dir.indiamart.com/impcat/mobile-phones.html'],
    ]
    wb = Workbook(); ws = wb.active; ws.title = 'suppliers'
    ws.append(headers)
    for cell in ws[1]: cell.font = HEAD_FONT; cell.fill = HEAD_FILL; cell.alignment = CENTER
    for r in rows: ws.append(r)
    widths = {'A': 30, 'B': 30, 'C': 16, 'D': 12, 'E': 14, 'F': 12, 'G': 12, 'H': 10,
              'I': 16, 'J': 18, 'K': 12, 'L': 8, 'M': 12, 'N': 14, 'O': 18, 'P': 24,
              'Q': 16, 'R': 34, 'S': 36, 'T': 46}
    for c, w in widths.items(): ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A2'
    wb.save(OUT / 'im_mobile_shops_bidar_SAMPLE.xlsx')


maps(); jd(); im()
print('Sample files written to', OUT)
for f in sorted(OUT.iterdir()):
    print(' -', f.name)