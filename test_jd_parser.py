import importlib.util
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).parent / '.scripts' / 'jd.py'
spec = importlib.util.spec_from_file_location('jd_scraper', MODULE_PATH)
jd = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(jd)


class JustDialParserTests(unittest.TestCase):
    def test_location_fallback_ignores_district_and_state_words(self):
        self.assertEqual(jd.fallback_location('bidar district karnataka'), ('Bidar', None))
        self.assertEqual(
            jd.fallback_location('koramangala bangalore'),
            ('Bangalore', 'Koramangala'),
        )

    def test_normalize_phone_accepts_indian_mobile_only(self):
        self.assertEqual(jd.normalize_phone('09606202743'), '+91 9606202743')
        self.assertEqual(jd.normalize_phone('+91 98450 12345'), '+91 9845012345')
        self.assertEqual(jd.normalize_phone('1234567890'), '')

    def test_dom_cards_keep_real_business_fields(self):
        rows = jd.parse_dom_cards([
            {
                'name': 'Sohel Mobile and Repair Centre',
                'rating': 'Ratings : 2.5',
                'reviewsText': '2 Ratings',
                'address': 'Chidri Ring Road Chidri, Bidar',
                'area': 'Chidri',
                'city': 'Bidar',
                'pincode': '',
                'phone': '09606202743',
                'categoryName': 'Mobile Phone Dealers',
                'url': 'https://www.justdial.com/Bidar/Sohel-Mobile/9999P8482_BZDET',
                'docid': '9999P8482',
            },
            {
                'name': 'IFSC Code guide',
                'address': '',
                'url': 'https://www.justdial.com/guides/ifsc-code/example',
            },
        ])

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['name'], 'Sohel Mobile and Repair Centre')
        self.assertEqual(rows[0]['address'], 'Chidri Ring Road Chidri, Bidar')
        self.assertEqual(rows[0]['phone'], '+91 9606202743')

    def test_stale_json_rows_without_addresses_are_rejected(self):
        self.assertFalse(jd.is_valid_listing({
            'name': 'IFSC Code: KKBK0009306',
            'address': '',
            'url': 'https://www.justdial.com/guides/ifsc-code/example',
        }))
        self.assertFalse(jd.is_valid_listing({
            'name': 'Mobile Phone Dealers in Bangalore',
            'address': '',
            'url': 'https://www.justdial.com/Bangalore/Mobile-Phone-Dealers/nct-1',
        }))


if __name__ == '__main__':
    unittest.main()
