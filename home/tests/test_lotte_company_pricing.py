import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.services.company_pricing import company_conditions, get_company_price, resolve_item_price


class CompanyPricingTests(unittest.TestCase):
    def test_company_and_method_must_match(self):
        for company, method in [('다른업체', 'MT'), ('세영정공', 'PT'), ('피테크산업', 'MT')]:
            self.assertEqual(company_conditions(company, method), ())
            self.assertIsNone(get_company_price(company, method, 'EA당', 1))

    def test_seyoung_explicit_length_conditions(self):
        for method in ('UT', 'MT'):
            for condition, price in zip(company_conditions('세영정공', method), (20000, 25000, 30000)):
                result = get_company_price('(주) 세영정공', method, condition, 2)
                self.assertEqual(result['fee'], price * 2)
        self.assertIsNone(get_company_price('세영정공', 'UT', '', 2))

    def test_haneun_lump_sum_and_meter_threshold(self):
        self.assertEqual(get_company_price('한은이앤씨', 'MT', '1m', 10)['fee'], 130000)
        self.assertEqual(get_company_price('한은이앤씨', 'MT', '1m', 12)['fee'], 132000)
        self.assertEqual(get_company_price('한은이앤씨', 'MT', '1식', 1)['fee'], 130000)

    def test_pitech_travel_is_added_once(self):
        result = get_company_price('피테크', 'MT', 'EA당', 3)
        self.assertEqual(result['price'], 13000)
        self.assertEqual(result['travel'], 150000)
        self.assertEqual(result['fee'], 189000)

    def test_invalid_quantity(self):
        for quantity in (-1, float('nan'), float('inf')):
            self.assertIsNone(get_company_price('피테크', 'MT', 'EA당', quantity))

    def test_ham_rates_and_travel(self):
        for method, prices in [('UT', (11000, 12000, 15000)), ('MT', (13000, 22000, 33000))]:
            for condition, price in zip(company_conditions('함감속기', method), prices):
                result = get_company_price('함감속기', method, condition, 2)
                self.assertEqual(result['fee'], 2 * price + 100000)

    def test_ilshin_only_applies_to_ut(self):
        self.assertEqual(company_conditions('일신기계', 'MT'), ())
        for condition, price in zip(company_conditions('일신기계', 'UT'), (275000, 375000)):
            self.assertEqual(get_company_price('일신기계', 'UT', condition, 1)['fee'], price)

    def test_diameter_boundaries(self):
        for diameter, price in [(150, 13000), (200, 13000), (200.1, 22000), (400, 22000), (400.1, 33000)]:
            self.assertEqual(resolve_item_price('함감속기', 'MT', diameter, '', 1)['price'], price)
        for diameter, price in [(650, 275000), (650.1, 375000)]:
            self.assertEqual(resolve_item_price('일신기계', 'UT', diameter, '', 1)['price'], price)

    def test_length_conditions_and_invalid_diameter(self):
        self.assertEqual(resolve_item_price('세영정공', 'MT', 100, '길이 없음, D100 이상', 2)['fee'], 40000)
        with self.assertRaises(ValueError):
            resolve_item_price('세영정공', 'UT', 99, '길이 1m, D100 이상', 1)
        for diameter in ('', 'abc', 'nan', 'inf', '-1', '0'):
            with self.assertRaises(ValueError):
                resolve_item_price('함감속기', 'MT', diameter, '', 1)

    def test_unmatched_method_keeps_manual_price(self):
        self.assertIsNone(resolve_item_price('일신기계', 'MT', '', '', 1))


if __name__ == '__main__':
    unittest.main()
