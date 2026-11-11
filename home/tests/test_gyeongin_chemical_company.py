from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
import pandas as pd
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.chemical_company import chemical_company_options


class ChemicalCompanyTests(unittest.TestCase):
    def test_options_use_manufacturers_and_exclude_clients(self):
        app = SimpleNamespace(companies=['경인사무소', '현대산업'], co_code_list=['약품사A'],
            materials_df=pd.DataFrame([{'품목명': '침투제', '제조사': '약품사B'},
                                      {'품목명': '검사장비', '제조사': '장비사'}]),
            daily_usage_df=pd.DataFrame([{'업체명': '현대산업', '회사코드': '현대산업', 'NDT_침투제': 1},
                                         {'업체명': '현대산업', '회사코드': '약품사C', 'NDT_침투제': 2}]))
        self.assertEqual(chemical_company_options(app), ['약품사A', '약품사B', '약품사C'])

    def test_empty_manufacturer_stays_empty(self):
        self.assertEqual(chemical_company_options(SimpleNamespace(companies=['경인사무소'])), [])


if __name__ == '__main__': unittest.main()
