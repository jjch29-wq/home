"""Billing imports must use the active ledger without silently dropping rows."""
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.billing_daily_import import load_billing_daily_records


class DailyBillingImportTest(unittest.TestCase):
    def test_numeric_blank_material_column_accepts_resolved_film_spec(self):
        app = SimpleNamespace(
            daily_usage_df=pd.DataFrame([{'Date': '2026-09-29', '검사방법': 'RT',
                                          'Material': float('nan'), 'MaterialID': 481.0, 'Usage': 6}]),
            materials_df=pd.DataFrame([{'MaterialID': 481, '품목명': 'Carestream T200-3⅓*6"'}]),
        )
        self.assertEqual(str(app.daily_usage_df['Material'].dtype), 'float64')
        result = load_billing_daily_records(app)
        self.assertEqual(result.iloc[0]['Material'], '3 1/3 x 6"')
        self.assertTrue(pd.isna(app.daily_usage_df.iloc[0]['Material']))

    def test_rt_film_size_is_resolved_from_material_id_in_active_ledger(self):
        app = SimpleNamespace(
            daily_usage_df=pd.DataFrame([{'Date': '2026-09-29', '검사방법': 'RT',
                                          'Material': None, 'MaterialID': 481.0, 'Usage': 6}]),
            materials_df=pd.DataFrame([{'MaterialID': 481, '품목명': 'Carestream T200-3⅓*6" [지급]'}]),
        )
        with patch('pandas.read_excel') as read:
            result = load_billing_daily_records(app)
        read.assert_not_called()
        self.assertEqual(result.iloc[0]['Material'], '3 1/3 x 6"')
        self.assertEqual(result.iloc[0]['Usage'], 6)

    def test_active_ledger_includes_recent_and_noncanonical_method_rows(self):
        original = pd.DataFrame([
            {'Date': '2026-10-22 23:59:59.999', 'Site': '중앙', '검사방법': ' paut ', 'Usage': 2},
            {'Date': '2026-10-21', 'Site': '중앙', '검사방법': 'RT', 'Usage': 3},
            {'Date': '2026-10-23', 'Site': '중앙', '검사방법': 'MT', 'Usage': 4},
            {'Date': '2026-10-21', 'Site': '타현장', '검사방법': 'PT', 'Usage': 5},
        ])
        app = SimpleNamespace(daily_usage_df=original, db_path='active-cloud.xlsx')
        with patch('pandas.read_excel') as read:
            result = load_billing_daily_records(app, '2026-10-21', '2026-10-22', '중앙')
        read.assert_not_called()
        self.assertEqual(result['검사방법'].tolist(), ['RT', 'PAUT'])
        self.assertEqual(result['Usage'].tolist(), [3, 2])
        self.assertEqual(original.iloc[0]['검사방법'], ' paut ')

    def test_disk_fallback_uses_selected_database(self):
        data = pd.DataFrame([{'Date': '2026-10-22', '검사방법': 'MT'}])
        with patch('pandas.read_excel', return_value=data) as read:
            result = load_billing_daily_records(SimpleNamespace(db_path='selected.xlsx'))
        read.assert_called_once_with('selected.xlsx', sheet_name='DailyUsage')
        self.assertEqual(len(result), 1)

    def test_invalid_date_is_reported_instead_of_silently_omitted(self):
        app = SimpleNamespace(daily_usage_df=pd.DataFrame([
            {'Date': 'bad date', '검사방법': 'RT'},
        ]))
        with self.assertRaisesRegex(ValueError, '날짜'):
            load_billing_daily_records(app)

    def test_empty_active_ledger_does_not_restore_stale_disk_records(self):
        app = SimpleNamespace(daily_usage_df=pd.DataFrame(columns=['Date', '검사방법']))
        with patch('pandas.read_excel') as read:
            self.assertTrue(load_billing_daily_records(app).empty)
        read.assert_not_called()


if __name__ == '__main__':
    unittest.main()
