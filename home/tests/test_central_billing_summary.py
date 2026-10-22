import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.billing_summary import (
    add_unmatched_record, allocate_summary_amounts, build_company_summaries,
)
from site_apps.central.src.services.billing_totals import grouped_supply_amount


class BillingSummaryTest(unittest.TestCase):
    def test_history_quantity_cannot_inflate_billing_quantity(self):
        records = [dict(date='2026-10-01', ndt_type='PAUT', material_type='300A 이상',
                        work_time='일반', qty=10, unit_price=201863, subtotal=2018630,
                        unit='M', company='한국지역난방공사', loc='열배관') for _ in range(2)]
        history = {'2026-10-01': {'ndt_results': [
            {'검사방법': 'PAUT', '관경': '300A', '업체': 'GS네오텍',
             '구간': 'Sec.1', '라인번호': 'L1', 'Joint No.': '1', 'PAUT': 30},
        ]}}
        summaries = build_company_summaries(records, history, {})
        rows = [r for s in summaries.values() for r in s.values()]
        self.assertEqual(sum(r['qty'] for r in rows), 20)
        self.assertEqual(sum(r['calculated_amt'] for r in rows), 4037260)
        self.assertEqual(rows[0]['places'], 1)
        self.assertNotIn('미지정', summaries)

    def test_missing_250a_is_included_even_when_other_paut_details_exist(self):
        record = dict(date='2026-10-19', ndt_type='PAUT', material_type='250A',
                      work_time='야간', qty=1.6802, unit_price=325451, subtotal=546822,
                      unit='M', company='한국지역난방공사', loc='열배관')
        history = {'2026-10-19': {'ndt_results': [
            {'검사방법': 'PAUT', '관경': '400A', '업체': 'GS네오텍', 'PAUT': 28.088},
        ]}}
        result = build_company_summaries([record], history, {})
        key, row = next(iter(result['미지정'].items()))
        self.assertIn('미연결', key[0])
        self.assertIsNone(row['places'])
        self.assertEqual(row['qty'], 1.6802)
        self.assertEqual(row['calculated_amt'], 546822)

    def test_rounding_is_applied_across_companies_for_each_contract_item(self):
        records = [dict(date=f'2026-10-0{i}', ndt_type='PAUT', material_type='300A 이상',
                        work_time='일반', qty=.5, unit_price=3, subtotal=1,
                        unit='M', company='한국지역난방공사', loc='열배관') for i in (1, 2)]
        history = {r['date']: {'ndt_results': [
            {'검사방법': 'PAUT', '관경': '300A', '업체': f'업체{i}', 'PAUT': .5},
        ]} for i, r in enumerate(records)}
        result = build_company_summaries(records, history, {})
        total = sum(r['calculated_amt'] for s in result.values() for r in s.values())
        self.assertEqual(total, 3)
        self.assertEqual(total, grouped_supply_amount(records, {}))

    def test_history_shortfall_is_preserved_without_double_counting(self):
        record = dict(date='2026-10-01', ndt_type='RT', material_type='3 1/3 x 12"',
                      work_time='일반', qty=12, unit_price=140023, subtotal=1680276,
                      unit='매', company='한국지역난방공사', loc='열배관')
        history = {'2026-10-01': {'ndt_results': [
            {'검사방법': 'RT', '업체': 'GS네오텍', 'RT_OR': 6},
        ]}}
        result = build_company_summaries([record], history, {})
        rows = [r for s in result.values() for r in s.values()]
        self.assertEqual(sum(r['qty'] for r in rows), 12)
        self.assertEqual(sum(r['calculated_amt'] for r in rows), 1680276)

    def test_missing_250a_night_record_preserves_exact_billed_amount(self):
        summary = {}
        record = {'qty': 1.6802, 'subtotal': 546822}
        add_unmatched_record(summary, record, 'PAUT 250A 야간', 'M', 325451)
        allocate_summary_amounts(summary, list(summary))
        key, row = next(iter(summary.items()))
        self.assertIn('미연결', key[0])
        self.assertEqual(key[1], '미지정')
        self.assertIsNone(row['places'])
        self.assertEqual(row['qty'], 1.6802)
        self.assertEqual(row['calculated_amt'], 546822)
        self.assertEqual(69439860 + row['calculated_amt'], 69986682)

    def test_unmatched_and_matched_rows_keep_independent_rounding(self):
        key1 = ('Sec.1', 'Line1', 'PAUT 250A 야간', 'M', 325451)
        key2 = ('Sec.2', 'Line2', 'PAUT 250A 야간', 'M', 325451)
        summary = {key1: {'qty': .5}, key2: {'qty': .5}}
        add_unmatched_record(summary, {'qty': 1.6802, 'subtotal': 546822}, key1[2], 'M', 325451)
        allocate_summary_amounts(summary, list(summary))
        self.assertEqual(summary[key1]['calculated_amt'] + summary[key2]['calculated_amt'], 325451)
        self.assertEqual(sum(r['calculated_amt'] for r in summary.values()), 872273)

    def test_multiple_unmatched_rows_are_accumulated_without_loss(self):
        summary = {}
        for _ in range(2):
            add_unmatched_record(summary, {'qty': 1.6802, 'subtotal': 546822}, 'PAUT 250A 야간', 'M', 325451)
        allocate_summary_amounts(summary, list(summary))
        row = next(iter(summary.values()))
        self.assertEqual(row['qty'], 3.3604)
        self.assertEqual(row['calculated_amt'], 1093644)

    def test_already_represented_rt_and_paut_are_not_added_as_unmatched(self):
        summary = {}
        for record, spec, price in (
            ({'qty': 12, 'subtotal': 1680276}, 'RT', 140023),
            ({'qty': 28.0880, 'subtotal': 5669927}, 'PAUT 300A 이상', 201863),
        ):
            self.assertFalse(add_unmatched_record(
                summary, record, spec, 'M', price, has_history_match=True))
        self.assertFalse(summary)
        self.assertTrue(add_unmatched_record(
            summary, {'qty': 1.6802, 'subtotal': 546822},
            'PAUT 250A 야간', 'M', 325451, has_history_match=False))
        allocate_summary_amounts(summary, list(summary))
        self.assertEqual(69439860 + sum(r['calculated_amt'] for r in summary.values()), 69986682)


if __name__ == '__main__':
    unittest.main()
