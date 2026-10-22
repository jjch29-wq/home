import copy
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.services.billing_work_log_import import build_work_log_billing_records
from site_apps.central.src.services.billing_summary import build_company_summaries
from site_apps.central.src.services.billing_totals import grouped_supply_amount


class WorkLogBillingImportTest(unittest.TestCase):
    def build(self, history, start='2026-08-01', end='2026-10-22', contracts=None, company_filter='전체'):
        return build_work_log_billing_records(
            history, start, end, material_costs={'PAUT_300A 이상': 201863,
                                               'PAUT_250A': 325451,
                                               'RT_3 1/3 x 6"': 138793,
                                               'RT_3 1/3 x 12"': 205412},
            labor_costs={}, overhead_rate=0, technical_fee_rate=0, contracts=contracts,
            company_filter=company_filter)

    def test_supervisor_details_are_the_only_quantity_source(self):
        history = {'2026-08-07': {'qty_data': {'PAUT_소계': {'금일작업': '999'}}, 'ndt_results': [
            {'업체': 'GS네오텍', '검사방법': 'PAUT', '관경': '400A', 'PAUT': '1.2767',
             '구간': 'Sec.13-2', '라인번호': 'L1'},
            {'업체': 'GS네오텍', '검사방법': 'PAUT', '관경': '400A', 'PAUT': '1.2767',
             '구간': 'Sec.13-2', '라인번호': 'L1'},
            {'검사방법': 'PAUT', 'PAUT': ''},
        ]}}
        original = copy.deepcopy(history)
        records = self.build(history)
        self.assertEqual(history, original)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['qty'], 2.5534)
        self.assertEqual(records[0]['company'], 'GS네오텍')
        self.assertEqual(records[0]['section'], 'Sec.13-2')
        self.assertEqual(records[0]['line'], 'L1')
        self.assertEqual(records[0]['history_rows'], [0, 1])
        summary = build_company_summaries(records, history, {})
        self.assertEqual(sum(v['calculated_amt'] for rows in summary.values() for v in rows.values()),
                         grouped_supply_amount(records, {}))
        self.assertNotIn('미지정', summary)

    def test_rt_total_is_not_increased_by_rejected_films_and_6inch_is_preserved(self):
        history = {'2026-09-29': {'ndt_results': [
            {'업체': '다림건설', '검사방법': 'RT', '규격': '3 1/3 X 6"',
             'RT_OR': '6', 'RT_RE': '2', '관경': '80A', '라인번호': 'B-87'},
        ]}}
        record = self.build(history)[0]
        self.assertEqual(record['qty'], 6)
        self.assertEqual(record['material_type'], '3 1/3 x 6"')
        self.assertEqual(record['subtotal'], 832758)

    def test_explicit_night_shift_and_date_range_are_preserved(self):
        detail = {'업체': 'GS네오텍', '검사방법': 'PAUT', '관경': '250A',
                  'PAUT': '1.6802', '근무구분': '야간'}
        history = {day: {'ndt_results': [dict(detail)]}
                   for day in ('2026-08-07', '2026-10-19', '2026-10-23')}
        records = self.build(history, '2026-10-19', '2026-10-22')
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['date'], '2026-10-19')
        self.assertEqual(records[0]['work_time'], '야간')
        self.assertEqual(records[0]['material_type'], '250A')
        self.assertEqual(records[0]['subtotal'], 546822)

    def test_missing_rt_spec_is_reported_instead_of_assumed_12inch(self):
        history = {'2026-10-06': {'ndt_results': [
            {'검사방법': 'RT', 'RT_OR': '3'},
        ]}}
        with self.assertRaisesRegex(ValueError, '2026-10-06.*1행.*규격'):
            self.build(history)

    def test_summary_uses_only_selected_source_rows(self):
        history = {'2026-10-06': {'ndt_results': [
            {'업체': '업체A', '검사방법': 'RT', '규격': '3 1/3 X 12"',
             'RT_OR': '3', '구간': 'Sec.1', '라인번호': 'L1'},
            {'업체': '업체B', '검사방법': 'RT', '규격': '3 1/3 X 6"',
             'RT_OR': '6', '구간': 'Sec.2', '라인번호': 'L2'},
        ]}}
        records = self.build(history)
        selected = [r for r in records if r['company'] == '업체B']
        summaries = build_company_summaries(selected, history, {})
        self.assertEqual(set(summaries), {'업체B'})
        self.assertEqual(next(iter(summaries['업체B']))[1], 'L2')
        self.assertEqual(sum(v['qty'] for v in summaries['업체B'].values()), 6)

    def test_contract_price_is_applied_and_cost_components_balance(self):
        history = {'2026-08-07': {'ndt_results': [
            {'검사방법': 'PAUT', '관경': '400A', 'PAUT': '2'},
        ]}}
        record = self.build(history, contracts={'열배관_일반_PAUT_300A 이상': {'c_price': 300000}})[0]
        self.assertEqual(record['subtotal'], 600000)
        self.assertEqual(sum(record[k] for k in ('mat_cost', 'lab_cost', 'overhead', 'tech')), 600000)

    def test_company_filter_uses_supervisor_log_contractor(self):
        history = {'2026-08-07': {'ndt_results': [
            {'업체': '업체A', '검사방법': 'PAUT', '관경': '400A', 'PAUT': '2'},
            {'업체': '업체B', '검사방법': 'PAUT', '관경': '400A', 'PAUT': '3'},
        ]}}
        records = self.build(history, company_filter='업체B')
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['qty'], 3)
        self.assertEqual(records[0]['company'], '업체B')


if __name__ == '__main__':
    unittest.main()
