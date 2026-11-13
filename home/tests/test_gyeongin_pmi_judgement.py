import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import openpyxl
from site_apps.gyeongin.src.pmi_judgement import judge_measurement, report_result
from site_apps.gyeongin.src.cert_exporter import CertificateExporter


def criterion(key='304'):
    return dict(id=key, label=key, source='Project criteria', ranges={'Ni': '8~11', 'Cr': '18~20', 'Mo': '-'})


class PmiJudgementTests(unittest.TestCase):
    def test_relative_tolerance_boundaries_and_missing_values(self):
        for ni, cr, expected in [('7.2', '16.2', 'PASS'), ('12.1', '22', 'PASS'),
                                 ('7.199', '18', 'FAIL'), ('8', '22.001', 'FAIL'),
                                 ('ND', '18', 'HOLD'), ('', '18', 'HOLD'),
                                 ('NaN', '18', 'HOLD'), ('Infinity', '18', 'HOLD'),
                                 ('ND', '30', 'FAIL')]:
            with self.subTest(ni=ni, cr=cr):
                row = dict(elements={'Ni': ni, 'Cr': cr, 'Mo': 'ND'})
                self.assertEqual(judge_measurement(row, [criterion()], ['Ni', 'Cr', 'Mo'], 10)[0], expected)
        self.assertEqual(judge_measurement(dict(elements={'Ni': '7.2', 'Cr': '18'}),
                                          [criterion()], ['Ni', 'Cr'], 0)[0], 'FAIL')

    def test_multiple_criteria_require_explicit_binding(self):
        row = dict(alloy='SS-304', elements={'Ni': '8', 'Cr': '18'})
        criteria = [criterion(), criterion('316')]
        self.assertEqual(judge_measurement(row, criteria, ['Ni', 'Cr'], 10)[0], 'HOLD')
        row['criterion_id'] = '304'
        self.assertEqual(judge_measurement(row, criteria, ['Ni', 'Cr'], 10)[0], 'PASS')
        row['criterion_id'] = 'deleted'
        self.assertEqual(judge_measurement(row, [criterion()], ['Ni', 'Cr'], 10)[0], 'HOLD')

    def test_one_sided_limits_and_invalid_tolerance(self):
        item = dict(id='limits', label='limits', ranges={'Mn': '<=2', 'Cr': '>=18'})
        row = dict(elements={'Mn': '2.2', 'Cr': '16.2'})
        self.assertEqual(judge_measurement(row, [item], ['Mn', 'Cr'], 10)[0], 'PASS')
        row['elements']['Mn'] = '2.2001'
        self.assertEqual(judge_measurement(row, [item], ['Mn', 'Cr'], 10)[0], 'FAIL')
        self.assertEqual(judge_measurement(row, [item], ['Mn', 'Cr'], 'invalid')[0], 'HOLD')

    def test_no_numeric_checks_or_no_criteria_cannot_pass(self):
        row = dict(elements={'Mo': '1'})
        self.assertEqual(judge_measurement(row, [criterion()], ['Mo'], 10)[0], 'HOLD')
        self.assertEqual(judge_measurement(row, [], ['Mo'], 10)[0], 'HOLD')
        self.assertEqual(report_result([]), 'HOLD')
        self.assertEqual(report_result(['PASS', 'HOLD']), 'HOLD')
        self.assertEqual(report_result(['HOLD', 'FAIL']), 'FAIL')

    def test_report_outputs_chemistry_verdicts_and_overall_result(self):
        rows = [dict(alloy='incorrect instrument grade', elements={'Ni': '8', 'Cr': '18'}),
                dict(alloy='SS-304', elements={'Ni': '1', 'Cr': '18'}),
                dict(elements={'Ni': 'ND', 'Cr': '18'})]
        batch = dict(method='PMI', date='2026-11-13', pmi_measurements=rows,
                     pmi_elements=['Ni', 'Cr'], pmi_criteria=[criterion()])
        with tempfile.TemporaryDirectory() as tmp:
            out = CertificateExporter().export_certificate(batch, Path(tmp) / 'report.xlsx')
            wb = openpyxl.load_workbook(out)
            sheet = wb['PMI(을)']
            self.assertEqual([sheet.cell(r, 14).value for r in range(20, 23)], ['PASS', 'FAIL', 'HOLD'])
            self.assertEqual(sheet['N17'].value, 'FAIL')
            self.assertIn('incorrect instrument grade', sheet['N20'].comment.text)
            self.assertIn('Ni', sheet['T21'].value)
            wb.close()


if __name__ == '__main__':
    unittest.main()
