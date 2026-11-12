from pathlib import Path
import sys
import unittest
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.ndt_material_choices import (
    material_choices, filtered_material_choices, report_material_presets,
)
from site_apps.gyeongin.src.cert_exporter import build_cert_batches


class NdtMaterialChoicesTests(unittest.TestCase):
    def test_report_evidence_is_available_for_each_method(self):
        data = report_material_presets()
        for method, value in [('MT', 'SCM420'), ('UT', 'STS310S'), ('PT', 'STS304'),
                              ('RT', 'STS316'), ('PAUT', 'TP347H'), ('PMI', 'A351 CF3M')]:
            self.assertIn(value, data['methods'][method])
            self.assertTrue(any(e['method'] == method and value in e['materials']
                                and e['source'] and e['cell'] for e in data['evidence']))
        all_values = {v for values in data['methods'].values() for v in values}
        for invalid in ('WP034', 'SPTG', 'HSN355', '1ST PINION', '12', 'SCH160'):
            self.assertNotIn(invalid, all_values)

    def test_company_method_first_with_other_materials_still_available(self):
        data = report_material_presets()
        preferred = data['companies']['함감속기제작소']['MT']
        values = material_choices('함감속기제작소', 'MT', history=['CUSTOM-MATERIAL'])
        self.assertEqual(values[:len(preferred)], preferred)
        self.assertIn('CUSTOM-MATERIAL', values)
        self.assertIn('STS304', values)
        self.assertEqual(len(values), len(set(values)))

    def test_search_including_multiple_material_entry(self):
        values = ['SCM420', 'SCM440', 'A182 F316L', 'STS304']
        self.assertEqual(filtered_material_choices(values, 'scm4'), ['SCM420', 'SCM440'])
        self.assertEqual(filtered_material_choices(values, 'A182-F316'), ['A182 F316L'])
        self.assertEqual(filtered_material_choices(values, 'SCM420, scm44'), ['SCM440'])
        self.assertEqual(filtered_material_choices(values, 'SCM420'), values)

    def test_refresh_keeps_direct_input_and_certificate_material(self):
        from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab
        class Field:
            def __init__(self, value):
                self.value, self.options = value, {}
            def get(self):
                return self.value
            def configure(self, **kwargs):
                self.options.update(kwargs)
        row = {key: Field(value) for key, value in dict(
            업체='함감속기제작소', 검사방법='MT', 재질='SCM420, CUSTOM-X', 요구재질='Special grade').items()}
        DailyWorkLogTab._sync_ndt_material_choices(SimpleNamespace(), row)
        self.assertEqual(row['재질'].get(), 'SCM420, CUSTOM-X')
        self.assertEqual(row['요구재질'].get(), 'Special grade')
        self.assertIn('SCM420', row['재질'].options['values'])
        batch = build_cert_batches([{key: field.get() for key, field in row.items()}], target_date='2026-11-13')[0]
        self.assertEqual(batch['material'], 'SCM420, CUSTOM-X')


if __name__ == '__main__':
    unittest.main()
