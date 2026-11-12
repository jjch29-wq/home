from pathlib import Path
import sys
import tempfile
import json
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.daily_certificate_context import daily_certificate_context, merge_daily_context
from site_apps.gyeongin.src.cert_exporter import build_cert_batches, find_daily_usage_context
from site_apps.gyeongin.src.pmi_criteria import save_report, restore_report, state_key


class DailyCertificateContextTests(unittest.TestCase):
    def test_imported_context_survives_history_roundtrip_without_daily_dataframe(self):
        source = dict(Date='2026-11-10', Site='Daily project', 장비명='NITON XL2 100G (SN-001)',
                      보조장비='Light', User='Inspector', 발주처='Customer')
        row = dict(업체='Builder', 검사방법='PMI', 검사품명='Structure', 성적서번호='PMI-001', 공사명='N/A')
        persisted = json.loads(json.dumps(merge_daily_context(row, daily_certificate_context(source))))
        batch = build_cert_batches([persisted], target_date='2026-11-13')[0]
        self.assertEqual(batch['project_name'], 'Daily project')
        self.assertEqual(batch['equipment'], source['장비명'])
        self.assertEqual(batch['aux_equipment'], 'Light')
        self.assertEqual(batch['inspector'], 'Inspector')

    def test_current_daily_header_wins_over_stale_saved_output_fields(self):
        row = dict(업체='Builder', 검사방법='PMI', 검사품명='Structure', 성적서번호='PMI-001')
        daily = pd.DataFrame([dict(Date='2026-11-13', 업체명='Builder', 검사방법='PMI',
                                  Site='Current project', 장비명='Current analyzer', 보조장비='Current light',
                                  User='Current inspector', 성적서번호='PMI-001')])
        batch = build_cert_batches([row], target_date='2026-11-13', daily_usage_df=daily)[0]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'state.json'
            stale = dict(batch, project_name='N/A', equipment='', aux_equipment='', inspector='Old inspector',
                         pmi_elements=['Mn'], pmi_criteria=[])
            save_report(stale, path, original_key=state_key(batch))
            restore_report(batch, path)
            self.assertEqual(batch['project_name'], 'Current project')
            self.assertEqual(batch['equipment'], 'Current analyzer')
            self.assertEqual(batch['aux_equipment'], 'Current light')
            self.assertEqual(batch['inspector'], 'Current inspector')
            self.assertEqual(batch['pmi_elements'], ['Mn'])

    def test_context_join_does_not_take_other_report_or_other_method(self):
        df = pd.DataFrame([
            dict(Date='2026-11-13', 업체명='Builder', 검사방법='PMI', 성적서번호='001', Site='A', 장비명='Device A'),
            dict(Date='2026-11-13', 업체명='Builder', 검사방법='PMI', 성적서번호='002', Site='B', 장비명='Device B'),
        ])
        self.assertEqual(find_daily_usage_context(df, '2026-11-13', 'Builder', 'PMI', '002')['equipment'], 'Device B')
        self.assertEqual(find_daily_usage_context(df, '2026-11-13', 'Builder', 'UT')['equipment'], '')

    def test_current_work_log_export_uses_attached_daily_context_after_row_edits(self):
        from types import SimpleNamespace
        from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab
        class Field:
            def __init__(self, value): self.value = value
            def get(self): return self.value
        row = {k: Field(v) for k, v in dict(업체='Builder', 검사방법='PMI', 검사품명='Edited item',
                                          성적서번호='Edited number', 공사명='N/A').items()}
        row['업체']._daily_usage_meta = daily_certificate_context(dict(Site='Original project', 장비명='Saved analyzer'))
        fake_tab = SimpleNamespace(ndt_grid_entries=[row], ndt_selection_vars=[])
        result = DailyWorkLogTab._current_ndt_rows_for_export(fake_tab)[0]
        self.assertEqual(result['공사명'], 'Original project')
        self.assertEqual(result['장비명'], 'Saved analyzer')
        self.assertEqual(result['검사품명'], 'Edited item')


if __name__ == '__main__':
    unittest.main()
