"""Mobile target/spec persistence without touching live inventory or network."""
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import openpyxl
import pandas as pd
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.mobile import server


class MobileTargetTests(unittest.TestCase):
    def request(self, **changes):
        data = dict(date='2026-11-11', method='RT', inspection_target='배관 외',
                    inspection_spec='두께 12 mm', pipe='250A', ori_joints=2,
                    rep_joints=1, qty=17, ori_qty=17, rep_qty=3, rej_qty=2)
        data.update(changes)
        return server.RecordSaveRequest(**data)

    def test_manual_quantities_and_spec_round_trip(self):
        row = server.build_row_data(self.request())
        self.assertEqual(row['관경(Inch)'], '')
        self.assertEqual(row['Usage'], 17)
        self.assertEqual(row['검사규격'], '두께 12 mm')
        record = server._row_to_record(0, pd.Series(row))
        self.assertEqual(record['inspection_target'], '배관 외')
        self.assertEqual(record['inspection_spec'], '두께 12 mm')
        self.assertEqual(record['rep_qty'], 3)
        self.assertEqual(record['rep_joints'], 1)
        self.assertEqual(record['rej_qty'], 2)

    def test_chemical_manufacturer_is_not_inspection_company(self):
        row = server.build_row_data(self.request(company='현대산업', company_code='약품제조사A'))
        self.assertEqual(row['업체명'], '현대산업')
        self.assertEqual(row['회사코드'], '약품제조사A')
        record = server._row_to_record(0, pd.Series(row))
        self.assertEqual(record['company_code'], '약품제조사A')
        blank = server.build_row_data(self.request(company='현대산업', company_code=''))
        self.assertEqual(blank['회사코드'], '')

    def test_pipe_diameter_and_quantity_still_saved(self):
        row = server.build_row_data(self.request(inspection_target='배관', qty=6))
        self.assertEqual(row['관경(Inch)'], '250A')
        self.assertEqual(row['Usage'], 6)

    def test_old_workbook_gains_fields_on_save_and_update(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'inventory.xlsx'
            wb = openpyxl.Workbook()
            ws = wb.active; ws.title = 'DailyUsage'
            ws.append(['Date', 'Usage', 'EntryTime', '관경(Inch)', '검사방법'])
            wb.save(path); wb.close()
            with patch.object(server, 'get_inventory_paths', return_value=[path]):
                server.save_record(self.request())
                update = server.RecordUpdateRequest(**{**self.request().model_dump(), 'id': 0,
                                                       'inspection_spec': '두께 20 mm', 'rep_qty': 4})
                server.update_record(update)
            wb = openpyxl.load_workbook(path)
            ws = wb['DailyUsage']
            headers = [c.value for c in ws[1]]
            row = dict(zip(headers, [c.value for c in ws[2]]))
            self.assertEqual(row['검사대상'], '배관 외')
            self.assertEqual(row['검사규격'], '두께 20 mm')
            self.assertEqual(row['재검사매수'], 4)
            self.assertEqual(ws.max_row, 2)
            wb.close()


if __name__ == '__main__':
    unittest.main()
