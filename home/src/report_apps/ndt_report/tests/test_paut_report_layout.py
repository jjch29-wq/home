import sys
from pathlib import Path
import unittest
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from paut_report_layout import prepare_indication_row


class IndicationLayoutTests(unittest.TestCase):
    def test_merged_template_accepts_all_indication_values(self):
        ws = openpyxl.Workbook().active
        ws.merge_cells('N9:X9')
        prepare_indication_row(ws, 9, True)
        values = {'N9': '203.5', 'R9': '5.8', 'T9': '1.175', 'V9': '0.83', 'AE9': 'Surface'}
        for address, value in values.items():
            ws[address] = value
        self.assertEqual({address: ws[address].value for address in values}, values)
        self.assertNotIn('N9:X9', {str(r) for r in ws.merged_cells.ranges})
        self.assertIn('W9:X9', {str(r) for r in ws.merged_cells.ranges})
        self.assertNotIn('V9:W9', {str(r) for r in ws.merged_cells.ranges})
        self.assertEqual(ws['V9'].border.right.style, 'dotted')
        self.assertEqual(ws['W9'].border.left.style, 'dotted')
        prepare_indication_row(ws, 9, False)
        ws['N9'] = 'NO RECORDABLE INDICATION'
        self.assertIn('N9:X9', {str(r) for r in ws.merged_cells.ranges})
