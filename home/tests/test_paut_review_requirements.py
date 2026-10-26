import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl

SOURCE = Path(__file__).resolve().parents[1] / 'src' / 'PAUT_검사설정표_생성기.py'
spec = importlib.util.spec_from_file_location('paut_review_generator', SOURCE)
app = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = app
spec.loader.exec_module(app)


class ReviewRequirementsTests(unittest.TestCase):
    def review_values(self, book):
        self.assertEqual(book.sheetnames, ['PAUT Setup'])
        sheet = book['PAUT Setup']
        self.assertEqual(sheet.page_setup.fitToHeight, 1)
        return {sheet.cell(row, label).value: sheet.cell(row, value).value
                for row in range(1, sheet.max_row + 1)
                for label, value in ((1, 4), (7, 9))}

    def test_five_thicknesses_keep_frequency_and_required_width(self):
        with tempfile.TemporaryDirectory() as folder:
            for thickness, width in [(15.87, 15.87), (22.23, 22.23), (25.48, 25),
                                     (28.55, 25), (42.88, 25), (200, 25), (201, 50)]:
                with self.subTest(thickness=thickness):
                    settings = app.ScanSettings(probe_model='2.25L64-A2', frequency_mhz=2.25,
                                                thickness_mm=thickness)
                    path = Path(folder) / f'{thickness}.xlsx'
                    app.create_workbook(settings, str(path))
                    book = openpyxl.load_workbook(path)
                    try:
                        self.assertEqual(book['PAUT Setup']['B14'].value, 2.25)
                        self.assertEqual(book['PAUT Setup']['B14'].number_format, '0.00')
                        review = self.review_values(book)
                        self.assertTrue(review['Additional width on EACH side'].startswith(f'{width:g} mm'))
                        self.assertIn('NOT ESTABLISHED', review['Actual HAZ width'])
                        self.assertIn('NOT DEMONSTRATED', review['Coverage verification'])
                        self.assertEqual(len(book['PAUT Setup']._images), 1)
                    finally:
                        book.close()

    def test_old_settings_accept_new_fields_and_roundtrip(self):
        settings = app.ScanSettings.from_dict({'probe_model': '5L64-A2'})
        self.assertEqual(settings.pulser_voltage, '')
        settings.pulser_voltage = '80 V; group A1 only; setup R1'
        restored = app.ScanSettings.from_dict(app.asdict(settings))
        self.assertEqual(restored.pulser_voltage, settings.pulser_voltage)

    def test_wrong_nominal_frequency_is_rejected(self):
        with self.assertRaisesRegex(ValueError, '2.25 MHz'):
            app.ScanSettings(probe_model='2.25L64-A2', frequency_mhz=2).validate()

    def test_filled_settings_preserve_values_without_claiming_approval(self):
        with tempfile.TemporaryDirectory() as folder:
            settings = app.ScanSettings()
            for key, _ in app.INSTRUMENT_FIELDS:
                setattr(settings, key, 'Example setting - not for examination')
            path = Path(folder) / 'filled.xlsx'
            app.create_workbook(settings, str(path))
            book = openpyxl.load_workbook(path)
            try:
                review = self.review_values(book)
                self.assertIn('SUBJECT TO PROCEDURE APPROVAL', review['Review status'])
                self.assertEqual(review['1. Net digitizing frequency / points / compression'], settings.net_digitizing)
            finally:
                book.close()


if __name__ == '__main__':
    unittest.main()
