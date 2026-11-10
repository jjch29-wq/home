import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.workbook_io import write_workbook, replace_workbook, copy_workbook


class WorkbookIOTests(unittest.TestCase):
    def test_complete_workbook_and_mirror(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'db.xlsx'
            mirror = Path(directory) / 'mirror.xlsx'
            write_workbook(target, {'DailyUsage': pd.DataFrame({'Usage': [16.9]}),
                                    'Materials': pd.DataFrame({'Stock': [10]})})
            copy_workbook(target, mirror)
            self.assertEqual(target.read_bytes(), mirror.read_bytes())
            with zipfile.ZipFile(target) as archive:
                self.assertIsNone(archive.testzip())
            self.assertEqual(pd.read_excel(target, sheet_name='Materials').iloc[0]['Stock'], 10)

    def test_failed_replacement_preserves_original_and_cleans_temporary(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'db.xlsx'
            write_workbook(target, {'DailyUsage': pd.DataFrame({'Usage': [1]})})
            original = target.read_bytes()
            with patch('site_apps.gyeongin.src.workbook_io.os.replace', side_effect=PermissionError):
                with self.assertRaises(PermissionError):
                    write_workbook(target, {'DailyUsage': pd.DataFrame({'Usage': [2]})})
            self.assertEqual(target.read_bytes(), original)
            self.assertEqual(list(Path(directory).iterdir()), [target])

    def test_invalid_archive_never_replaces_original(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'db.xlsx'
            source = Path(directory) / 'broken.xlsx'
            target.write_bytes(b'original')
            source.write_bytes(b'invalid')
            with self.assertRaises(zipfile.BadZipFile):
                replace_workbook(source, target)
            self.assertEqual(target.read_bytes(), b'original')


if __name__ == '__main__':
    unittest.main()
