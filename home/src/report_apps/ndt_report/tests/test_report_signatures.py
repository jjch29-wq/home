import io
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
import openpyxl
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from report_signatures import append_rt_signatures


class SignatureTests(unittest.TestCase):
    def test_adds_both_signatures_without_removing_original_logo(self):
        def png():
            data = io.BytesIO()
            PILImage.new('RGBA', (40, 20), (0, 0, 0, 255)).save(data, format='PNG')
            data.seek(0)
            return data
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'report.xlsx'
            book = openpyxl.Workbook()
            book.active.add_image(Image(png()), 'A1')
            book.active['D32'] = '우명광'
            book.active['D35'] = '주진철'
            book.save(path)
            with zipfile.ZipFile(path) as archive:
                original_logo = archive.read('xl/media/image1.png')
            append_rt_signatures(path, [(png(), 'E32', 65, 23), (png(), 'E35', 65, 23)])
            with zipfile.ZipFile(path) as archive:
                self.assertEqual(archive.read('xl/media/image1.png'), original_logo)
                self.assertIsNone(archive.testzip())
            result = openpyxl.load_workbook(path)
            self.assertEqual(len(result.active._images), 3)
            self.assertEqual(result.active['D32'].value, '우명광')
            result.close()
