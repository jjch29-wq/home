from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.cert_exporter import CertificateExporter, build_cert_batches
from site_apps.gyeongin.src.certificate_photos import batch_photos


class CertificatePhotoTests(unittest.TestCase):
    def test_row_links_and_ambiguous_legacy_photos(self):
        row = {'업체': 'A', '검사방법': 'MT', 'Joint No.': '1', 'photo_row_id': 'r1'}
        other = dict(row, photo_row_id='r2', 업체='B')
        batch = {'rows': [row], 'date': '2026-11-13', 'method': 'MT', 'company': 'A'}
        history = {'2026-11-13': {'process_photos': [
            {'id': '1', 'row_id': 'r1', 'company': 'A', 'process': 'MT', 'file_path': '1.png'},
            {'id': '2', 'row_id': 'r2', 'company': 'B', 'process': 'MT', 'file_path': '2.png'},
            {'id': '3', 'joint_no': '1', 'process': 'MT', 'file_path': '3.png'},
        ]}}
        photos = batch_photos(history, batch, '.', [row, other])
        self.assertEqual([p['id'] for p in photos], ['1', '3'])
        self.assertEqual([p['include'] for p in photos], [True, False])

    def test_mt_pt_appendices_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            image = root / 'photo.png'
            Image.new('RGB', (800, 400), 'red').save(image)
            for method in ('MT', 'PT'):
                for count in (0, 1, 4, 5):
                    with self.subTest(method=method, count=count):
                        batch = build_cert_batches([{'업체': 'A', '검사방법': method, '검사품명': 'SHAFT', 'photo_row_id': 'r1'}], target_date='2026-11-13')[0]
                        batch['certificate_photos'] = [dict(file_path=str(image), description=f'Photo {i}', include=True) for i in range(count)]
                        output = root / f'{method}{count}.xlsx'
                        CertificateExporter().export_certificate(batch, output)
                        wb = openpyxl.load_workbook(output)
                        pages = [ws for ws in wb if '(사진 ' in ws.title]
                        self.assertEqual(len(pages), (count + 3) // 4)
                        self.assertEqual(sum(len([p for p in ws._images if p.anchor._from.row >= 9]) for ws in pages), count)
                        for ws in pages:
                            col = 19 if method == 'MT' else 15
                            self.assertEqual(ws.cell(6, col).value, batch['report_no'])
                            for picture in ws._images:
                                if picture.anchor._from.row >= 9:
                                    self.assertAlmostEqual(picture.anchor.ext.cx / picture.anchor.ext.cy, 2, places=4)
                        wb.close()


if __name__ == '__main__':
    unittest.main()
