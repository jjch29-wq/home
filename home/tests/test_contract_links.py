import csv
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.contract_links import connect_contract, allowed


class ContractLinksTests(unittest.TestCase):
    def test_report_membership_and_repair_photo_use_exact_inspection(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reports = root / '보고서/RT'
            reports.mkdir(parents=True)
            (reports / 'SIS-RT-002.xlsx').touch()
            images = root / '영상화장비사진/중앙지사/영상화_결과'
            (images / 'LINE').mkdir(parents=True)
            (images / 'LINE/R10R(3-1).JPG').touch()
            with (images / '영상화_작업기록.csv').open('w', newline='', encoding='utf-8-sig') as stream:
                writer = csv.DictWriter(stream, fieldnames=['line_no', 'joint_no', 'exam_date', 'report_no', 'section', 'output_name'])
                writer.writeheader()
                writer.writerow(dict(line_no='LINE', joint_no='R10R', exam_date='2026-10-19', report_no='SIS/RT-002', section='(3-1)', output_name='LINE/R10R(3-1).JPG'))
            issue = root / 'history.json'
            issue.write_text(json.dumps([dict(mode='RT', report_no='SIS/RT-002', items=[dict(Dwg='LINE', Joint='R10R', Date='2026-10-19')])]), encoding='utf-8')
            row = {'보고서종류': 'RT', '검사일자': '2026-10-19', '라인번호': 'LINE', '용접부번호': 'R10R'}
            result = connect_contract([row, dict(row, 검사일자='2026-10-20'), dict(row, 용접부번호='R10')], root, issue)
            self.assertEqual(result[0]['보고서번호'], 'SIS/RT-002')
            self.assertEqual(result[0]['RT영상']['3-1'], ['영상화장비사진/중앙지사/영상화_결과/LINE/R10R(3-1).JPG'])
            self.assertNotIn('RT영상', result[1])
            self.assertNotIn('RT영상', result[2])
            self.assertTrue(allowed(root, reports / 'SIS-RT-002.xlsx', 'RT', '보고서'))
            self.assertFalse(allowed(root, reports / 'SIS-RT-002.xlsx', 'PAUT', '보고서'))
            self.assertNotIn('보고서번호', row)

    def test_paut_uses_exact_line_and_joint_and_preserves_multiple_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            folder = root / 'PAUT DATA/LINE'
            folder.mkdir(parents=True)
            for name in ('LINE_S01_90.opd', 'LINE_S01_270.opd', 'LINE_S010.opd', 'LINE_S01R.opd'):
                (folder / name).touch()
            row = {'보고서종류': 'PAUT', '검사일자': '2026-10-19', '라인번호': 'LINE', '용접부번호': 'S01'}
            result = connect_contract([row], root)[0]
            self.assertEqual(len(result['데이터파일']), 2)
            result = connect_contract([row, dict(row, 검사일자='2026-10-20')], root)
            self.assertTrue(all(not r.get('데이터') for r in result))


if __name__ == '__main__':
    unittest.main()
