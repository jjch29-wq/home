import json
import unittest
import tempfile
from pathlib import Path
from report_link import load_reports, suggested_count, build_slots, map_photos
from core import render_filename


def report():
    return {"mode": "RT", "report_no": "RT/001", "items": [
        {"Dwg": "LINE-B", "Joint": "S01", "Welder": "W01", "Size": "100A", "RT_OR": "3"},
        {"Dwg": "LINE-A", "Joint": "R06R", "RT_RE": "1"}]}


class ReportLinkTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.tmp_path = Path(temp.name)

    def test_history_keeps_order_and_ignores_deleted_and_other_modes(self):
        tmp_path = self.tmp_path
        path = tmp_path / 'history.json'
        path.write_text(json.dumps([dict(report(), deleted_at='today'), {"mode": "PT"}, report()]), encoding='utf-8-sig')
        data = load_reports(path)
        assert len(data) == 1
        assert [i['Dwg'] for i in data[0]['items']] == ['LINE-B', 'LINE-A']


    def test_preview_json_and_wrong_format(self):
        tmp_path = self.tmp_path
        path = tmp_path / 'preview.json'
        path.write_text(json.dumps(report()['items']), encoding='utf-8')
        assert load_reports(path)[0]['items'] == report()['items']
        path.write_text('{}')
        with self.assertRaises(ValueError): load_reports(path)


    def test_counts_and_repair_never_infer_section(self):
        r = report()
        assert [suggested_count(i) for i in r['items']] == [3, 1]
        assert suggested_count({'Joint': 'R06R'}) == 0
        slots = build_slots(r, [3, 1])
        assert [s['section'] for s in slots] == ['(1-2)', '(2-3)', '(3-1)', '']
        assert slots[-1]['joint_no'] == 'R06R'


    def test_missing_middle_photo_does_not_shift_manual_assignment(self):
        slots = build_slots(report(), [3, 1])
        records = map_photos(slots, ['a.jpg', 'b.jpg', 'repair.jpg', 'extra.jpg'], [0, None, 1, 2])
        assert [r.source for r in records] == ['a.jpg', 'b.jpg', 'repair.jpg']
        assert [r.sequence for r in records] == [1, 3, 1]
        assert records[1].section == '(3-1)'
        assert records[-1].report_item == '2'
        assert render_filename('{report}_{line}_{joint}_{seq}', records[0]) == 'RT-001_LINE-B_S01_1.JPG'


    def test_duplicate_photo_and_invalid_counts_block_apply(self):
        slots = build_slots(report(), [1, 1])
        with self.assertRaises(ValueError): map_photos(slots, ['a.jpg'], [0, 0])
        with self.assertRaises(ValueError): build_slots(report(), [-1, 1])
        with self.assertRaises(ValueError): map_photos(slots, ['a.jpg'], [0, 9])


    def test_one_two_four_shots_require_manual_sections(self):
        for count in [1, 2, 4]:
            assert not any(s['section'] for s in build_slots(report(), [count, 0]))
