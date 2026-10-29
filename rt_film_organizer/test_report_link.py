import json
import unittest
import tempfile
from pathlib import Path
from report_link import load_reports, suggested_count, build_slots, map_photos, ReportBatch
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

    def test_repair_row_uses_saved_count_without_becoming_original(self):
        item = {'Dwg': 'LINE-A', 'Joint': 'R10R', 'RT_OR': '1', 'RT_RE': '0'}
        self.assertEqual(suggested_count(item), 1)
        slots = build_slots({'items': [item]}, [suggested_count(item)])
        self.assertEqual(len(slots), 1)
        self.assertEqual(slots[0]['joint_no'], 'R10R')
        self.assertEqual(slots[0]['section'], '')
        self.assertEqual(suggested_count(dict(item, RT_RE='2')), 2)
        self.assertEqual(suggested_count(dict(item, RT_OR='0')), 0)

    def test_preview_import_respects_saved_line_and_date_filters(self):
        path = self.tmp_path / 'preview.json'
        rows = [dict(report()['items'][0], drawing_filtered=True),
                dict(report()['items'][1], drawing_filtered=False),
                dict(report()['items'][0], date_filtered=False),
                dict(report()['items'][0], selected=False)]
        path.write_text(json.dumps(rows), encoding='utf-8')
        self.assertEqual(load_reports(path)[0]['items'], rows[:1])
        path.write_text(json.dumps([dict(report(), items=[rows[1]])]), encoding='utf-8')
        self.assertEqual(load_reports(path)[0]['items'], [rows[1]])

    def test_preview_with_no_visible_rows_is_rejected(self):
        path = self.tmp_path / 'empty.json'
        path.write_text(json.dumps([dict(report()['items'][0], drawing_filtered=False)]), encoding='utf-8')
        with self.assertRaises(ValueError):
            load_reports(path)

    def test_saved_repair_sections_preserve_line_joint_and_photo_order(self):
        rows = [dict(Dwg='LINE-A', Joint='R06R', Loc=',N/1', RT_OR='1', RT_RE='0'),
                dict(Dwg='LINE-A', Joint='R10R', Loc=',,N/1', RT_OR='1', RT_RE='0')]
        slots = build_slots({'report_no': 'RT-002', 'items': rows}, [1, 1])
        records = map_photos(slots, ['first.jpg', 'second.jpg'], [0, 1])
        self.assertEqual([(r.source, r.line_no, r.joint_no, r.section) for r in records],
                         [('first.jpg', 'LINE-A', 'R06R', '(2-3)'),
                          ('second.jpg', 'LINE-A', 'R10R', '(3-1)')])
        self.assertTrue(all(r.section_source == '보고서 저장 구간' for r in records))
        self.assertEqual(render_filename('{joint}{section}', records[1]), 'R10R(3-1).JPG')

    def test_saved_sections_support_two_shots_and_reject_count_mismatch(self):
        row = dict(Dwg='LINE-A', Joint='R01R', Loc='N/1,,N/1')
        slots = build_slots({'items': [row]}, [2])
        self.assertEqual([s['section'] for s in slots], ['(1-2)', '(3-1)'])
        slots = build_slots({'items': [row]}, [3])
        self.assertEqual([s['section'] for s in slots], ['', '', ''])


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

    def make_batch(self, counts):
        batch = ReportBatch([f'{i+1:03}.jpg' for i in range(sum(counts))])
        for index, count in enumerate(counts):
            batch.add({'id': str(index), 'mode': 'RT', 'report_no': f'RT-{index+1}',
                       'items': [{'Dwg': f'LINE-{index+1}', 'Joint': 'R01', 'RT_OR': str(count)}]})
        return batch

    def test_six_reports_follow_requested_order_and_boundaries(self):
        batch = self.make_batch([9,6,12,3,9,6])
        batch.move(5, -1)
        batch.generate_all()
        self.assertEqual(len(batch.records()), 45)
        self.assertEqual(batch.entries[0].assignments, list(range(9)))
        self.assertEqual(batch.entries[1].assignments, list(range(9,15)))
        self.assertEqual(batch.entries[4].report['report_no'], 'RT-6')
        self.assertEqual(batch.records()[-1].report_no, 'RT-5')

    def test_local_missing_and_count_change_do_not_shift_other_reports(self):
        batch = self.make_batch([3,3,3]); batch.generate_all()
        before = list(batch.entries[1].assignments)
        batch.entries[0].assignments[1] = None
        self.assertEqual(batch.entries[1].assignments, before)
        batch.entries[0].counts[0] = 2; batch.entries[0].dirty = True
        with self.assertRaises(ValueError): batch.records()
        batch.generate(0)
        self.assertEqual(batch.entries[1].assignments, before)
        self.assertEqual(len(batch.records()), 8)

    def test_cross_report_duplicate_photo_is_rejected(self):
        batch = self.make_batch([3,3]); batch.generate_all()
        batch.entries[1].assignments[0] = 0
        with self.assertRaises(ValueError): batch.records()

    def test_move_keeps_manual_assignment_and_remove_does_not_shift_others(self):
        batch = self.make_batch([3,3,3]); batch.generate_all()
        entry = batch.entries[0]; entry.assignments[1] = None
        batch.move(0,1)
        self.assertIs(batch.entries[1],entry)
        self.assertEqual(entry.assignments,[0,None,2])
        del batch.entries[1]
        self.assertEqual(batch.entries[1].assignments,[6,7,8])
        self.assertEqual(len(batch.records()),6)

    def test_same_report_cannot_be_added_twice(self):
        batch = self.make_batch([3])
        with self.assertRaises(ValueError): batch.add(batch.entries[0].report)
