import ast
import json
from pathlib import Path
import sys
import tempfile
import unittest

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
from report_issue_history import record_issue, load_history, item_values, edit_issue, delete_selections, report_output_path


class IssueHistoryTests(unittest.TestCase):
    def test_report_filename_uses_number_without_overwriting(self):
        with tempfile.TemporaryDirectory() as directory:
            first = Path(report_output_path(directory, 'SIS/GI-KDHC-JY-RT-001', '.xlsx'))
            self.assertEqual(first.name, 'SIS-GI-KDHC-JY-RT-001.xlsx')
            first.write_bytes(b'keep')
            second = Path(report_output_path(directory, 'SIS/GI-KDHC-JY-RT-001', '.xlsx'))
            self.assertEqual(second.name, 'SIS-GI-KDHC-JY-RT-001 (2).xlsx')
            self.assertEqual(first.read_bytes(), b'keep')
            self.assertEqual(Path(report_output_path(directory, 'CON', '.xlsm')).name, '_CON.xlsm')

    def test_bulk_delete_same_issue_and_other_issue_is_atomic(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'history.json'
            output = Path(directory) / 'report.xlsx'
            output.touch()
            first = record_issue('RT', 'R1', output, [{'Joint': str(i)} for i in range(4)], path)
            second = record_issue('RT', 'R2', output, [{'Joint': 'A'}], path)
            stale = dict(second, report_no='stale')
            before = path.read_bytes()
            with self.assertRaises(ValueError):
                delete_selections([(first, 0), (stale, 0)], path=path)
            self.assertEqual(path.read_bytes(), before)
            delete_selections([(first, 0), (first, 2), (second, 0)], path=path)
            result = load_history(path)
            self.assertEqual([r['Joint'] for r in result[0]['items']], ['1', '3'])
            self.assertTrue(result[1]['deleted_at'])
            delete_selections([(result[0], 0), (result[0], 1)], whole_issue=True, path=path)
            self.assertTrue(load_history(path)[0]['deleted_at'])

    def test_edit_delete_and_conflict_preserve_original_files(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'history.json'
            output = Path(directory) / 'report.xlsx'
            output.write_bytes(b'original report')
            first = record_issue('RT', 'R1', output,
                                 [{'Joint': 'S01', '_source_id': 'id1'}, {'Joint': 'S02'}], path)
            other = record_issue('RT', 'R2', output, [{'Joint': 'S03'}], path)
            edited = edit_issue(first['id'], 0, {'report_no': 'R1-REV1'},
                                {'Joint': 'R01'}, path=path, expected=first)
            self.assertEqual(edited['report_no'], 'R1-REV1')
            self.assertEqual(edited['items'][0]['Joint'], 'R01')
            self.assertEqual(edited['items'][0]['_source_id'], 'id1')
            self.assertEqual(edited['revisions'][0]['before']['items'][0]['Joint'], 'S01')
            with self.assertRaises(ValueError):
                edit_issue(first['id'], 0, {'report_no': 'stale'}, path=path, expected=first)
            removed = edit_issue(first['id'], 0, delete=True, path=path, expected=edited)
            self.assertEqual([r['Joint'] for r in removed['items']], ['S02'])
            deleted = edit_issue(first['id'], delete=True, path=path, expected=removed)
            self.assertTrue(deleted['deleted_at'])
            self.assertEqual(load_history(path)[1], other)
            self.assertEqual(output.read_bytes(), b'original report')

    def test_snapshot_and_reissue_keep_report_membership(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'history.json'
            output = Path(directory) / 'report.xlsx'
            output.write_bytes(b'report')
            items = [{'Date': '2026-10-21', 'Dwg': 'L1', 'Joint': 'S01',
                      '_source_site': 'central', '_source_id': 'source-1'}]
            first = record_issue('RT', 'RT-001', output, items, path)
            items[0]['Joint'] = 'S02'
            record_issue('RT', 'RT-002', output, items, path)
            records = load_history(path)
            self.assertEqual(len(records), 2)
            self.assertEqual(records[0]['items'][0]['Joint'], 'S01')
            self.assertEqual(records[1]['items'][0]['Joint'], 'S02')
            self.assertNotEqual(records[0]['id'], records[1]['id'])
            values = item_values(first, first['items'][0])
            self.assertEqual(values[1], 'RT-001')
            self.assertIn('2026-10-21', values)
            self.assertIn('L1', values)

    def test_missing_output_and_corrupt_history_are_not_recorded(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'history.json'
            output = Path(directory) / 'report.xlsx'
            with self.assertRaises(FileNotFoundError):
                record_issue('RT', 'R1', output, [], path)
            self.assertFalse(path.exists())
            output.touch()
            path.write_text('{broken', encoding='utf-8')
            with self.assertRaises(ValueError):
                record_issue('RT', 'R1', output, [], path)
            self.assertEqual(path.read_text(encoding='utf-8'), '{broken')

    def test_generation_records_only_after_output_is_saved(self):
        tree = ast.parse((SRC / '비파괴검사보고서.py').read_text(encoding='utf-8-sig'))
        for name in ('_run_rt_process', '_run_pt_process', '_run_mt_process', '_run_pmi_process', '_generate_paut_report'):
            method = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and n.name == name)
            calls = [n for n in ast.walk(method) if isinstance(n, ast.Call)
                     and isinstance(n.func, ast.Attribute)]
            saved = next(n.lineno for n in calls if n.func.attr == 'save')
            recorded = next(n.lineno for n in calls if n.func.attr == '_record_report_issue')
            self.assertLess(saved, recorded)
            if name == '_run_rt_process':
                injected = next(n.lineno for n in calls if n.func.attr == 'inject_drawing_layer')
                self.assertLess(injected, recorded)


if __name__ == '__main__':
    unittest.main()
