import ast
import hashlib
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

REPORT_ROOT = Path(__file__).resolve().parents[1]
SRC = REPORT_ROOT / 'src'
sys.path.insert(0, str(SRC))
from ndt_history_import import SITES, history_path, read_results


class HistoryImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)

    def write(self, site, content):
        path = history_path(site, self.base)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(content, ensure_ascii=False), encoding='utf-8')
        return path

    def test_each_site_uses_its_own_work_log(self):
        for index, site in enumerate(SITES):
            self.write(site, {'2026-09-05': {'ndt_results': [{'검사방법': 'RT', 'Joint No.': str(index)}]}})
        for index, site in enumerate(SITES):
            rows = read_results(site, 'KOGAS', self.base)
            self.assertEqual(rows[0]['Joint No.'], str(index))
            self.assertEqual(rows[0]['_source_site'], site)

    def test_no_shared_file_fallback(self):
        (self.base / 'daily_work_history.json').write_text('{}', encoding='utf-8')
        with self.assertRaises(FileNotFoundError):
            read_results(next(iter(SITES)), 'RT', self.base)

    def test_method_filter_and_read_only(self):
        site = next(iter(SITES))
        path = self.write(site, {'2026-09-05': {'ndt_results': [
            {'검사방법': 'MT', '결과': '합격'}, {'검사방법': 'PMI'},
            {'검사방법': 'PT/RT'}, {'검사방법': 'PAUT'}, {}, {'검사방법': None}]}})
        before = hashlib.sha256(path.read_bytes()).digest()
        for mode in ('MT', 'PMI', 'PT', 'RT', 'PAUT'):
            rows = read_results(site, mode, self.base)
            self.assertEqual(len(rows), 1)
            rows[0]['Date'] = 'modified snapshot'
        self.assertEqual(before, hashlib.sha256(path.read_bytes()).digest())

    def test_malformed_data_never_returns_partial_rows(self):
        site = next(iter(SITES))
        for value in ([], {'date': []}, {'date': {'ndt_results': {}}},
                      {'date': {'ndt_results': [{'검사방법': 'RT'}, None]}}):
            self.write(site, value)
            with self.assertRaises(ValueError):
                read_results(site, 'RT', self.base)
        path = history_path(site, self.base)
        path.write_text('{', encoding='utf-8')
        with self.assertRaises(ValueError):
            read_results(site, 'RT', self.base)

    def test_empty_history(self):
        site = next(iter(SITES))
        self.write(site, {})
        self.assertEqual(read_results(site, 'RT', self.base), [])

    def report_method(self, name='load_daily_work_history', messagebox=None):
        path = next(SRC.glob('*검사보고서.py'))
        source = path.read_text(encoding='utf-8-sig')
        compile(source, str(path), 'exec')
        tree = ast.parse(source)
        method = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and n.name == name)
        messages = []
        if messagebox is None:
            messagebox = SimpleNamespace(showinfo=lambda *args: messages.append(args),
                                         showerror=lambda *args: self.fail(str(args)),
                                         showwarning=lambda *args: messages.append(args),
                                         askyesno=lambda *args: True)
        context = {'messagebox': messagebox,
                   'os': SimpleNamespace(path=SimpleNamespace(exists=lambda path: True))}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(path), 'exec'), context)
        return context[name]

    def test_cancel_does_not_touch_report(self):
        app = SimpleNamespace(root=None)
        with patch('ndt_history_import.select_history_rows', return_value=[]):
            self.report_method()(app)
        self.assertFalse(hasattr(app, 'data'))

    def test_selected_rt_row_preserves_existing_mapping(self):
        data, indices, inserted = [], [], []
        tree = SimpleNamespace(insert=lambda *args, **kw: inserted.append(kw))
        notebook = SimpleNamespace(select=lambda: 'tab', tab=lambda *args: 'RT')
        keys = ['Date', 'Sec', 'Loc', 'Joint', 'Size', 'Result', 'ISO']
        app = SimpleNamespace(root=None, mode_notebook=notebook, rt_preview_nb=notebook,
            _get_mode_info=lambda mode: (tree, indices, data, keys), update_date_listbox=lambda **kw: None)
        row = {'Date': '2026-09-05', '검사방법': 'RT', '구간': 'A', '구간정보': 'B',
               'Joint No.': 'J1', '라인번호': 'L1', '결과': '합격', '_source_site': '롯데건설 바이오로직스',
               '_source_path': '/test/history.json'}
        with patch('ndt_history_import.select_history_rows', return_value=[row]):
            self.report_method()(app)
        self.assertEqual((data[0]['Sec'], data[0]['Loc'], data[0]['Joint']), ('A', 'B', 'J1'))
        self.assertEqual(data[0]['_source_site'], row['_source_site'])
        self.assertEqual(indices, [0])
        self.assertEqual(inserted[0]['tags'], ('pass',))

    def test_get_mode_info_supports_mt(self):
        app = SimpleNamespace(mt_preview_tree='tree', mt_item_idx_map='idx',
                              mt_extracted_data='data', mt_column_keys='keys')
        self.assertEqual(self.report_method('_get_mode_info')(app, 'MT'),
                         ('tree', 'idx', 'data', 'keys'))

    def test_duplicate_source_id_is_not_imported_twice(self):
        data, indices, inserted = [], [], []
        tree = SimpleNamespace(insert=lambda *args, **kw: inserted.append(kw))
        notebook = SimpleNamespace(select=lambda: 'tab', tab=lambda *args: 'PT')
        keys = ['Date', 'Joint No.', "Th'k(mm)", 'Remarks', 'Result']
        app = SimpleNamespace(root=None, mode_notebook=notebook,
            _get_mode_info=lambda mode: (tree, indices, data, keys), update_date_listbox=lambda **kw: None)
        row = {'Date': '2026-09-05', '검사방법': 'PT', 'Joint No.': 'J1', '관경': '1234',
               '결과': '합격', '_source_site': '한국지역난방 중앙지사',
               '_source_path': '/test/history.json', '_source_id': 'same-row'}
        with patch('ndt_history_import.select_history_rows', return_value=[row]):
            self.report_method()(app)
            self.report_method()(app)
        self.assertEqual(len(data), 1)
        self.assertEqual(len(inserted), 1)
        self.assertEqual(data[0]["Th'k(mm)"], '')
        self.assertEqual(data[0]['Remarks'], '두께 확인 필요')

    def test_mixed_site_import_can_be_rejected(self):
        data, indices, inserted = [{'_source_site': '한국지역난방 중앙지사'}], [], []
        tree = SimpleNamespace(insert=lambda *args, **kw: inserted.append(kw))
        notebook = SimpleNamespace(select=lambda: 'tab', tab=lambda *args: 'PT')
        app = SimpleNamespace(root=None, mode_notebook=notebook,
            _get_mode_info=lambda mode: (tree, indices, data, ['Date', 'Result']),
            update_date_listbox=lambda **kw: self.fail('should not update'))
        row = {'Date': '2026-09-05', '검사방법': 'PT', '결과': '합격',
               '_source_site': '롯데건설 바이오로직스', '_source_path': '/test/history.json',
               '_source_id': 'lotte-row'}
        messagebox = SimpleNamespace(showinfo=lambda *args: None,
                                     showerror=lambda *args: self.fail(str(args)),
                                     askyesno=lambda *args: False)
        with patch('ndt_history_import.select_history_rows', return_value=[row]):
            self.report_method(messagebox=messagebox)(app)
        self.assertEqual(len(data), 1)
        self.assertEqual(inserted, [])

    def test_run_process_uses_pt_and_mt_targets(self):
        calls = []
        notebook = SimpleNamespace(select=lambda: 'tab', tab=lambda *args: 'PT')
        app = SimpleNamespace(config={}, save_settings=lambda: None, mode_notebook=notebook,
            pt_target_file_path=SimpleNamespace(get=lambda: 'pt-data.xlsx'),
            pt_template_file_path=SimpleNamespace(get=lambda: 'pt-template.xlsx'),
            mt_target_file_path=SimpleNamespace(get=lambda: 'mt-data.xlsx'),
            mt_template_file_path=SimpleNamespace(get=lambda: 'mt-template.xlsx'),
            rt_target_file_path=SimpleNamespace(get=lambda: ''),
            rt_template_file_path=SimpleNamespace(get=lambda: ''),
            kogas_target_file_path=SimpleNamespace(get=lambda: ''),
            kogas_template_file_path=SimpleNamespace(get=lambda: ''),
            target_file_path=SimpleNamespace(get=lambda: ''),
            template_file_path=SimpleNamespace(get=lambda: ''),
            pt_extracted_data=[{'selected': True, 'date_filtered': True}],
            mt_extracted_data=[{'selected': True, 'date_filtered': True}],
            rt_extracted_data=[], kogas_extracted_data=[], extracted_data=[],
            _run_pt_process=lambda data, template: calls.append(('PT', data, template)),
            _run_mt_process=lambda data, template: calls.append(('MT', data, template)),
            _run_rt_process=lambda *args, **kw: self.fail('wrong mode'),
            _run_pmi_process=lambda *args: self.fail('wrong mode'))
        self.report_method('run_process')(app)
        notebook.tab = lambda *args: 'MT'
        self.report_method('run_process')(app)
        self.assertEqual(calls[0][0], 'PT')
        self.assertEqual(calls[0][2], 'pt-template.xlsx')
        self.assertEqual(calls[1][0], 'MT')
        self.assertEqual(calls[1][2], 'mt-template.xlsx')


if __name__ == '__main__':
    unittest.main()
