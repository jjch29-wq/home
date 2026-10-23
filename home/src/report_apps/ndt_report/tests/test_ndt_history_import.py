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
from ndt_history_import import SITES, history_path, read_results, rt_film_findings, rt_report_quantities


class HistoryImportTests(unittest.TestCase):
    def test_photo_order_groups_lines_and_sorts_joint_numbers(self):
        from ndt_history_import import sort_photo_paths
        paths = ['G:/L1_S10_capture.png', 'G:/L2_R01.png', 'G:/L1_R10.png',
                 'G:/L1_S2_capture.png', 'G:/L1_R02.png']
        self.assertEqual(sort_photo_paths(paths), [paths[4], paths[2], paths[3], paths[0], paths[1]])
        self.assertEqual(len(paths), 5)

    def test_report_quantities_use_all_selected_source_counts(self):
        items = [
            {'Loc': 'N/1,이물질/4,N/1', 'RT_OR': '3', 'RT_RE': '1'},
            {'Loc': ',N/1', 'RT_OR': '1', 'RT_RE': '0'},
            {'Loc': 'N/1,N/1,P/3', 'RT_OR': '3', 'RT_RE': '1'},
            {'Loc': ',,N/1', 'RT_OR': '1', 'RT_RE': '0'},
        ]
        self.assertEqual(rt_report_quantities(items), (8, 2, 0))
        self.assertEqual(rt_report_quantities([{'Loc': 'N/1,P/3', 'Rej': 'X'}]), (2, 1, 1))

    def test_rt_film_findings_preserve_grade_and_defect(self):
        self.assertEqual(rt_film_findings('P/2,N/1,N/1'), {'D5': '√', 'Deg': '2'})
        self.assertEqual(rt_film_findings('N/1,N/1,N/1'), {'Deg': '1'})
        self.assertEqual(rt_film_findings('IP/3, P/2'), {'D2': '√', 'D5': '√', 'Deg': '3'})
        self.assertEqual(rt_film_findings('1-2,Sec.01'), {})

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

    def test_central_reads_latest_configured_history(self):
        site = next(iter(SITES))
        local = self.write(site, {'2026-10-07': {'ndt_results': []}})
        cloud = self.base / 'cloud' / 'central'
        cloud.mkdir(parents=True)
        latest = cloud / 'daily_work_history.json'
        latest.write_text(json.dumps({'2026-10-21': {'ndt_results': [
            {'검사방법': 'RT', 'Joint No.': 'R21', '재촬영구간': ['A-B']}]}}), encoding='utf-8')
        (local.parent / 'sync_config.json').write_text(
            json.dumps({'data_dir': str(cloud)}), encoding='utf-8')
        before = latest.read_bytes()
        rows = read_results(site, 'RT', self.base)
        self.assertEqual(history_path(site, self.base), latest)
        self.assertEqual(rows[0]['Date'], '2026-10-21')
        self.assertEqual(rows[0]['Joint No.'], 'R21')
        self.assertEqual(latest.read_bytes(), before)
        latest.unlink()
        with self.assertRaises(FileNotFoundError):
            read_results(site, 'RT', self.base)

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

    def test_legacy_wide_columns_are_imported(self):
        site = next(iter(SITES))
        self.write(site, {'2026-09-08': {'ndt_results': [
            {'Joint No.': 'S01', 'PAUT': '0.6795', 'PT': '', 'MT': ''},
            {'Joint No.': 'R01', 'RT_OR': '1', 'RT_RE': ''},
        ]}})
        paut = read_results(site, 'PAUT', self.base)
        rt = read_results(site, 'RT', self.base)
        self.assertEqual(len(paut), 1)
        self.assertEqual(paut[0]['Date'], '2026-09-08')
        self.assertEqual(paut[0]['검사방법'], 'PAUT')
        self.assertEqual(paut[0]['검사길이'], '0.6795')
        self.assertEqual(rt[0]['Joint No.'], 'R01')

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
        keys = ['Date', 'Sec', 'Loc', 'Joint', 'Size', 'Result', 'ISO', 'Deg', 'D5']
        app = SimpleNamespace(root=None, mode_notebook=notebook, rt_preview_nb=notebook,
            _normalize_daily_photo_welder=lambda value: value,
            _get_mode_info=lambda mode: (tree, indices, data, keys), update_date_listbox=lambda **kw: None,
            populate_preview=lambda *args, **kw: None)
        row = {'Date': '2026-09-05', '검사방법': 'RT', '구간': 'A', '구간정보': 'P/2,N/1,N/1',
               'Joint No.': 'J1', '라인번호': 'L1', '결과': '합격', '_source_site': '롯데건설 바이오로직스',
               '_source_path': '/test/history.json'}
        with patch('ndt_history_import.select_history_rows', return_value=[row]):
            self.report_method()(app)
        self.assertEqual((data[0]['Sec'], data[0]['Loc'], data[0]['Joint']), ('A', 'P/2,N/1,N/1', 'J1'))
        self.assertEqual((data[0]['D5'], data[0]['Deg']), ('√', '2'))
        self.assertEqual(data[0]['_source_site'], row['_source_site'])
        self.assertEqual(indices, [0])
        self.assertEqual(inserted[0]['tags'], ('pass',))

    def test_get_mode_info_supports_mt(self):
        app = SimpleNamespace(mt_preview_tree='tree', mt_item_idx_map='idx',
                              mt_extracted_data='data', mt_column_keys='keys')
        self.assertEqual(self.report_method('_get_mode_info')(app, 'MT'),
                         ('tree', 'idx', 'data', 'keys'))

    def test_mt_check_actions_only_change_matching_mt_rows(self):
        rows = [{'selected': False}, {'selected': False, 'drawing_filtered': False},
                {'selected': False, 'date_filtered': False}]
        other = [{'selected': False}]
        app = SimpleNamespace(mt_extracted_data=rows, extracted_data=other,
                              populate_preview=lambda *args, **kw: None)
        self.report_method('select_all')(app, 'MT')
        self.assertEqual([r['selected'] for r in rows], [True, False, False])
        self.assertFalse(other[0]['selected'])
        self.report_method('deselect_all')(app, 'MT')
        self.assertFalse(rows[0]['selected'])

    def test_mt_add_row_uses_mt_columns(self):
        rows, other = [], []
        app = SimpleNamespace(mt_extracted_data=rows, extracted_data=other,
                              mt_column_keys=['Dwg', 'Material', 'TestItem'],
                              populate_preview=lambda *args, **kw: None, log=lambda *args: None)
        self.report_method('add_item')(app, 'MT')
        self.assertEqual(len(rows), 1)
        self.assertIn('Material', rows[0])
        self.assertEqual(other, [])

    def test_rt_and_kogas_check_actions_respect_drawing_and_date(self):
        for mode in ('RT', 'KOGAS'):
            rows = [{'selected': False}, {'selected': False, 'drawing_filtered': False},
                    {'selected': False, 'date_filtered': False}]
            other = [{'selected': False}]
            app = SimpleNamespace(rt_extracted_data=rows if mode == 'RT' else other,
                                  kogas_extracted_data=rows if mode == 'KOGAS' else other,
                                  extracted_data=other, populate_preview=lambda *args, **kw: None)
            self.report_method('select_all')(app, mode)
            self.assertEqual([row['selected'] for row in rows], [True, False, False])
            self.assertFalse(other[0]['selected'])

    def test_rt_and_kogas_display_report_number_after_date(self):
        class Tree(dict):
            def __bool__(self):
                return True
            def heading(self, *args, **kwargs):
                pass
            def column(self, *args, **kwargs):
                pass
        for mode in ('RT', 'KOGAS'):
            tree = Tree()
            app = SimpleNamespace(rt_preview_tree=tree, kogas_preview_tree=tree,
                                  rt_column_keys=['selected', 'No', 'Date', 'Dwg'],
                                  kogas_column_keys=['selected', 'No', 'Date', 'Dwg'],
                                  rt_extracted_data=[], kogas_extracted_data=[], config={}, setting_vars={})
            self.report_method('_update_rt_preview_columns')(app, mode)
            self.assertEqual(tree['columns'][:4], ('selected', 'No', 'Date', 'Report No'))

    def test_paut_daily_measurements_reach_preview_even_when_accepted(self):
        from ndt_history_import import paut_flaw_findings
        section = 'X:203.5-209.0,L:5.8,D:1.175~1.259,H:0.83,Surface'
        data, indices, inserted = [], [], []
        tree = SimpleNamespace(insert=lambda *args, **kw: inserted.append(kw))
        notebook = SimpleNamespace(select=lambda: 'tab', tab=lambda *args: 'PAUT')
        keys = ['Joint No.', "Th'k(mm)", 'Start', 'End', 'Length(mm)', 'Upper',
                'Lower', 'Height(mm)', 'Type of Flaw', 'a/l', 'a/t', 'Acc', 'Evaluation', 'Remarks']
        app = SimpleNamespace(root=None, mode_notebook=notebook,
            _normalize_daily_photo_welder=lambda value: value,
            _get_mode_info=lambda mode: (tree, indices, data, keys), update_date_listbox=lambda **kw: None,
            populate_preview=lambda *args, **kw: None)
        row = {'Joint No.': 'R01', '두께': '4.9', '구간정보': section, '결과': '합격',
               '_source_site': '한국지역난방 중앙지사', '_source_path': '/test/history.json'}
        with patch('ndt_history_import.select_history_rows', return_value=[row]):
            self.report_method()(app)
        self.assertEqual([data[0][k] for k in keys[2:8]],
                         ['203.5', '209.0', '5.8', '1.175', '1.259', '0.83'])
        self.assertEqual((data[0]['Acc'], data[0]['a/l'], data[0]['a/t']), ('√', '0.143', '0.169'))
        self.assertEqual(data[0]['Remarks'], 'Surface')
        self.assertEqual(data[0]['_source_paut_section'], section)
        self.assertEqual(inserted[0]['values'][2], '203.5')
        self.assertEqual(paut_flaw_findings(''), {})
        self.assertEqual(paut_flaw_findings('관련지시 없음')['Remarks'], '관련지시 없음')
        self.assertEqual(paut_flaw_findings(section.replace('Surface', 'Subsurface'))['Remarks'], 'Subsurface')
        self.assertEqual(paut_flaw_findings(section.replace('Surface', 'SUBSURFACE'))['Remarks'], 'Subsurface')
        self.assertEqual(data[0]['Type of Flaw'], '')

    def test_duplicate_source_id_is_not_imported_twice(self):
        data, indices, inserted = [], [], []
        tree = SimpleNamespace(insert=lambda *args, **kw: inserted.append(kw))
        notebook = SimpleNamespace(select=lambda: 'tab', tab=lambda *args: 'PT')
        keys = ['Date', 'Joint No.', "Th'k(mm)", 'Remarks', 'Result']
        app = SimpleNamespace(root=None, mode_notebook=notebook,
            _normalize_daily_photo_welder=lambda value: value,
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
