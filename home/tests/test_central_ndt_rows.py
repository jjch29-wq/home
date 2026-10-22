"""Exercise row expansion without opening or modifying a user's live session."""
import ast
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock


class NdtRowsTests(unittest.TestCase):
    def test_section_batch_replaces_exact_matches_or_checked_rows_only(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        method = next(node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)
                      and node.name == '_replace_ndt_section')
        namespace = {}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), namespace)
        class Section(dict):
            def __init__(self, value):
                self.value = value
            def get(self):
                return self.value
            def set(self, value):
                self.value = value
        rows = [{'구간': Section(s), '라인번호': 'L1', 'Joint No.': f'S{i}', 'RT_OR': '3'}
                for i, s in enumerate(('Sec.15', 'Sec.15', 'Sec.15-3', ''))]
        app = SimpleNamespace(ndt_grid_entries=rows, history_sections=['', 'Sec.15', 'Sec.15-3'],
                              ndt_selection_vars=[Section(v) for v in (False, True, True, False)])
        replace = namespace['_replace_ndt_section']
        self.assertEqual(replace(app, ' Sec.15 ', ' Sec.15-4 '), 2)
        self.assertEqual([r['구간'].get() for r in rows], ['Sec.15-4', 'Sec.15-4', 'Sec.15-3', ''])
        self.assertEqual(replace(app, '', 'Sec.16', True), 2)
        self.assertEqual([r['구간'].get() for r in rows], ['Sec.15-4', 'Sec.16', 'Sec.16', ''])
        self.assertEqual([r['Joint No.'] for r in rows], ['S0', 'S1', 'S2', 'S3'])
        self.assertTrue(all(r['라인번호'] == 'L1' and r['RT_OR'] == '3' for r in rows))
        self.assertIn('Sec.16', app.history_sections)
        for before, after in (('', 'Sec.1'), ('Sec.16', ''), ('Sec.16', 'Sec.16'), ('missing', 'Sec.1')):
            with self.assertRaises(ValueError):
                replace(app, before, after)

    def test_line_replacement_changes_only_exact_matches_and_preserves_joint_data(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        cls = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == 'DailyWorkLogTab')
        method = next(node for node in cls.body if isinstance(node, ast.FunctionDef) and node.name == '_replace_ndt_line')
        namespace = {}
        exec(compile(ast.fix_missing_locations(ast.Module(body=[method], type_ignores=[])), str(source), 'exec'), namespace)

        class Line(dict):
            def __init__(self, value):
                self.value = value
            def get(self):
                return self.value
            def set(self, value):
                self.value = value

        old, new = 'JA2026-017-001-0-1-100A', 'JA2026-017-003-0-1-100A'
        rows = [{'라인번호': Line(line), 'Joint No.': f'S{index}', 'PAUT': '0.3591'}
                for index, line in enumerate((new, old, old, old + '-extra', ''))]
        app = SimpleNamespace(ndt_grid_entries=rows, history_lines=['', old, new])
        replace = namespace['_replace_ndt_line']
        self.assertEqual(replace(app, old, new), 2)
        self.assertEqual([row['라인번호'].get() for row in rows], [new, new, new, old + '-extra', ''])
        self.assertEqual([row['Joint No.'] for row in rows], [f'S{index}' for index in range(5)])
        self.assertTrue(all(row['PAUT'] == '0.3591' for row in rows))
        self.assertEqual(replace(app, old, new), 0)
        for before, after in (('', new), (old, ''), (new, new)):
            with self.assertRaises(ValueError):
                replace(app, before, after)

    def test_rt_film_spec_is_removed_from_paut_and_rt_options_remain(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        cls = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == 'DailyWorkLogTab')
        method = next(node for node in cls.body if isinstance(node, ast.FunctionDef) and node.name == '_sync_ndt_spec')
        namespace = {}
        exec(compile(ast.fix_missing_locations(ast.Module(body=[method], type_ignores=[])), str(source), 'exec'), namespace)

        class Spec(dict):
            def __init__(self, value):
                self.value = value
            def get(self):
                return self.value
            def set(self, value):
                self.value = value

        for inspection in ('PAUT', 'UT', 'MT', 'PT', 'RT'):
            for value in ('3 1/3 X 12"', '31/3 X6"', '직접 입력한 규격'):
                spec = Spec(value)
                row = {'검사방법': SimpleNamespace(get=lambda: inspection), '규격': spec}
                namespace['_sync_ndt_spec'](None, row)
                expected = value if inspection == 'RT' or value == '직접 입력한 규격' else ''
                self.assertEqual(spec.get(), expected)
                self.assertEqual(len(spec['values']), 3 if inspection == 'RT' else 1)

    def test_append_and_restore_capacity_preserve_existing_rows(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/daily_work_log_tab.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        cls = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == 'DailyWorkLogTab')
        methods = [node for node in cls.body if isinstance(node, ast.FunctionDef)
                   and node.name in ('_append_ndt_rows', '_ensure_ndt_capacity')]
        namespace = {}
        exec(compile(ast.fix_missing_locations(ast.Module(body=methods, type_ignores=[])), str(source), 'exec'), namespace)
        existing = [object() for _ in range(30)]
        app = SimpleNamespace(ndt_grid_entries=list(existing), _focus_ndt_joint=Mock())
        created = []

        def create(index):
            created.append(index)
            app.ndt_grid_entries.append(object())

        app._create_ndt_row = create
        app._append_ndt_rows = lambda count=10, focus=False: namespace['_append_ndt_rows'](app, count, focus)
        app._append_ndt_rows(10, focus=True)
        self.assertEqual(created, list(range(31, 41)))
        self.assertEqual(app.ndt_grid_entries[:30], existing)
        app._focus_ndt_joint.assert_called_once_with(30)
        namespace['_ensure_ndt_capacity'](app, 75)
        self.assertEqual(len(app.ndt_grid_entries), 80)
        self.assertEqual(app.ndt_grid_entries[:30], existing)
        namespace['_ensure_ndt_capacity'](app, 80)
        namespace['_ensure_ndt_capacity'](app, 30)
        self.assertEqual(len(app.ndt_grid_entries), 80)


if __name__ == '__main__':
    unittest.main()
