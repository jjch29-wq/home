"""Run with: python home/tests/test_site_isolation.py (no GUI/dependencies needed)."""
import ast
import importlib
import importlib.util
import os
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
ROOT = SRC / 'site_apps'
SITES = ('central', 'kogas', 'lotte')


class SiteIsolationTests(unittest.TestCase):
    def test_all_site_python_files_compile(self):
        for path in ROOT.rglob('*.py'):
            with self.subTest(path=path):
                compile(path.read_text(encoding='utf-8-sig'), str(path), 'exec')

    def test_local_imports_never_escape_site(self):
        local_roots = {p.stem for p in SRC.glob('*.py')}
        local_roots.update(p.name for p in SRC.iterdir() if p.is_dir())
        for site in SITES:
            for path in (ROOT / site / 'src').rglob('*.py'):
                tree = ast.parse(path.read_text(encoding='utf-8-sig'))
                for node in ast.walk(tree):
                    names = []
                    if isinstance(node, ast.Import):
                        names = [alias.name for alias in node.names]
                    elif isinstance(node, ast.ImportFrom) and not node.level:
                        names = [node.module]
                    for name in names:
                        with self.subTest(path=path, module=name):
                            if name.startswith('site_apps.'):
                                self.assertTrue(name.startswith(f'site_apps.{site}.src.'))
                                spec = importlib.util.find_spec(name)
                                self.assertIsNotNone(spec)
                                self.assertTrue(Path(spec.origin).is_relative_to(ROOT / site))
                            else:
                                self.assertNotIn(name.split('.')[0], local_roots)

    def test_inventory_and_app_config_paths_are_separate(self):
        paths = []
        for site in SITES:
            app_path = ROOT / site / 'src' / 'app.py'
            tree = ast.parse(app_path.read_text(encoding='utf-8'))
            manager = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'MaterialManager')
            init = next(n for n in manager.body if isinstance(n, ast.FunctionDef) and n.name == '__init__')
            # Execute only the actual path-selection branches; do not start UI or load/save data.
            branches = [n for n in init.body if isinstance(n, ast.If)
                        and "getattr(sys, 'frozen', False)" == ast.unparse(n.test)]
            self.assertEqual(len(branches), 2)
            obj = SimpleNamespace()
            context = {'self': obj, 'sys': SimpleNamespace(frozen=False), 'os': os, '__file__': str(app_path)}
            exec(compile(ast.Module(body=branches, type_ignores=[]), str(app_path), 'exec'), context)
            for value in (obj.db_path, obj.config_path):
                value = Path(value).resolve()
                self.assertTrue(value.is_relative_to(ROOT / site))
                self.assertTrue(value.exists(), value)
                paths.append(value)
        self.assertEqual(len(paths), len(set(paths)))

    def test_billing_config_and_history_paths_are_separate(self):
        for site in SITES:
            folder = ROOT / site / 'src'
            billing = folder / ('ndt_billing_kogas_tab.py' if site == 'kogas' else 'ndt_billing_tab.py')
            tree = ast.parse(billing.read_text(encoding='utf-8'))
            assignments = [n for n in tree.body if isinstance(n, ast.Assign)
                           and any(isinstance(t, ast.Name) and t.id in ('SCRIPT_DIR', 'CONFIG_FILE') for t in n.targets)]
            context = {'os': os, '__file__': str(billing)}
            exec(compile(ast.Module(body=assignments, type_ignores=[]), str(billing), 'exec'), context)
            self.assertEqual(Path(context['CONFIG_FILE']).parent, folder)
            tab = folder / ('kogas_daily_work_log_tab.py' if site == 'kogas' else 'daily_work_log_tab.py')
            tree = ast.parse(tab.read_text(encoding='utf-8'))
            assignment = next(n for n in ast.walk(tree) if isinstance(n, ast.Assign)
                and any(isinstance(t, ast.Attribute) and t.attr == 'history_path' for t in n.targets))
            obj = SimpleNamespace()
            exec(compile(ast.Module(body=[assignment], type_ignores=[]), str(tab), 'exec'),
                 {'self': obj, 'os': os, '__file__': str(tab)})
            expected = {'central': 'daily_work_history.json', 'kogas': 'kogas_daily_work_history.json',
                        'lotte': 'daily_work_history.json'}[site]
            self.assertEqual(Path(obj.history_path), folder / expected)

    def test_billing_results_and_module_identity(self):
        modules = [importlib.import_module(f'site_apps.{site}.src.services.ndt_calculator') for site in SITES]
        self.assertEqual(len({id(m) for m in modules}), 3)
        for module in modules:
            result = module.calculate_billing(quantity=2, adjusted_quantity=3,
                material_key='RT_FILM', ndt_type='RT', location='plant', work_time='day',
                material_costs={'RT_FILM': 100}, labor_costs={'plant': {'day': {'RT_FILM': 200}}},
                overhead_rate=0.1, technical_fee_rate=0.1)
            self.assertEqual(result, {'mat_cost': 200, 'lab_cost': 600, 'overhead': 60,
                'tech': 66, 'subtotal': 926, 'vat': 92, 'total_amount': 1018})

    def test_compatibility_launchers_target_the_correct_site(self):
        entries = {'central': '한국지역난방 중앙지사.py', 'kogas': '가스공사 가산~가평.py',
                   'lotte': '롯데건설 바이오로직스.py'}
        for site, filename in entries.items():
            source = (SRC / filename).read_text(encoding='utf-8')
            self.assertIn(f'from site_apps.{site}.src.app import MaterialManager', source)
            compile(source, filename, 'exec')


if __name__ == '__main__':
    unittest.main()
