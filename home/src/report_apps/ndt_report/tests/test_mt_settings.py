import ast
from pathlib import Path
from types import SimpleNamespace
import unittest


SOURCE = Path(__file__).resolve().parents[1] / 'src' / '비파괴검사보고서.py'


class MTSettingsTests(unittest.TestCase):
    def test_saved_mt_paths_are_restored_to_ui_on_startup(self):
        module = ast.parse(SOURCE.read_text(encoding='utf-8'))
        loader = next(n for n in ast.walk(module) if isinstance(n, ast.FunctionDef) and n.name == 'load_settings')
        restored = {}
        app = SimpleNamespace(config={'MT_TEMPLATE_PATH': 'templates/mt.xlsx', 'MT_TARGET_PATH': 'input/mt.xlsx'},
            mt_template_file_path=SimpleNamespace(set=lambda value: restored.update(template=value)),
            mt_target_file_path=SimpleNamespace(set=lambda value: restored.update(target=value)))
        statements = [n for n in ast.walk(loader) if isinstance(n, ast.Expr)
                      and isinstance(n.value, ast.Call) and isinstance(n.value.func, ast.Attribute)
                      and n.value.func.attr == 'set' and isinstance(n.value.func.value, ast.Attribute)
                      and n.value.func.value.attr in ('mt_template_file_path', 'mt_target_file_path')]
        self.assertEqual(len(statements), 2)
        exec(compile(ast.Module(body=statements, type_ignores=[]), str(SOURCE), 'exec'), {'self': app})
        self.assertEqual(restored, {'template': 'templates/mt.xlsx', 'target': 'input/mt.xlsx'})

    def test_template_save_persists_main_settings_first(self):
        module = ast.parse(SOURCE.read_text(encoding='utf-8'))
        method = next(n for n in ast.walk(module) if isinstance(n, ast.FunctionDef) and n.name == 'save_template_specific_config')
        first_action = method.body[1]
        self.assertIsInstance(first_action, ast.Expr)
        self.assertEqual(first_action.value.func.attr, 'save_settings')
