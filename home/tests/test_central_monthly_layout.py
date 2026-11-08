"""Verify monthly sash restoration without loading app data."""
import ast
from pathlib import Path
from types import SimpleNamespace
import unittest

SOURCE = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
TREE = ast.parse(SOURCE.read_text(encoding='utf-8'))
METHODS = {node.name: node for node in ast.walk(TREE) if isinstance(node, ast.FunctionDef)}


def method(name):
    namespace = {}
    exec(compile(ast.Module(body=[METHODS[name]], type_ignores=[]), str(SOURCE), 'exec'), namespace)
    return namespace[name]


class MonthlyLayoutTests(unittest.TestCase):
    def test_drag_restart_and_resize(self):
        class Pane:
            height = 800
            positions = [100, 200]

            def winfo_exists(self): return True
            def update_idletasks(self): pass
            def winfo_height(self): return self.height
            def sashpos(self, index, value=None):
                if value is not None:
                    self.positions[index] = value
                return self.positions[index]

        pane = Pane()
        app = SimpleNamespace(monthly_paned=pane, tab_config={}, is_ready=True)
        saved = {}
        app.save_tab_config = lambda: saved.update(app.tab_config)
        pane.sashpos(0, 230)
        pane.sashpos(1, 490)
        method('_on_monthly_sash_changed')(app)
        self.assertTrue(saved['monthly_sash_ratios'])
        # Simulate a newly created pane at its default positions after restart.
        pane.sashpos(0, 100)
        pane.sashpos(1, 200)
        app.tab_config = dict(saved)
        method('_ensure_monthly_sash_visibility')(app)
        self.assertAlmostEqual(pane.sashpos(0), 230, delta=1)
        self.assertAlmostEqual(pane.sashpos(1), 490, delta=1)
        pane.height = 1000
        method('_ensure_monthly_sash_visibility')(app)
        for i, ratio in enumerate(saved['monthly_sash_ratios']):
            self.assertAlmostEqual(pane.sashpos(i), pane.winfo_height() * ratio, delta=1)
        app.is_ready = False
        pane.sashpos(0, 100)
        method('_on_monthly_sash_changed')(app)
        self.assertEqual(app.tab_config, saved)

    def test_general_save_preserves_preferences(self):
        # General saves must never sample temporary monthly pane geometry.
        node = METHODS['save_tab_config']
        reads = [n for n in ast.walk(node) if isinstance(n, ast.Attribute)
                 and isinstance(n.value, ast.Attribute)
                 and n.value.attr == 'monthly_paned']
        self.assertEqual(reads, [])


if __name__ == '__main__':
    unittest.main()
