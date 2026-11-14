"""Build the real worker controls without opening or saving business data."""
import ast
import math
from pathlib import Path
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.components import WorkerDataGroup


class WorkerLayoutTests(unittest.TestCase):
    def test_layout_builds_and_resizes_without_losing_worker_values(self):
        root = tk.Tk()
        root.withdraw()
        self.addCleanup(root.destroy)
        page = ttk.Frame(root)
        app = SimpleNamespace(root=root, theme_bg='#f0f0f0', users=[], worktimes=[],
                              open_list_management_dialog=lambda *args: None)
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/views/daily_usage_view.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        setup = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'setup_daily_usage_tab_impl')
        start = next(i for i, n in enumerate(setup.body) if isinstance(n, ast.Assign)
                     and any(isinstance(t, ast.Name) and t.id == 'workers_box_frame' for t in n.targets))
        end = next(i for i, n in enumerate(setup.body) if isinstance(n, ast.FunctionDef) and n.name == 'layout_workers')
        wrapper = ast.parse('def build(self, workers_page):\n    pass').body[0]
        wrapper.body = setup.body[start:end + 1] + ast.parse('return layout_workers, crew_rows, crew_header_blocks').body
        scope = dict(tk=tk, ttk=ttk, math=math, messagebox=messagebox, WorkerDataGroup=WorkerDataGroup)
        exec(compile(ast.fix_missing_locations(ast.Module(body=[wrapper], type_ignores=[])), str(source), 'exec'), scope)
        layout, rows, headers = scope['build'](app, page)
        for i in range(1, 11):
            worker = getattr(app, f'worker_group{i}')
            worker.set_worker(f'Worker {i}')
            worker.ent_worktime.set('18:00~24:00')
            worker.participation_var.set(i % 2 == 1)
        root.update_idletasks()
        for width, columns in ((1900, 2), (1000, 1), (1900, 2)):
            page.winfo_width = lambda width=width: width
            layout()
            self.assertEqual([int(row.grid_info()['column']) for row in rows],
                             [i // 5 if columns == 2 else 0 for i in range(10)])
            self.assertEqual(bool(headers[1].winfo_manager()), columns == 2)
            for i in range(1, 11):
                worker = getattr(app, f'worker_group{i}')
                self.assertEqual(worker.get_worker(), f'Worker {i}')
                self.assertEqual(worker.ent_worktime.get(), '18:00~24:00')
                self.assertEqual(worker.participation_var.get(), i % 2 == 1)


if __name__ == '__main__':
    unittest.main()
