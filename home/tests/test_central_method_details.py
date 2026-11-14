"""Method details follow restored values without modifying the entered data."""
import sys
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk
from tkinter import ttk
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.daily_usage_view import refresh_method_details


class MethodDetailsTests(unittest.TestCase):
    def test_programmatic_method_changes_show_panels_and_keep_values(self):
        root = tk.Tk()
        root.withdraw()
        self.addCleanup(root.destroy)
        value = tk.StringVar(root)
        field = ttk.Combobox(root, textvariable=value)
        app = SimpleNamespace(root=root, cb_daily_test_method=field, daily_method_var=value,
                              _ensure_canvas_scroll_region=lambda: None)
        for name in ('ndt_calc_frame', 'ndt_frame', 'rtk_grid',
                     'empty_guide_frame', 'daily_method_hint'):
            panel = ttk.Label(root, text='unchanged')
            panel.grid()
            setattr(app, name, panel)
        entered = tk.StringVar(root, value='123.45')
        ttk.Entry(app.ndt_calc_frame, textvariable=entered)
        value.trace_add('write', lambda *args: refresh_method_details(app))
        for method, expected in (
                ('', (False, False, False, False, True)),
                ('RT', (True, False, True, False, False)),
                ('MT', (True, True, False, False, False)),
                ('PT', (True, True, False, False, False)),
                ('UT', (True, False, False, True, False)),
                ('PAUT', (True, False, False, True, False)),
                ('ETC', (False, False, False, False, True))):
            value.set(method)
            actual = tuple(bool(getattr(app, n).winfo_manager()) for n in (
                'ndt_calc_frame', 'ndt_frame', 'rtk_grid', 'empty_guide_frame', 'daily_method_hint'))
            self.assertEqual(actual, expected, method)
            self.assertEqual(entered.get(), '123.45')


if __name__ == '__main__':
    unittest.main()
