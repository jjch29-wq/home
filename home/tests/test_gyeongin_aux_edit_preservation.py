import ast
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.aux_equipment import restore, value


class Field:
    def __init__(self, text=''):
        self.text = text

    def get(self):
        return self.text

    def set(self, text):
        self.text = text

    def delete(self, *args):
        self.text = ''

    def insert(self, position, text):
        self.text = text

    def __setitem__(self, key, value):
        pass


class AuxEditTests(unittest.TestCase):
    def test_restored_custom_or_empty_aux_survives_same_method_events(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/gyeongin/src/views/daily_usage_view.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        callback = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)
                        and n.name == 'on_method_change_auto_unit_logic')
        for saved in ('Custom UV meter (123); Custom probe (456)', ''):
            with self.subTest(saved=saved):
                app = SimpleNamespace(cb_daily_test_method=Field('MT'), cb_daily_unit=Field(),
                                      ent_daily_equip=Field(), daily_aux_equipment_var=Field(),
                                      equipments=[])
                restore(app, saved, method='MT')
                scope = {'self': app, 'ttk': SimpleNamespace(Combobox=Field), 'tk': SimpleNamespace(END='end')}
                exec(compile(ast.Module(body=[callback], type_ignores=[]), str(source), 'exec'), scope)
                scope['on_method_change_auto_unit_logic']()
                self.assertEqual(value(app), saved)
                # Switching to another method still installs that method's defaults.
                app.cb_daily_test_method.set('RT')
                scope['on_method_change_auto_unit_logic']()
                self.assertNotEqual(value(app), saved)
                # Editing a loaded record must not reselect cleared auxiliaries,
                # even when a method event is fired after restoration.
                app._daily_edit_source = (0, {})
                app.daily_aux_equipment_var.set('')
                app.cb_daily_test_method.set('MT')
                scope['on_method_change_auto_unit_logic']()
                self.assertEqual(value(app), '')


if __name__ == '__main__':
    unittest.main()
