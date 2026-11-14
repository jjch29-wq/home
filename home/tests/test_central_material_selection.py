import ast
import datetime
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock
import pandas as pd

from test_central_inspection_batch import Field, fixture, SRC
from site_apps.central.src.material_selection import choose_daily_material, resolve_daily_material, remember_selection, selected_id


def sample():
    app = SimpleNamespace(cb_daily_material=Field(), materials_df=pd.DataFrame([
        {'MaterialID': 1, '품목명': 'PT약품', '모델명': '세척제', 'Active': 1},
        {'MaterialID': 2, '품목명': 'PT약품', '모델명': '침투제', 'Active': 1},
        {'MaterialID': 3, '품목명': 'MT약품', '모델명': '백색페인트', '규격': '자동등록', 'Active': 1},
        {'MaterialID': 4, '품목명': 'Carestream T200', '모델명': '', '규격': '자동등록', 'Active': 1},
        {'MaterialID': 5, '품목명': 'PT약품', '모델명': '폐기품', 'Active': 0},
    ]))
    def display(mid):
        row = app.materials_df[app.materials_df.MaterialID == int(mid)].iloc[0]
        model = row.get('모델명', '')
        return row['품목명'] + (f' ({model})' if model else '')
    app.get_material_display_name = display
    return app


class MaterialSelectionTests(unittest.TestCase):
    def test_choice_preserves_model_and_exact_id(self):
        app = sample()
        choose_daily_material(app, '2')
        self.assertEqual(app.cb_daily_material.get(), 'PT약품 (침투제)')
        self.assertEqual(resolve_daily_material(app), 2)

        self.assertEqual(selected_id(app), '2')

    def test_manual_edit_invalidates_previous_choice(self):
        app = sample()
        choose_daily_material(app, 2)
        app.cb_daily_material.set('PT약품 (세척제)')
        self.assertIsNone(selected_id(app))
        self.assertEqual(resolve_daily_material(app), 1)

    def test_changed_or_inactive_selected_record_requires_reselection(self):
        for field, value in [('Active', 0), ('모델명', '변경된 약품')]:
            app = sample()
            choose_daily_material(app, 2)
            app.materials_df.loc[app.materials_df.MaterialID == 2, field] = value
            with self.assertRaises(ValueError):
                resolve_daily_material(app)

    def test_duplicate_display_uses_selected_id_and_snapshot_identity(self):
        app = sample()
        app.materials_df.loc[len(app.materials_df)] = dict(app.materials_df.iloc[1], MaterialID=6)
        app.cb_daily_material.set('PT약품 (침투제)')
        with self.assertRaises(ValueError):
            resolve_daily_material(app)
        choose_daily_material(app, 6)
        self.assertEqual(resolve_daily_material(app), 6)
        snapshot_id = selected_id(app)
        app._daily_material_selection = None
        remember_selection(app, snapshot_id)
        self.assertEqual(resolve_daily_material(app), 6)

    def test_actual_search_callback_includes_auto_registered_stock_and_selects_full_label(self):
        app = sample()
        class Tree:
            def __init__(self): self.rows = []
            def get_children(self): return []
            def insert(self, *args, **kwargs): self.rows.append(kwargs['values'])
            def tag_configure(self, *args, **kwargs): pass
            def selection(self): return ['chosen']
            def item(self, item): return {'values': next(row for row in self.rows if row[0] == '2')}
        source = ast.parse((SRC / 'site_apps/central/src/app.py').read_text(encoding='utf-8'))
        outer = next(n for n in ast.walk(source) if isinstance(n, ast.FunctionDef) and n.name == 'open_material_search_dialog')
        functions = [n for n in outer.body if isinstance(n, ast.FunctionDef) and n.name in ('refresh_list', 'on_select')]
        tree = Tree()
        scope = dict(self=app, pd=pd, tk=SimpleNamespace(END='end'), tree=tree,
                     search_var=SimpleNamespace(get=lambda: ''), target_form='daily_usage', dlg=SimpleNamespace(destroy=Mock()))
        exec(compile(ast.fix_missing_locations(ast.Module(body=functions, type_ignores=[])), '<actual material search>', 'exec'), scope)
        scope['refresh_list']()
        self.assertEqual({row[0] for row in tree.rows}, {'1', '2', '3', '4'})
        scope['on_select']()
        self.assertEqual(app.cb_daily_material.get(), 'PT약품 (침투제)')
        self.assertEqual(resolve_daily_material(app), 2)
        app.cb_daily_test_method = Field('')
        scope['include_unknown'] = Field(False)
        for method, expected in [('MT', {'3'}), ('PT', {'1', '2'}), ('RT', {'4'})]:
            app.cb_daily_test_method.set(method)
            tree.rows.clear()
            scope['refresh_list']()
            self.assertEqual({row[0] for row in tree.rows}, expected)

    def test_exact_selected_id_reaches_actual_record_builder(self):
        app = fixture()
        app.materials_df = pd.DataFrame([
            {'MaterialID': 10, '품목명': 'Carestream T200', '모델명': 'A', 'Active': 1},
            {'MaterialID': 20, '품목명': 'Carestream T200', '모델명': 'B', 'Active': 1}])
        app.get_material_display_name = lambda mid: f'Carestream T200 ({"A" if int(mid) == 10 else "B"})'
        app.cb_daily_test_method.set('RT')
        app.cb_daily_unit.set('매')
        app.ent_daily_test_amount.set('54')
        app.ndt_ori_qty_var.set('54')
        choose_daily_material(app, 20)
        plan = app._add_single_usage_record_logic(resolve_daily_material(app), datetime.date(2026, 11, 14), '중앙지사', prepare_only=True)
        self.assertEqual(plan['common_data']['MaterialID'], 20)


if __name__ == '__main__':
    unittest.main()
