from pathlib import Path
import sys
import unittest
import ast
from types import SimpleNamespace
import datetime
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.services.ndt_record_edit import chemical_changes


class ChemicalEditTests(unittest.TestCase):
    def test_actual_save_handler_updates_selected_mt_record(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/lotte/src/app.py'
        tree = ast.parse(source.read_text(encoding='utf-8-sig'))
        method = next(node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)
                      and node.name == 'add_daily_usage_entry')
        messages = []
        namespace = {'pd': pd, 'worker_names': lambda record: '',
                     'messagebox': SimpleNamespace(showinfo=lambda *args: messages.append(args),
                                                   showerror=lambda *args: messages.append(args))}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), namespace)
        field = lambda value: SimpleNamespace(get=lambda: value)
        date = datetime.date(2026, 10, 14)
        frame = pd.DataFrame([{'Date': date, 'Site': '함감속기', '검사방법': 'MT',
                               '회사코드': float('nan'), 'NDT_형광자분': 0.0},
                              {'Date': date, 'Site': '함감속기', '검사방법': 'UT',
                               '회사코드': float('nan'), 'NDT_형광자분': 0.0}])
        saves, deltas = [], []
        app = SimpleNamespace(daily_usage_df=frame, _loaded_daily_record_index=0,
            fixed_vehicle_widget=SimpleNamespace(_source_record_index=1, get_data=lambda: {}),
            cb_daily_material=field(''), cb_daily_site=field('함감속기'), cb_daily_test_method=field('MT'),
            ent_daily_date=SimpleNamespace(get_date=lambda: date), ent_daily_inspection_item=field('Gear'),
            ndt_company_entries=[{'_company': field(''), '형광자분': field('1')}],
            clean_nan=lambda value: '' if pd.isna(value) else str(value).strip(),
            save_data=lambda: saves.append(frame.copy()), update_daily_usage_view=lambda: None,
            update_recent_entries_view=lambda: None,
            _auto_reconcile_and_register_ndt=lambda *args, **kwargs: deltas.append(args[2]))
        namespace['add_daily_usage_entry'](app)
        self.assertEqual(len(saves), 1, messages)
        self.assertEqual(saves[0].at[0, 'NDT_형광자분'], 1.0)
        self.assertEqual(saves[0].at[1, 'NDT_형광자분'], 0.0)
        self.assertIn('형광자분 저장량: 1', messages[-1][1])
        self.assertEqual(deltas, [{'형광자분': 1.0}])

    def test_mt_fluorescent_powder_can_be_saved_to_existing_record(self):
        updates, differences = chemical_changes({'NDT_형광자분': 0}, 'MT', {'형광자분': '1'})
        self.assertEqual(updates, {'NDT_형광자분': 1.0})
        self.assertEqual(differences, {'형광자분': 1.0})

    def test_saving_again_does_not_deduct_stock_twice(self):
        self.assertEqual(chemical_changes({'NDT_형광자분': 1}, 'MT', {'형광자분': '1'}), ({}, {}))

    def test_reducing_quantity_returns_only_difference(self):
        updates, differences = chemical_changes({'NDT_형광자분': 2}, 'MT', {'형광자분': '1'})
        self.assertEqual(differences['형광자분'], -1)

    def test_pt_does_not_use_mt_powder(self):
        updates, differences = chemical_changes({}, 'PT', {'형광자분': '1'})
        self.assertNotIn('NDT_형광자분', updates)


if __name__ == '__main__':
    unittest.main()
