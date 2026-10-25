"""Exercise real Tk vehicle tabs without opening the live application or database."""
import datetime
import ast
from pathlib import Path
import sys
import tkinter as tk
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.vehicle_panel import VehiclePanel
from site_apps.central.src.inspection_batch import InspectionBatch
from test_central_inspection_batch import Field, fixture


class VehiclePanelTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f'Tk runtime unavailable: {exc}')
        self.root.withdraw()
        self.app = SimpleNamespace(
            theme_bg='#f0f0f0', vehicles=['90서8616', '95가0175'],
            vehicle_inspections={}, vehicle_boxes=[],
            ent_daily_date=SimpleNamespace(get_date=lambda: datetime.date(2026, 10, 24)),
            cb_daily_site=Mock(get=lambda: '중앙지사'), cb_daily_company=Field('발주처'),
            ent_daily_travel_cost=Field('0'), ent_daily_meal_cost=Field('0'),
            save_single_vehicle_data=Mock(return_value=True), open_list_management_dialog=Mock())
        self.panel = VehiclePanel(self.root, self.app)
        self.app.vehicle_panel = self.panel
        self.panel.pack()

    def tearDown(self):
        self.root.destroy()

    def two_vehicles(self):
        first = self.app.fixed_vehicle_widget
        first.set_data({'vehicle_info': '90서8616', 'mileage': '100', 'out_exterior': '양호',
                        'remarks': '첫 차량', 'photo_path': 'first-photo.jpg'})
        second = self.panel.add({'vehicle_info': '95가0175', 'mileage': '80',
                                 'out_exterior': '불량', 'remarks': '추가 차량',
                                 'photo_path': 'second-photo.jpg'})
        return first, second

    def test_switching_tabs_preserves_checks_distance_and_photos(self):
        first, second = self.two_vehicles()
        self.panel.notebook.select(first)
        self.panel.notebook.select(second)
        self.assertEqual(first.get_data()['mileage'], '100')
        self.assertEqual(second.get_data()['out_exterior'], '불량')
        self.assertEqual(first.photo_path, 'first-photo.jpg')
        self.assertEqual(second.photo_path, 'second-photo.jpg')
        self.assertEqual(self.app.vehicle_boxes, [second])
        self.assertIn(second, self.app.vehicle_inspections.values())
        self.assertNotIn('place', (first.winfo_manager(), second.winfo_manager()))

    def test_serialization_restores_both_vehicles_and_clears_old_photos(self):
        self.two_vehicles()
        data = self.panel.serialize()
        self.panel.restore(data)
        self.assertEqual(self.panel.serialize(), data)
        self.panel.restore([])
        self.assertEqual(len(self.panel.editors), 1)
        self.assertEqual(self.app.fixed_vehicle_widget.photo_path, '')
        self.assertEqual(self.app.vehicle_boxes, [])
        self.assertEqual(self.app.vehicle_inspections, {})

    def test_delete_primary_promotes_remaining_vehicle_without_losing_data(self):
        first, second = self.two_vehicles()
        self.panel.remove(first)
        self.assertIs(self.app.fixed_vehicle_widget, second)
        self.assertEqual(self.app.vehicle_boxes, [])
        self.assertEqual(second.get_data()['mileage'], '80')
        self.panel.remove(second)
        self.assertEqual(len(self.panel.editors), 1)
        self.assertEqual(second.photo_path, '')

    def test_status_changes_only_after_successful_save_and_edits_invalidate_it(self):
        first, _ = self.two_vehicles()
        self.app.save_single_vehicle_data.return_value = False
        self.panel.save_one(first)
        self.assertNotIn(first, self.panel.saved)
        self.app.save_single_vehicle_data.return_value = True
        self.panel.save_one(first)
        self.assertIn('저장 완료', self.panel.notebook.tab(first, 'text'))
        with patch('site_apps.central.src.views.vehicle_panel.messagebox.showinfo'):
            self.panel.save_one(first)
        self.assertEqual(self.app.save_single_vehicle_data.call_count, 2)
        first.ent_mileage.insert('end', '1')
        self.assertIn('작성 중', self.panel.notebook.tab(first, 'text'))

    def test_hidden_tabs_are_included_in_single_and_batch_records(self):
        first, second = self.two_vehicles()
        batch = SimpleNamespace(app=self.app)
        shared = InspectionBatch.shared(batch)
        self.assertEqual(shared['차량번호'], '90서8616 || 95가0175')
        self.assertEqual(shared['주행거리'], '100 || 80')
        self.assertIn('photo_path:second-photo.jpg', shared['차량점검'])
        builder = fixture()
        builder.fixed_vehicle_widget = first
        builder.vehicle_boxes = [second]
        plan = builder._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 24),
                                                       '중앙지사', prepare_only=True)
        self.assertEqual(plan['records_to_save'][0]['차량번호'], '90서8616 || 95가0175')


class Editor:
    def __init__(self, **data):
        self.data = data
        self._panel_key = data.get('vehicle_info', '')

    def get_data(self):
        data = dict(self.data)
        data['_raw_mileage'] = str(data.get('mileage', '')).replace(',', '')
        return data

    def reset_fields(self): self.data = {}
    def clear_photo(self): self.data['photo_path'] = ''
    def winfo_exists(self): return True
    def destroy(self): pass


class VehiclePanelStateTests(unittest.TestCase):
    def panel(self):
        panel = object.__new__(VehiclePanel)
        first = Editor(vehicle_info='90서8616', mileage='100', out_exterior='양호',
                       photo_path='first.jpg')
        second = Editor(vehicle_info='95가0175', mileage='80', out_exterior='불량',
                        photo_path='second.jpg')
        panel.editors = [first, second]
        panel.saved = {}
        panel.sources = {}
        panel.app = SimpleNamespace(
            vehicle_boxes=[second], vehicle_inspections={second._panel_key: second},
            fixed_vehicle_widget=first, vehicle_widget=first,
            ent_daily_date=SimpleNamespace(get_date=lambda: datetime.date(2026, 10, 24)),
            cb_daily_site=Field('중앙지사'), cb_daily_company=Field('발주처'),
            ent_daily_travel_cost=Field('0'), ent_daily_meal_cost=Field('0'),
            save_single_vehicle_data=Mock(return_value=True))
        panel.notebook = Mock()
        panel.summary = Mock()
        return panel, first, second

    def test_close_x_targets_clicked_inactive_tab(self):
        panel, first, second = self.panel()
        panel.notebook.identify.return_value = 'VehicleTabs.close'
        panel.notebook.tabs.return_value = [str(first), str(second)]
        panel.notebook.index.return_value = 1
        panel.notebook.select.return_value = str(first)
        event = SimpleNamespace(x=210, y=12)
        self.assertEqual(panel.on_tab_close_press(event), 'break')
        with patch('site_apps.central.src.views.vehicle_panel.messagebox.askyesno', return_value=True):
            self.assertEqual(panel.on_tab_close_release(event), 'break')
        self.assertEqual(panel.editors, [first])
        self.assertEqual(panel.app.vehicle_boxes, [])

    def test_dragging_off_close_x_does_not_close_tab(self):
        panel, first, second = self.panel()
        panel.notebook.identify.return_value = 'VehicleTabs.close'
        panel.notebook.tabs.return_value = [str(first), str(second)]
        panel.notebook.index.return_value = 1
        event = SimpleNamespace(x=210, y=12)
        panel.on_tab_close_press(event)
        panel.notebook.identify.return_value = 'Notebook.label'
        with patch('site_apps.central.src.views.vehicle_panel.messagebox.askyesno') as confirm:
            panel.on_tab_close_release(event)
        confirm.assert_not_called()
        self.assertEqual(panel.editors, [first, second])

    def test_cancel_close_preserves_unsaved_checks_and_photo(self):
        panel, first, second = self.panel()
        before = panel.serialize()
        with patch('site_apps.central.src.views.vehicle_panel.messagebox.askyesno', return_value=False):
            panel.request_close(second)
        self.assertEqual(panel.serialize(), before)
        self.assertEqual(len(panel.editors), 2)

    def test_empty_tab_closes_without_confirmation(self):
        panel, first, second = self.panel()
        second.reset_fields()
        with patch('site_apps.central.src.views.vehicle_panel.messagebox.askyesno') as confirm:
            panel.request_close(second)
        confirm.assert_not_called()
        self.assertEqual(panel.editors, [first])

    def test_saved_status_does_not_erase_vehicle_and_changes_invalidate_it(self):
        panel, first, second = self.panel()
        before = panel.serialize()
        panel.app.save_single_vehicle_data.return_value = False
        panel.save_one(second)
        self.assertNotIn(second, panel.saved)
        panel.app.save_single_vehicle_data.return_value = True
        panel.save_one(second)
        self.assertEqual(panel.serialize(), before)
        self.assertIn('저장 완료', panel.notebook.tab.call_args.kwargs['text'])
        with patch('site_apps.central.src.views.vehicle_panel.messagebox.showinfo'):
            panel.save_one(second)
        self.assertEqual(panel.app.save_single_vehicle_data.call_count, 2)
        second.data['mileage'] = '81'
        panel.refresh(second)
        self.assertIn('작성 중', panel.notebook.tab.call_args.kwargs['text'])

    def test_deleting_primary_keeps_remaining_vehicle_and_reset_clears_photo(self):
        panel, first, second = self.panel()
        panel.remove(first)
        self.assertIs(panel.app.fixed_vehicle_widget, second)
        self.assertEqual(panel.app.vehicle_boxes, [])
        self.assertEqual(panel.app.vehicle_inspections, {})
        self.assertEqual(second.data['mileage'], '80')
        panel.reset()
        self.assertEqual(len(panel.editors), 1)
        self.assertEqual(second.data, {'photo_path': ''})

    def test_all_vehicles_are_collected_in_batch_and_single_record(self):
        panel, first, second = self.panel()
        batch = SimpleNamespace(app=panel.app)
        shared = InspectionBatch.shared(batch)
        self.assertEqual(shared['차량번호'], '90서8616 || 95가0175')
        self.assertEqual(shared['주행거리'], '100 || 80')
        self.assertIn('photo_path:second.jpg', shared['차량점검'])
        builder = fixture()
        builder.fixed_vehicle_widget = first
        builder.vehicle_boxes = [second]
        plan = builder._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 24),
                                                       '중앙지사', prepare_only=True)
        self.assertEqual(plan['records_to_save'][0]['차량번호'], '90서8616 || 95가0175')

    def save_method(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'MaterialManager')
        method = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == 'save_single_vehicle_data')
        namespace = {'pd': pd, 'datetime': datetime, 'messagebox': Mock()}
        namespace['messagebox'].askyesno.return_value = True
        exec(compile(ast.fix_missing_locations(ast.Module(body=[method], type_ignores=[])),
                     '<actual vehicle save>', 'exec'), namespace)
        return namespace['save_single_vehicle_data']

    def save_app(self, panel):
        app = panel.app
        app.vehicle_panel = panel
        app.cb_daily_equip = Field('OMNI')
        app.cb_daily_test_method = Field('PAUT')
        app.daily_usage_df = pd.DataFrame()
        app.save_data = Mock(return_value=True)
        app.update_daily_usage_view = Mock()
        app.update_recent_entries_view = Mock()
        app.refresh_inquiry_filters = Mock()
        app.daily_save_status = Mock()
        return app

    def test_actual_individual_save_keeps_tabs_and_persists_selected_vehicle_only(self):
        panel, first, second = self.panel()
        app = self.save_app(panel)
        self.assertTrue(self.save_method()(app, second))
        self.assertEqual(app.daily_usage_df['차량번호'].tolist(), ['95가0175'])
        self.assertEqual(app.daily_usage_df['주행거리'].tolist(), ['80'])
        self.assertEqual(len(panel.editors), 2)
        self.assertEqual(first.data['mileage'], '100')
        self.assertIn(second, panel.saved)

    def test_actual_failed_individual_save_rolls_back_and_keeps_both_tabs(self):
        panel, first, second = self.panel()
        app = self.save_app(panel)
        app.save_data.return_value = False
        self.assertIsNone(self.save_method()(app, second))
        self.assertTrue(app.daily_usage_df.empty)
        self.assertNotIn(second, panel.saved)
        self.assertEqual(len(panel.editors), 2)


if __name__ == '__main__':
    unittest.main()
