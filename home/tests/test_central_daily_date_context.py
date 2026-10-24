"""Changing the entry date must not silently move a previous job or batch."""
import ast
import datetime
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import pandas as pd

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
from site_apps.central.src.views.daily_date_context import change_daily_date, has_daily_input, show_recent_scope


class Field:
    def __init__(self, value=''): self.value = value
    def get(self): return self.value
    def get_date(self): return self.value
    def set(self, value): self.value = value
    def set_date(self, value): self.value = value
    def delete(self, *args): self.value = ''


def fixture(data):
    crew = SimpleNamespace(cb_name=Field('작업자'), participation_var=Field(True),
        cb_shift=Field('야간'), ent_worktime=Field('22:00~익일09:00'),
        ent_ot=Field('77500'), ent_meal=Field('1000'))
    return SimpleNamespace(ent_daily_date=Field(datetime.date(2026, 10, 24)),
        _last_daily_date=datetime.date(2026, 10, 23), cb_daily_company=Field('한국지역난방공사'),
        cb_daily_site=Field('중앙지사'), root=Mock(), _get_usage_session_data=lambda:data,
        inspection_batch=Mock(), clear_daily_usage_form_all=Mock(),
        worker_group1=crew, main_memo_text=Mock(), daily_save_status=Mock(),
        update_recent_entries_view=Mock())


class DailyDateContextTests(unittest.TestCase):
    def test_record_scope_switch_does_not_change_date_or_discard_input(self):
        app = fixture({'amount':'18'})
        app._daily_recent_date_filter = app.ent_daily_date.get_date()
        show_recent_scope(app)
        self.assertIsNone(app._daily_recent_date_filter)
        self.assertTrue(app._daily_recent_show_all)
        show_recent_scope(app, selected_only=True)
        self.assertEqual(app._daily_recent_date_filter, datetime.date(2026,10,24))
        self.assertFalse(app._daily_recent_show_all)
        self.assertEqual(app.worker_group1.ent_ot.get(), '77500')
        self.assertEqual(app.ent_daily_date.get_date(), datetime.date(2026,10,24))
        app.inspection_batch.clear.assert_not_called()
        app.clear_daily_usage_form_all.assert_not_called()
        self.assertEqual(app.update_recent_entries_view.call_count,2)

    def test_cancel_restores_date_and_preserves_input_and_batch(self):
        app = fixture({'amount':'18', 'inspection_batch':[{'work':'RT'}]})
        with patch('site_apps.central.src.views.daily_date_context.messagebox.askyesno', return_value=False):
            change_daily_date(app)
        self.assertEqual(app.ent_daily_date.get_date(), datetime.date(2026,10,23))
        self.assertEqual(app.worker_group1.ent_ot.get(), '77500')
        app.clear_daily_usage_form_all.assert_not_called()
        app.inspection_batch.clear.assert_not_called()
        app.update_recent_entries_view.assert_not_called()

    def test_confirm_starts_clean_date_and_retains_site_and_roster(self):
        app = fixture({'amount':'18'})
        with patch('site_apps.central.src.views.daily_date_context.messagebox.askyesno', return_value=True):
            change_daily_date(app)
        self.assertEqual(app.cb_daily_site.get(), '중앙지사')
        self.assertEqual(app.cb_daily_company.get(), '한국지역난방공사')
        self.assertEqual(app.worker_group1.cb_name.get(), '작업자')
        self.assertFalse(app.worker_group1.participation_var.get())
        self.assertEqual(app.worker_group1.ent_worktime.get(), '')
        self.assertEqual(app.worker_group1.ent_ot.get(), '')
        self.assertEqual(app.worker_group1.ent_meal.get(), '')
        self.assertEqual(app._daily_recent_date_filter, datetime.date(2026,10,24))
        app.inspection_batch.clear.assert_called_once()
        app.clear_daily_usage_form_all.assert_called_once_with(keep_date=True, keep_workers=True)
        app.update_recent_entries_view.assert_called_once()

    def test_empty_entry_switches_without_prompt_and_same_date_is_noop(self):
        app = fixture({})
        app.worker_group1.ent_meal.set('')
        with patch('site_apps.central.src.views.daily_date_context.messagebox.askyesno') as ask:
            change_daily_date(app)
            change_daily_date(app)
        ask.assert_not_called()
        app.inspection_batch.clear.assert_called_once()

    def test_vehicle_or_worker_meal_alone_needs_confirmation(self):
        app = fixture({})
        self.assertTrue(has_daily_input(app))
        app.worker_group1.ent_meal.set('')
        app.fixed_vehicle_widget = SimpleNamespace(get_data=lambda:{'vehicle_info':'81두1580'})
        self.assertTrue(has_daily_input(app))

    def test_selected_date_records_keep_original_ids_and_do_not_load_form(self):
        tree = ast.parse((SRC / 'site_apps/central/src/app.py').read_text(encoding='utf-8-sig'))
        cls = next(n for n in tree.body if isinstance(n,ast.ClassDef) and n.name == 'MaterialManager')
        fn = next(n for n in cls.body if isinstance(n,ast.FunctionDef) and n.name == 'update_recent_entries_view')
        ns = {'pd':pd}
        exec(compile(ast.Module(body=[fn],type_ignores=[]),'<recent records>','exec'),ns)
        tv = Mock()
        tv.get_children.return_value = ['old']
        app = SimpleNamespace(tv_recent=tv, recent_frame=Mock(), recent_scope_status=Mock(), cb_daily_site=Field('중앙지사'),
            _daily_recent_date_filter=datetime.date(2026,10,24),
            daily_usage_df=pd.DataFrame([
                {'Date':'2026-10-23','Site':'중앙지사'},
                {'Date':'2026-10-24','Site':'다른 현장'},
                {'Date':'2026-10-24','Site':'중앙지사'}],index=[7,8,9]))
        original = app.daily_usage_df.copy(deep=True)
        ns['update_recent_entries_view'](app)
        self.assertEqual(tv.insert.call_count,1)
        self.assertEqual(tv.insert.call_args.kwargs['values'][0],9)
        app.recent_frame.config.assert_called_with(text='2026-10-24 저장 내역 · 1건')
        app.recent_scope_status.config.assert_called_with(text='2026-10-24 · 중앙지사 · 1건')
        pd.testing.assert_frame_equal(app.daily_usage_df,original)
        app._daily_recent_date_filter = datetime.date(2026,10,25)
        tv.insert.reset_mock()
        ns['update_recent_entries_view'](app)
        tv.insert.assert_not_called()
        app.recent_frame.config.assert_called_with(text='2026-10-25 저장 내역 · 0건')
        app.recent_scope_status.config.assert_called_with(text='2026-10-25 · 중앙지사 · 0건')
        app._daily_recent_date_filter = None
        app._daily_recent_show_all = True
        ns['update_recent_entries_view'](app)
        app.recent_scope_status.config.assert_called_with(text='전체 · 3건')
        app._daily_recent_show_all = False
        ns['update_recent_entries_view'](app)
        app.recent_scope_status.config.assert_called_with(text='최근 · 3건 / 전체 3건')


if __name__ == '__main__':
    unittest.main()
