"""Recent-record deletion must preserve unrelated stock records and drafts."""
import ast
from pathlib import Path
import sys
from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest
import pandas as pd

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
from site_apps.central.src.views.daily_date_context import load_recent_record


class RecentHistoryTests(unittest.TestCase):
    def setUp(self):
        source = SRC / 'site_apps/central/src/app.py'
        fn = next(n for n in ast.walk(ast.parse(source.read_text(encoding='utf-8')))
                  if isinstance(n, ast.FunctionDef) and n.name == 'delete_recent_entry')
        self.messages = Mock()
        ns = {'messagebox':self.messages}
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(source),'exec'),ns)
        self.delete = ns['delete_recent_entry']
        tv = Mock()
        tv.selection.return_value = ['row']
        tv.item.return_value = ('7',)
        self.app = SimpleNamespace(root=Mock(), tv_recent=tv,
            daily_usage_df=pd.DataFrame([{'Date':'2026-10-23','Site':'A'},
                                        {'Date':'2026-10-23','Site':'A'}],index=[7,8]),
            transactions_df=pd.DataFrame([{'Note':'A 現場 manual'}, {'Note':'auto'}]),
            save_data=Mock(return_value=True), update_daily_usage_view=Mock(),
            update_recent_entries_view=Mock(), update_stock_view=Mock(),
            update_transaction_view=Mock(), refresh_inquiry_filters=Mock(), update_monthly_usage_view=Mock())

    def test_delete_keeps_other_job_ids_and_all_transactions(self):
        before = self.app.transactions_df.copy(deep=True)
        self.delete(self.app)
        self.assertEqual(list(self.app.daily_usage_df.index), [8])
        pd.testing.assert_frame_equal(self.app.transactions_df, before)
        self.app.save_data.assert_called_once_with(require_primary=True)
        self.app.update_monthly_usage_view.assert_called_once_with()

    def test_save_failure_or_exception_restores_originals(self):
        for failure in (False, OSError('locked')):
            with self.subTest(failure=failure):
                before = self.app.daily_usage_df.copy(deep=True)
                self.app.save_data.side_effect = failure if isinstance(failure, Exception) else None
                self.app.save_data.return_value = False
                self.delete(self.app)
                pd.testing.assert_frame_equal(self.app.daily_usage_df, before)
        self.app.update_stock_view.assert_not_called()

    def test_cancel_keeps_records(self):
        self.messages.askyesno.return_value = False
        self.delete(self.app)
        self.app.save_data.assert_not_called()
        self.assertEqual(list(self.app.daily_usage_df.index), [7,8])

    def test_load_cancel_keeps_input_and_batch(self):
        app = SimpleNamespace(root=Mock(), inspection_batch=Mock(),
            clear_daily_usage_form_all=Mock(), load_daily_usage_to_form=Mock())
        with patch('site_apps.central.src.views.daily_date_context.has_daily_input',return_value=True), \
             patch('site_apps.central.src.views.daily_date_context.messagebox.askyesno',return_value=False):
            self.assertFalse(load_recent_record(app, {'Date':'2026-10-23'}))
        app.inspection_batch.clear.assert_not_called()
        app.clear_daily_usage_form_all.assert_not_called()
        app.load_daily_usage_to_form.assert_not_called()

    def test_confirm_clears_old_batch_before_loading(self):
        app = SimpleNamespace(root=Mock(), inspection_batch=Mock(),
            clear_daily_usage_form_all=Mock(), load_daily_usage_to_form=Mock())
        with patch('site_apps.central.src.views.daily_date_context.has_daily_input',return_value=True), \
             patch('site_apps.central.src.views.daily_date_context.messagebox.askyesno',return_value=True):
            self.assertTrue(load_recent_record(app, {'Date':'2026-10-23'}))
        app.inspection_batch.clear.assert_called_once()
        app.load_daily_usage_to_form.assert_called_once_with({'Date':'2026-10-23'})


if __name__ == '__main__':
    unittest.main()
