"""Exercise history actions without opening the app or writing business data."""
import ast
import datetime
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock
import pandas as pd

SOURCE = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
TREE = ast.parse(SOURCE.read_text(encoding='utf-8'))


class InoutHistoryTests(unittest.TestCase):
    def setUp(self):
        self.messages, self.dialog = Mock(), Mock()
        self.ns = {'pd':pd, 'messagebox':self.messages, 'filedialog':self.dialog, 'datetime':datetime}
        for name in ('delete_transaction_entry', 'export_transaction_history'):
            fn = next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef) and n.name == name)
            exec(compile(ast.Module(body=[fn],type_ignores=[]), str(SOURCE), 'exec'), self.ns)
        tv = Mock()
        tv.selection.return_value = ['item']
        tv.item.return_value = ('7', 'DAILY')
        self.app = SimpleNamespace(root=Mock(), inout_tree=tv,
            transactions_df=pd.DataFrame({'Quantity':[2]},index=[4]),
            daily_usage_df=pd.DataFrame({'Usage':[5], 'DailyMemo':['memo']}, index=[7]),
            save_data=Mock(return_value=True), update_transaction_view=Mock(), update_stock_view=Mock(),
            update_recent_entries_view=Mock(), update_daily_usage_view=Mock(), update_monthly_usage_view=Mock())

    def test_failed_save_restores_records_and_does_not_report_success(self):
        before = self.app.daily_usage_df.copy(deep=True)
        self.app.save_data.return_value = False
        self.ns['delete_transaction_entry'](self.app)
        pd.testing.assert_frame_equal(self.app.daily_usage_df, before)
        self.messages.showerror.assert_called_once()
        self.messages.showinfo.assert_not_called()
        self.app.update_stock_view.assert_not_called()
        self.app.save_data.assert_called_once_with(require_primary=True)

    def test_exception_restores_records(self):
        self.app.save_data.side_effect = OSError('locked')
        self.ns['delete_transaction_entry'](self.app)
        self.assertEqual(list(self.app.daily_usage_df.index), [7])
        self.messages.showinfo.assert_not_called()

    def test_cancel_keeps_original_and_explains_daily_record_deletion(self):
        self.messages.askyesno.return_value = False
        self.ns['delete_transaction_entry'](self.app)
        self.assertIn('원본 일일 작업 기록 전체', self.messages.askyesno.call_args.args[1])
        self.app.save_data.assert_not_called()
        self.assertEqual(list(self.app.daily_usage_df.index), [7])

    def test_success_refreshes_record_views(self):
        self.ns['delete_transaction_entry'](self.app)
        self.assertTrue(self.app.daily_usage_df.empty)
        self.app.update_monthly_usage_view.assert_called_once_with()
        self.app.update_recent_entries_view.assert_called_once()
        self.messages.showinfo.assert_called_once()

    def test_export_reaches_save_and_preserves_all_displayed_rows(self):
        tv = self.app.inout_tree
        tv.__getitem__ = Mock(return_value=['Date', 'Quantity'])
        tv.heading.side_effect = lambda col, *args: col
        tv.get_children.return_value = list(range(501))
        tv.item.side_effect = lambda item, *args: ('2026-10-23', str(item))
        self.dialog.asksaveasfilename.return_value = 'unused.xlsx'
        self.app.clean_df_export = lambda frame: frame
        self.app.save_df_to_excel_autofit = Mock()
        self.ns['export_transaction_history'](self.app)
        self.dialog.asksaveasfilename.assert_called_once()
        frame, path, sheet = self.app.save_df_to_excel_autofit.call_args.args
        self.assertEqual(len(frame), 501)
        self.assertEqual(path, 'unused.xlsx')


if __name__ == '__main__':
    unittest.main()
