"""Inquiry note flags and popups must use the same date/site memo contents."""
from pathlib import Path
import sys
import ast
from types import SimpleNamespace
import unittest
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.memo_details import daily_memo_lookup, record_notes


class MemoDetailsTests(unittest.TestCase):
    def test_date_site_lookup_preserves_full_text_and_deduplicates(self):
        frame = pd.DataFrame([
            {'Date': '2026-10-23', 'Site': 'A', 'DailyMemo': 'first\nsecond'},
            {'Date': pd.Timestamp('2026-10-23'), 'Site': 'A', 'DailyMemo': 'first\nsecond'},
            {'Date': '2026-10-23', 'Site': 'A', 'DailyMemo': 'another'},
            {'Date': '2026-10-23', 'Site': 'B', 'DailyMemo': 'other site'},
            {'Date': '2026-10-24', 'Site': 'A', 'DailyMemo': 'other day'},
        ])
        lookup = daily_memo_lookup(frame)
        self.assertEqual(record_notes({'Date':'2026-10-23', 'Site':'A', 'Note':'record'}, lookup),
                         ('record', 'first\nsecond\n\nanother'))
        self.assertEqual(record_notes({'Date':'2026-10-23', 'Site':'B'}, lookup), ('', 'other site'))

    def test_history_uses_all_distinct_memos_for_selected_site(self):
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
        tree = ast.parse(source.read_text(encoding='utf-8'))
        fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)
                  and n.name == 'open_usage_memo_history')
        body = []
        for node in fn.body:
            if isinstance(node, ast.Import) and any(a.name == 'tkinter' for a in node.names):
                break
            body.append(node)
        fn.body = body + [ast.Return(value=ast.Tuple(elts=[
            ast.Name(id='full_memos', ctx=ast.Load()),
            ast.Name(id='memo_rows', ctx=ast.Load())], ctx=ast.Load()))]
        ast.fix_missing_locations(fn)
        ns = {}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), str(source), 'exec'), ns)
        frame = pd.DataFrame([
            {'Date':'2026-10-23', 'Site':'A', 'DailyMemo':'first\nline'},
            {'Date':'2026-10-23 12:00:00', 'Site':'A', 'DailyMemo':'second'},
            {'Date':'2026-10-23', 'Site':'A', 'DailyMemo':'first\nline'},
            {'Date':'2026-10-23', 'Site':'B', 'DailyMemo':'other site'},
            {'Date':'2026-10-24', 'Site':'A', 'DailyMemo':'   '},
        ])
        original = frame.copy(deep=True)
        app = SimpleNamespace(daily_usage_df=frame,
                              cb_daily_site=SimpleNamespace(get=lambda:'A'))
        full, rows = ns['open_usage_memo_history'](app)
        self.assertEqual(full, {'2026-10-23':'first\nline\n\nsecond'})
        self.assertEqual(rows, [('2026-10-23','first line second')])
        pd.testing.assert_frame_equal(frame, original)

    def test_blank_missing_nan_and_invalid_dates_do_not_create_flags(self):
        frame = pd.DataFrame([
            {'Date':'2026-10-23', 'Site':'A', 'DailyMemo':float('nan')},
            {'Date':'2026-10-24', 'Site':'A', 'DailyMemo':'   '},
            {'Date':'invalid', 'Site':'A', 'DailyMemo':'ignored'},
        ])
        lookup = daily_memo_lookup(frame)
        self.assertEqual(lookup, {})
        self.assertEqual(record_notes({'Note':float('nan')}, lookup), ('', ''))
        self.assertEqual(record_notes({'Note':'record only'}, lookup), ('record only', ''))


if __name__ == '__main__':
    unittest.main()
