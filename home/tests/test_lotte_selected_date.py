"""선택한 날짜의 저장 기록 조회 검증 (GUI 실행 불필요)."""
import ast
import datetime
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

import pandas as pd

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
from site_apps.lotte.src.worker_assignment import worker_names


class Tree:
    def __init__(self):
        self.rows = {}

    def get_children(self):
        return list(self.rows)

    def delete(self, item):
        del self.rows[item]

    def insert(self, parent, position, **kwargs):
        self.rows[len(self.rows)] = kwargs

    def see(self, item):
        pass


class SelectedDateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        source = SRC / 'site_apps/lotte/src/app.py'
        tree = ast.parse(source.read_text(encoding='utf-8-sig'))
        method = next(node for node in ast.walk(tree)
                      if isinstance(node, ast.FunctionDef) and node.name == 'update_recent_entries_view')
        ns = dict(pd=pd, datetime=datetime, worker_names=worker_names)
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), ns)
        cls.refresh = staticmethod(ns['update_recent_entries_view'])

    def app(self, records):
        return SimpleNamespace(
            daily_usage_df=pd.DataFrame(records), tv_recent=Tree(),
            ent_daily_date=SimpleNamespace(get_date=lambda: datetime.date(2026, 1, 7)),
            recent_frame=SimpleNamespace(configure=lambda **kwargs: None))

    def test_selected_date_only_and_original_index_is_preserved(self):
        app = self.app([dict(Date='2026-01-06', User='전일'),
                        dict(Date='2026-01-07 12:00:00', User2='검사자'),
                        dict(Date='invalid', User='오류')])
        self.refresh(app)
        self.assertEqual(len(app.tv_recent.rows), 1)
        row = next(iter(app.tv_recent.rows.values()))
        self.assertEqual(row['values'][0], 1)
        self.assertEqual(row['tags'], ('1',))
        self.assertEqual(row['values'][8], '검사자')

    def test_date_without_records_shows_empty_list(self):
        app = self.app([dict(Date='2026-01-06')])
        app.tv_recent.insert('', 'end', values=['old'])
        self.refresh(app)
        self.assertEqual(app.tv_recent.rows, {})

    def test_all_selected_date_records_are_available(self):
        app = self.app([dict(Date='2026-01-07', EntryTime=f'2026-01-07 10:{i:02d}:00')
                        for i in range(35)])
        self.refresh(app)
        self.assertEqual(len(app.tv_recent.rows), 35)
        self.assertEqual(app.tv_recent.rows[0]['values'][0], 34)


if __name__ == '__main__':
    unittest.main()
