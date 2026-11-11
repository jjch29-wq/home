from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace, MethodType
from unittest.mock import Mock, patch
import sys
import unittest
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.company_memos import memo_snapshot, memo_history_rows, report_memo_text
from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab
from site_apps.gyeongin.src.daily_work_log_exporter import DailyWorkLogExporter


class Field:
    def __init__(self, text=''):
        self.text = text
        self.options = {}
    def get(self, *args): return self.text
    def set(self, text): self.text = text
    def configure(self, **kwargs): self.options.update(kwargs)
    def delete(self, *args): self.text = ''
    def insert(self, index, text): self.text = text


class CompanyMemoTests(unittest.TestCase):
    def tab(self):
        app = SimpleNamespace(company_memos={}, _company_memo_company=None,
            company_memo_text=Field(), company_memo_selector=Field(), company_memo_hint=Mock(),
            memo_notebook=Mock(), company_memo_page=object(), date_entry=Field('2026-11-12'),
            remarks_text=Field('공통 지시'), weather_entry=Field(), qty_entries={}, equip_entries={},
            personnel_entries={}, ndt_grid_entries=[],
            company_summary_rows=[dict(company='업체A', method='RT'), dict(company='업체A', method='MT'), dict(company='업체B')])
        for method in ('_company_memo_snapshot', '_refresh_company_memo_choices', '_select_company_memo', '_load_company_memos'):
            setattr(app, method, MethodType(getattr(DailyWorkLogTab, method), app))
        return app

    def test_switching_company_keeps_drafts_and_same_company_shares_methods(self):
        tab = self.tab()
        tab._select_company_memo('업체A')
        tab.company_memo_text.text = '성적서 오전 제출'
        tab._select_company_memo('업체A')
        self.assertEqual(tab.company_memo_text.text, '성적서 오전 제출')
        tab._select_company_memo('업체B')
        tab.company_memo_text.text = '추가 검사 예정'
        tab._select_company_memo('업체A')
        self.assertEqual(tab.company_memo_text.text, '성적서 오전 제출')
        self.assertEqual(tab._company_memo_snapshot(), {'업체A': '성적서 오전 제출', '업체B': '추가 검사 예정'})
        self.assertEqual(tab.remarks_text.text, '공통 지시')

    def test_loading_another_date_never_copies_previous_notes(self):
        tab = self.tab()
        tab._load_company_memos({'업체A': '어제 메모'})
        tab._select_company_memo('업체A')
        tab._load_company_memos({})
        tab._select_company_memo('업체A')
        self.assertEqual(tab.company_memo_text.text, '')
        self.assertEqual(tab._company_memo_snapshot(), {})

    def test_deleting_text_removes_only_selected_company(self):
        notes = {'업체A': 'A', '업체B': 'B'}
        self.assertEqual(memo_snapshot(notes, '업체A', ''), {'업체B': 'B'})
        self.assertEqual(notes, {'업체A': 'A', '업체B': 'B'})

    def test_unsaved_company_edit_changes_dirty_snapshot(self):
        tab = self.tab()
        before = DailyWorkLogTab._capture_log_input(tab)
        tab._select_company_memo('업체A')
        tab.company_memo_text.text = '미저장 메모'
        self.assertNotEqual(DailyWorkLogTab._capture_log_input(tab), before)

    def test_cancel_date_change_keeps_unsaved_company_text(self):
        tab = self.tab()
        tab._capture_log_input = MethodType(DailyWorkLogTab._capture_log_input, tab)
        tab._log_baseline = tab._capture_log_input()
        tab._loaded_log_date = '2026-11-12'
        tab._selected_date_key = lambda: '2026-11-13'
        tab.date_entry.set_date = Mock()
        tab.load_history = Mock()
        tab._select_company_memo('업체A')
        tab.company_memo_text.text = '아직 저장하지 않은 메모'
        with patch('site_apps.gyeongin.src.daily_work_log_tab.messagebox.askyesno', return_value=False):
            DailyWorkLogTab.on_date_change(tab)
        tab.load_history.assert_not_called()
        self.assertEqual(tab.company_memo_text.text, '아직 저장하지 않은 메모')
        tab.date_entry.set_date.assert_called_once_with('2026-11-12')

    def test_choices_preserve_notes_even_when_company_not_in_current_query(self):
        tab = self.tab()
        tab._load_company_memos({'이전업체': '남겨 둘 메모'})
        self.assertEqual(tab.company_memo_selector.options['values'], ['업체A', '업체B', '이전업체'])

    def test_history_includes_common_and_company_notes_from_legacy_and_new_days(self):
        history = {'2026-11-11': {'remarks': '기존 공통'},
                   '2026-11-12': {'remarks': '전체 공지', 'company_memos': {'업체A': '제출 예정'}}}
        before = deepcopy(history)
        rows = memo_history_rows(history)
        self.assertIn(('2026-11-12', '업체A', '제출 예정'), rows)
        self.assertIn(('2026-11-11', '', '기존 공통'), rows)
        self.assertEqual(history, before)

    def test_export_distinguishes_companies_and_preserves_full_long_text(self):
        notes = {'업체A': '성적서 오전 제출', '업체B': '다음 작업\n' * 100}
        self.assertIn('업체별 메모', report_memo_text('전체 공지', notes))
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'memo.xlsx'
            DailyWorkLogExporter().generate_excel(dict(date='2026-11-12', company_summary=[], remarks='전체 공지', company_memos=notes), path)
            book = openpyxl.load_workbook(path)
            try:
                sheet = book['업체별 메모']
                self.assertEqual(sheet['B2'].value, '공통')
                self.assertEqual(sheet['B3'].value, '업체A')
                self.assertEqual(sheet['C4'].value, notes['업체B'])
                self.assertIn('업체별 메모', book.active['L16'].value)
            finally:
                book.close()


if __name__ == '__main__': unittest.main()
