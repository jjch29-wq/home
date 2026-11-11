from copy import deepcopy
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.worker_review import review_participants


class WorkerReviewTests(unittest.TestCase):
    def row(self, index, name='', time='', selected=True):
        return dict(index=index, name=name, time=time, selected=selected)

    def test_initial_checked_empty_rows_do_not_raise_missing_time_warnings(self):
        review = review_participants([self.row(i) for i in range(1, 11)])
        self.assertEqual(review, {'count': 0, 'names': [], 'warnings': []})

    def test_only_current_participants_are_reviewed(self):
        rows = [self.row(1, '김민수', '09:00~18:00'),
                self.row(2, '박영희', selected=False), self.row(3, ' ')]
        before = deepcopy(rows)
        self.assertEqual(review_participants(rows),
                         {'count': 1, 'names': ['김민수'], 'warnings': []})
        self.assertEqual(rows, before)

    def test_missing_time_and_duplicate_include_row_identifiers(self):
        review = review_participants([self.row(1, '김민수', '09:00~18:00'),
                                      self.row(6, ' 김민수 ', ' '), self.row(7, '박영희')])
        self.assertEqual(review['count'], 2)
        self.assertEqual(review['warnings'],
                         ['작업시간 미입력: 06 김민수, 07 박영희', '중복 선택: 김민수 (01·06번)'])

    def test_warnings_clear_after_time_edit_and_unchecking_duplicate(self):
        rows = [self.row(1, '김민수'), self.row(6, '김민수', '09:00~18:00')]
        self.assertEqual(len(review_participants(rows)['warnings']), 2)
        rows[0]['time'] = '09:00~18:00'
        rows[1]['selected'] = False
        self.assertEqual(review_participants(rows)['warnings'], [])


if __name__ == '__main__':
    unittest.main()
