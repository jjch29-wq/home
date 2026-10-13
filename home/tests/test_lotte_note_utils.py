from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.note_utils import strip_generated_rej_note


class NoteUtilsTests(unittest.TestCase):
    def test_generated_rej_note_only_becomes_empty(self):
        self.assertEqual(strip_generated_rej_note('[당일 불량(REJ) 0.0개 발생]'), '')

    def test_user_note_is_preserved(self):
        self.assertEqual(
            strip_generated_rej_note('현장 확인 [당일 불량(REJ) 2개 발생]'),
            '현장 확인')

    def test_regular_note_is_unchanged(self):
        self.assertEqual(strip_generated_rej_note('장비 점검 완료'), '장비 점검 완료')


if __name__ == '__main__':
    unittest.main()
