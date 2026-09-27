import json
import sys
import tempfile
import unittest
from pathlib import Path


DESKTOP_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(DESKTOP_DIR))

from glossary_data import GLOSSARY
from problem_summary_data import EXACT_SUMMARIES, summary_for_problem
from standard_references import references_for_problem
from study_content import WEEKLY_CONTENT
import app


class ContentIntegrityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.problems = json.loads((DESKTOP_DIR / "problem_index.json").read_text(encoding="utf-8"))
        cls.study_data = json.loads((DESKTOP_DIR / "study_data.json").read_text(encoding="utf-8"))

    def test_problem_index_is_complete(self):
        self.assertEqual(len(self.problems), 370)
        numbers = [problem["number"] for problem in self.problems]
        self.assertEqual(len(numbers), len(set(numbers)))
        for problem in self.problems:
            self.assertTrue({"number", "title", "week", "study_group", "book_page"}.issubset(problem))
            self.assertIn(problem["week"], range(1, 13))
            self.assertGreater(problem["book_page"], 0)

    def test_weekly_core_answers(self):
        self.assertEqual(len(WEEKLY_CONTENT), 12)
        self.assertEqual(sum(len(item["lessons"]) for item in WEEKLY_CONTENT), 48)
        self.assertTrue(all(len(item["lessons"]) == 4 for item in WEEKLY_CONTENT))

    def test_every_problem_has_a_summary(self):
        for problem in self.problems:
            summary, basis = summary_for_problem(problem)
            self.assertTrue(summary.strip(), problem["number"])
            self.assertTrue(basis.strip(), problem["number"])

    def test_review_statuses_are_valid_and_review_target_is_reached(self):
        allowed = {"검토 완료", "규격 확인 필요", "원문 확인 필요", "초안", "미검토"}
        self.assertEqual(len(EXACT_SUMMARIES), 370)
        for problem in self.problems:
            self.assertIn(app.problem_review_status(problem), allowed)
        counts = app.content_review_counts()
        self.assertEqual(sum(counts.values()), len(self.problems))
        self.assertEqual(counts["초안"], 0)
        self.assertEqual(counts["미검토"], 0)
        self.assertEqual(counts["원문 확인 필요"], 6)

    def test_glossary_terms_are_unique_and_complete(self):
        terms = [entry["term"] for entry in GLOSSARY]
        self.assertEqual(len(terms), len(set(terms)))
        required = {"term", "aliases", "category", "definition", "relation", "formula", "related", "questions"}
        for entry in GLOSSARY:
            self.assertTrue(required.issubset(entry), entry.get("term"))

    def test_standard_reference_matching(self):
        problem = {"title": "STB-A2와 RB-4의 비교", "study_group": "탐촉자·교정", "tags": []}
        references = references_for_problem(problem)
        self.assertTrue(any("KS B 0896" in item for item in references))

    def test_study_data_shape(self):
        required = {"weeks", "questions", "tasks", "answers", "mistakes", "question_notes"}
        self.assertTrue(required.issubset(self.study_data))
        self.assertIsInstance(self.study_data["answers"], dict)
        self.assertIsInstance(self.study_data["question_notes"], list)

    def test_question_notes_link_to_problem_with_new_and_legacy_metadata(self):
        notes = [
            {"problem_number": 143, "source": ""},
            {"source": "143. STB-A2와 RB-4의 비교 · 교재 p.176"},
            {"source": "교재 p.176"},
        ]
        linked = app.question_notes_for_problem(notes, 143)
        self.assertEqual(len(linked), 2)
        self.assertEqual(app.question_note_problem_number(notes[1]), 143)

    def test_atomic_save_creates_previous_version_backup(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            original_data_file = app.DATA_FILE
            original_backup_file = app.DATA_BACKUP_FILE
            try:
                app.DATA_FILE = Path(temp_dir) / "study_data.json"
                app.DATA_BACKUP_FILE = Path(temp_dir) / "study_data.backup.json"
                app.DATA_FILE.write_text('{"version": 1}', encoding="utf-8")
                holder = type("Holder", (), {"data": {"version": 2}})()
                app.StudyApp.save(holder)
                self.assertEqual(json.loads(app.DATA_FILE.read_text(encoding="utf-8"))["version"], 2)
                self.assertEqual(json.loads(app.DATA_BACKUP_FILE.read_text(encoding="utf-8"))["version"], 1)
                self.assertFalse(app.DATA_FILE.with_suffix(".json.tmp").exists())
            finally:
                app.DATA_FILE = original_data_file
                app.DATA_BACKUP_FILE = original_backup_file

    def test_load_uses_backup_when_primary_is_damaged(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            original_data_file = app.DATA_FILE
            original_backup_file = app.DATA_BACKUP_FILE
            try:
                app.DATA_FILE = Path(temp_dir) / "study_data.json"
                app.DATA_BACKUP_FILE = Path(temp_dir) / "study_data.backup.json"
                app.DATA_FILE.write_text("{damaged", encoding="utf-8")
                app.DATA_BACKUP_FILE.write_text('{"weeks": [1]}', encoding="utf-8")
                self.assertEqual(app.load_data()["weeks"], [1])
            finally:
                app.DATA_FILE = original_data_file
                app.DATA_BACKUP_FILE = original_backup_file


if __name__ == "__main__":
    unittest.main()
