# -*- coding: utf-8 -*-
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from site_apps.central.src.reports.ndt_history_import import read_unissued_results


class TestCentralNDTHistoryImport(unittest.TestCase):
    def test_already_issued_source_is_excluded(self):
        rows = [
            {"_source_id": "issued-id", "_source_site": "한국지역난방 중앙지사"},
            {"_source_id": "pending-id", "_source_site": "한국지역난방 중앙지사"},
        ]
        history = [
            {
                "id": "issue-1",
                "mode": "PAUT",
                "report_no": "SIS/GI-KDHC-JY-PAUT-012",
                "items": [dict(rows[0])],
            }
        ]
        with tempfile.TemporaryDirectory() as directory:
            issue_path = Path(directory) / "report_issue_history.json"
            issue_path.write_text(json.dumps(history, ensure_ascii=False), encoding="utf-8")
            with patch(
                "site_apps.central.src.reports.ndt_history_import.read_results",
                return_value=[dict(row) for row in rows],
            ):
                pending, excluded = read_unissued_results(
                    "한국지역난방 중앙지사", "PAUT", issue_path=issue_path
                )

        self.assertEqual(excluded, 1)
        self.assertEqual([row["_source_id"] for row in pending], ["pending-id"])

    def test_empty_or_deleted_issue_number_does_not_exclude_source(self):
        row = {"_source_id": "source-id", "_source_site": "한국지역난방 중앙지사"}
        history = [
            {"id": "empty", "mode": "PAUT", "report_no": "", "items": [dict(row)]},
            {
                "id": "deleted",
                "mode": "PAUT",
                "report_no": "SIS/GI-KDHC-JY-PAUT-013",
                "deleted_at": "2026-11-26T12:00:00+09:00",
                "items": [dict(row)],
            },
        ]
        with tempfile.TemporaryDirectory() as directory:
            issue_path = Path(directory) / "report_issue_history.json"
            issue_path.write_text(json.dumps(history, ensure_ascii=False), encoding="utf-8")
            with patch(
                "site_apps.central.src.reports.ndt_history_import.read_results",
                return_value=[dict(row)],
            ):
                pending, excluded = read_unissued_results(
                    "한국지역난방 중앙지사", "PAUT", issue_path=issue_path
                )

        self.assertEqual(excluded, 0)
        self.assertEqual(len(pending), 1)


if __name__ == "__main__":
    unittest.main()
