import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "home" / "src"))
CENTRAL = ROOT / "home" / "src" / "site_apps" / "central"


class TestCentralMobileProgress(unittest.TestCase):
    def test_reference_snapshot_contains_progress_history(self):
        ref_path = CENTRAL / "google_apps_script" / "Reference.gs"
        self.assertTrue(ref_path.exists(), "Reference.gs should exist")
        text = ref_path.read_text(encoding="utf-8")
        self.assertIn("daily_progress_history", text)
        self.assertIn("default_expected_qty", text)

    def test_mobile_package_contains_progress_feature(self):
        code_path = ROOT / "outputs" / "central-mobile-gas" / "Code.gs"
        self.assertTrue(code_path.exists(), "outputs/central-mobile-gas/Code.gs should exist")
        code = code_path.read_text(encoding="utf-8")

        # Backend logic verification
        self.assertIn("function getProgressSummary", code)
        self.assertIn("classifyMethodSpec", code)
        self.assertIn("api === 'progress'", code)

        # Frontend UI verification
        self.assertIn("공정 현황", code)
        self.assertIn("workspace-progress", code)
        self.assertIn("workspaceLoadProgress", code)
        self.assertIn("progress-kpi-grid", code)
        self.assertIn("progress-table", code)

    def test_qty_row_keys_match_pc_daily_work_log(self):
        from home.src.site_apps.central.src.daily_work_log_tab import DailyWorkLogTab

        expected_qty_keys = [
            ("PAUT", "300A이상"), ("PAUT", "300A이상-야간"), ("PAUT", "250A"), ("PAUT", "250A-야간"),
            ("PAUT", "200A"), ("PAUT", "200A-야간"), ("PAUT", "150A~125A"), ("PAUT", "150A~125A-야간"),
            ("PAUT", "100A이하"), ("PAUT", "100A이하-야간"), ("PAUT", "소계"),
            ("RT", "150A~100A"), ("RT", "150A~100A-야간"), ("RT", "80A이하"), ("RT", "80A이하-야간"),
            ("RT", "소계"),
            ("MT", "전체(주간)"), ("MT", "전체(야간)"),
            ("PT", "전체(주간)"), ("PT", "전체(야간)")
        ]

        # Inspect Workspace.gs rows
        workspace_path = CENTRAL / "google_apps_script" / "Workspace.gs"
        text = workspace_path.read_text(encoding="utf-8")
        for method, spec in expected_qty_keys:
            comp_key = f"{method}_{spec}"
            self.assertIn(comp_key, text, f"Missing key in Workspace.gs: {comp_key}")

    def test_billing_records_integrated_in_mobile(self):
        ref_path = CENTRAL / "google_apps_script" / "Reference.gs"
        ref_text = ref_path.read_text(encoding="utf-8")
        self.assertIn("billing_records", ref_text)

        code_path = ROOT / "outputs" / "central-mobile-gas" / "Code.gs"
        code = code_path.read_text(encoding="utf-8")
        self.assertIn("billing_records", code)
        self.assertIn("workspace-progress-billing", code)
        self.assertIn("일별 기성 정산 상세 목록", code)
        self.assertIn("누적 기성 공급가액", code)


if __name__ == "__main__":
    unittest.main()
