import shutil
import sys
import tempfile
import unittest
from pathlib import Path

SRC_ROOT = Path(__file__).resolve().parents[1] / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from site_apps.gyeongin.src.sync_gas_assets import (
    CODE_GS_PATH,
    INDEX_HTML_PATH,
    build_gas_master_blocks,
    sync_gas_masters,
)


class TestGyeonginGasSync(unittest.TestCase):
    def test_live_gas_assets_are_in_sync_with_pc_masters(self):
        result = sync_gas_masters(write=False)
        self.assertEqual(
            result["status"],
            "in_sync",
            f"GAS master blocks drifted from PC sources: {result['changed_files']}. "
            "Run: python -m site_apps.gyeongin.src.sync_gas_assets",
        )
        self.assertEqual(result["changed_files"], [])

    def test_sync_detects_drift_and_repairs_temp_copies(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            gs_copy = tmp / "Code.gs"
            html_copy = tmp / "index.html"
            shutil.copy2(CODE_GS_PATH, gs_copy)
            shutil.copy2(INDEX_HTML_PATH, html_copy)

            blocks = build_gas_master_blocks()
            self.assertIn("APP_CONFIG_CORE", blocks)
            self.assertIn("COMPANY_NDT_PRESETS", blocks)
            self.assertIn("MAIN_EQUIP_BY_METHOD", blocks)
            self.assertIn("AUX_MASTER_DATA", blocks)

            # Introduce drift in copied index.html
            corrupted = html_copy.read_text(encoding="utf-8").replace("PAUT MX2(OMNI2-103085)", "STALE_EQUIP")
            html_copy.write_text(corrupted, encoding="utf-8")

            check_res = sync_gas_masters(code_gs_path=gs_copy, index_html_path=html_copy, write=False)
            self.assertEqual(check_res["status"], "drifted")
            self.assertEqual(check_res["changed_files"], [str(html_copy)])

            repair_res = sync_gas_masters(code_gs_path=gs_copy, index_html_path=html_copy, write=True)
            self.assertEqual(repair_res["status"], "updated")

            verify_res = sync_gas_masters(code_gs_path=gs_copy, index_html_path=html_copy, write=False)
            self.assertEqual(verify_res["status"], "in_sync")

    def test_gyeongin_runtime_paths_do_not_fallback_to_central(self):
        gyeongin_src = Path(__file__).resolve().parents[1] / "src" / "site_apps" / "gyeongin" / "src"
        app_text = (gyeongin_src / "app.py").read_text(encoding="utf-8")
        self.assertNotIn("'Documents', 'MaterialManager', 'central'", app_text)
        billing_text = (gyeongin_src / "section_line_billing_dialog.py").read_text(encoding="utf-8")
        self.assertNotIn(r"PMI_Data\central", billing_text)


if __name__ == "__main__":
    unittest.main()

