"""Tests for multi-site / annual-contract / SMB profiles & shared-inventory isolation."""
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from site_apps.central.src import data_sync
from site_apps.central.src import site_profile
from site_apps.central.src.services.smb_statement_exporter import (
    build_statement_rows,
    export_smb_statement_excel,
)


class CentralMultiSiteProfileTests(unittest.TestCase):
    def setUp(self):
        self._prev_env = os.environ.get(site_profile.PROFILE_ENV_VAR)

    def tearDown(self):
        if self._prev_env is None:
            os.environ.pop(site_profile.PROFILE_ENV_VAR, None)
        else:
            os.environ[site_profile.PROFILE_ENV_VAR] = self._prev_env

    def test_builtin_profiles_discoverable_and_central_is_first(self):
        profiles = site_profile.list_site_profiles()
        ids = [p.profile_id for p in profiles]
        self.assertGreaterEqual(len(profiles), 3)
        self.assertEqual(ids[0], "central")
        self.assertIn("annual_contract", ids)
        self.assertIn("smb_general", ids)

        central = site_profile.load_site_profile("central")
        self.assertTrue(central.is_central)
        self.assertTrue(central.enable_hot_tapping)
        self.assertEqual(central.enabled_tabs, site_profile.CENTRAL_DEFAULT_TAB_KEYS)

        annual = site_profile.load_site_profile("annual_contract")
        self.assertEqual(annual.site_type, "annual_contract")
        self.assertEqual(annual.inventory_mode, "shared")
        self.assertFalse(annual.enable_hot_tapping)
        self.assertTrue(annual.is_tab_enabled("ndt_billing"))

        smb = site_profile.load_site_profile("smb_general")
        self.assertEqual(smb.site_type, "smb_general")
        self.assertEqual(smb.inventory_mode, "shared")
        self.assertFalse(smb.is_tab_enabled("work_log"))
        self.assertFalse(smb.is_tab_enabled("ndt_billing"))
        self.assertTrue(smb.is_tab_enabled("daily_usage"))
        self.assertTrue(smb.is_tab_enabled("monthly"))
        self.assertTrue(smb.is_tab_enabled("ndt_report"))

    def test_shared_inventory_with_isolated_operational_data(self):
        shared_db = r"C:\shared\Material_Inventory.xlsx"
        default_cfg = r"C:\central\Material_Manager_Config.json"
        default_billing = r"C:\central\config.json"

        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(site_profile, "get_profiles_root", return_value=tmp):
                smb = site_profile.create_site_profile(
                    "client_alpha",
                    "알파기계 건별",
                    site_type="smb_general",
                    inventory_mode="shared",
                    companies=["알파기계"],
                )
                iso = site_profile.create_site_profile(
                    "branch_isolated",
                    "독립재고 지사",
                    site_type="annual_contract",
                    inventory_mode="isolated",
                )
                central = site_profile.load_site_profile("central")

                # Shared profiles reuse the single company inventory DB
                self.assertEqual(site_profile.resolve_inventory_path(central, shared_db), shared_db)
                self.assertEqual(site_profile.resolve_inventory_path(smb, shared_db), shared_db)

                # Isolated profile gets its own inventory file inside profiles/<id>/data
                iso_db = site_profile.resolve_inventory_path(iso, shared_db)
                self.assertNotEqual(iso_db, shared_db)
                self.assertTrue(iso_db.endswith(os.path.join("branch_isolated", "data", "Material_Inventory.xlsx")))

                # Operational configs & history are isolated per non-central profile
                self.assertEqual(site_profile.resolve_config_path(central, default_cfg), default_cfg)
                smb_cfg = site_profile.resolve_config_path(smb, default_cfg)
                self.assertNotEqual(smb_cfg, default_cfg)
                self.assertIn("client_alpha", smb_cfg)

                smb_bill = site_profile.resolve_billing_config_path(smb, default_billing)
                self.assertNotEqual(smb_bill, default_billing)
                self.assertIn("client_alpha", smb_bill)

                # data_sync.get_history_path() follows active non-central profile
                site_profile.set_active_profile_id("client_alpha")
                hist_path = data_sync.get_history_path()
                photos_dir = data_sync.get_process_photos_dir()
                self.assertTrue(hist_path.endswith(os.path.join("client_alpha", "data", "daily_work_history.json")))
                self.assertTrue(photos_dir.endswith(os.path.join("client_alpha", "data", "process_photos")))

    def test_smb_statement_exporter_calculates_totals_and_writes_xlsx(self):
        df = pd.DataFrame(
            [
                {
                    "Date": "2026-11-05",
                    "Site": "A공장 배관",
                    "업체명": "알파기계",
                    "검사방법": "RT",
                    "검사품명": "SUS 100A",
                    "검사량": 10,
                    "단가": 25000,
                    "출장비": 50000,
                    "일식": 0,
                    "검사비": 0,
                    "Note": "1차 촬영",
                },
                {
                    "Date": "2026-11-12",
                    "Site": "A공장 배관",
                    "업체명": "알파기계",
                    "검사방법": "MT",
                    "검사품명": "탱크 용접부",
                    "검사량": 1,
                    "단가": 0,
                    "출장비": 0,
                    "일식": 400000,
                    "검사비": 0,
                    "Note": "일식 정산",
                },
                {
                    "Date": "2026-10-20",
                    "Site": "B공장",
                    "업체명": "베타산업",
                    "검사방법": "PT",
                    "검사품명": "노즐",
                    "검사량": 5,
                    "단가": 10000,
                    "출장비": 0,
                    "일식": 0,
                    "검사비": 0,
                    "Note": "전월 기록",
                },
            ]
        )

        # Row 1: 10*25,000 + 50,000 = 300,000 | Row 2: 400,000 -> Total 700,000, VAT 70,000, Grand 770,000
        summary = build_statement_rows(df, year=2026, month=11, site="A공장 배관")
        self.assertEqual(len(summary["rows"]), 2)
        self.assertEqual(summary["supply_total"], 700000)
        self.assertEqual(summary["vat_total"], 70000)
        self.assertEqual(summary["grand_total"], 770000)

        with tempfile.TemporaryDirectory() as tmp:
            out_path = os.path.join(tmp, "statement_2026_11.xlsx")
            exported = export_smb_statement_excel(
                df,
                out_path,
                year=2026,
                month=11,
                site="A공장 배관",
                company="알파기계",
                profile_name="중소업체 건별/일식",
            )
            self.assertEqual(exported["grand_total"], 770000)
            self.assertTrue(os.path.isfile(out_path))

            wb = openpyxl.load_workbook(out_path)
            ws = wb["거래명세서"]
            self.assertEqual(ws["B5"].value, 700000)
            self.assertEqual(ws["E5"].value, 70000)
            self.assertEqual(ws["H5"].value, 770000)
            self.assertEqual(ws["A8"].value, "2026-11-05")
            self.assertEqual(ws["J8"].value, 300000)
            self.assertEqual(ws["J9"].value, 400000)
            self.assertEqual(ws["J10"].value, 700000)


if __name__ == "__main__":
    unittest.main()

