"""Tests for NDT-Pro multi-site app stages 1-4:
1. Dedicated brand & launcher title formatting
2. Client & Contract master atomic persistence
3. Flexible contract billing calculator & 2-sheet Excel export
4. Per-client report/photo template library with default fallback
5. ClientContractTab UI instantiation & refresh
"""
import os
from pathlib import Path
import sys
import tempfile
import tkinter as tk
import unittest
from unittest.mock import patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from site_apps.central.src import site_profile
from site_apps.central.src.multi_app_launcher import APP_BRAND_NAME, format_window_title
from site_apps.central.src.services.client_contract_store import (
    delete_client,
    delete_contract,
    get_client_by_id,
    get_contract_by_id,
    load_client_contract_store,
    upsert_client,
    upsert_contract,
)
from site_apps.central.src.services.flexible_billing_service import (
    calculate_flexible_billing,
    export_flexible_billing_excel,
)
from site_apps.central.src.services.template_library import (
    DEFAULT_CLIENT_KEY,
    list_templates,
    register_template,
    remove_template,
    resolve_template,
)
from site_apps.central.src.views.client_contract_tab import ClientContractTab


class MultiNDTAppFeaturesTests(unittest.TestCase):
    def test_brand_window_title_and_clients_tab_enablement(self):
        central = site_profile.load_site_profile("central")
        smb = site_profile.load_site_profile("smb_general")
        annual = site_profile.load_site_profile("annual_contract")

        self.assertEqual(format_window_title(central), "자재 및 소모품 관리 시스템 (Material Manager)")
        self.assertIn(APP_BRAND_NAME, format_window_title(smb))
        self.assertIn("공용재고", format_window_title(smb))
        self.assertTrue(smb.is_tab_enabled("clients_contracts"))
        self.assertTrue(annual.is_tab_enabled("clients_contracts"))
        self.assertFalse(central.is_tab_enabled("clients_contracts"))

    def test_client_contract_crud_and_cascade_delete(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(site_profile, "get_profiles_root", return_value=tmp):
                prof = site_profile.create_site_profile(
                    "smb_test", "테스트 중소업체", site_type="smb_general", inventory_mode="shared"
                )
                client = upsert_client(
                    {
                        "name": "한빛기계(주)",
                        "biz_no": "123-45-67890",
                        "ceo": "김대표",
                        "contact_name": "이과장",
                        "contact_phone": "010-1111-2222",
                        "payment_terms": "월말 마감 익월 10일",
                    },
                    profile=prof,
                )
                self.assertEqual(get_client_by_id(client["client_id"], profile=prof)["name"], "한빛기계(주)")

                contract = upsert_contract(
                    {
                        "client_id": client["client_id"],
                        "contract_name": "2026 연간 비파괴검사 단가계약",
                        "contract_type": "annual_unit",
                        "start_date": "2026-01-01",
                        "end_date": "2026-12-31",
                        "budget_amount": 50000000,
                        "default_trip_fee": 50000,
                        "rate_items": [
                            {"code": "RT-100", "name": "배관 RT 100A", "method": "RT", "spec": "100A", "unit": "매", "unit_price": 25000, "contract_qty": 200},
                            {"code": "MT-GEN", "name": "자분탐상", "method": "MT", "spec": "일반", "unit": "m", "unit_price": 30000, "contract_qty": 100},
                        ],
                    },
                    profile=prof,
                )
                loaded = get_contract_by_id(contract["contract_id"], profile=prof)
                self.assertEqual(len(loaded["rate_items"]), 2)
                self.assertEqual(loaded["rate_items"][0]["unit_price"], 25000)

                # Deleting contract works
                self.assertTrue(delete_contract(contract["contract_id"], profile=prof))
                self.assertIsNone(get_contract_by_id(contract["contract_id"], profile=prof))

                # Re-add contract and verify cascade delete with client
                upsert_contract(
                    {"client_id": client["client_id"], "contract_name": "건별 계약", "contract_type": "per_job"},
                    profile=prof,
                )
                self.assertTrue(delete_client(client["client_id"], profile=prof))
                store = load_client_contract_store(prof)
                self.assertEqual(store["clients"], [])
                self.assertEqual(store["contracts"], [])

    def test_flexible_billing_calculation_and_excel_export(self):
        client = {
            "name": "성진플랜트",
            "biz_no": "222-33-44444",
            "ceo": "박대표",
            "contact_name": "최대리",
            "contact_phone": "010-9999-8888",
            "payment_terms": "익월 15일 현금",
        }
        contract = {
            "contract_name": "2026 플랜트 연간계약",
            "start_date": "2026-01-01",
            "end_date": "2026-12-31",
            "budget_amount": 30000000,
            "rate_items": [
                {"code": "RT-01", "name": "RT 촬영", "method": "RT", "spec": "100A", "unit": "매", "unit_price": 20000, "contract_qty": 100},
                {"code": "PT-01", "name": "PT 검사", "method": "PT", "spec": "일반", "unit": "m", "unit_price": 15000, "contract_qty": 50},
            ],
        }
        usage = [
            {"code": "RT-01", "qty": 30, "trip_fee": 50000},     # 600,000 + 50,000 = 650,000
            {"code": "PT-01", "qty": 10},                         # 150,000
            {"name": "긴급 야간 출동 일식", "qty": 0, "lump_sum": 200000},  # 200,000
        ]
        res = calculate_flexible_billing(contract, usage, prev_cum_by_code={"RT-01": 20})
        self.assertEqual(res["supply_total"], 1000000)
        self.assertEqual(res["vat_total"], 100000)
        self.assertEqual(res["grand_total"], 1100000)

        rt_line = res["lines"][0]
        self.assertEqual(rt_line["prev_qty"], 20.0)
        self.assertEqual(rt_line["curr_qty"], 30.0)
        self.assertEqual(rt_line["total_qty"], 50.0)
        self.assertEqual(rt_line["remain_qty"], 50.0)
        self.assertEqual(rt_line["progress_pct"], 50.0)

        with tempfile.TemporaryDirectory() as tmp:
            out_path = os.path.join(tmp, "flexible_billing.xlsx")
            export_flexible_billing_excel(contract, client, res, out_path, period_label="2026-11", round_no=3)
            self.assertTrue(os.path.isfile(out_path))

            wb = openpyxl.load_workbook(out_path)
            self.assertEqual(wb.sheetnames, ["청구서(요약)", "기성·거래내역"])
            ws_sum = wb["청구서(요약)"]
            self.assertEqual(ws_sum["B3"].value, "성진플랜트")
            self.assertEqual(ws_sum["D10"].value, 1000000)
            self.assertEqual(ws_sum["D11"].value, 100000)
            self.assertEqual(ws_sum["D12"].value, 1100000)

            ws_det = wb["기성·거래내역"]
            self.assertEqual(ws_det["A2"].value, "RT-01")
            self.assertEqual(ws_det["L2"].value, 50.0)
            self.assertEqual(ws_det["P2"].value, 650000)
            self.assertEqual(ws_det["P5"].value, 1000000)

    def test_template_library_register_copy_resolve_fallback_and_remove(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(site_profile, "get_profiles_root", return_value=tmp):
                prof = site_profile.create_site_profile("tpl_test", "템플릿 테스트", site_type="smb_general")
                src_default = os.path.join(tmp, "std_rt.xlsx")
                src_client = os.path.join(tmp, "alpha_rt.xlsx")
                openpyxl.Workbook().save(src_default)
                openpyxl.Workbook().save(src_client)

                register_template(src_default, "RT", client_id=DEFAULT_CLIENT_KEY, description="기본 RT", profile=prof)
                register_template(src_client, "RT", client_id="CLI-ALPHA", description="알파 전용 RT", profile=prof)

                # Client-specific resolves first
                resolved_alpha = resolve_template("RT", client_id="CLI-ALPHA", profile=prof)
                self.assertIsNotNone(resolved_alpha)
                self.assertEqual(resolved_alpha["client_id"], "CLI-ALPHA")
                self.assertTrue(os.path.isfile(resolved_alpha["path"]))

                # Unknown client falls back to default
                resolved_beta = resolve_template("RT", client_id="CLI-BETA", profile=prof)
                self.assertIsNotNone(resolved_beta)
                self.assertEqual(resolved_beta["client_id"], DEFAULT_CLIENT_KEY)

                self.assertEqual(len(list_templates(profile=prof)), 2)
                self.assertTrue(remove_template("RT", client_id="CLI-ALPHA", profile=prof))
                self.assertEqual(resolve_template("RT", client_id="CLI-ALPHA", profile=prof)["client_id"], DEFAULT_CLIENT_KEY)

    def test_client_contract_tab_ui_renders_clients_and_contracts(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(site_profile, "get_profiles_root", return_value=tmp):
                prof = site_profile.create_site_profile("ui_test", "UI 테스트", site_type="smb_general")
                cli = upsert_client({"name": "UI거래처", "biz_no": "111-22-33333"}, profile=prof)
                upsert_contract(
                    {"client_id": cli["client_id"], "contract_name": "UI계약", "budget_amount": 1000000},
                    profile=prof,
                )
                root = tk.Tk()
                root.withdraw()
                try:
                    class DummyApp:
                        site_profile = prof
                        companies = []
                    tab = ClientContractTab(root, main_app=DummyApp())
                    self.assertEqual(len(tab.client_tree.get_children()), 1)
                    self.assertEqual(len(tab.contract_tree.get_children()), 1)
                finally:
                    root.destroy()


if __name__ == "__main__":
    unittest.main()

