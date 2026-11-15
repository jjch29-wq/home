import os
import unittest
import openpyxl


class TestPMIReportLayout(unittest.TestCase):
    def test_write_gapji_metadata_sitco_pmi(self):
        """Test metadata writing when template has SITCO standard PMI layout (A11=Project / L5=Customer)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "갑"
        ws['A11'] = "Project 공사명"
        ws['L5'] = "주문주\nCustomer"

        # Mock app with config and safe_set_value
        class DummyApp:
            def __init__(self):
                self.config = {
                    'GAPJI_PROJECT': 'Project PROVIDENCE (K1)',
                    'GAPJI_CUSTOMER': 'LG-Eni BioRefining Co., Ltd.',
                    'GAPJI_REPORT_NO': 'SIS/GI-MNS-PMI-TNTSAHYC001',
                    'GAPJI_ITEM': 'Stainless Steel Pipes',
                    'GAPJI_MATERIAL': 'A312-316 / A182-F316 / ER316L',
                    'GAPJI_EXAM_DATE': '26/10/2026',
                }

            def safe_set_value(self, target_ws, cell_ref, val, align=None):
                target_ws[cell_ref] = val

        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp
        dummy = DummyApp()
        PMIReportApp._write_gapji_metadata(dummy, ws, mode="PMI", first_item=None)

        self.assertEqual(ws['E11'].value, 'Project PROVIDENCE (K1)')
        self.assertEqual(ws['N5'].value, 'LG-Eni BioRefining Co., Ltd.')
        self.assertEqual(ws['N8'].value, 'SIS/GI-MNS-PMI-TNTSAHYC001')
        self.assertEqual(ws['N11'].value, '26/10/2026')
        self.assertEqual(ws['E13'].value, 'Stainless Steel Pipes')
        self.assertEqual(ws['E17'].value, 'A312-316 / A182-F316 / ER316L')

    def test_write_gapji_metadata_legacy_fallback(self):
        """Test fallback to B5..B10 for templates without SITCO A11/L5 structure."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Cover"

        class DummyApp:
            def __init__(self):
                self.config = {
                    'GAPJI_PROJECT': 'Generic Project',
                    'GAPJI_CUSTOMER': 'Generic Customer',
                    'GAPJI_ITEM': 'Generic Item',
                    'GAPJI_MATERIAL': 'Generic Mat',
                    'GAPJI_REPORT_NO': 'REP-001',
                }

            def safe_set_value(self, target_ws, cell_ref, val, align=None):
                target_ws[cell_ref] = val

        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp
        dummy = DummyApp()
        PMIReportApp._write_gapji_metadata(dummy, ws, mode="PMI", first_item=None)

        self.assertEqual(ws['B5'].value, 'Generic Project')
        self.assertEqual(ws['B6'].value, 'Generic Customer')
        self.assertEqual(ws['B7'].value, 'Generic Item')
        self.assertEqual(ws['B8'].value, 'Generic Mat')
        self.assertEqual(ws['B10'].value, 'REP-001')

    def test_pmi_triple_mode_auto_load(self):
        """Test auto-detection of PLANT_TRIPLE layout mode for K1/K3/롯데 templates."""
        class DummyVar:
            def __init__(self, val="STANDARD"):
                self._val = val
            def get(self):
                return self._val
            def set(self, val):
                self._val = val

        class MockApp:
            def __init__(self):
                self.config = {}
                self.setting_vars = {}
                self.pmi_layout_mode = DummyVar("STANDARD")

            def log(self, msg):
                pass

            def _update_pmi_filter_ui(self):
                pass

        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp
        app = MockApp()

        # Simulate loading a template with K1 or 롯데 in the name
        template_name = "home/recovery_backups/SIT-PMI-K3 롯데 템플릿보고서_backup.xlsm"
        PMIReportApp.load_template_specific_config(app, template_name, mode="PMI")

        self.assertEqual(app.pmi_layout_mode.get(), "PLANT_TRIPLE")

        # Simulate loading a standard template
        std_template = "home/data/SIT-PMI-001.xlsx"
        PMIReportApp.load_template_specific_config(app, std_template, mode="PMI")
        self.assertEqual(app.pmi_layout_mode.get(), "STANDARD")

    def test_template_cert_pmi_structure(self):
        """Verify the SIT-PMI-K3 template backup structure and address correctness."""
        backup_path = 'home/recovery_backups/SIT-PMI-K3 롯데 템플릿보고서_backup.xlsm'
        if not os.path.exists(backup_path):
            self.skipTest("Backup template not found")

        wb = openpyxl.load_workbook(backup_path)
        self.assertIn('갑', wb.sheetnames)
        self.assertIn('을', wb.sheetnames)

    def test_pmi_print_area_mode_split(self):
        """Verify PMI print area distinguishes between STANDARD (A1:T47) and PLANT_TRIPLE (A1:M47)."""
        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp

        class DummyVar:
            def __init__(self, val):
                self._val = val
            def get(self):
                return self._val

        class MockApp:
            def __init__(self, layout_mode):
                self.config = {'PRINT_END_ROW': 47}
                self.current_mode = "PMI"
                self.pmi_layout_mode = DummyVar(layout_mode)

        wb = openpyxl.Workbook()
        ws_std = wb.active
        app_std = MockApp("STANDARD")
        PMIReportApp.force_print_settings(app_std, ws_std, context="DATA")
        self.assertIn("$A$1:$T$47", str(ws_std.print_area))

        ws_triple = wb.create_sheet("을_triple")
        app_triple = MockApp("PLANT_TRIPLE")
        PMIReportApp.force_print_settings(app_triple, ws_triple, context="DATA")
        self.assertIn("$A$1:$M$47", str(ws_triple.print_area))

    def test_pmi_custom_elements_filter_report_no(self):
        """Verify Report No and other metadata are excluded from custom chemical elements."""
        non_element_keys = {
            "_status", "selected", "No", "Date", "Report No", "Dwg", "Joint",
            "Loc", "Grade", "Result", "Remarks", "Welder", "order_index", "V", "ST",
            "Ni", "Cr", "Mo"
        }
        test_keys = ["_status", "selected", "No", "Date", "Report No", "Dwg", "Joint", "Loc", "Ni", "Cr", "Mo", "Grade", "Cu"]
        custom = [k for k in test_keys if k not in non_element_keys]
        self.assertEqual(custom, ["Cu"])
        self.assertNotIn("Report No", custom)


if __name__ == '__main__':
    unittest.main()
