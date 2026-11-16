import unittest
import tkinter as tk
import openpyxl
from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp


class TestPMIMaterialPresetAndCompliance(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.root = tk.Tk()
            cls.root.withdraw()
            cls.app = PMIReportApp(cls.root)
        except Exception as e:
            cls.app = None
            cls.skip_gui = True

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, 'root') and cls.root:
            try:
                cls.root.destroy()
            except:
                pass

    def test_lookup_pmi_material_spec(self):
        """Verify alias and spec lookup for various material strings."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # Exact match
        spec_316 = self.app._lookup_pmi_material_spec("A312-TP316")
        self.assertIsNotNone(spec_316)
        self.assertEqual(spec_316["name"], "A312-TP316")
        self.assertIn("Mo", spec_316["limits"])

        # Alias match
        spec_tp316 = self.app._lookup_pmi_material_spec("TP316")
        self.assertIsNotNone(spec_tp316)
        self.assertEqual(spec_tp316["name"], "A312-TP316")

        spec_ss304 = self.app._lookup_pmi_material_spec("SS304")
        self.assertIsNotNone(spec_ss304)
        self.assertEqual(spec_ss304["name"], "A312-TP304")

        spec_duplex = self.app._lookup_pmi_material_spec("DUPLEX")
        self.assertIsNotNone(spec_duplex)
        self.assertEqual(spec_duplex["name"], "DUPLEX (S31803)")

        # Structural steel match (COLUMN, GIRDER, BEAM)
        spec_a992 = self.app._lookup_pmi_material_spec("ASTM A992")
        self.assertIsNotNone(spec_a992)
        self.assertEqual(spec_a992["name"], "ASTM A992")
        self.assertIn("Mn", spec_a992["limits"])
        self.assertIn("V", spec_a992["limits"])

        spec_beam = self.app._lookup_pmi_material_spec("BEAM")
        self.assertIsNotNone(spec_beam)
        self.assertEqual(spec_beam["name"], "ASTM A992")

        spec_sm355 = self.app._lookup_pmi_material_spec("SM355")
        self.assertIsNotNone(spec_sm355)
        self.assertEqual(spec_sm355["name"], "SM355 (SM490)")

        # Project PROVIDENCE field specs (A351-CF8M, A403-WP316L, ER309L)
        spec_cf8m = self.app._lookup_pmi_material_spec("A351-CF8M")
        self.assertIsNotNone(spec_cf8m)
        self.assertEqual(spec_cf8m["name"], "A351-CF8M")
        self.assertIn("Mo", spec_cf8m["limits"])

        spec_valve = self.app._lookup_pmi_material_spec("VALVE")
        self.assertIsNotNone(spec_valve)
        self.assertEqual(spec_valve["name"], "A351-CF8M")

        spec_wp316l = self.app._lookup_pmi_material_spec("WP316L")
        self.assertIsNotNone(spec_wp316l)
        self.assertEqual(spec_wp316l["name"], "A403-WP316L")

        spec_er309 = self.app._lookup_pmi_material_spec("ER309L")
        self.assertIsNotNone(spec_er309)
        self.assertEqual(spec_er309["name"], "ER309L (이재질)")

    def test_apply_pmi_material_preset(self):
        """Verify applying material preset updates chemical elements, specs, and filters."""
        if not self.app:
            self.skipTest("GUI not initialized")

        success = self.app.apply_pmi_material_preset("A312-TP316L")
        self.assertTrue(success)

        # Check elements
        elem_vals = [var.get() for var in self.app.pmi_chem_elements]
        self.assertIn("Ni", elem_vals)
        self.assertIn("Cr", elem_vals)
        self.assertIn("Mo", elem_vals)

        # Check specs
        spec_vals = [var.get() for var in self.app.pmi_chem_specs]
        self.assertTrue(any("ER316L" in s for s in spec_vals))

        # Check element filters
        filter_keys = [f["key"].get() for f in self.app.element_filters]
        self.assertIn("Ni", filter_keys)
        self.assertIn("Cr", filter_keys)
        self.assertIn("Mo", filter_keys)

    def test_evaluate_pmi_data_compliance(self):
        """Verify compliance evaluation assigns ALL MATCH for valid items and MISMATCH for deficient items."""
        if not self.app:
            self.skipTest("GUI not initialized")

        self.app.apply_pmi_material_preset("A312-TP316")

        # Mock extracted data: Item 1 = Pass, Item 2 = Mo deficient (0.1%), Item 3 = Cr deficient (12.0%)
        self.app.extracted_data = [
            {"No": "1", "Cr": 17.20, "Ni": 11.50, "Mo": 2.25, "Grade": "", "Result": ""},
            {"No": "2", "Cr": 17.50, "Ni": 11.00, "Mo": 0.12, "Grade": "", "Result": ""},
            {"No": "3", "Cr": 12.00, "Ni": 11.00, "Mo": 2.20, "Grade": "", "Result": ""},
        ]

        self.app.evaluate_pmi_data_compliance(show_msg=False)

        # Item 1 should pass
        self.assertEqual(self.app.extracted_data[0]["Result"], "ALL MATCH")
        self.assertEqual(self.app.extracted_data[0]["Grade"], "ALL MATCH")
        self.assertFalse(self.app.extracted_data[0]["is_deficient"])

        # Item 2 should fail due to Mo
        self.assertEqual(self.app.extracted_data[1]["Result"], "MISMATCH")
        self.assertIn("Mo미달", self.app.extracted_data[1]["Grade"])
        self.assertTrue(self.app.extracted_data[1]["is_deficient"])

        # Item 3 should fail due to Cr
        self.assertEqual(self.app.extracted_data[2]["Result"], "MISMATCH")
        self.assertIn("Cr미달", self.app.extracted_data[2]["Grade"])
        self.assertTrue(self.app.extracted_data[2]["is_deficient"])

    def test_write_gapji_metadata_updates_acceptance_criteria(self):
        """Verify cover sheet Acceptance Criteria table (A24:G30) is dynamically updated with chosen material."""
        if not self.app:
            self.skipTest("GUI not initialized")

        wb = openpyxl.load_workbook("home/data/PMI-템플릿보고서.xlsx")
        ws = wb["PMI(갑)"]

        self.app.gapji_material.set("A312-TP316L")
        self.app.gapji_item.set("PIPE")

        self.app._write_gapji_metadata(ws, mode="PMI", first_item=None)

        # Check A24 title
        self.assertIn("ASTM A312 TP316L", str(ws["A24"].value))
        self.assertIn("PIPE", str(ws["A24"].value))

        # Check elements and composition ranges
        self.assertEqual(ws["A28"].value, "Chromium")
        self.assertEqual(ws["G28"].value, "16.00~18.00")
        self.assertEqual(ws["A29"].value, "Nickel")
        self.assertEqual(ws["G29"].value, "10.00~14.00")
        self.assertEqual(ws["A30"].value, "Molybdenum")
        self.assertEqual(ws["G30"].value, "2.00~3.00")

        # Test Structural steel A992 updates cover sheet with Manganese, Silicon, Vanadium
        self.app.gapji_material.set("ASTM A992")
        self.app.gapji_item.set("H-BEAM")
        self.app._write_gapji_metadata(ws, mode="PMI", first_item=None)

        self.assertIn("ASTM A992", str(ws["A24"].value))
        self.assertIn("H-BEAM", str(ws["A24"].value))
        self.assertEqual(ws["A28"].value, "Manganese")
        self.assertEqual(ws["G28"].value, "0.50~1.60")
        self.assertEqual(ws["A29"].value, "Silicon")
        self.assertEqual(ws["G29"].value, "0.15~0.40")
        self.assertEqual(ws["A30"].value, "Vanadium")
        self.assertEqual(ws["G30"].value, "Max 0.15")


if __name__ == "__main__":
    unittest.main()

