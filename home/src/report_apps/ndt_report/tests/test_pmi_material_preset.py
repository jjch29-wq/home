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

    def setUp(self):
        if self.app:
            self.app.clear_all_pmi_specs()
            # 슬롯 4~6이 남아있다면 3개까지 축소 정리
            while len(self.app.pmi_chem_specs) > 3:
                self.app.remove_pmi_spec_slot(len(self.app.pmi_chem_specs))

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

        # Fitting items (ELBOW, TEE, REDUCER) and A403-WP304L
        spec_wp304l = self.app._lookup_pmi_material_spec("WP304L")
        self.assertIsNotNone(spec_wp304l)
        self.assertEqual(spec_wp304l["name"], "A403-WP304L")
        self.assertEqual(spec_wp304l["limits"]["Cr"], (18.0, 20.0))
        self.assertEqual(spec_wp304l["limits"]["Ni"], (9.0, 13.0))

        spec_elbow = self.app._lookup_pmi_material_spec("ELBOW")
        self.assertIsNotNone(spec_elbow)
        self.assertEqual(spec_elbow["name"], "A403-WP316L")

        spec_tee = self.app._lookup_pmi_material_spec("티")
        self.assertIsNotNone(spec_tee)
        self.assertEqual(spec_tee["name"], "A403-WP316L")

        spec_reducer = self.app._lookup_pmi_material_spec("레듀샤")
        self.assertIsNotNone(spec_reducer)
        self.assertEqual(spec_reducer["name"], "A403-WP316L")

        spec_er309 = self.app._lookup_pmi_material_spec("ER309L")
        self.assertIsNotNone(spec_er309)
        self.assertEqual(spec_er309["name"], "ER309L (이재질)")

    def test_apply_pmi_material_preset(self):
        """Verify applying material preset updates chemical elements, specs, and filters."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # 1. Clear all specs first
        self.app.clear_all_pmi_specs()
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[2].get(), "")

        # 2. Applying base metal (A312-TP316L) fills ONLY criteria 1; criteria 2 & 3 stay blank!
        success = self.app.apply_pmi_material_preset("A312-TP316L")
        self.assertTrue(success)

        # Check elements
        elem_vals = [var.get() for var in self.app.pmi_chem_elements]
        self.assertIn("Ni", elem_vals)
        self.assertIn("Cr", elem_vals)
        self.assertIn("Mo", elem_vals)

        # Check criteria 1 is populated and criteria 2 & 3 stay blank
        self.assertTrue("A312" in self.app.pmi_chem_specs[0].get())
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[2].get(), "")

        # 3. Applying welding consumable (ER316L) populates criteria 2 (Weld)
        self.app.apply_pmi_material_preset("ER316L")
        self.assertTrue("ER316L" in self.app.pmi_chem_specs[1].get())
        self.assertEqual(self.app.pmi_chem_specs[2].get(), "")

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

    def test_show_pmi_element_selector_dialog_creates_elements_without_error(self):
        """Verify element selector dialog opens cleanly with element checkboxes and bottom buttons."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # Mock extracted data with detected elements
        self.app.extracted_data = [
            {"No": "1", "Cr": 18.5, "Ni": 11.2, "Mo": 2.1, "Ti": 0.05, "Mn": 1.2}
        ]

        # Call dialog - must not throw TclError or any exception
        try:
            self.app.show_pmi_element_selector_dialog()
        except Exception as e:
            self.fail(f"show_pmi_element_selector_dialog raised an exception: {e}")

    def test_item_and_material_bidirectional_sync(self):
        """Verify selecting item updates material and vice-versa, without mutating chem specs."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # Set baseline specs to test that changing item NEVER changes bottom specs
        self.app.clear_all_pmi_specs()
        self.app.pmi_chem_specs[0].set("Original Spec 1")

        # 1. Setting item to FLANGE updates empty material to A182-F316L and limits values to flange materials
        self.app.gapji_material.set("")
        self.app.gapji_item.set("FLANGE")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "A182-F316L")
        self.assertEqual(list(self.app.pmi_material_combobox['values']), ["A182-F316L", "A182-F304L"])
        # Bottom spec 1 MUST remain unchanged!
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "Original Spec 1")

        # 2. Setting item to VALVE updates empty material to A351-CF8M and limits values to valve materials
        self.app.gapji_material.set("")
        self.app.gapji_item.set("VALVE")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "A351-CF8M")
        self.assertEqual(list(self.app.pmi_material_combobox['values']), ["A351-CF8M", "A351-CF3M", "A351-CF8"])
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "Original Spec 1")

        # 3. Setting item to FITTING updates empty material to A403-WP316L
        self.app.gapji_material.set("")
        self.app.gapji_item.set("FITTING")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "A403-WP316L")
        self.assertEqual(list(self.app.pmi_material_combobox['values']), ["A403-WP316L", "A403-WP304L"])

        # 3-1. Setting item to ELBOW updates empty material to A403-WP316L
        self.app.gapji_material.set("")
        self.app.gapji_item.set("ELBOW")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "A403-WP316L")
        self.assertEqual(list(self.app.pmi_material_combobox['values']), ["A403-WP316L", "A403-WP304L"])

        # 3-2. Setting item to TEE updates empty material to A403-WP316L
        self.app.gapji_material.set("")
        self.app.gapji_item.set("TEE")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "A403-WP316L")

        # 3-3. Setting item to REDUCER updates empty material to A403-WP316L
        self.app.gapji_material.set("")
        self.app.gapji_item.set("REDUCER")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "A403-WP316L")
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "Original Spec 1")

        # 3-4. Setting material to A403-WP304L retains existing fitting item (e.g. REDUCER)
        self.app.gapji_material.set("A403-WP304L")
        self.app._on_pmi_material_changed()
        self.assertEqual(self.app.gapji_item.get(), "REDUCER")

        # 3-5. Setting material to A403-WP316L when item is PIPE defaults to ELBOW
        self.app.gapji_item.set("PIPE")
        self.app.gapji_material.set("A403-WP316L")
        self.app._on_pmi_material_changed()
        self.assertEqual(self.app.gapji_item.get(), "ELBOW")

        # 4. Setting material to ASTM A992 updates item to H-BEAM
        self.app.gapji_material.set("ASTM A992")
        self.app._on_pmi_material_changed()
        self.assertEqual(self.app.gapji_item.get(), "H-BEAM")

        # 5. Setting material to A312-TP316L updates item to PIPE
        self.app.gapji_material.set("A312-TP316L")
        self.app._on_pmi_material_changed()
        self.assertEqual(self.app.gapji_item.get(), "PIPE")

        # 6. Setting item to WELD updates material to ER316L when material is empty
        self.app.gapji_material.set("")
        self.app.gapji_item.set("WELD")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.gapji_material.get(), "ER316L")

        # 7. Setting material to ER308L updates item to WELD
        self.app.gapji_material.set("ER308L")
        self.app._on_pmi_material_changed()
        self.assertEqual(self.app.gapji_item.get(), "WELD")

        # 8. Setting material to ER70S-6 updates item to WELD
        self.app.gapji_material.set("ER70S-6")
        self.app._on_pmi_material_changed()
        self.assertEqual(self.app.gapji_item.get(), "WELD")

    def test_welding_material_presets_limits(self):
        """Verify welding consumables (ER316L, ER308L, ER2209, ER70S-6) have correct limits and specs."""
        if not self.app:
            self.skipTest("GUI not initialized")

        spec_316 = self.app._lookup_pmi_material_spec("ER316L")
        self.assertIsNotNone(spec_316)
        self.assertEqual(spec_316["limits"]["Cr"], (18.0, 20.0))
        self.assertEqual(spec_316["limits"]["Ni"], (11.0, 14.0))

        spec_308 = self.app._lookup_pmi_material_spec("ER308L")
        self.assertIsNotNone(spec_308)
        self.assertEqual(spec_308["limits"]["Cr"], (19.5, 22.0))
        self.assertEqual(spec_308["limits"]["Ni"], (9.0, 11.0))

        spec_2209 = self.app._lookup_pmi_material_spec("ER2209")
        self.assertIsNotNone(spec_2209)
        self.assertEqual(spec_2209["limits"]["Cr"], (21.5, 23.5))

        spec_70s = self.app._lookup_pmi_material_spec("ER70S-6")
        self.assertIsNotNone(spec_70s)
        self.assertEqual(spec_70s["limits"]["Mn"], (0.80, 1.85))

    def test_spec_combobox_individual_selection(self):
        """Verify selecting each spec 1, 2, 3 combobox individually populates the respective chem spec text."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # 1. Spec 1 selection
        self.app.pmi_spec1_mat.set("A312-TP304")
        self.app._on_pmi_spec1_selected()
        self.assertIn("A312-TP304", self.app.pmi_chem_specs[0].get())
        self.assertIn("Cr:18.00 ~ 20.00%", self.app.pmi_chem_specs[0].get())

        # 2. Spec 2 selection (e.g. switch to ER309L dissimilar)
        self.app.pmi_spec2_mat.set("ER309L (이재질)")
        self.app._on_pmi_spec2_selected()
        self.assertIn("ER309L", self.app.pmi_chem_specs[1].get())
        self.assertIn("Cr:23.00 ~ 25.00%", self.app.pmi_chem_specs[1].get())

        # 3. Spec 3 selection (e.g. switch to A403-WP316L fitting)
        self.app.pmi_spec3_mat.set("A403-WP316L (피팅/엘보)")
        self.app._on_pmi_spec3_selected()
        self.assertIn("A403-WP316L", self.app.pmi_chem_specs[2].get())
        self.assertIn("Fitting", self.app.pmi_chem_specs[2].get())

        # 4. Spec 3 empty option
        self.app.pmi_spec3_mat.set("(공란 / 비움)")
        self.app._on_pmi_spec3_selected()
        self.assertEqual(self.app.pmi_chem_specs[2].get(), "")
        self.assertEqual(self.app.pmi_spec3_lbl.cget("text"), "기준 3:")

    def test_dynamic_spec_labels_and_apply_current(self):
        """Verify labels dynamically display selected item name (not raw material) and apply_current_to_spec works."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # 1. Clear all
        self.app.clear_all_pmi_specs()
        self.assertEqual(self.app.pmi_spec1_lbl.cget("text"), "기준 1:")
        self.assertEqual(self.app.pmi_spec2_lbl.cget("text"), "기준 2:")
        self.assertEqual(self.app.pmi_spec3_lbl.cget("text"), "기준 3:")

        # 2. Select Spec 1 (A312-TP304) -> label updates dynamically with item name "PIPE"
        self.app.gapji_item.set("PIPE")
        self.app.pmi_spec1_mat.set("A312-TP304")
        self.app._on_pmi_spec1_selected()
        self.assertEqual(self.app.pmi_spec1_lbl.cget("text"), "기준 1 (PIPE):")

        # 3. Apply current upper material (e.g. ELBOW / A403-WP316L) to Spec 2 slot -> label displays "ELBOW"
        self.app.gapji_item.set("ELBOW")
        self.app.gapji_material.set("A403-WP316L")
        self.app.apply_current_to_spec(2)
        self.assertIn("A403-WP316L", self.app.pmi_chem_specs[1].get())
        self.assertEqual(self.app.pmi_spec2_lbl.cget("text"), "기준 2 (ELBOW):")

        # 4. Select welding wire for Spec 2 -> label displays "용접봉"
        self.app.pmi_spec2_mat.set("ER308L (용접봉)")
        self.app._on_pmi_spec2_selected()
        self.assertEqual(self.app.pmi_spec2_lbl.cget("text"), "기준 2 (용접봉):")

        # 5. Select flange for Spec 3 -> label displays "FLANGE"
        self.app.pmi_spec3_mat.set("A182-F304L (플랜지)")
        self.app._on_pmi_spec3_selected()
        self.assertEqual(self.app.pmi_spec3_lbl.cget("text"), "기준 3 (FLANGE):")

        # 6. Clear slot 2 -> label reverts to "기준 2:"
        self.app.clear_pmi_spec(2)
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")
        self.assertEqual(self.app.pmi_spec2_lbl.cget("text"), "기준 2:")

    def test_clear_pmi_spec_buttons(self):
        """Verify clear_pmi_spec individual and all buttons clear the fields to blank."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # Set some values
        self.app.pmi_chem_specs[0].set("Spec 1")
        self.app.pmi_chem_specs[1].set("Spec 2")
        self.app.pmi_chem_specs[2].set("Spec 3")

        # Clear 1
        self.app.clear_pmi_spec(1)
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "Spec 2")

        # Clear 2
        self.app.clear_pmi_spec(2)
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")

        # Clear all
        self.app.pmi_chem_specs[0].set("Spec 1")
        self.app.pmi_chem_specs[1].set("Spec 2")
        self.app.clear_all_pmi_specs()
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[2].get(), "")

    def test_sequential_item_change_filling_next_slot(self):
        """Verify browsing items/materials does not mutate specs, clicking add_current_to_next_spec fills slots 1 -> 2 -> 3 sequentially, and apply_current_to_spec edits specific slot."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # 1. Clear all specs initially
        self.app.clear_all_pmi_specs()
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")
        self.assertEqual(self.app.pmi_chem_specs[2].get(), "")

        # 2. Browsing multiple items/materials does NOT mutate specs automatically
        self.app.gapji_item.set("PIPE")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "")

        self.app.gapji_item.set("ELBOW")
        self.app._on_pmi_item_changed()
        self.assertEqual(self.app.pmi_chem_specs[0].get(), "")

        # 3. Choose PIPE / A312-TP304 and click add_current_to_next_spec -> fills Slot 1
        self.app.gapji_item.set("PIPE")
        self.app.gapji_material.set("A312-TP304")
        self.app.add_current_to_next_spec()
        self.assertIn("A312-TP304", self.app.pmi_chem_specs[0].get())
        self.assertEqual(self.app.pmi_chem_specs[1].get(), "")
        self.assertEqual(self.app.pmi_spec1_lbl.cget("text"), "기준 1 (PIPE):")

        # 4. Choose ELBOW / A403-WP316L and click add_current_to_next_spec -> fills Slot 2 (Slot 1 preserved!)
        self.app.gapji_item.set("ELBOW")
        self.app.gapji_material.set("A403-WP316L")
        self.app.add_current_to_next_spec()
        # Slot 1 must remain intact with PIPE
        self.assertIn("A312-TP304", self.app.pmi_chem_specs[0].get())
        self.assertEqual(self.app.pmi_spec1_lbl.cget("text"), "기준 1 (PIPE):")
        # Slot 2 must be populated with ELBOW
        self.assertIn("A403-WP316L", self.app.pmi_chem_specs[1].get())
        self.assertEqual(self.app.pmi_spec2_lbl.cget("text"), "기준 2 (ELBOW):")

        # 5. Choose WELD / ER308L and click add_current_to_next_spec -> fills Slot 3
        self.app.gapji_item.set("WELD")
        self.app.gapji_material.set("ER308L")
        self.app.add_current_to_next_spec()
        self.assertIn("ER308L", self.app.pmi_chem_specs[2].get())
        self.assertEqual(self.app.pmi_spec3_lbl.cget("text"), "기준 3 (용접봉):")

        # 6. Re-editing: Choose FLANGE / A182-F304L and edit Slot 1 using apply_current_to_spec(1)
        self.app.gapji_item.set("FLANGE")
        self.app.gapji_material.set("A182-F304L")
        self.app.apply_current_to_spec(1)
        # Slot 1 is now overwritten with FLANGE, while Slot 2 and Slot 3 remain intact
        self.assertIn("A182-F304L", self.app.pmi_chem_specs[0].get())
        self.assertEqual(self.app.pmi_spec1_lbl.cget("text"), "기준 1 (FLANGE):")
        self.assertIn("A403-WP316L", self.app.pmi_chem_specs[1].get())
        self.assertIn("ER308L", self.app.pmi_chem_specs[2].get())

    def test_dynamic_slot_addition_and_removal(self):
        """Verify dynamic slot addition up to 6 and safe removal."""
        if not self.app:
            self.skipTest("GUI not initialized")

        # Initial slots = 3
        self.assertEqual(len(self.app.pmi_chem_specs), 3)

        # Add slot 4
        s4 = self.app.add_pmi_spec_slot()
        self.assertEqual(s4, 4)
        self.assertEqual(len(self.app.pmi_chem_specs), 4)

        # Add slot 5
        s5 = self.app.add_pmi_spec_slot()
        self.assertEqual(s5, 5)
        self.assertEqual(len(self.app.pmi_chem_specs), 5)

        # Add slot 6
        s6 = self.app.add_pmi_spec_slot()
        self.assertEqual(s6, 6)
        self.assertEqual(len(self.app.pmi_chem_specs), 6)

        # 7th addition must be blocked
        s7 = self.app.add_pmi_spec_slot(show_msg=False)
        self.assertIsNone(s7)
        self.assertEqual(len(self.app.pmi_chem_specs), 6)

        # Remove slot 6
        self.app.remove_pmi_spec_slot(6)
        self.assertEqual(len(self.app.pmi_chem_specs), 5)

        # Auto slot expansion via add_current_to_next_spec when all slots full
        # Fill slot 1~5
        for i in range(1, 6):
            self.app.pmi_chem_specs[i-1].set("DUMMY: Cr:18.00 ~ 20.00%")
        self.app.gapji_item.set("VALVE")
        self.app.gapji_material.set("A351-CF8M")
        self.app.add_current_to_next_spec(show_msg=False)
        # Should have automatically expanded to slot 6 and filled it!
        self.assertEqual(len(self.app.pmi_chem_specs), 6)
        self.assertIn("CF8M", self.app.pmi_chem_specs[5].get())

    def test_acceptance_criteria_grouping_and_matrix_layout(self):
        """Verify grouping of identical limits (TEE + REDUCER) and Option B matrix generation."""
        if not self.app:
            self.skipTest("GUI not initialized")

        wb = openpyxl.load_workbook("home/data/PMI-템플릿보고서.xlsx")
        ws = wb["PMI(갑)"]

        # Slot 1: PIPE / A312-TP304 (Cr 18~20, Ni 8~11)
        self.app.pmi_chem_specs[0].set("A312-TP304 (PIPE): Cr:18.00 ~ 20.00%, Ni:8.00 ~ 11.00%")
        self.app.pmi_spec1_mat.set("A312-TP304")

        # Slot 2: TEE / A403-WP304L (Cr 18~20, Ni 9~13)
        self.app.pmi_chem_specs[1].set("A403-WP304L (TEE): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%")
        self.app.pmi_spec2_mat.set("A403-WP304L (피팅/엘보)")

        # Slot 3: REDUCER / A403-WP304L (Same limits as Slot 2 -> should be grouped!)
        self.app.pmi_chem_specs[2].set("A403-WP304L (REDUCER): Cr:18.00 ~ 20.00%, Ni:9.00 ~ 13.00%")
        self.app.pmi_spec3_mat.set("A403-WP304L (피팅/엘보)")

        # 1. Test grouping logic
        grouped = self.app._collect_pmi_active_criteria()
        self.assertEqual(len(grouped), 2, "3 criteria with 2 identical limits must be grouped into 2 items")
        self.assertIn("PIPE", grouped[0]["item"])
        self.assertIn("TEE", grouped[1]["item"])
        self.assertIn("REDUCER", grouped[1]["item"])

        # 2. Test Option B matrix layout application
        self.app._write_gapji_metadata(ws, mode="PMI")

        # Title: 1. Acceptance Criteria (%) across A24:T26
        self.assertEqual(ws["A24"].value, "1. Acceptance Criteria (%)")

        # Row 27 Header: Element, Group 1, Group 2
        self.assertEqual(ws["A27"].value, "Element")
        self.assertIn("PIPE", ws["G27"].value)
        self.assertIn("TEE, REDUCER", ws["N27"].value)

        # Row 28 (Chromium)
        self.assertEqual(ws["A28"].value, "Chromium")
        self.assertEqual(ws["G28"].value, "18.00~20.00")
        self.assertEqual(ws["N28"].value, "18.00~20.00")

        # Row 29 (Nickel)
        self.assertEqual(ws["A29"].value, "Nickel")
        self.assertEqual(ws["G29"].value, "8.00~11.00")
        self.assertEqual(ws["N29"].value, "9.00~13.00")

        # Row 31 (Tolerance)
        self.assertEqual(ws["A31"].value, "* Tolerance ±10%")

        # Row 27 Header Height & Font Size Check (Prevent text clipping)
        self.assertEqual(ws.row_dimensions[27].height, 32.0, "Row 27 height must be 32.0 to prevent 2-line header text clipping")
        self.assertEqual(ws["G27"].font.size, 9, "Header font size should be 9pt when multiple criteria present")
        self.assertAlmostEqual(ws.row_dimensions[32].height, 3.77, places=2, msg="Row 32 should offset Row 27 height expansion to preserve total vertical height")


if __name__ == "__main__":
    unittest.main()

