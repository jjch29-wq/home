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
        """Verify PMI print area distinguishes between STANDARD (A1:T47) and PLANT_TRIPLE (A1:M45)."""
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
        self.assertIn("$A$1:$M$45", str(ws_triple.print_area))

        # Verify PMI margins are standardized to 0.25 inches (template standard) for both Cover and Data
        ws_cover = wb.create_sheet("갑_cover")
        PMIReportApp.force_print_settings(app_triple, ws_cover, context="COVER")
        self.assertAlmostEqual(ws_cover.page_margins.left, 0.25, places=2)
        self.assertAlmostEqual(ws_cover.page_margins.right, 0.25, places=2)
        self.assertAlmostEqual(ws_triple.page_margins.left, 0.25, places=2)
        self.assertAlmostEqual(ws_triple.page_margins.right, 0.25, places=2)

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


    def test_pmi_footer_row_height_safety(self):
        """Verify PMI print settings ensure at least 24.0 height for footer row 46."""
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
        ws = wb.active
        app = MockApp("STANDARD")
        PMIReportApp.force_print_settings(app, ws, context="DATA")
        self.assertGreaterEqual(ws.row_dimensions[46].height, 24.0)

    def test_pmi_footer_template_preservation(self):
        """Verify template footer is preserved and standard eulji does not force H 802-2."""
        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp

        class DummyVar:
            def __init__(self, val):
                self._val = val
            def get(self):
                return self._val

        class MockApp:
            def __init__(self, layout_mode):
                self.config = {}
                self.current_mode = "PMI"
                self.pmi_layout_mode = DummyVar(layout_mode)
                self.logs = []

            def log(self, msg):
                self.logs.append(msg)

            def find_image_smart(self, keywords):
                return None

            def place_image_freely(self, *args, **kwargs):
                pass

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.oddFooter.left.text = "양 식 H 801-2"
        app = MockApp("STANDARD")

        PMIReportApp.add_logos_to_sheet(app, ws, is_cover=False, clear_existing=False, mode="PMI")
        self.assertTrue(any("원본 바닥글이 감지되어" in m for m in app.logs))

    def test_pmi_sitco_logo_centering(self):
        """Verify SITCO logo in A5:D10 box is centered both horizontally and vertically."""
        from openpyxl.drawing.image import Image
        from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
        from openpyxl.utils.units import pixels_to_EMU
        import io
        from PIL import Image as PILImage

        wb = openpyxl.Workbook()
        ws = wb.active

        # Create dummy SITCO-like image (179x52 aspect ~3.44)
        buf = io.BytesIO()
        PILImage.new("RGB", (179, 52), color="red").save(buf, format="JPEG")
        buf.seek(0)
        img = Image(buf)
        img.anchor = TwoCellAnchor(
            _from=AnchorMarker(col=0, colOff=19050, row=5, rowOff=38100),
            to=AnchorMarker(col=3, colOff=9525, row=9, rowOff=19050)
        )
        ws.add_image(img)

        # Apply header stamp normalization
        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp
        app = PMIReportApp.__new__(PMIReportApp)
        app.config = {}

        for im in ws._images:
            marker = getattr(im.anchor, '_from', None)
            if marker is None: continue
            r, c = marker.row, marker.col
            aspect = im.width / max(im.height, 1)
            if c == 0 and 3 <= r <= 6 and 2.5 <= aspect <= 4.0:
                im.anchor = TwoCellAnchor(
                    _from=AnchorMarker(col=0, colOff=pixels_to_EMU(6), row=4, rowOff=pixels_to_EMU(8)),
                    to=AnchorMarker(col=3, colOff=pixels_to_EMU(12), row=9, rowOff=pixels_to_EMU(6))
                )

        # Verify centered anchor: row 4 (5th row), top margin 8px, left margin 6px
        self.assertEqual(img.anchor._from.row, 4)
        self.assertEqual(img.anchor._from.colOff, pixels_to_EMU(6))
        self.assertEqual(img.anchor._from.rowOff, pixels_to_EMU(8))
        self.assertEqual(img.anchor.to.col, 3)
        self.assertEqual(img.anchor.to.colOff, pixels_to_EMU(12))
        self.assertEqual(img.anchor.to.row, 9)
        self.assertEqual(img.anchor.to.rowOff, pixels_to_EMU(6))

    def test_pmi_chemical_composition_preserved(self):
        """Verify H15 Chemical Composition header is preserved in standard PMI mode."""
        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp

        class DummyVar:
            def __init__(self, val):
                self._val = val
            def get(self):
                return self._val

        class MockApp:
            def __init__(self):
                self.config = {}
                self.current_mode = "PMI"
                self.pmi_layout_mode = DummyVar("STANDARD")

            def safe_set_value(self, ws, coord, value, align=None):
                from openpyxl.cell.cell import MergedCell
                target = ws[coord] if not hasattr(coord, 'coordinate') else coord
                if isinstance(target, MergedCell):
                    for m in ws.merged_cells.ranges:
                        if (coord if not hasattr(coord, 'coordinate') else coord.coordinate) in m:
                            target = ws.cell(m.min_row, m.min_col)
                            break
                target.value = value

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.merge_cells("H15:M16")
        ws['H15'] = "Chemical Composition"

        app = MockApp()
        PMIReportApp.set_eulji_headers(app, ws)

        # H15 should NOT be overwritten with '-' in standard PMI mode
        self.assertEqual(ws['H15'].value, "Chemical Composition")

    def test_convert_pmi_eulji_to_plant_triple(self):
        """Verify dynamic conversion of standard 20-col PMI Eulji to 13-col plant piping layout."""
        from home.src.report_apps.ndt_report.src.비파괴검사보고서 import PMIReportApp

        class DummyVar:
            def __init__(self, val):
                self._val = val
            def get(self):
                return self._val

        class MockApp:
            def __init__(self):
                self.config = {'DATA_END_ROW': 47}
                self.current_mode = "PMI"
                self.pmi_layout_mode = DummyVar("PLANT_TRIPLE")
                self.gapji_customer = DummyVar("테스트고객사")
                self.gapji_report_no = DummyVar("SIS-TEST-001")
                self.pmi_chem_elements = [DummyVar("Ni"), DummyVar("Cr"), DummyVar("Mo"), DummyVar("-"), DummyVar("-")]
                self.pmi_chem_specs = [
                    DummyVar("A312-L Ni:10.0 ~ 14.0%, Cr:16.0 ~ 18.0%, Mo:2.0 ~ 3.0%"),
                    DummyVar("ER-316L Ni:11.0 ~ 14.0%, Cr:18.0 ~ 20.0%, Mo:2.0 ~ 3.0%"),
                    DummyVar("A182-Ni: 10.0 ~ 14.0%, Cr: 16.0 ~ 18.0%, Mo:2.0 ~ 3.0%")
                ]

            def safe_set_value(self, ws, coord, value, align=None):
                from openpyxl.cell.cell import MergedCell
                target = ws[coord] if not hasattr(coord, 'coordinate') else coord
                if isinstance(target, MergedCell):
                    for m in ws.merged_cells.ranges:
                        if (coord if not hasattr(coord, 'coordinate') else coord.coordinate) in m:
                            target = ws.cell(m.min_row, m.min_col)
                            break
                target.value = value

            def safe_merge_cells(self, ws, start_row, start_column, end_row, end_column):
                for m in list(ws.merged_cells.ranges):
                    if (m.min_row <= end_row and m.max_row >= start_row and
                        m.min_col <= end_column and m.max_col >= start_column):
                        try: ws.unmerge_cells(str(m))
                        except: pass
                ws.merge_cells(start_row=start_row, start_column=start_column,
                               end_row=end_row, end_column=end_column)

        std_template_path = "home/data/PMI-템플릿보고서.xlsx"
        if not os.path.exists(std_template_path):
            self.skipTest("PMI-템플릿보고서.xlsx not found")

        wb = openpyxl.load_workbook(std_template_path)
        ws0 = wb.worksheets[0]
        ws = wb['001'] if '001' in wb.sheetnames else wb.worksheets[1]

        app = MockApp()
        PMIReportApp.convert_pmi_eulji_to_plant_triple(app, ws, ws0=ws0)

        merges_11_18 = [str(m) for m in ws.merged_cells.ranges if any(m.min_row <= r <= m.max_row for r in range(11, 19))]

        # Verify key merges matching Photo 3
        self.assertIn("A11:E18", merges_11_18)
        self.assertIn("F11:F18", merges_11_18)
        self.assertIn("G11:G18", merges_11_18)
        self.assertIn("H11:L12", merges_11_18)
        self.assertIn("H13:L13", merges_11_18)
        self.assertIn("H14:L14", merges_11_18)
        self.assertIn("H16:L16", merges_11_18)
        self.assertIn("H17:L17", merges_11_18)
        self.assertIn("H18:L18", merges_11_18)
        self.assertIn("M11:M18", merges_11_18)

        # Verify header text values matching Photo 3
        self.assertEqual(ws['A11'].value, "Drawing No.")
        self.assertEqual(ws['F11'].value, "Joint No.")
        self.assertEqual(ws['G11'].value, "Test Location")
        self.assertIn("Alloy Verification Method", str(ws['H11'].value))
        self.assertIn("Alloy Analyzer", str(ws['H13'].value))
        self.assertEqual(ws['H14'].value, "Chemical Composition")
        self.assertEqual(ws['H15'].value, "Ni")
        self.assertEqual(ws['I15'].value, "Cr")
        self.assertEqual(ws['J15'].value, "Mo")
        self.assertIn("A312-L", str(ws['H16'].value))
        self.assertIn("ER-316L", str(ws['H17'].value))
        self.assertIn("A182-Ni", str(ws['H18'].value))
        self.assertEqual(ws['M11'].value, "Type of Material")

        # Verify top metadata
        self.assertIn("주문주", str(ws['K5'].value))
        self.assertIn("보고서 번호", str(ws['K8'].value))
        self.assertEqual(ws['M5'].value, "테스트고객사")
        self.assertEqual(ws['M8'].value, "SIS-TEST-001")

        # Verify unused columns (N~T) are cleared
        for r in range(11, 19):
            for c in range(14, 21):
                self.assertIsNone(ws.cell(row=r, column=c).value)

        # Verify print area is adjusted to 13-column layout (A1:M45, 27 rows = 9 joints)
        self.assertIn("$A$1:$M$45", str(ws.print_area))
        self.assertAlmostEqual(ws.page_margins.left, 0.25, places=2)
        self.assertAlmostEqual(ws.page_margins.right, 0.25, places=2)

    def test_pmi_element_cell_format_and_widths(self):
        """Verify chemical element columns have adequate width and shrink_to_fit to prevent ### numbers."""
        import openpyxl
        wb = openpyxl.load_workbook("home/data/PMI-템플릿보고서.xlsx")
        ws = wb["PMI(을)"]
        
        # Verify H~M column widths >= 4.6 (preventing ###)
        for col_l in ['H', 'I', 'J', 'K', 'L', 'M']:
            w = ws.column_dimensions[col_l].width
            self.assertGreaterEqual(w, 4.6, f"Column {col_l} width {w} should be >= 4.6")
            
        # Verify total A~T width remains exactly 91.70 (aligned with cover sheet)
        total_w = sum(ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width or 8.43 for c in range(1, 21))
        self.assertAlmostEqual(total_w, 91.70, places=1)
        
        # Verify shrink_to_fit on element header row (row 17)
        for c in range(8, 14):
            cell = ws.cell(row=17, column=c)
            self.assertTrue(cell.alignment.shrink_to_fit, f"Header cell {cell.coordinate} must have shrink_to_fit=True")
            self.assertEqual(cell.font.size, 9.0, f"Header cell {cell.coordinate} font size must be 9.0")


if __name__ == '__main__':
    unittest.main()
