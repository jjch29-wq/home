import importlib.util
import datetime
from pathlib import Path
import unittest

import pandas as pd
import openpyxl


MODULE_PATH = Path(__file__).resolve().parents[1] / "src" / "비파괴검사보고서.py"
SPEC = importlib.util.spec_from_file_location("ndt_report_app", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class PtColumnDetectionTests(unittest.TestCase):
    class DummyVar:
        def __init__(self, value=""):
            self.value = value

        def get(self):
            return self.value

        def set(self, value):
            self.value = value

    def test_no_column_does_not_match_spool_or_request_number(self):
        df = pd.DataFrame(columns=["ISO", "Spool No", "Joint", "Request No", "Film No"])

        self.assertIsNone(MODULE.find_exact_column(df, ["NO.", "NO", "SEQ", "ITEM", "순번"]))

    def test_standalone_no_header_is_detected(self):
        df = pd.DataFrame(columns=["ISO", " No. ", "Spool No", "Joint"])

        self.assertEqual(
            MODULE.find_exact_column(df, ["NO.", "NO", "SEQ", "ITEM", "순번"]),
            " No. ",
        )

    def test_supported_sequence_aliases_are_detected(self):
        for header in ("NO", "SEQ", "ITEM", "순번"):
            with self.subTest(header=header):
                df = pd.DataFrame(columns=["ISO", header, "Joint"])
                self.assertEqual(
                    MODULE.find_exact_column(df, ["NO.", "NO", "SEQ", "ITEM", "순번"]),
                    header,
                )

    def test_request_date_header_is_detected_without_matching_other_dates(self):
        df = pd.DataFrame(columns=["Report Date", "Test Date", "Request Date", "Weld Date"])

        self.assertEqual(
            MODULE.find_exact_column(df, ["REQUEST DATE"]),
            "Request Date",
        )

    def test_repeated_request_date_from_merged_headers_is_detected(self):
        df = pd.DataFrame(
            columns=["Report Date Report Date", "Request Date Request Date", "Weld Date Weld Date"]
        )

        self.assertEqual(
            MODULE.find_exact_column(df, ["REQUEST DATE"]),
            "Request Date Request Date",
        )

    def test_thickness_range_uses_smallest_and_largest_values(self):
        rows = [
            {"Thk.": "2.77"},
            {"Thk.": 1.65},
            {"Thk.": "7.15mm"},
            {"Thk.": "SCH40"},
            {"Thk.": ""},
        ]

        self.assertEqual(MODULE.format_thickness_range(rows), "1.65 ~ 7.15mm")

    def test_thickness_range_is_empty_without_numeric_values(self):
        self.assertEqual(
            MODULE.format_thickness_range([{"Thk.": "SCH40"}, {"Thk.": ""}]),
            "",
        )

    def test_first_report_date_uses_extracted_request_date(self):
        result = MODULE.first_report_date([{"Date": "2026-09-04"}])

        self.assertEqual(result, datetime.datetime(2026, 9, 4))

    def test_pt_exam_date_field_value_is_parseable_for_cover(self):
        result = MODULE.first_report_date([{"Date": "2026-09-05"}])

        self.assertEqual(result, datetime.datetime(2026, 9, 5))

    def test_first_report_date_skips_invalid_values(self):
        result = MODULE.first_report_date(
            [{"Date": ""}, {"Date": "not-a-date"}, {"Date": "2026.09.04"}]
        )

        self.assertEqual(result, datetime.datetime(2026, 9, 4))

    def test_unused_template_sheet_is_removed_from_page_count(self):
        workbook = openpyxl.Workbook()
        cover = workbook.active
        unused_data_sheet = workbook.create_sheet("001")

        MODULE.remove_unused_report_sheets(workbook, [cover])

        self.assertEqual(workbook.worksheets, [cover])
        self.assertNotIn(unused_data_sheet, workbook.worksheets)

    def test_pt_report_info_is_stored_and_restored_per_mode(self):
        app = MODULE.PMIReportApp.__new__(MODULE.PMIReportApp)
        app.config = {}
        app.gapji_project = self.DummyVar("PT Project")
        app.gapji_customer = self.DummyVar("PT Customer")
        app.gapji_item = self.DummyVar("PIPE")
        app.gapji_material = self.DummyVar("S/S")
        app.gapji_report_no = self.DummyVar("PT-0072")
        app.gapji_exam_date = self.DummyVar("2026-07-30")
        app._initialize_report_info_by_mode()

        app._store_report_info_for_mode("PT")
        for var in (
            app.gapji_project, app.gapji_customer, app.gapji_item,
            app.gapji_material, app.gapji_report_no, app.gapji_exam_date,
        ):
            var.set("")
        app._load_report_info_for_mode("PT")

        self.assertEqual(app.gapji_project.get(), "PT Project")
        self.assertEqual(app.gapji_customer.get(), "PT Customer")
        self.assertEqual(app.gapji_item.get(), "PIPE")
        self.assertEqual(app.gapji_material.get(), "S/S")
        self.assertEqual(app.gapji_report_no.get(), "PT-0072")
        self.assertEqual(app.gapji_exam_date.get(), "2026-07-30")
        self.assertEqual(app.config["PT_PROJECT"], "PT Project")


if __name__ == "__main__":
    unittest.main()
