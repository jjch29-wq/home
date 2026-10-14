from pathlib import Path
import sys
import unittest
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.services.daily_report_records import select_daily_report_records


class DailyReportRecordsTests(unittest.TestCase):
    def test_display_name_change_does_not_change_source_records(self):
        frame = pd.DataFrame([
            dict(Date='2026-10-14', Site='함감속기', 검사방법='MT', NDT_형광자분=1),
            dict(Date='2026-10-14', Site='함감속기', 검사방법='UT', NDT_형광자분=0),
            dict(Date='2026-10-14', Site='N/A', 검사방법='PT', NDT_형광자분=0)])
        rows = select_daily_report_records(frame, '2026-10-14', 'N/A', loaded_index=1)
        self.assertEqual(rows['검사방법'].tolist(), ['MT', 'UT'])
        self.assertEqual(rows['NDT_형광자분'].sum(), 1)
        self.assertEqual(frame.at[1, 'Site'], '함감속기')

    def test_any_method_uses_whole_day_and_site(self):
        frame = pd.DataFrame([
            dict(Date='2026-10-14', Site='함감속기', 검사방법='MT', Usage=2),
            dict(Date='2026-10-14', Site='함감속기', 검사방법='UT', Usage=6),
            dict(Date='2026-10-14', Site='함감속기', 검사방법='UT', Usage=6),
            dict(Date='2026-10-13', Site='함감속기', 검사방법='MT', Usage=9),
            dict(Date='2026-10-14', Site='다른현장', 검사방법='MT', Usage=9)])
        for selected in frame.iloc[:3].to_dict('records'):
            rows = select_daily_report_records(frame, selected['Date'], selected['Site'])
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows.groupby('검사방법')['Usage'].sum().to_dict(), {'MT': 2, 'UT': 12})


if __name__ == '__main__':
    unittest.main()
