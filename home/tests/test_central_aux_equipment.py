"""Auxiliary equipment stays separate from the instrument and stock materials."""
import datetime
import unittest

from test_central_inspection_batch import Field, fixture
from site_apps.central.src.aux_equipment import names, value, restore, report_names
from site_apps.central.src.inspection_batch import commit_plans


class AuxiliaryEquipmentTests(unittest.TestCase):
    def test_multiple_jobs_keep_their_own_auxiliary_equipment(self):
        app = fixture()
        app.daily_aux_equipment_var = Field('JIREH MICROBE BTG006; Probe A; JIREH MICROBE BTG006')
        day = datetime.date(2026, 10, 24)
        first = app._add_single_usage_record_logic('OMNI', day, '중앙지사', prepare_only=True)
        restore(app, 'Probe B')
        second = app._add_single_usage_record_logic('OMNI', day, '중앙지사', prepare_only=True)
        self.assertEqual(first['common_data']['장비명'], 'OMNI')
        self.assertEqual(first['common_data']['보조장비'], 'JIREH MICROBE BTG006; Probe A')
        self.assertEqual(second['common_data']['보조장비'], 'Probe B')
        commit_plans(app, [first, second])
        self.assertEqual(set(app.daily_usage_df['보조장비']),
                         {'JIREH MICROBE BTG006; Probe A', 'Probe B'})
        self.assertEqual(report_names(app.daily_usage_df),
                         ['OMNI', 'JIREH MICROBE BTG006', 'Probe A', 'Probe B'])

    def test_old_records_and_sessions_default_to_empty(self):
        app = fixture()
        self.assertEqual(value(app), '')
        app.daily_aux_equipment_var = Field('old scanner')
        restore(app, float('nan'))
        self.assertEqual(value(app), '')
        self.assertEqual(names(None), [])
        plan = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 24),
                                                  '중앙지사', prepare_only=True)
        self.assertEqual(plan['common_data']['보조장비'], '')


if __name__ == '__main__':
    unittest.main()
