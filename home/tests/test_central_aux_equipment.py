"""Auxiliary equipment stays separate from the instrument and stock materials."""
import datetime
import unittest
from types import SimpleNamespace
import pandas as pd

from test_central_inspection_batch import Field, fixture
from site_apps.central.src.aux_equipment import names, value, restore, report_names, equipment_candidates, equipment_catalog, set_selection
from site_apps.central.src.inspection_batch import commit_plans


class AuxiliaryEquipmentTests(unittest.TestCase):
    def test_spelling_aliases_merge_and_existing_checks_follow_canonical_label(self):
        app = SimpleNamespace(
            equipments=['COBRA COBRA Manual Scanner', 'COBRA Manual Scanner',
                        'COBRA Manual Scanner (COBRA) [자동등록]',
                        'PAUT MX2(OMNI2-103085)', 'PAUT장비-OMNI2-103085 (MX2)'],
            materials_df=pd.DataFrame([
                {'품목명': 'PAUT장비', '모델명': 'MX2', 'SN': 'OMNI2-103085', 'Active': 1}]))
        candidates, aliases = equipment_catalog(app, 'main', 'COBRA COBRA Manual Scanner')
        self.assertEqual(len(candidates), 2)
        self.assertIn('COBRA Manual Scanner', candidates)
        self.assertEqual(aliases['COBRA COBRA Manual Scanner'], 'COBRA Manual Scanner')
        self.assertEqual(aliases['PAUT장비-OMNI2-103085 (MX2)'], aliases['PAUT MX2(OMNI2-103085)'])

    def test_ambiguous_inventory_names_and_distinct_serials_are_not_collapsed(self):
        app = SimpleNamespace(
            equipments=['Probe', 'Probe (SN1)', 'Probe (SN2)', 'Probe (OMNI2-123)', 'Probe (OMNI123-2)'],
            materials_df=pd.DataFrame([
                {'품목명': 'Probe', '모델명': '', 'SN': 'SN1', 'Active': 1},
                {'품목명': 'Probe', '모델명': '', 'SN': 'SN2', 'Active': 1}]))
        candidates = equipment_candidates(app, 'main')
        self.assertEqual(len(candidates), 5)
        self.assertTrue({'Probe (SN1)', 'Probe (SN2)', 'Probe (OMNI2-123)', 'Probe (OMNI123-2)'} <= set(candidates))

    def test_candidates_keep_serials_custom_values_and_exclude_inactive_stock(self):
        app = SimpleNamespace(
            equipments=['OMNI (SN001); OMNI (SN002)'],
            aux_equipments=[{'name': 'Probe (P01)'}],
            daily_usage_df=pd.DataFrame([{'장비명': 'OMNI (SN001); Legacy', '보조장비': 'Probe (P02)'}]),
            materials_df=pd.DataFrame([
                {'품목명': 'OMNI', '모델명': 'MX2', 'SN': 'SN003', 'Active': 1},
                {'품목명': 'retired', '모델명': '', 'SN': 'SN004', 'Active': 0},
                {'품목명': 'chemical', '모델명': '', 'SN': '', 'Active': 1}]),
            _is_consumable_material=lambda name, method: name == 'chemical')
        candidates = equipment_candidates(app, 'main', 'Custom (X01)')
        self.assertTrue({'OMNI (SN001)', 'OMNI (SN002)', 'MX2 OMNI (SN003)', 'Legacy', 'Custom (X01)'} <= set(candidates))
        self.assertFalse(any('retired' in item or 'chemical' in item for item in candidates))
        self.assertIn('Probe (P01)', equipment_candidates(app, 'aux'))
        self.assertIn('Probe (P02)', equipment_candidates(app, 'aux'))

    def test_main_and_aux_multiple_selections_survive_save_and_report(self):
        app = fixture()
        set_selection(app.cb_daily_equip, ['OMNI (SN1)', 'OMNI (SN2)', 'OMNI (SN1)'])
        app.daily_aux_equipment_var = Field('')
        set_selection(app.daily_aux_equipment_var, ['Probe (P1)', 'Scanner (S1)'])
        plan = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 24),
                                                 '중앙지사', prepare_only=True)
        self.assertEqual(plan['common_data']['장비명'], 'OMNI (SN1); OMNI (SN2)')
        self.assertEqual(plan['common_data']['보조장비'], 'Probe (P1); Scanner (S1)')
        commit_plans(app, [plan])
        self.assertEqual(report_names(app.daily_usage_df),
                         ['OMNI (SN1)', 'OMNI (SN2)', 'Probe (P1)', 'Scanner (S1)'])

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
