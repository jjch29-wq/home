"""Verify Step 3 field separation and legacy compatibility in the central app."""
import datetime
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'home/src'))

from site_apps.central.src.app import MaterialManager
from site_apps.central.src.inspection_batch import combine_plans
from site_apps.central.src.models.material_model import _sync_dataframe_schema_impl
from site_apps.central.src.services.data_loader import load_data_impl
from site_apps.central.src.views.daily_usage_form_controller import load_daily_usage_to_form_impl


class DummyField:
    def __init__(self, value=''):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = str(value)

    def delete(self, *_args):
        self.value = ''

    def insert(self, _index, value):
        self.value = str(value)

    def cget(self, key):
        if key == 'values':
            return ('한국지역난방공사', '한국지역난방공사 중앙지사')
        return ''

    def event_generate(self, _event):
        return None


class DummyText:
    def __init__(self, text=''):
        self.text = text

    def get(self, *_args):
        return self.text

    def delete(self, *_args):
        self.text = ''

    def insert(self, _index, value):
        self.text += str(value)


class DummyDateField:
    def __init__(self, dt=None):
        self.dt = dt or datetime.date(2026, 11, 26)

    def set_date(self, val):
        self.dt = pd.to_datetime(val).date()

    def get_date(self):
        return self.dt


def build_form_app():
    ndt_names = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
    ndt_entries = [{'_company': DummyField(''), **{n: DummyField('') for n in ndt_names}}]

    def add_section():
        ndt_entries.append({'_company': DummyField(''), **{n: DummyField('') for n in ndt_names}})

    def remove_section():
        if len(ndt_entries) > 1:
            ndt_entries.pop()

    app = SimpleNamespace(
        ent_daily_date=DummyDateField(),
        cb_daily_client=DummyField('한국지역난방공사'),
        cb_daily_company=DummyField(''),
        cb_daily_site=DummyField(''),
        ndt_loc_type_var=DummyField(''),
        ent_daily_applied_code=DummyField(''),
        cb_daily_equip=DummyField(''),
        ent_daily_report_no=DummyField(''),
        ent_daily_inspection_item=DummyField(''),
        cb_daily_test_method=DummyField(''),
        cb_daily_unit=DummyField(''),
        hot_tap_stage_var=DummyField('1차'),
        hot_tap_size_var=DummyField('600A'),
        hot_tap_location_var=DummyField(''),
        ent_daily_test_amount=DummyField(''),
        ent_daily_unit_price=DummyField(''),
        ent_daily_travel_cost=DummyField(''),
        ent_daily_test_fee=DummyField(''),
        ent_daily_meal_cost=DummyField(''),
        ndt_work_time_var=DummyField(''),
        ndt_source_var=DummyField(''),
        ndt_pipe_var=DummyField(''),
        ndt_thickness_var=DummyField(''),
        ndt_overhead_var=DummyField('110'),
        ndt_tech_var=DummyField('20'),
        ndt_report_pipe_var=DummyField(''),
        ndt_rep_joint_var=DummyField(''),
        ndt_rep_qty_var=DummyField(''),
        ndt_ori_joint_var=DummyField(''),
        ndt_ori_qty_var=DummyField(''),
        ndt_rej_joint_var=DummyField(''),
        cb_daily_material=DummyField(''),
        rtk_entries={},
        calculate_rtk_total=lambda: None,
        ndt_company_entries=ndt_entries,
        add_ndt_company_section=add_section,
        remove_last_ndt_company=remove_section,
        ent_daily_note=DummyField(''),
        main_memo_text=DummyText('기존메모'),
        daily_usage_df=pd.DataFrame(),
        clean_nan=lambda v: '' if v is None or (isinstance(v, float) and pd.isna(v)) or str(v) == 'nan' else str(v),
        get_material_display_name=lambda m: str(m or ''),
    )
    return app


class FieldSeparationTests(unittest.TestCase):
    def test_schema_and_batch_separate_client_from_chemical_manufacturer(self):
        synced = _sync_dataframe_schema_impl(SimpleNamespace(), pd.DataFrame([{'Site': '중앙'}]), 'DailyUsage')
        self.assertIn('발주처', synced.columns)
        self.assertIn('회사코드', synced.columns)

        plans = [{
            'common_data': {'Date': '2026-11-26', 'Site': 'A', '업체명': '다림건설(주)', 'User': '홍길동', 'WorkTime': '09:00~18:00', 'OT': '', 'Meal': ''},
            'records_to_save': [{'Date': '2026-11-26', 'Site': 'A', '업체명': '다림건설(주)', '회사코드': 'NAWOO', '검사비': 1000.0}],
        }]
        shared = {'Date': '2026-11-26', 'Site': '중앙지사', '발주처': '한국지역난방공사 중앙지사', '업체명': '다림건설(주)', '출장비': 0.0, '일식': 0.0}
        combined = combine_plans(plans, shared)
        row = combined[0]['records_to_save'][0]
        self.assertEqual(row['발주처'], '한국지역난방공사 중앙지사')
        self.assertEqual(row['회사코드'], 'NAWOO')

    def test_form_restores_client_and_manufacturer_including_legacy_fallback(self):
        app = build_form_app()
        # 1) New separated record: 발주처 has client, 회사코드 has NDT manufacturer
        load_daily_usage_to_form_impl(app, {
            'Date': '2026-11-26',
            'Site': '중앙지사',
            '발주처': '한국지역난방공사 중앙지사',
            '업체명': '다림건설(주)',
            '회사코드': 'NAWOO',
            'NDT_백색페인트': 2.0,
            'Note': '',
            'DailyMemo': '',
        })
        self.assertEqual(app.cb_daily_client.get(), '한국지역난방공사 중앙지사')
        self.assertEqual(app.ndt_company_entries[0]['_company'].get(), 'NAWOO')

        # 2) Legacy record without 발주처 where 회사코드 stored client name
        load_daily_usage_to_form_impl(app, {
            'Date': '2026-11-25',
            'Site': '중앙지사',
            '업체명': '다림건설(주)',
            '회사코드': '한국지역난방공사 중앙지사',
            'Note': '',
            'DailyMemo': '',
        })
        self.assertEqual(app.cb_daily_client.get(), '한국지역난방공사 중앙지사')
        self.assertEqual(app.ndt_company_entries[0]['_company'].get(), '')

    def test_note_and_system_tags_never_pollute_daily_memo_on_load(self):
        app = build_form_app()
        load_daily_usage_to_form_impl(app, {
            'Date': '2026-11-26',
            'Site': '중앙지사',
            'Note': '용접부 재확인 [핫태핑 S01 400A 1차]',
            'DailyMemo': '',
        })
        self.assertEqual(app.main_memo_text.get(), '')

        # Previously polluted DailyMemo starting with system tag is suppressed
        load_daily_usage_to_form_impl(app, {
            'Date': '2026-11-26',
            'Site': '중앙지사',
            'Note': '[핫태핑 S01 400A 1차]',
            'DailyMemo': '[핫태핑 S01 400A 1차]',
        })
        self.assertEqual(app.main_memo_text.get(), '')

        # Explicit user DailyMemo restores cleanly
        load_daily_usage_to_form_impl(app, {
            'Date': '2026-11-26',
            'Site': '중앙지사',
            'Note': '개별 작업 비고',
            'DailyMemo': '당일 현장 특이사항 공유',
        })
        self.assertEqual(app.main_memo_text.get(), '당일 현장 특이사항 공유')

    def test_edit_synchronizes_usage_amount_qty_and_preserves_ot_strings(self):
        daily_df = pd.DataFrame([
            {
                'Date': pd.Timestamp('2026-11-26'),
                'Site': '중앙지사',
                '검사방법': 'PAUT',
                '검사구분': 'ORI',
                'Usage': 5.0,
                '검사량': 5.0,
                '수량': 5.0,
                'OT': '45,000원',
                'MaterialID': '',
            },
            {
                'Date': pd.Timestamp('2026-11-26'),
                'Site': '중앙지사',
                '검사방법': 'RT',
                '검사구분': 'REP',
                'Usage': 0.0,
                '검사량': 0.0,
                '수량': 0.0,
                '재검사매수': 3.0,
                'OT': '',
                'MaterialID': '',
            },
        ])
        app = SimpleNamespace(
            daily_usage_df=daily_df,
            transactions_df=pd.DataFrame(),
            materials_df=pd.DataFrame(columns=['MaterialID', '품목명', 'Active']),
            monthly_usage_df=pd.DataFrame(),
            save_data=Mock(return_value=True),
            get_material_display_name=lambda _m: '',
            update_daily_usage_view=Mock(),
            refresh_inquiry_filters=Mock(),
            update_stock_view=Mock(),
            update_transaction_view=Mock(),
            update_material_combo=Mock(),
            update_recent_entries_view=Mock(),
            update_monthly_usage_view=Mock(),
            _is_consumable_material=lambda *_a: False,
            _auto_reconcile_and_register_ndt=Mock(),
        )
        app._safe_set_daily_df = MaterialManager._safe_set_daily_df.__get__(app, SimpleNamespace)

        self.assertTrue(MaterialManager.save_daily_usage_edits(app, 0, {
            'Date': '2026-11-26',
            'Site': '중앙지사',
            '검사방법': 'PAUT',
            '검사량': 12.0,
            'OT': '60,000원',
        }))
        self.assertEqual(app.daily_usage_df.loc[0, 'Usage'], 12.0)
        self.assertEqual(app.daily_usage_df.loc[0, '검사량'], 12.0)
        self.assertEqual(app.daily_usage_df.loc[0, '수량'], 12.0)
        self.assertEqual(app.daily_usage_df.loc[0, 'OT'], '60,000원')

        # RT REP row keeps Usage / 검사량 / 수량 at 0.0 so totals are not double-counted
        self.assertTrue(MaterialManager.save_daily_usage_edits(app, 1, {
            'Date': '2026-11-26',
            'Site': '중앙지사',
            '검사방법': 'RT',
            '검사량': 4.0,
            'OT': '',
        }))
        self.assertEqual(app.daily_usage_df.loc[1, 'Usage'], 0.0)
        self.assertEqual(app.daily_usage_df.loc[1, '검사량'], 0.0)
        self.assertEqual(app.daily_usage_df.loc[1, '수량'], 0.0)

    def test_load_data_strips_empty_unnamed_ghost_columns(self):
        with tempfile.TemporaryDirectory(dir=ROOT / 'tmp') as tmpdir:
            db_path = Path(tmpdir) / 'ghost.xlsx'
            with pd.ExcelWriter(db_path) as writer:
                pd.DataFrame([{'MaterialID': 1, '품목명': '세척제', '수량': 5, 'Active': 1, 'Unnamed: 9': None}]).to_excel(
                    writer, sheet_name='Materials', index=False
                )
                pd.DataFrame(columns=['Date', 'MaterialID', 'Site', 'Type', 'Quantity', 'Note', 'User']).to_excel(
                    writer, sheet_name='Transactions', index=False
                )
                pd.DataFrame(columns=['MaterialID', 'Year', 'Month', 'Site', 'Usage', 'Note', 'EntryDate']).to_excel(
                    writer, sheet_name='MonthlyUsage', index=False
                )
                pd.DataFrame([{
                    'Date': '2026-11-26',
                    'Site': '중앙지사',
                    'EntryTime': '2026-11-26 09:00:00',
                    'Usage': 3,
                    'Unnamed: 107': None,
                }]).to_excel(writer, sheet_name='DailyUsage', index=False)
                pd.DataFrame(columns=['Site', 'Revenue']).to_excel(
                    writer, sheet_name='Budget', index=False
                )

            app = SimpleNamespace(
                db_path=str(db_path),
                bundle_dir=tmpdir,
                config_path=str(Path(tmpdir) / 'cfg.json'),
                ndt_materials_all=["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"],
                extract_sn_from_model=lambda m, s: (m, s),
                refresh_inquiry_filters=lambda: None,
                save_data=Mock(return_value=True),
            )
            app._sync_dataframe_schema = lambda df, name: _sync_dataframe_schema_impl(app, df, name)
            load_data_impl(app)

            self.assertFalse(any(str(c).startswith('Unnamed:') for c in app.daily_usage_df.columns))
            self.assertFalse(any(str(c).startswith('Unnamed:') for c in app.materials_df.columns))
            self.assertIn('발주처', app.daily_usage_df.columns)


if __name__ == '__main__':
    unittest.main()
