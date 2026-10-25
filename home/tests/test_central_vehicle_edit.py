"""Editing a saved vehicle must preserve other vehicles, work rows and costs."""
from copy import deepcopy
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.vehicle_records import (
    record_choices, update_vehicle, persist_vehicle_edit, slots,
)
import test_central_vehicle_panel as panel_tests


def records():
    base = dict(Date='2026-10-24', Site='중앙지사', EntryTime='2026-10-24 10:00:00',
                MaterialID='OMNI', Usage=12.3, 검사방법='PAUT', 검사비=99900,
                차량번호='90서8616 || 95가0175', 주행거리='100 || 80',
                차량점검='out_exterior:양호|photo_path:first.jpg || out_exterior:불량|photo_path:second.jpg',
                차량비고='첫 차량 || 추가 차량', Note='작업 비고')
    return pd.DataFrame([dict(base, 검사구분='ORI'), dict(base, 검사구분='REP', Usage=2.0),
                         dict(base, EntryTime='2026-10-24 12:00:00', 검사구분='ORI')])


class VehicleEditTests(unittest.TestCase):
    def selection(self, frame):
        return next(c for c in record_choices(frame) if c['data']['vehicle_info'] == '95가0175')

    def test_updates_selected_slot_and_linked_ori_rep_without_changing_costs(self):
        frame = records()
        before = frame.copy(deep=True)
        choice = self.selection(frame)
        self.assertEqual(len(choice['source']['fingerprints']), 2)
        data = dict(choice['data'], vehicle_info='95가0176', mileage='1,234', remarks='',
                    out_exterior='양호', photo_path='new-photo.jpg')
        updated, _ = update_vehicle(frame, choice['source'], data)
        pd.testing.assert_frame_equal(frame, before)
        self.assertEqual(len(updated), 3)
        self.assertEqual(updated.iloc[0]['차량번호'], '90서8616 || 95가0176')
        self.assertEqual(updated.iloc[1]['주행거리'], '100 || 1234')
        self.assertEqual(updated.iloc[0]['차량비고'], '첫 차량 || ')
        self.assertEqual(slots(updated.iloc[0])[0][1]['photo_path'], 'first.jpg')
        self.assertEqual(slots(updated.iloc[0])[1][1]['photo_path'], 'new-photo.jpg')
        pd.testing.assert_series_equal(updated.iloc[2], before.iloc[2])
        pd.testing.assert_frame_equal(updated.drop(columns=['차량번호', '주행거리', '차량점검', '차량비고']),
                                      before.drop(columns=['차량번호', '주행거리', '차량점검', '차량비고']))

    def test_row_reordering_does_not_change_the_target(self):
        frame = records()
        choice = self.selection(frame)
        reordered = frame.iloc[[2, 1, 0]].reset_index(drop=True)
        result, _ = update_vehicle(reordered, choice['source'], dict(choice['data'], mileage='90'))
        self.assertEqual(result['주행거리'].tolist(), ['100 || 80', '100 || 90', '100 || 90'])

    def test_stale_or_deleted_record_is_rejected_without_mutating_frame(self):
        frame = records()
        choice = self.selection(frame)
        frame.loc[0, '차량비고'] = '다른 편집자가 수정함'
        before = frame.copy(deep=True)
        with self.assertRaisesRegex(ValueError, '변경되었거나 삭제'):
            update_vehicle(frame, choice['source'], dict(choice['data'], mileage='90'))
        pd.testing.assert_frame_equal(frame, before)

    def test_failed_persistence_rolls_back_and_success_can_be_edited_again(self):
        frame = records()
        choice = self.selection(frame)
        app = SimpleNamespace(daily_usage_df=frame, save_data=Mock(return_value=False))
        with self.assertRaises(RuntimeError):
            persist_vehicle_edit(app, choice['source'], dict(choice['data'], mileage='90'))
        self.assertIs(app.daily_usage_df, frame)
        app.save_data.return_value = True
        source = persist_vehicle_edit(app, choice['source'], dict(choice['data'], mileage='90'))
        persist_vehicle_edit(app, source, dict(choice['data'], mileage='95'))
        self.assertEqual(app.daily_usage_df.iloc[0]['주행거리'], '100 || 95')
        self.assertEqual(len(app.daily_usage_df), 3)

    def test_blank_legacy_columns_and_serialized_sources(self):
        import json
        frame = pd.DataFrame([{'Date': '2026-10-24', 'Site': '중앙지사', '차량번호': '한대',
                               '주행거리': float('nan'), '차량점검': 'out_exterior,out_locking'}])
        choice = record_choices(frame)[0]
        self.assertEqual(choice['data']['out_exterior'], '양호')
        self.assertEqual(choice['data']['out_locking'], '잠금')
        source = json.loads(json.dumps(choice['source']))
        result, _ = update_vehicle(frame, source, dict(choice['data'], mileage='0', remarks='수정'))
        self.assertEqual(result.iloc[0]['차량비고'], '수정')

    def test_invalid_mileage_is_rejected(self):
        frame = records()
        choice = self.selection(frame)
        for mileage in ('-1', 'nan', 'inf', 'abc'):
            with self.subTest(mileage=mileage), self.assertRaises(ValueError):
                update_vehicle(frame, choice['source'], dict(choice['data'], mileage=mileage))

    def test_panel_edit_uses_existing_record_path_and_keeps_source_after_save(self):
        panel, first, second = panel_tests.VehiclePanelStateTests().panel()
        choice = self.selection(records())
        panel.sources[second] = choice['source']
        panel.app.daily_usage_df = records()
        panel.app.save_data = Mock(return_value=True)
        for name in ('update_daily_usage_view', 'update_recent_entries_view', 'refresh_inquiry_filters'):
            setattr(panel.app, name, Mock())
        second.data = deepcopy(choice['data'])
        second.data['mileage'] = '90'
        with self.assertRaisesRegex(ValueError, '수정 저장'):
            panel.check_pending_edits()
        panel.save_one(second)
        panel.check_pending_edits()
        panel.app.save_single_vehicle_data.assert_not_called()
        self.assertEqual(len(panel.app.daily_usage_df), 3)
        self.assertEqual(panel.app.daily_usage_df.iloc[0]['주행거리'], '100 || 90')
        self.assertIn('_edit_source', panel.serialize()[1])

    def test_completed_edit_is_not_duplicated_in_new_inspection_records(self):
        from site_apps.central.src.inspection_batch import InspectionBatch
        from test_central_inspection_batch import fixture
        panel, first, second = panel_tests.VehiclePanelStateTests().panel()
        panel.sources[second] = self.selection(records())['source']
        panel.mark_saved(second)
        shared = InspectionBatch.shared(SimpleNamespace(app=SimpleNamespace(
            **vars(panel.app), vehicle_panel=panel)))
        self.assertEqual(shared['차량번호'], '90서8616')
        app = fixture()
        app.vehicle_panel = panel
        app.fixed_vehicle_widget = first
        app.vehicle_boxes = [second]
        import datetime
        plan = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 24),
                                                  '중앙지사', prepare_only=True)
        self.assertEqual(plan['records_to_save'][0]['차량번호'], '90서8616')


if __name__ == '__main__':
    unittest.main()
