from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.lotte.src.vehicle_assignment import (
    update_vehicle_in_record, validate_vehicle_methods, vehicle_applies_to_method,
)


class VehicleAssignmentTests(unittest.TestCase):
    def test_vehicle_is_attached_only_to_its_method(self):
        vehicle = {'vehicle_info': '91주8839', 'inspection_method': 'RT'}
        self.assertTrue(vehicle_applies_to_method(vehicle, 'rt'))
        self.assertFalse(vehicle_applies_to_method(vehicle, 'PAUT'))

    def test_legacy_blank_method_works_for_single_method(self):
        vehicle = {'vehicle_info': '91주8839', 'inspection_method': ''}
        self.assertTrue(vehicle_applies_to_method(vehicle, 'RT'))
        self.assertEqual(validate_vehicle_methods([vehicle], {'RT'}), '')

    def test_multi_method_requires_each_vehicle_method(self):
        error = validate_vehicle_methods(
            [{'vehicle_info': '91주8839', 'inspection_method': ''}], {'RT', 'PAUT'})
        self.assertIn('선택', error)

    def test_method_must_exist_in_multi_items(self):
        error = validate_vehicle_methods(
            [{'vehicle_info': '91주8839', 'inspection_method': 'UT'}], {'RT', 'PAUT'})
        self.assertIn('UT', error)

    def test_missing_mileage_is_updated_in_selected_record(self):
        record = {
            '차량번호': '91주8839 || 90너4889',
            '주행거리': '50 || ',
            '차량점검': 'inspection_method:RT || inspection_method:RT',
            '차량비고': '첫차 || ',
        }
        changed = update_vehicle_in_record(record, {
            'inspection_method': 'RT', 'vehicle_info': '90너4889',
            'mileage': '75', '_raw_mileage': '75', 'remarks': '보완',
            'out_exterior': '양호', 'photo_path': '',
        })
        self.assertTrue(changed)
        self.assertEqual(record['주행거리'], '50 || 75')
        self.assertEqual(record['차량비고'], '첫차 || 보완')
        self.assertIn('out_exterior:양호', record['차량점검'].split('||')[1])

    def test_unrelated_vehicle_is_not_inserted_into_selected_record(self):
        record = {'차량번호': '91주8839', '주행거리': '', '차량점검': '', '차량비고': ''}
        self.assertFalse(update_vehicle_in_record(
            record, {'vehicle_info': '90너4889', '_raw_mileage': '75'}))


if __name__ == '__main__':
    unittest.main()
