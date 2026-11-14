import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.vehicle_catalog import vehicle_label, restore_vehicle_catalog
from site_apps.central.mobile import server


class PeopleVehicleTests(unittest.TestCase):
    def test_deleted_default_vehicles_remain_deleted_after_restore(self):
        saved={'vehicles':['71고4405 (스타렉스 12인승)', '자가용/대중교통']}
        restored=restore_vehicle_catalog(saved)
        self.assertEqual(set(restored),set(saved['vehicles']))
        self.assertEqual(restore_vehicle_catalog({'vehicles':[]}),[])
        self.assertEqual(len(restore_vehicle_catalog({})),12)
    def test_descriptions_survive_restore(self):
        for value in ('71고4405 (스타렉스 12인승)', '자가용/대중교통',
                      '개인차량-이병국 (스타렉스)'):
            self.assertEqual(vehicle_label(value), value)
        self.assertEqual(vehicle_label('81두1580'), '81두1580 (탑차)')

    def test_mobile_reads_updates_and_preserves_explicit_empty_list(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            with patch.object(server, 'CONFIG_PATH', path):
                self.assertEqual(server.get_people_vehicle_choices('users', ['fallback']), ['fallback'])
                path.write_text(json.dumps({'users':['A','B'], 'vehicles':['자가용/대중교통']}), encoding='utf-8')
                self.assertEqual(server.get_config()['users'], ['A','B'])
                self.assertEqual(server.get_config()['vehicles'], ['자가용/대중교통'])
                path.write_text('{"users":[]}', encoding='utf-8')
                self.assertEqual(server.get_people_vehicle_choices('users', ['fallback']), [])


if __name__ == '__main__': unittest.main()
