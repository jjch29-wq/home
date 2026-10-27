import csv
import json
from pathlib import Path
import tempfile
import unittest
import numpy as np
from core import FilmRecord, imwrite_jpg, export_records, save_project, load_project


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory(); self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.source = self.root / 'original.jpg'
        imwrite_jpg(self.source, np.zeros((100, 200, 3), dtype=np.uint8))
        self.record = FilmRecord(source=str(self.source), line_no='LINE-A', joint_no='R01',
                                 section='(1-2)', crop_rect=(10, 10, 100, 50), report_no='RT001', report_item='1')

    def test_resave_keeps_previous_log_and_images(self):
        output = self.root / 'out'
        first = export_records([self.record], output, (0,0,200,100), '{joint}{section}')
        log = output / '영상화_작업기록.csv'; before = log.read_bytes()
        second = export_records([self.record], output, (0,0,200,100), '{joint}{section}')
        self.assertEqual(log.read_bytes(), before)
        self.assertTrue(first[0].exists()); self.assertTrue(second[0].exists())
        self.assertNotEqual(first[0], second[0])
        self.assertEqual(len(list(output.glob('*.csv'))), 2)

    def test_partial_failure_keeps_log_of_successful_files(self):
        missing = FilmRecord(source=str(self.root/'missing.jpg'))
        output = self.root/'out'
        with self.assertRaises(OSError):
            export_records([self.record, missing], output, (0,0,200,100), '{joint}{section}')
        with (output/'영상화_작업기록.csv').open(encoding='utf-8-sig', newline='') as handle:
            rows = list(csv.DictReader(handle))
        self.assertEqual(len(rows), 1)
        self.assertTrue((output/rows[0]['output_name']).exists())

    def test_project_roundtrip_preserves_crop_and_report_mapping(self):
        path = self.root/'work.rtfilm.json'
        save_project(path, [self.record], common_rect=(0,0,200,100), initial_rect=(0,0,180,90),
                     source_photos=[str(self.source)], template='{joint}{section}')
        data = load_project(path)
        self.assertEqual(data['records'][0], self.record)
        self.assertEqual(data['common_rect'], (0,0,200,100))
        self.assertEqual(data['initial_rect'], (0,0,180,90))

    def test_invalid_project_crop_is_rejected(self):
        path = self.root/'work.rtfilm.json'
        save_project(path, [self.record], common_rect=(0,0,-1,100))
        with self.assertRaises(ValueError): load_project(path)

    def test_missing_project_source_is_rejected(self):
        path = self.root/'work.rtfilm.json'
        save_project(path, [self.record])
        self.source.unlink()
        with self.assertRaises(ValueError): load_project(path)
