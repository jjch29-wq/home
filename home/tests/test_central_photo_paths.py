import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src import data_sync


class PhotoPathTest(unittest.TestCase):
    def test_cloud_history_reuses_exact_central_local_photo(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            local_src = root / 'central/src'
            local_src.mkdir(parents=True)
            photo = root / 'central/data/process_photos/2026-08-07/PAUT_uuid.jpg'
            photo.parent.mkdir(parents=True)
            photo.write_bytes(b'fixture')
            history = root / 'cloud/central/daily_work_history.json'
            stored = '../data/process_photos/2026-08-07/PAUT_uuid.jpg'
            with patch.object(data_sync, '__file__', str(local_src / 'data_sync.py')):
                self.assertEqual(data_sync.resolve_process_photo_path(str(history), stored), str(photo))
                self.assertIsNone(data_sync.resolve_process_photo_path(str(history), stored.replace('uuid', 'other')))
                # A current cloud copy always wins over the historical copy.
                current = history.parent / 'data/process_photos/2026-08-07/PAUT_uuid.jpg'
                current.parent.mkdir(parents=True)
                current.write_bytes(b'new')
                self.assertEqual(data_sync.resolve_process_photo_path(str(history), stored), str(current))


if __name__ == '__main__':
    unittest.main()
