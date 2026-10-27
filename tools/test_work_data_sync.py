import importlib.util
from pathlib import Path
import tempfile
import unittest

spec = importlib.util.spec_from_file_location('sync_tool', Path(__file__).with_name('work_data_sync.py'))
tool = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tool)


class SyncTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.a, self.b, self.cloud = (root / name for name in ('a', 'b', 'cloud'))
        for directory in (self.a, self.b, self.cloud):
            directory.mkdir()

    def run_sync(self, root, name):
        return tool.sync(root, self.cloud, 'development', name)

    def test_roundtrip_multiple_edits_without_stale_conflicts(self):
        a, b = self.a / 'settings.json', self.b / 'settings.json'
        a.write_text('{"value":1}')
        self.assertEqual(self.run_sync(self.a, 'a')['uploaded'], 1)
        self.assertEqual(self.run_sync(self.b, 'b')['downloaded'], 1)
        a.write_text('{"value":2}')
        self.run_sync(self.a, 'a')
        self.assertEqual(self.run_sync(self.b, 'b')['downloaded'], 1)
        self.assertFalse(self.run_sync(self.b, 'b')['conflicts'])
        b.write_text('{"value":3}')
        self.run_sync(self.b, 'b')
        self.assertEqual(self.run_sync(self.a, 'a')['downloaded'], 1)
        self.assertEqual(a.read_bytes(), b.read_bytes())
        self.assertFalse(self.run_sync(self.a, 'a')['conflicts'])

    def test_accept_unchanged_local_remains_available_to_other_pc(self):
        (self.a / 'settings.json').write_text('{"value":1}')
        self.run_sync(self.a, 'a')
        tool.sync(self.a, self.cloud, 'development', 'a', accept_local=['settings.json'])
        self.assertEqual(self.run_sync(self.a, 'a')['uploaded'], 0)
        self.assertEqual(self.run_sync(self.b, 'b')['downloaded'], 1)
        self.assertEqual((self.b / 'settings.json').read_text(), '{"value":1}')

    def test_conflicting_edits_keep_both(self):
        a, b = self.a / 'settings.json', self.b / 'settings.json'
        a.write_text('{"value":1}')
        self.run_sync(self.a, 'a')
        self.run_sync(self.b, 'b')
        a.write_text('{"value":2}')
        b.write_text('{"value":3}')
        self.run_sync(self.a, 'a')
        result = self.run_sync(self.b, 'b')
        self.assertEqual(result['conflicts'], ['settings.json'])
        self.assertEqual(b.read_text(), '{"value":3}')
        self.assertTrue(list((self.b / '.work-sync/conflicts').rglob('settings.json')))
        tool.sync(self.b, self.cloud, 'development', 'b', accept_local=['settings.json'])
        self.assertEqual(self.run_sync(self.a, 'a')['downloaded'], 1)
        self.assertEqual(a.read_text(), '{"value":3}')

    def test_credentials_code_and_runtime_are_excluded(self):
        (self.a / 'app.py').write_text('print(1)')
        (self.a / 'credentials.json').write_text('{}')
        (self.a / 'config.json').write_text('{"api_key":"sensitive"}')
        (self.a / 'tmp').mkdir()
        (self.a / 'tmp/cache.json').write_text('{}')
        result = self.run_sync(self.a, 'a')
        self.assertEqual(result['uploaded'], 0)
        self.assertEqual(result['skipped'], ['config.json'])

    def test_identical_files_have_independent_version_history(self):
        for name in ('one.json', 'two.json'):
            (self.a / name).write_text('{}')
        self.run_sync(self.a, 'a')
        (self.a / 'one.json').write_text('{"value":2}')
        self.run_sync(self.a, 'a')
        result = self.run_sync(self.b, 'b')
        self.assertEqual(result['downloaded'], 2)
        self.assertFalse(result['conflicts'])
        self.assertEqual((self.b / 'two.json').read_text(), '{}')

    def test_dry_run_and_offline_do_not_write(self):
        (self.a / 'settings.json').write_text('{}')
        tool.sync(self.a, self.cloud, 'dev', 'a', dry_run=True)
        self.assertFalse((self.a / '.work-sync').exists())
        self.assertEqual(list(self.cloud.iterdir()), [])
        with self.assertRaises(OSError):
            tool.sync(self.a, self.cloud / 'missing', 'dev', 'a')

    def test_branch_change_cannot_silently_overwrite_cloud(self):
        target = self.a / 'settings.json'
        target.write_text('{"value":1}')
        self.run_sync(self.a, 'a')
        target.write_text('{"value":0}')
        result = tool.sync(self.a, self.cloud, 'old-branch', 'a')
        self.assertEqual(result['conflicts'], ['settings.json'])
        self.assertEqual(target.read_text(), '{"value":0}')

    def test_missing_object_recovers_from_verified_snapshot(self):
        (self.a / 'guide.pdf').write_bytes(b'%PDF-1.4 original')
        self.run_sync(self.a, 'a')
        blob = next((self.cloud / 'objects').iterdir())
        blob.unlink()
        result = self.run_sync(self.b, 'b')
        self.assertFalse(result['errors'])
        self.assertEqual(result['downloaded'], 1)
        self.assertEqual(blob.read_bytes(), (self.b / 'guide.pdf').read_bytes())

    def test_missing_object_dry_run_does_not_restore_blob(self):
        (self.a / 'guide.pdf').write_bytes(b'%PDF-1.4 original')
        self.run_sync(self.a, 'a')
        blob = next((self.cloud / 'objects').iterdir())
        blob.unlink()
        result = tool.sync(self.b, self.cloud, 'development', 'b', dry_run=True)
        self.assertFalse(result['errors'])
        self.assertFalse(blob.exists())
        self.assertFalse((self.b / 'guide.pdf').exists())

    def test_mismatched_snapshot_is_never_restored(self):
        (self.a / 'guide.pdf').write_bytes(b'%PDF-1.4 original')
        self.run_sync(self.a, 'a')
        blob = next((self.cloud / 'objects').iterdir())
        blob.unlink()
        snapshot = next((self.cloud / 'computers').rglob('guide.pdf'))
        snapshot.write_bytes(b'different content')
        result = self.run_sync(self.b, 'b')
        self.assertEqual(len(result['errors']), 1)
        self.assertFalse(blob.exists())
        self.assertFalse((self.b / 'guide.pdf').exists())


if __name__ == '__main__':
    unittest.main()
