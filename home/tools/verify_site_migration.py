"""Audit the initial migration; run before making further business-code changes."""
import importlib.util
import json
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('site_migration', TOOLS / 'isolate_sites.py')
migration = importlib.util.module_from_spec(spec)
spec.loader.exec_module(migration)
HOME = migration.HOME
ROOT = migration.ROOT


def verify():
    manifest = json.loads((ROOT / 'migration_manifest.json').read_text(encoding='utf-8'))
    for record in manifest['copies']:
        if Path(record['source']).name not in manifest['launchers']:
            assert migration.digest(HOME / record['source']) == record['sha256'], record['source']
        assert migration.digest(HOME / record['target']) == record['sha256'], record['target']
    for record in manifest['code']:
        original = HOME / record['source']
        if original.name in manifest['launchers']:
            original = HOME / 'archive' / 'site_isolation_originals' / original.name
        assert migration.digest(original) == record['original_sha256'], str(original)
        target = HOME / record['target']
        site = target.relative_to(ROOT).parts[0]
        expected = migration.transform(original.read_text(encoding='utf-8-sig'), site,
                                       Path(record['source']).relative_to('src'))
        assert target.read_text(encoding='utf-8').rstrip() == expected.rstrip(), str(target)
        assert migration.digest(target) == record['sha256'], str(target)
    print(f'Audit passed: {len(manifest["code"])} code files and {len(manifest["copies"])} copies; originals unchanged.')


if __name__ == '__main__':
    verify()
