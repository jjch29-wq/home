"""Conservative two-way sync of business artifacts via Google Drive Desktop.

No deletions are propagated. Divergent edits are retained for manual review.
Run while business applications are closed, or use --watch between saves.
"""
import argparse
import hashlib
import json
import os
import re
from pathlib import Path
import socket
import subprocess
import tempfile
import time
import zipfile

ROOT = Path(__file__).resolve().parents[1]
EXTENSIONS = {'.json', '.xlsx', '.xls', '.xlsm', '.xltx', '.xltm', '.csv', '.tsv',
              '.db', '.sqlite', '.sqlite3', '.png', '.jpg', '.jpeg', '.bmp', '.gif',
              '.svg', '.ico', '.pdf', '.docx', '.hwp', '.hwpx', '.dxf', '.dwg'}
EXCLUDED = {'.git', '.venv', '.work-sync', '.aws', '.codex', '.agents', 'node_modules',
            '__pycache__', '.pytest_cache', 'archive', 'site_isolation_originals',
            'branch-switch-backup', 'tmp', 'temp', 'build', 'dist'}
SECRET_WORDS = ('credential', 'secret', 'token', 'password', 'api_key', 'apikey', 'private_key')


def digest(path):
    if not path.is_file():
        return None
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write_atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    name = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, delete=False) as stream:
            name = stream.name
            stream.write(data)
        os.replace(name, path)
    finally:
        if name and os.path.exists(name):
            os.unlink(name)


def read_json(path, default):
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else default


def eligible(relative):
    return (not any(part.lower() in EXCLUDED for part in relative.parts)
            and not relative.name.startswith(('~$', '.'))
            and relative.suffix.lower() in EXTENSIONS
            and relative.name not in {'sync_config.json', 'work_data_sync.json'}
            and not any(word in relative.name.lower() for word in SECRET_WORDS))


def inventory(root):
    """Include ignored/untracked artifacts, pruning runtimes and temporary trees."""
    found = set()
    for current, directories, files in os.walk(root, followlinks=False):
        directories[:] = [name for name in directories
                          if name.lower() not in EXCLUDED and not name.startswith('.')
                          and not (Path(current) / name).is_symlink()
                          and not name.startswith(('review_', 'temp_'))]
        for name in files:
            path = Path(current) / name
            relative = path.relative_to(root)
            if not path.is_symlink() and eligible(relative):
                found.add(relative.as_posix())
    return found


def safe_path(root, name):
    relative = Path(name)
    if relative.is_absolute() or '..' in relative.parts or not eligible(relative):
        raise ValueError(f'허용되지 않은 공유 경로: {name}')
    target = root / relative
    if not target.resolve().is_relative_to(root.resolve()):
        raise ValueError(f'공유 경로가 작업 폴더를 벗어납니다: {name}')
    return target


def contains_secret(data, suffix):
    if suffix.lower() != '.json':
        return False
    try:
        value = json.loads(data)
    except (ValueError, UnicodeError):
        return False
    def walk(item):
        if isinstance(item, dict):
            return any((any(word in key.lower() for word in SECRET_WORDS) and bool(val))
                       or walk(val) for key, val in item.items())
        return isinstance(item, list) and any(walk(val) for val in item)
    return walk(value)


def read_object(cloud, name, expected, dry_run=False):
    blob = cloud / 'objects' / expected
    try:
        data = blob.read_bytes()
    except FileNotFoundError:
        # Snapshots preserve the original bytes if a Drive object is missing.
        for computer in (cloud / 'computers').glob('*'):
            snapshot = safe_path(computer, name)
            if snapshot.is_file():
                candidate = snapshot.read_bytes()
                if hashlib.sha256(candidate).hexdigest() == expected:
                    data = candidate
                    if not dry_run:
                        write_atomic(blob, data)
                    break
        else:
            raise FileNotFoundError(f'공유 원본 및 백업을 찾을 수 없습니다: {name} ({expected})')
    if hashlib.sha256(data).hexdigest() != expected:
        raise ValueError(f'공유 파일 검증 실패: {name}')
    return data


def branch_name(root):
    return subprocess.check_output(['git', 'branch', '--show-current'], cwd=root,
                                   encoding='utf-8').strip() or 'detached'


def archive_stashes(root, cloud, machine):
    """Git stash is local-only; preserve business data without applying it."""
    stashes = subprocess.check_output(['git', 'stash', 'list', '--format=%H'], cwd=root,
                                      encoding='ascii').splitlines()
    count = 0
    for commit in stashes:
        target = cloud / 'stash_backups' / hashlib.sha256(machine.encode()).hexdigest() / commit
        if (target / 'complete.json').exists():
            continue
        names = subprocess.check_output(['git', 'diff', '--name-only', '-z', commit + '^1', commit],
                                         cwd=root).decode('utf-8').split('\0')
        saved = []
        for name in filter(None, names):
            if not eligible(Path(name)):
                continue
            result = subprocess.run(['git', 'show', commit + ':' + name], cwd=root,
                                    capture_output=True)
            if result.returncode or contains_secret(result.stdout, Path(name).suffix):
                continue
            write_atomic(safe_path(target, name), result.stdout)
            saved.append(name)
            count += 1
        write_atomic(target / 'complete.json', json.dumps({'commit': commit, 'files': saved},
                     ensure_ascii=False, indent=2).encode('utf-8'))
    return count


def sync(root, cloud, branch, machine, dry_run=False, accept_local=()):
    # Drive Desktop may fail to open object paths containing sync_tools/..
    # even when the same canonical object path exists.
    root = root.resolve()
    cloud = cloud.resolve()
    if not cloud.is_dir():
        raise OSError(f'구글 드라이브 공유 폴더에 연결할 수 없습니다: {cloud}')
    state_path = root / '.work-sync' / (hashlib.sha256(branch.encode()).hexdigest() + '.json')
    state = read_json(state_path, {})
    # One manifest per PC avoids losing another PC's newly uploaded records.
    manifests = list((cloud / 'manifests').glob('*.json'))
    remote = {}
    superseded = {}
    for path in manifests:
        manifest = read_json(path, {})
        for name, hashes in (manifest.get('superseded') or {}).items():
            # A manifest cannot retire the version it currently publishes.
            # Older clients could do this when accepting an unchanged local file.
            superseded.setdefault(name, set()).update(set(hashes) - set(manifest.get('files', {}).get(name, [])))
        for name, hashes in manifest.get('files', {}).items():
            if any(not isinstance(h, str) or not re.fullmatch('[0-9a-f]{64}', h) for h in hashes):
                raise ValueError(f'공유 파일 해시 형식 오류: {name}')
            remote.setdefault(name, set()).update(hashes)
    remote = {name: hashes - superseded.get(name, set()) for name, hashes in remote.items()}
    own_path = cloud / 'manifests' / (hashlib.sha256(machine.encode()).hexdigest() + '.json')
    own_manifest = read_json(own_path, {})
    own = own_manifest.get('files', {})
    retired = {name: set(hashes) for name, hashes in (own_manifest.get('superseded') or {}).items()}
    result = {'uploaded': 0, 'downloaded': 0, 'conflicts': [], 'skipped': [], 'errors': []}
    local_names = inventory(root)
    for name in sorted(local_names | set(remote)):
        try:
            local = safe_path(root, name)
            if name not in local_names:
                candidates = remote.get(name, set())
                if len(candidates) != 1:
                    result['conflicts'].append(name)
                    continue
                incoming = next(iter(candidates))
                data = read_object(cloud, name, incoming, dry_run)
                if hashlib.sha256(data).hexdigest() != incoming:
                    raise ValueError(f'공유 파일 검증 실패: {name}')
                if not dry_run:
                    write_atomic(local, data)
                    own[name] = [incoming]
                    state[name] = incoming
                result['downloaded'] += 1
                continue
            payload = local.read_bytes()
            if payload and local.suffix.lower() == '.json':
                json.loads(payload.decode('utf-8-sig'))
            if local.suffix.lower() in {'.xlsx', '.xlsm', '.docx', '.hwpx', '.xltx', '.xltm'}:
                with zipfile.ZipFile(local) as archive:
                    archive.infolist()
            if local.suffix.lower() in {'.db', '.sqlite', '.sqlite3'} and any(
                    Path(str(local) + suffix).exists() for suffix in ('-wal', '-journal')):
                raise OSError('데이터베이스를 사용 중입니다. 앱을 닫고 다시 동기화해주세요.')
            local_hash = hashlib.sha256(payload).hexdigest()
            if contains_secret(payload, local.suffix):
                result['skipped'].append(name)
                continue
            prior = state.get(name)
            candidates = set(remote.get(name, []))
            if prior:
                candidates.discard(prior)
            candidates.discard(local_hash)
            if name in accept_local or not remote.get(name) or (prior and local_hash != prior and not candidates):
                action = 'upload'
            elif not candidates:
                action = 'same'
            elif prior == local_hash and len(candidates) == 1:
                action = 'download'
            else:
                action = 'conflict'
            if action in ('upload', 'conflict') and not dry_run:
                blob = cloud / 'objects' / local_hash
                if not blob.exists():
                    write_atomic(blob, payload)
                # Preserve original paths for convenient browsing and recovery.
                snapshot = safe_path(cloud / 'computers' / hashlib.sha256(machine.encode()).hexdigest(), name)
                write_atomic(snapshot, payload)
            if action == 'upload':
                result['uploaded'] += 1
                if not dry_run:
                    if prior and prior != local_hash:
                        retired.setdefault(name, set()).add(prior)
                    if name in accept_local:
                        retired.setdefault(name, set()).update(remote.get(name, set()) - {local_hash})
                    own[name] = [local_hash]
                    retired.setdefault(name, set()).discard(local_hash)
                    state[name] = local_hash
            elif action == 'download':
                incoming = next(iter(candidates))
                data = read_object(cloud, name, incoming, dry_run)
                if hashlib.sha256(data).hexdigest() != incoming:
                    raise ValueError(f'공유 파일 검증 실패: {name}')
                if not dry_run:
                    if digest(local) != local_hash:
                        raise OSError(f'동기화 도중 파일이 변경되었습니다: {name}')
                    write_atomic(root / '.work-sync' / 'backups' / local_hash / name, payload)
                    write_atomic(local, data)
                    own[name] = [incoming]
                    retired.setdefault(name, set()).discard(incoming)
                    state[name] = incoming
                result['downloaded'] += 1
            elif action == 'same' and not dry_run:
                own[name] = [local_hash]
                retired.setdefault(name, set()).discard(local_hash)
                state[name] = local_hash
            elif action == 'conflict':
                result['conflicts'].append(name)
                if not dry_run:
                    own[name] = sorted(set(own.get(name, [])) | {local_hash})
                    for incoming in candidates:
                        data = read_object(cloud, name, incoming, dry_run)
                        if hashlib.sha256(data).hexdigest() != incoming:
                            raise ValueError(f'충돌 파일 검증 실패: {name}')
                        write_atomic(root / '.work-sync' / 'conflicts' / incoming / name, data)
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            result['errors'].append({'path': name, 'error': str(exc)})
    if not dry_run:
        manifest = {'files': own, 'superseded': {name: sorted(hashes) for name, hashes in retired.items()}}
        write_atomic(own_path, json.dumps(manifest, ensure_ascii=False, indent=2).encode('utf-8'))
        write_atomic(state_path, json.dumps(state, ensure_ascii=False, indent=2).encode('utf-8'))
        report = dict(result, machine=machine, branch=branch, checked_at=time.strftime('%Y-%m-%d %H:%M:%S'))
        data = json.dumps(report, ensure_ascii=False, indent=2).encode('utf-8')
        write_atomic(root / '.work-sync' / 'latest-report.json', data)
        write_atomic(cloud / 'status' / (hashlib.sha256(machine.encode()).hexdigest() + '.json'), data)
    return result


def main():
    global ROOT
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cloud', type=Path)
    parser.add_argument('--workspace', type=Path)
    parser.add_argument('--watch', action='store_true')
    parser.add_argument('--dry-run', action='store_true')
    parser.add_argument('--accept-local', action='append', default=[],
                        help='검토·병합을 마친 특정 상대 경로를 공유본으로 확정')
    args = parser.parse_args()
    if args.workspace:
        ROOT = args.workspace.resolve()
    config = read_json(ROOT / 'work_data_sync.json', {})
    cloud = args.cloud or Path(config['cloud_dir'])
    if not cloud.is_dir():
        # Accommodate a different Drive Desktop drive letter on the other PC.
        for letter in 'GHIJKLFMN':
            alternate = Path(letter + ':' + str(cloud)[2:])
            if alternate.is_dir():
                cloud = alternate
                break
    while True:
        try:
            result = sync(ROOT, cloud, branch_name(ROOT), socket.gethostname(), args.dry_run,
                          args.accept_local)
            if not args.dry_run:
                result['stash_backups'] = archive_stashes(ROOT, cloud, socket.gethostname())
            print(json.dumps(result, ensure_ascii=False), flush=True)
            failed = bool(result['conflicts'] or result['errors'])
        except (OSError, ValueError, KeyError) as exc:
            print(f'동기화 실패: {exc}', flush=True)
            failed = True
        if not args.watch:
            return 1 if failed else 0
        time.sleep(30)


if __name__ == '__main__':
    raise SystemExit(main())
