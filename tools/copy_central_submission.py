"""Copy central contract deliverables; preserve originals and verify every copy."""
import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import shutil


PARTS = ('보고서', 'PAUT DATA', '영상화장비사진/중앙지사/영상화_결과')


def submission_path(value):
    prefix = PARTS[2]
    return 'RT영상화' + value[len(prefix):] if value == prefix or value.startswith(prefix + '/') else value


def remap_connections(value):
    if isinstance(value, str):
        return submission_path(value)
    if isinstance(value, list):
        return [remap_connections(item) for item in value]
    if isinstance(value, dict):
        return {key: remap_connections(item) for key, item in value.items()}
    return value


def digest(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    args = parser.parse_args()
    source, destination = args.source.resolve(), args.destination.resolve()
    if source == destination or destination.is_relative_to(source) or source.is_relative_to(destination):
        raise ValueError('Source and destination must be separate folders')
    stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    plan = []
    for part in PARTS:
        folder = source / part
        if not folder.is_dir():
            raise FileNotFoundError(folder)
        for path in sorted(folder.rglob('*')):
            if not path.is_file() or path.name.startswith('~$'):
                continue
            if not path.resolve().is_relative_to(source):
                raise ValueError('Source escapes the contract folder')
            relative = Path(submission_path(path.relative_to(source).as_posix()))
            target = (destination / relative).resolve()
            if not target.is_relative_to(destination):
                raise ValueError('Target escapes the submission folder')
            plan.append((path, target, relative))
    # Preflight all collisions before starting. Never overwrite a differing file.
    for path, target, _ in plan:
        if target.exists() and (not target.is_file() or digest(path) != digest(target)):
            raise FileExistsError('Different existing submission file: ' + str(target))
    destination.mkdir(parents=True, exist_ok=True)
    manifest = []
    for index, (path, target, relative) in enumerate(plan, 1):
        expected = digest(path)
        copied = not target.exists()
        if copied:
            target.parent.mkdir(parents=True, exist_ok=True)
            temporary = target.with_name(target.name + '.copying-' + stamp)
            if temporary.exists():
                raise FileExistsError(temporary)
            try:
                shutil.copy2(path, temporary)
                if digest(temporary) != expected or digest(path) != expected:
                    raise OSError('Copy verification failed or source changed: ' + str(path))
                if target.exists():
                    raise FileExistsError(target)
                temporary.rename(target)
            finally:
                temporary.unlink(missing_ok=True)
        elif digest(target) != expected:
            raise OSError('Existing file changed during copy: ' + str(target))
        manifest.append(dict(path=relative.as_posix(), bytes=path.stat().st_size,
                             sha256=expected, copied=copied))
        if index % 20 == 0 or index == len(plan):
            print(f'Verified {index}/{len(plan)} files', flush=True)
    source_cache = source / '.central_attribute_connections.json'
    if source_cache.is_file():
        data = remap_connections(json.loads(source_cache.read_text(encoding='utf-8-sig')))
        # All file links remain relative to the same three preserved subtrees.
        for row in [*data.get('rows', []), *data.get('unresolved', [])]:
            links = [row.get(field) for field in ('보고서', '필름영상화', '데이터')]
            links += list(row.get('데이터파일', []))
            links += [value for values in (row.get('RT영상') or {}).values()
                      for value in ([values] if isinstance(values, str) else values)]
            for value in filter(None, links):
                if Path(value).is_absolute() or not (destination / value).resolve().is_relative_to(destination):
                    raise ValueError('Connection is not submission-relative: ' + value)
                if not (destination / value).exists():
                    raise FileNotFoundError('Copied connection target missing: ' + value)
        cache = destination / source_cache.name
        if cache.exists():
            backup = cache.with_name(cache.name + '.backup-' + stamp)
            shutil.copy2(cache, backup)
        temporary = cache.with_name(cache.name + '.tmp-' + stamp)
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(cache)
    report = destination / ('자료복사_검증기록_' + stamp + '.json')
    report.write_text(json.dumps(dict(source=str(source), destination=str(destination),
                                     files=len(manifest), bytes=sum(r['bytes'] for r in manifest),
                                     verified=manifest), ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'COMPLETE: {len(manifest)} files verified. Manifest: {report}', flush=True)


if __name__ == '__main__':
    main()
