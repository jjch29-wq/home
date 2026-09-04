"""One-time, non-destructive migration of the three site applications.

Run from any directory. Existing destinations are never overwritten.
The manifest records original hashes and migration copies for audit/rollback.
"""
import ast
import hashlib
import json
from pathlib import Path
import shutil

HOME = Path(__file__).resolve().parents[1]
SRC = HOME / 'src'
ROOT = SRC / 'site_apps'
SITES = {
    'central': ('한국지역난방 중앙지사.py', 'Material_Inventory.xlsx', 'Material_Manager_Config.json', 'daily_work_history.json'),
    'kogas': ('가스공사 가산~가평.py', 'Kogas_Material_Inventory.xlsx', 'Kogas_Material_Manager_Config.json', 'kogas_daily_work_history.json'),
    'lotte': ('롯데건설 바이오로직스.py', 'Lotte_Material_Inventory.xlsx', 'Lotte_Material_Manager_Config.json', 'lotte_daily_work_history.json'),
}

def resolve(name):
    path = SRC.joinpath(*name.split('.'))
    if path.with_suffix('.py').is_file():
        return path.with_suffix('.py')
    if (path / '__init__.py').is_file():
        return path / '__init__.py'

def dependencies(entry):
    pending = [SRC / entry, SRC / '문서_통합_관리_허브.py']
    found = set()
    while pending:
        path = pending.pop()
        if path in found:
            continue
        found.add(path)
        for node in ast.walk(ast.parse(path.read_text(encoding='utf-8-sig'))):
            names = []
            if isinstance(node, ast.Import):
                names = [n.name for n in node.names]
            elif isinstance(node, ast.ImportFrom) and not node.level:
                names = [node.module] + [f'{node.module}.{n.name}' for n in node.names]
            for name in names:
                target = resolve(name)
                if target:
                    pending.append(target)
                    for parent in target.parents:
                        if parent == SRC:
                            break
                        if (parent / '__init__.py').is_file():
                            pending.append(parent / '__init__.py')
    return found

def qualify_imports(source, prefix):
    """Change only local import statements, retaining all other source text."""
    lines = source.splitlines(keepends=True)
    offsets, offset = [], 0
    for line in lines:
        offsets.append(offset)
        offset += len(line.encode('utf-8'))
    edits = []
    for node in ast.walk(ast.parse(source)):
        replacement = None
        if isinstance(node, ast.ImportFrom) and not node.level:
            if resolve(node.module) or (SRC / node.module.replace('.', '/')).is_dir():
                names = ', '.join(n.name + (f' as {n.asname}' if n.asname else '') for n in node.names)
                replacement = f'from {prefix}.{node.module} import {names}'
        elif isinstance(node, ast.Import) and any(resolve(n.name) for n in node.names):
            parts = []
            for name in node.names:
                if resolve(name.name):
                    if '.' in name.name and not name.asname:
                        raise ValueError(f'Unaliased dotted local import: {name.name}')
                    parts.append(f'import {prefix}.{name.name} as {name.asname or name.name}')
                else:
                    parts.append('import ' + name.name + (f' as {name.asname}' if name.asname else ''))
            replacement = '; '.join(parts)
        if replacement:
            edits.append((offsets[node.lineno - 1] + node.col_offset,
                          offsets[node.end_lineno - 1] + node.end_col_offset,
                          replacement.encode('utf-8')))
    raw = source.encode('utf-8')
    for start, end, replacement in sorted(edits, reverse=True):
        raw = raw[:start] + replacement + raw[end:]
    return raw.decode('utf-8')

def transform(source, site, relative):
    source = qualify_imports(source, f'site_apps.{site}.src')
    entry, db, config, history = SITES[site]
    # Lotte's copied shared widgets must use Lotte's inventory and history.
    if relative.name in ('daily_work_log_tab.py', 'ndt_billing_tab.py', '문서_통합_관리_허브.py'):
        source = source.replace("'Material_Inventory.xlsx'", repr(db))
    if relative.name == '지역난방_안전관리교육.py':
        source = source.replace('CONFIG_FILE = "config_district_heating_safety_training.json"',
            'CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_district_heating_safety_training.json")')
    if relative.name in ('monthly_report_manager.py', 'kogas_monthly_report_manager.py'):
        source = source.replace("            if r'c:\\Users\\-\\PMI\\home\\src' not in sys.path:\n                sys.path.append(r'c:\\Users\\-\\PMI\\home\\src')\n", '')
        source = source.replace('f"c:\\\\Users\\\\-\\\\PMI\\\\home\\\\src\\\\지역난방_안전관리_{year}{month:02d}*.xlsx"',
            'os.path.join(os.path.dirname(os.path.dirname(__file__)), f"지역난방_안전관리_{year}{month:02d}*.xlsx")')
    # Development-only examples also stay within the owning site.
    for old in ('c:\\\\Users\\\\jjch2\\\\Desktop\\\\PMI\\\\home\\\\src\\\\daily_work_history.json',):
        source = source.replace('"' + old + '"', f'os.path.join(os.path.dirname(__file__), {history!r})')
        source = source.replace("'" + old + "'", f'os.path.join(os.path.dirname(__file__), {history!r})')
    if relative.name == entry:
        source = source.replace("'Documents', 'MaterialManager'", f"'Documents', 'MaterialManager', {site!r}")
    compile(source, str(relative), 'exec')
    return source

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def migrate():
    if ROOT.exists():
        raise SystemExit('site_apps already exists; refusing to overwrite site code or data.')
    # Parse every dependency before making any changes.
    graph = {site: dependencies(values[0]) for site, values in SITES.items()}
    transformed = {(site, path): transform(path.read_text(encoding='utf-8-sig'), site, path.relative_to(SRC))
                   for site, paths in graph.items() for path in paths}
    backup = HOME / 'archive' / 'site_isolation_originals'
    if backup.exists():
        raise SystemExit('Migration backup already exists; refusing to overwrite it.')
    backup.mkdir(parents=True)
    manifest = {'version': 1, 'code': [], 'copies': [], 'launchers': {}}

    def copy(source, target):
        if not source.is_file():
            return
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open('xb') as output:
            output.write(source.read_bytes())
        manifest['copies'].append({'source': str(source.relative_to(HOME)),
                                   'target': str(target.relative_to(HOME)), 'sha256': digest(source)})

    ROOT.mkdir(parents=True)
    (ROOT / '__init__.py').write_text('"""Independent site applications; no business logic is shared."""\n', encoding='utf-8')
    for site, (entry, db, config, history) in SITES.items():
        site_root = ROOT / site
        site_src = site_root / 'src'
        site_src.mkdir(parents=True)
        for path in sorted(graph[site]):
            relative = path.relative_to(SRC)
            target = site_src / ('app.py' if path.name == entry else relative)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(transformed[site, path], encoding='utf-8', newline='\n')
            manifest['code'].append({'source': str(path.relative_to(HOME)),
                'target': str(target.relative_to(HOME)), 'original_sha256': digest(path), 'sha256': digest(target)})
        for folder in [site_root, site_src] + [p for p in site_src.rglob('*') if p.is_dir()]:
            init = folder / '__init__.py'
            if not init.exists():
                init.write_text('"""Site-owned module package."""\n', encoding='utf-8')
        # Keep the old src/data/resources relationship, but under each site.
        for filename in {config, history, 'daily_work_history.json', 'config.json', 'kogas_config.json',
                         'codebook_db.json', 'config_district_heating_safety_training.json'}:
            copy(SRC / filename, site_src / filename)
        for folder in ('resources', 'src/signs', 'src/templates'):
            origin = HOME / folder
            if origin.is_dir():
                for path in origin.rglob('*'):
                    if path.is_file() and '__pycache__' not in path.parts:
                        copy(path, site_root / folder / path.relative_to(origin))
        # Use exactly the inventory that the old entry point would have selected.
        inventory = next((p for p in (SRC / db, HOME / 'data' / db) if p.is_file()), None)
        (site_root / 'data').mkdir(exist_ok=True)
        if inventory:
            copy(inventory, site_root / 'data' / db)
        photos = HOME / 'data' / 'process_photos'
        if photos.is_dir():
            for path in photos.rglob('*'):
                if path.is_file():
                    copy(path, site_root / 'data' / 'process_photos' / path.relative_to(photos))
        copy(SRC / entry, backup / entry)
        manifest['launchers'][entry] = digest(SRC / entry)
    # All copies must be complete and verified before switching entry points.
    for record in manifest['copies']:
        assert digest(HOME / record['target']) == record['sha256']
    (ROOT / 'migration_manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    for site, (entry, *_rest) in SITES.items():
        (SRC / entry).write_text(
            f'"""Compatibility launcher. Edit site_apps/{site}/src for this site."""\n'
            'from pathlib import Path\nimport sys\n\n'
            '_src = str(Path(__file__).resolve().parent)\n'
            'if _src not in sys.path:\n    sys.path.insert(0, _src)\n\n'
            f'from site_apps.{site}.src.app import MaterialManager\n\n'
            'def main():\n    import tkinter as tk\n    root = tk.Tk()\n'
            '    app = MaterialManager(root)\n    try:\n        root.mainloop()\n'
            '    except KeyboardInterrupt:\n        try:\n            root.destroy()\n'
            '        except Exception:\n            pass\n\n'
            'if __name__ == "__main__":\n    main()\n', encoding='utf-8')
    print(f'Isolated {len(SITES)} sites, {len(manifest["code"])} modules, {len(manifest["copies"])} preserved copies.')

if __name__ == '__main__':
    migrate()
