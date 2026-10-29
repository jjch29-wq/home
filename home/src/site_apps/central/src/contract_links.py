"""Read existing central contract files without moving or modifying them."""
import copy
import csv
import json
import re
from collections import defaultdict
from pathlib import Path


def report_key(value):
    return re.sub(r'[/\\]', '-', str(value or '').strip()).casefold()


def method_roots(root, method, field):
    root = Path(root).resolve()
    bases = [root / method]
    if field == '보고서':
        bases.append(root / '보고서' / method)
    elif method == 'PAUT' and field == '데이터':
        bases.append(root / 'PAUT DATA')
    elif method == 'RT':
        bases.append(root / '영상화장비사진' / '중앙지사' / '영상화_결과')
    return bases


def allowed(root, target, method, field):
    target = Path(target).resolve()
    return any(target.is_relative_to(base) for base in method_roots(root, method, field))


def connect_contract(rows, root, issue_path=None):
    root = Path(root).resolve()
    result = copy.deepcopy(rows)

    def relative(path):
        path = path.resolve()
        if not path.is_relative_to(root):
            raise ValueError('계약 폴더 밖의 자료입니다: ' + str(path))
        return path.relative_to(root).as_posix()

    def key(method, date, line, joint):
        return tuple(str(v or '').strip() for v in (method, date, line, joint))

    # Saved report snapshots establish membership; never infer it from a shared line.
    numbers = defaultdict(set)
    if issue_path and Path(issue_path).is_file():
        entries = json.loads(Path(issue_path).read_text(encoding='utf-8-sig'))
        for entry in entries:
            if entry.get('deleted_at') or not entry.get('report_no'):
                continue
            for item in entry.get('items', []):
                line = item.get('Line No.') or item.get('Dwg') or item.get('라인번호')
                joint = item.get('Joint') or item.get('Joint No.') or item.get('용접부번호')
                numbers[key(entry.get('mode'), item.get('Date') or item.get('검사일자'), line, joint)].add(entry['report_no'])

    reports = defaultdict(list)
    for method in ('RT', 'PAUT', 'MT', 'PT'):
        folder = root / '보고서' / method
        if folder.is_dir():
            for path in folder.rglob('*'):
                if path.is_file() and path.suffix.lower() in {'.xlsx', '.xls', '.pdf', '.docx', '.hwp', '.hwpx'} and not path.name.startswith('~$'):
                    reports[(method, report_key(path.stem))].append(path)

    images = defaultdict(list)
    image_root = root / '영상화장비사진' / '중앙지사' / '영상화_결과'
    if image_root.is_dir():
        for log in image_root.glob('영상화_작업기록*.csv'):
            with log.open(encoding='utf-8-sig', newline='') as stream:
                for record in csv.DictReader(stream):
                    target = (image_root / record.get('output_name', '')).resolve()
                    if not target.is_relative_to(image_root) or not target.is_file():
                        continue
                    section = record.get('section', '').strip('() ')
                    if not re.fullmatch(r'\d+-\d+', section):
                        continue
                    k = key('RT', record.get('exam_date'), record.get('line_no'), record.get('joint_no'))
                    images[k].append((section, relative(target), record.get('report_no', '')))

    paut_root = root / 'PAUT DATA'
    paut_files = list(paut_root.rglob('*.opd')) if paut_root.is_dir() else []
    groups = defaultdict(list)
    for row in result:
        groups[key(row.get('보고서종류'), row.get('검사일자'), row.get('라인번호'), row.get('용접부번호'))].append(row)
    inspections = defaultdict(set)
    for k in groups:
        inspections[(k[0], k[2], k[3])].add(k[1])
    for k, group in groups.items():
        if len(group) != 1:
            for row in group:
                row.setdefault('확인사항', []).append('계약 자료 연결 보류: 같은 날짜·라인·JOINT 기록 중복')
            continue
        row = group[0]
        method, date, line, joint = k
        notes = row.setdefault('확인사항', [])
        candidates = numbers.get(k, set())
        if not row.get('보고서번호') and len(candidates) == 1:
            row['보고서번호'] = next(iter(candidates))
        elif not row.get('보고서번호') and len(candidates) > 1:
            notes.append('계약 자료 연결 보류: 발행 보고서 번호가 여러 개입니다.')
        if not row.get('보고서') and row.get('보고서번호'):
            choices = reports.get((method, report_key(row['보고서번호'])), [])
            if len(choices) == 1:
                row['보고서'] = relative(choices[0])
            elif len(choices) > 1:
                notes.append('계약 자료 연결 보류: 같은 번호의 보고서 파일이 여러 개입니다.')
        if method == 'RT' and not row.get('RT영상') and not row.get('필름영상화'):
            matching = images.get(k, [])
            if row.get('보고서번호'):
                matching = [v for v in matching if report_key(v[2]) == report_key(row['보고서번호'])]
            elif len({report_key(v[2]) for v in matching}) > 1:
                matching = []
                notes.append('RT 영상 연결 보류: 발행 보고서 번호 확인 필요')
            mapping = defaultdict(list)
            for section, path, _ in matching:
                if path not in mapping[section]:
                    mapping[section].append(path)
            if mapping:
                row['RT영상'] = dict(mapping)
                row['필름영상화'] = relative(image_root / line)
                row['자동연결영상'] = True
        elif method == 'PAUT' and not row.get('데이터') and line and joint:
            # Exact line folder plus a bounded JOINT token avoids S01/S010/R01R collisions.
            token = re.compile(r'(?:^|_)' + re.escape(joint) + r'(?=_|$)', re.I)
            matches = [p for p in paut_files if line in p.relative_to(paut_root).parts[:-1]
                       and token.search(p.stem)]
            if len(inspections[(method, line, joint)]) > 1:
                notes.append('PAUT 데이터 연결 보류: 동일 라인·JOINT의 검사일이 여러 개입니다.')
            elif matches:
                row['데이터파일'] = [relative(p) for p in sorted(matches)]
                row['데이터'] = row['데이터파일'][0]
        row['확인사항'] = list(dict.fromkeys(notes))
    return result
