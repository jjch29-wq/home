"""Read existing central contract files without moving or modifying them."""
import copy
import csv
import json
import re
import zipfile
from collections import defaultdict
from pathlib import Path

DEFECT_CODE_MAP = {
    'P': 'P',       # 기공 (Porosity)
    'SI': 'SI',     # 슬래그 혼입 (Slag Inclusion)
    'CL': 'SI',     # 클러스터/슬래그
    '이물': 'SI',
    '이물질': 'SI',
    'SLAG': 'SI',
    'LF': 'LF',     # 융합 불량 (Lack of Fusion)
    'IP': 'IP',     # 용입 부족 (Incompleted Penetration)
    'LP': 'IP',
    'RC': 'RC',     # 루트 요면 (Root Concavity)
    'EC': 'EC',     # 외부 요면 (External Concavity)
    'EP': 'EP',     # 과도 용입 (Excessive Penetration)
    'MI': 'MI',     # 단차 (Misalignment)
    'TI': 'TI',     # 텅스텐 혼입 (Tungsten Inclusion)
    'BT': 'BT',     # 용락 (Burn Through)
    'IUC': 'IUC',   # 루트언더컷 (Internal Under Cut)
    'EUC': 'EUC',   # 외부언더컷 (External Under Cut)
    'UC': 'UC',     # 언더컷 (Under Cut)
    'AS': 'AS',     # 아크 스트라이크 (Arc Strike)
    'OL': 'OL',     # 오버랩 (Over Lap)
    'C': 'C',       # 균열 (Crack)
    'CR': 'C',
    'CRACK': 'C',
    'G': 'G',       # 사상불량 (Grinding)
    'M': 'M',       # 모재결함 (Material)
}

_RT_REPORT_DEFECT_CACHE = {}


def parse_defects(raw_value):
    """Parse defect codes from interpretation text or list of point readings per [별표 3]."""
    if not raw_value:
        return ''
    if isinstance(raw_value, (list, tuple, set)):
        items = list(raw_value)
    else:
        items = re.split(r'[,;]+', str(raw_value).strip())
    codes = []
    for item in items:
        item = str(item).strip()
        if not item:
            continue
        parts = item.split('/')
        code_part = parts[0].strip()
        if not code_part or code_part.upper().startswith('N') or code_part.upper() in ('OK', 'ACC', 'N/A'):
            continue
        mapped = DEFECT_CODE_MAP.get(code_part.upper(), DEFECT_CODE_MAP.get(code_part, code_part.upper()))
        if mapped and mapped not in codes:
            codes.append(mapped)
    return '/'.join(codes)


def extract_rt_defects_from_workbook(report_path):
    p = Path(report_path).resolve()
    if p in _RT_REPORT_DEFECT_CACHE:
        return _RT_REPORT_DEFECT_CACHE[p]
    mapping = {}
    if p.suffix.lower() not in ('.xlsx', '.xlsm') or not p.is_file():
        _RT_REPORT_DEFECT_CACHE[p] = mapping
        return mapping
    try:
        from openpyxl import load_workbook
        wb = load_workbook(p, data_only=True, read_only=True)
        try:
            for sname in wb.sheetnames:
                if sname.startswith('R1'):
                    continue
                ws = wb[sname]
                sheet_line = ''
                point_cols = {}
                for row in ws.iter_rows(values_only=True):
                    if not row:
                        continue
                    for cell in row:
                        if cell and ('JA2026' in str(cell) or 'Sec.' in str(cell)):
                            m = re.search(r'(JA2026[0-9A-Za-z_-]+)', str(cell))
                            if m:
                                sheet_line = m.group(1).strip()
                    for c_idx, cell in enumerate(row):
                        if cell and re.fullmatch(r'\d+\s*-\s*\d+', str(cell).strip()):
                            point_cols[c_idx] = str(cell).strip().replace(' ', '')
                    if row[0]:
                        val = str(row[0]).strip()
                        if not val or any(k in val for k in ('JOINT', 'REPORT', 'TOTAL', '총', '~')):
                            continue
                        joint_no = val
                        interps = []
                        if point_cols:
                            for c_idx in point_cols:
                                if c_idx < len(row) and row[c_idx] is not None:
                                    s = str(row[c_idx]).strip()
                                    if s:
                                        interps.append(s)
                        else:
                            for cell in row[3:12]:
                                if cell is not None and ('/' in str(cell) or str(cell).strip() in ('N', 'P', 'IP', 'LF', 'UC', 'SI', '이물질')):
                                    interps.append(str(cell).strip())
                        code = parse_defects(interps)
                        if sheet_line:
                            mapping[(sheet_line.casefold(), joint_no.casefold())] = code
                        mapping[('', joint_no.casefold())] = code
        finally:
            wb.close()
    except Exception:
        pass
    _RT_REPORT_DEFECT_CACHE[p] = mapping
    return mapping


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
        bases.append(root / 'RT영상화')
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
    image_root = root / 'RT영상화'
    if not image_root.is_dir():
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
    file_token_cache = {}
    def get_report_text(path):
        if path not in file_token_cache:
            text = ''
            if path.suffix.lower() in ('.xlsx', '.xlsm'):
                try:
                    with zipfile.ZipFile(path, 'r') as z:
                        for name in z.namelist():
                            if name == 'xl/sharedStrings.xml' or (name.startswith('xl/worksheets/') and name.endswith('.xml')):
                                text += ' ' + z.read(name).decode('utf-8', errors='ignore')
                except Exception:
                    pass
            elif path.suffix.lower() in ('.pdf', '.hwp', '.hwpx', '.docx', '.txt', '.csv'):
                try:
                    raw = path.read_bytes()
                    text = raw.decode('utf-8', errors='ignore') + ' ' + raw.decode('euc-kr', errors='ignore')
                except Exception:
                    pass
            file_token_cache[path] = text
        return file_token_cache[path]

    def report_matches(path, *tokens):
        content = get_report_text(path)
        return all(str(t).strip() in content for t in tokens if str(t).strip())

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
            verified = []
            for candidate in candidates:
                cand_reports = reports.get((method, report_key(candidate)), [])
                for rep_path in cand_reports:
                    if joint and report_matches(rep_path, line, joint):
                        verified.append((candidate, rep_path))
                        break
                    elif report_matches(rep_path, line):
                        verified.append((candidate, rep_path))
                        break
            if len(verified) == 1:
                row['보고서번호'] = verified[0][0]
                row['보고서'] = relative(verified[0][1])
            elif len(verified) > 1:
                notes.append('계약 자료 연결 보류: 발행 보고서 번호가 여러 개입니다.')
            else:
                notes.append('계약 자료 연결 보류: 발행 보고서 번호가 여러 개입니다.')
        if not row.get('보고서') and row.get('보고서번호'):
            choices = reports.get((method, report_key(row['보고서번호'])), [])
            if len(choices) == 1:
                row['보고서'] = relative(choices[0])
            elif len(choices) > 1:
                notes.append('계약 자료 연결 보류: 같은 번호의 보고서 파일이 여러 개입니다.')
        if method == 'RT' and not row.get('결함종류'):
            if row.get('보고서'):
                try:
                    rep_file = (root / row['보고서']).resolve()
                    if rep_file.is_file():
                        defects = extract_rt_defects_from_workbook(rep_file)
                        df = defects.get((line.casefold(), joint.casefold())) or defects.get(('', joint.casefold()))
                        if df:
                            row['결함종류'] = df
                except Exception:
                    pass
            if not row.get('결함종류'):
                raw_sec = row.get('구간정보') or (row.get('원본값') or {}).get('구간정보')
                df = parse_defects(raw_sec)
                if df:
                    row['결함종류'] = df
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
            clean_line = line.strip().rstrip('_')
            matches = [p for p in paut_files if (any(clean_line in part for part in p.relative_to(paut_root).parts[:-1]) or clean_line in p.stem)
                       and token.search(p.stem)]
            if len(inspections[(method, line, joint)]) > 1:
                notes.append('PAUT 데이터 연결 보류: 동일 라인·JOINT의 검사일이 여러 개입니다.')
            elif matches:
                row['데이터파일'] = [relative(p) for p in sorted(matches)]
                row['데이터'] = row['데이터파일'][0]
        if row.get('보고서'):
            notes = [n for n in notes if '보고서 후보' not in n and '보고서 번호가 여러 개' not in n and '보고서 파일이 여러 개' not in n]
        if row.get('RT영상') or row.get('필름영상화'):
            notes = [n for n in notes if 'RT 영상 연결 보류' not in n]
        if row.get('데이터'):
            notes = [n for n in notes if 'PAUT 데이터 연결 보류' not in n]
        row['확인사항'] = list(dict.fromkeys(notes))
    return result
