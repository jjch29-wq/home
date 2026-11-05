"""Central branch submission metadata; source records and files are read only.

Run through 중앙지사_속성정보_생성.py. No changes to either business app.
"""
from __future__ import annotations

import datetime as dt
import copy
import hashlib
import json
import os
import re
import tempfile
from pathlib import Path

METHODS = ('RT', 'PAUT', 'MT', 'PT')
FIELDS = ('공사명', '시공사', '하도사', '용역사', '도면번호', '라인번호',
          '용접부번호', '용접부특성', '검사일자', '결함종류', '보고서종류', '보고서 매수')
LINKS = ('보고서', '필름영상화', '데이터')
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.dcm', '.dicom', '.pdf'}
POINT_PATTERN = re.compile(r'(?<![0-9A-Za-z])(\d+\s*-\s*\d+)(?![0-9A-Za-z])')
CACHE_NAME = '.central_attribute_connections.json'
REPORT_EXTENSIONS = {'.pdf', '.xlsx', '.xls', '.xlsm', '.doc', '.docx', '.hwp', '.hwpx'}
SUPPLEMENT_FIELDS = (*FIELDS, *LINKS, '검사결과', '보고서번호', 'RT영상', '촬영포인트', '검사차수', '자동연결영상', '데이터파일')


def parse_points(value):
    values = value if isinstance(value, list) else re.split(r'[,;/\n]+', text(value))
    values = [text(v).replace(' ', '') for v in values if text(v)]
    return list(dict.fromkeys(values)) if values and all(re.fullmatch(r'\d+-\d+', v) for v in values) else []


def positive_quantity(value):
    try:
        return float(text(value).replace(',', '')) > 0
    except ValueError:
        return False


def text(value):
    return '' if value is None else str(value).strip()


def first(record, *keys):
    return next((text(record.get(k)) for k in keys if text(record.get(k))), '')


def date_text(value):
    if isinstance(value, (dt.date, dt.datetime)):
        return value.strftime('%Y-%m-%d')
    value = text(value)
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d', '%Y-%m-%d %H:%M:%S'):
        try:
            return dt.datetime.strptime(value, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return value


def normalize(record, method='', date='', source=''):
    row = {key: text(record.get(key)) for key in FIELDS}
    defect_val = first(record, '결함종류', 'Type of Flaw')
    if not defect_val and (first(record, '보고서종류', '검사방법') or method).upper() == 'RT':
        from site_apps.central.src.contract_links import parse_defects
        defect_val = parse_defects(record.get('구간정보'))
    row.update({
        '검사일자': date_text(first(record, '검사일자', 'Date') or date),
        '보고서종류': first(record, '보고서종류', '검사방법') or method,
        '라인번호': first(record, '라인번호', 'Line No.', 'Line'),
        '도면번호': first(record, '도면번호', 'Dwg', 'ISO'),
        '용접부번호': first(record, '용접부번호', 'Joint No.', 'Joint No', 'Joint'),
        '결함종류': defect_val,
        '용접부특성': first(record, '용접부특성'),
        '출처': source,
        '확인사항': [],
        '검사결과': first(record, '검사결과', '결과', 'Evaluation', 'Result'),
        '보고서번호': first(record, '보고서번호', '보고서 번호', 'Report No.', 'Report No', 'ReportNo', 'report_no'),
        'RT영상': record.get('RT영상', {}),
        '촬영포인트': parse_points(record.get('촬영포인트') or first(record, '구간정보', 'Loc')) if (first(record, '보고서종류', '검사방법') or method).upper() == 'RT' else [],
        '검사차수': first(record, '검사차수', '재검사차수'),
    })
    if not row['용역사']:
        row['용역사'] = first(record, '용역사') or '서울검사'
    if not row['시공사'] and record.get('업체'):
        row['시공사'] = text(record['업체'])
    if not row['공사명']:
        row['공사명'] = first(record, '공사명', '공사', 'Project') or '2026년 중앙지사 열수송관 비파괴검사 단가계약'
    joint_no = text(row.get('용접부번호'))
    if re.search(r'\d+R$', joint_no, re.I):
        row['용접부특성'] = 'REPAIR'
    elif not row['용접부특성'] and (positive_quantity(record.get('RT_OR')) or positive_quantity(record.get('RT_RE'))):
        ori = positive_quantity(record.get('RT_OR'))
        if ori:
            row['용접부특성'] = 'O'
        else:
            row['확인사항'].append('RT_RE는 불량매수: 전체 RT_OR 및 용접부특성 확인')
    for key in LINKS:
        row[key] = first(record, key + '경로', key + '파일')
    return row


def identity(row):
    # Do not equate drawings and line numbers, or remove identifier punctuation.
    return tuple(text(row.get(k)).upper() for k in
                 ('검사일자', '보고서종류', '라인번호', '용접부번호', '용접부특성'))


def load_history(path):
    with open(path, encoding='utf-8-sig') as stream:
        history = json.load(stream)
    if not isinstance(history, dict):
        raise ValueError('중앙지사 이력은 날짜별 객체 형식이어야 합니다.')
    if 'version' in history and 'rows' in history:
        raise ValueError('선택한 파일은 검사 이력이 아닌 속성정보 연결정보입니다. 검사 이력에는 daily_work_history.json을 선택하세요.')
    rows = []
    for date, day in history.items():
        if not isinstance(day, dict):
            raise ValueError(f'잘못된 일자 데이터: {date}')
        for index, item in enumerate(day.get('ndt_results', [])):
            if not isinstance(item, dict):
                raise ValueError(f'{date} 검사 기록 형식 오류')
            if text(item.get('검사방법')).upper() in METHODS:
                row = normalize(item, date=date, source=f'중앙지사 {date} #{index + 1}')
                row['보고서종류'] = row['보고서종류'].upper()
                row['원본값'] = {key: copy.deepcopy(row.get(key, '')) for key in SUPPLEMENT_FIELDS}
                row['id'] = hashlib.sha256((str(Path(path).resolve()) + '|' + date + '|' + str(index)).encode()).hexdigest()[:20]
                rows.append(row)
    return rows


def load_report_data(path, method):
    if method not in METHODS:
        raise ValueError('RT, PAUT, MT, PT 중 검사종류를 선택하세요.')
    if Path(path).suffix.lower() == '.json':
        with open(path, encoding='utf-8-sig') as stream:
            data = json.load(stream)
    else:
        from openpyxl import load_workbook
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = book.active
            values = iter(sheet.values)
            headers = [text(v) for v in next(values)]
            data = [dict(zip(headers, values)) for values in values if any(v is not None for v in values)]
        finally:
            book.close()
    if not isinstance(data, list) or any(not isinstance(item, dict) for item in data):
        raise ValueError('보고서 앱의 데이터 저장(JSON) 또는 표 형태의 속성정보 XLSX를 선택하세요.')
    rows = []
    for index, item in enumerate(data):
        if item.get('selected') is False or item.get('date_filtered') is False:
            continue
        row = normalize(item, method=method, source=f'{Path(path).name} #{index + 1}')
        if not row['용접부번호']:
            raise ValueError(f'{index + 1}번째 보고서 데이터에 용접부번호가 없습니다.')
        if row['보고서종류'].upper() != method:
            raise ValueError(f'{index + 1}번째 데이터의 검사종류가 선택한 {method}와 다릅니다.')
        row['보고서종류'] = method
        row['id'] = hashlib.sha256((str(Path(path).resolve()) + '|' + str(index)).encode()).hexdigest()[:20]
        rows.append(row)
    return rows


def merge_rows(history, reports, unmatched=None):
    """Central history owns the row list; unmatched reports are diagnostics only."""
    rows = copy.deepcopy(history)
    groups = {}
    for row in rows:
        groups.setdefault(identity(row), []).append(row)
    for group in groups.values():
        if len(group) > 1:
            for row in group:
                row['확인사항'].append('중앙지사 대조 키 중복: 자료 자동 병합·연결 복원 제외')
    report_groups = {}
    for row in reports:
        report_groups.setdefault(identity(row), []).append(row)
    for key, group in report_groups.items():
        matches = groups.get(key, [])
        complete = all(key[:4])  # Empty weld characteristic can match empty only.
        if complete and len(matches) == len(group) == 1:
            target, report = matches[0], group[0]
            for field in (*FIELDS, *LINKS, '검사결과', '보고서번호'):
                new, old = text(report.get(field)), text(target.get(field))
                if new and old and new != old:
                    target['확인사항'].append(f'{field} 불일치: 중앙지사={old}, 보고서={new}')
                elif new:
                    target[field] = new
            target['출처'] += ' / ' + report['출처']
            target['확인사항'].extend(report['확인사항'])
            if report.get('RT영상'):
                if target.get('RT영상') and target['RT영상'] != report['RT영상']:
                    target['확인사항'].append('RT 포인트 영상 연결 불일치')
                else:
                    target['RT영상'] = report['RT영상']
            for field in ('촬영포인트', '검사차수'):
                if not target.get(field) and report.get(field):
                    target[field] = copy.deepcopy(report[field])
                elif report.get(field) and target.get(field) != report[field]:
                    target['확인사항'].append(field + ' 보고서 불일치')
        else:
            reason = '대조 키 누락' if not complete else ('중복 후보' if matches else '중앙지사 일치 기록 없음')
            for report in group:
                if unmatched is not None:
                    unmatched.append({**copy.deepcopy(report), '대조결과': reason})
            for target in matches:
                target['확인사항'].append('보고서 대조 확인: ' + reason)
    return rows


def restore_connections(current, saved):
    """Restore supplements only to a unique unchanged source identity.

    Source-filled values always win. Old rows never become current inspection rows.
    Duplicate or incomplete identities cannot safely recover file attachments.
    """
    result = copy.deepcopy(current)
    current_groups, saved_groups = {}, {}
    for row in result:
        current_groups.setdefault(identity(row.get('원본값', row)), []).append(row)
    for row in saved:
        saved_groups.setdefault(identity(row.get('원본값', row)), []).append(row)
    for key, group in current_groups.items():
        previous = saved_groups.get(key, [])
        if not all(key[:4]) or len(group) != 1 or len(previous) != 1:
            continue
        target, old = group[0], previous[0]
        # Saved warnings remain until their underlying discrepancy is reviewed.
        target['확인사항'] = list(dict.fromkeys([*target.get('확인사항', []), *old.get('확인사항', [])]))
        source = target.get('원본값', target)
        previous_source = old.get('원본값', {})
        for field in ('검사결과', '도면번호', '촬영포인트', '검사차수'):
            if field in previous_source and previous_source[field] != source.get(field):
                target['확인사항'].append(f'중앙지사 {field} 변경: 보고서/영상 연결 재확인')
        for field in SUPPLEMENT_FIELDS:
            if not source.get(field) and old.get(field):
                target[field] = copy.deepcopy(old[field])
        if target.get('보고서'):
            target['확인사항'] = [n for n in target['확인사항'] if '보고서 후보' not in n and '보고서 번호가 여러 개' not in n and '보고서 파일이 여러 개' not in n]
        if target.get('RT영상') or target.get('필름영상화'):
            target['확인사항'] = [n for n in target['확인사항'] if 'RT 영상 연결 보류' not in n]
        if target.get('데이터'):
            target['확인사항'] = [n for n in target['확인사항'] if 'PAUT 데이터 연결 보류' not in n]
    return result


def read_connections(path):
    with open(path, encoding='utf-8-sig') as stream:
        data = json.load(stream)
    rows = data.get('rows') if isinstance(data, dict) and data.get('version') in (1, 2, 3) else None
    if isinstance(rows, list):
        pending = data.get('unresolved', [])
        if not isinstance(pending, list):
            raise ValueError('연결정보 복원보류 목록 형식 오류')
        rows = rows + pending
    if not isinstance(rows, list) or any(not isinstance(r, dict) or not all(k in r for k in (*FIELDS, 'id', '확인사항')) for r in rows):
        raise ValueError('올바른 속성정보 연결정보 파일이 아닙니다.')
    return rows


def write_connections(path, rows, root):
    if not root:
        raise ValueError('제출 폴더를 먼저 선택하세요.')
    path = Path(path)
    previous = read_connections(path) if path.exists() else []
    saved = copy.deepcopy(rows)
    for row in saved:
        for field in LINKS:
            target = contained_path(root, row.get(field))
            if target:
                row[field] = target.relative_to(Path(root).resolve()).as_posix()
        if row.get('보고서종류') == 'RT':
            # Save explicit attachments; folder-derived points are rescanned next time.
            mapping = row.get('RT영상') or {}
            row['RT영상'] = {point: [contained_path(root, value).relative_to(Path(root).resolve()).as_posix()
                                    for value in values] for point, values in point_images({**row, '필름영상화': ''}, root).items()} if mapping else {}
    current_groups, previous_groups = {}, {}
    for row in saved:
        current_groups.setdefault(identity(row.get('원본값', row)), []).append(row)
    for row in previous:
        previous_groups.setdefault(identity(row.get('원본값', row)), []).append(row)
    unresolved = []
    seen = set()
    for key, group in previous_groups.items():
        if all(key[:4]) and len(group) == 1 and len(current_groups.get(key, [])) == 1:
            continue
        for row in group:
            source = row.get('원본값', {})
            if not any(row.get(field) and row.get(field) != source.get(field) for field in SUPPLEMENT_FIELDS):
                continue
            signature = json.dumps(row, ensure_ascii=False, sort_keys=True)
            if signature not in seen:
                unresolved.append(copy.deepcopy(row))
                seen.add(signature)
    descriptor, temporary = tempfile.mkstemp(prefix='.connections_', suffix='.json', dir=path.parent)
    os.close(descriptor)
    try:
        Path(temporary).write_text(json.dumps({'version': 3, 'rows': saved, 'unresolved': unresolved}, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


def contained_path(root, value):
    if not value:
        return None
    root = Path(root).resolve()
    path = Path(value)
    path = (path if path.is_absolute() else root / path).resolve()
    if not path.is_relative_to(root):
        raise ValueError('링크 대상은 제출 폴더 안에 있어야 합니다: ' + str(path))
    return path


def inventory(root):
    root = Path(root).resolve()
    bases = [root / method for method in METHODS]
    bases += [root / '보고서', root / 'PAUT DATA', root / 'RT영상화', root / '영상화장비사진' / '중앙지사' / '영상화_결과']
    return sorted(path for base in bases if base.is_dir()
                  for path in base.rglob('*')
                  if path.is_file() and path.resolve().is_relative_to(root) and not path.name.startswith('~$'))


def folder_component(value):
    value = text(value)
    if not value:
        raise ValueError('폴더 생성에 필요한 라인번호·용접부번호·검사일자가 없습니다.')
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', value).rstrip(' .')
    reserved = re.fullmatch(r'(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?', clean, re.I)
    if clean != value or not clean or len(clean) > 90 or reserved:
        clean = (clean[:75] or '항목') + '_' + hashlib.sha256(value.encode()).hexdigest()[:10]
    return clean


def submission_plan(rows):
    plans, grouped = [], {}
    for row in rows:
        method = row.get('보고서종류')
        if method not in METHODS:
            raise ValueError('지원하지 않는 검사종류: ' + text(method))
        date = dt.date.fromisoformat(row['검사일자']).isoformat()
        kind = {'O': '최초', 'REPAIR': '보수'}.get(row.get('용접부특성'), '검사구분미확인')
        if row.get('검사차수'):
            kind += '_' + folder_component(row['검사차수'])
        path = Path(method) / folder_component(row['라인번호']) / folder_component(row['용접부번호']) / (date + '_' + kind)
        grouped.setdefault(path.as_posix().casefold(), []).append((row, path))
    for group in grouped.values():
        for row, path in group:
            if len(group) > 1:
                token = hashlib.sha256(json.dumps([row['id'], row.get('원본값', identity(row))], sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:12]
                path = path.with_name(path.name + '_기록-' + token)
            plans.append((row, path))
    return plans


def create_submission_folders(rows, root):
    if not rows:
        raise ValueError('중앙지사 이력을 먼저 불러오세요.')
    root = Path(root).resolve()
    plans = submission_plan(rows)  # Validate the entire plan before creating anything.
    targets = []
    for method in METHODS:
        targets.append(root / method)
        targets.append(root / method / '보고서')  # Shared reports need only one copy.
    for row, path in plans:
        targets.append(root / path / '보고서')
        if row['보고서종류'] == 'RT':
            targets.append(root / path / '필름영상화')
        elif row['보고서종류'] == 'PAUT':
            targets.append(root / path / '데이터')
    for target in targets:
        contained_path(root, str(target))
        if target.exists() and not target.is_dir():
            raise ValueError('폴더 위치에 파일이 있습니다: ' + str(target))
    for target in targets:
        target.mkdir(parents=True, exist_ok=True)
    return len(plans)


def auto_connect(rows, root):
    """Resolve only the generated record folder or an exact shared report number."""
    root = Path(root).resolve()
    if (root / '보고서').is_dir() or (root / 'PAUT DATA').is_dir():
        from site_apps.central.src.contract_links import connect_contract
        from site_apps.central.src.data_sync import get_data_dir
        issue_path = os.environ.get('CENTRAL_ATTRIBUTE_ISSUES') or Path(get_data_dir()).parent / 'ndt_report' / 'report_issue_history.json'
        result = connect_contract(rows, root, issue_path)
    else:
        result = copy.deepcopy(rows)
        for row, relative in submission_plan(result):
            directory = contained_path(root, relative)
            notes = row.setdefault('확인사항', [])
            if not row.get('보고서'):
                own = directory / '보고서'
                candidates = [p for p in own.iterdir() if p.is_file() and p.suffix.lower() in REPORT_EXTENSIONS and not p.name.startswith('~$')] if own.is_dir() else []
                if not candidates and row.get('보고서번호'):
                    shared = contained_path(root, Path(row['보고서종류']) / '보고서')
                    candidates = [p for p in shared.iterdir() if p.is_file() and p.suffix.lower() in REPORT_EXTENSIONS
                                  and p.stem.casefold() == row['보고서번호'].casefold()] if shared.is_dir() else []
                if row.get('보고서번호'):
                    candidates = [p for p in candidates if p.stem.casefold() == row['보고서번호'].casefold()]
                if len(candidates) == 1:
                    target = contained_path(root, str(candidates[0]))
                    row['보고서'] = target.relative_to(root).as_posix()
                    if not row.get('보고서번호'):
                        row['보고서번호'] = target.stem
                elif len(candidates) > 1:
                    message = '자동 연결 보류: 보고서 후보가 여러 개입니다.'
                    if message not in notes:
                        notes.append(message)
            if row['보고서종류'] == 'RT' and (row.get('자동연결영상') or (not row.get('RT영상') and not row.get('필름영상화'))):
                images = directory / '필름영상화'
                if images.is_dir():
                    # This exact path was derived from this record, not discovered by point alone.
                    mapping = {}
                    for target in sorted(images.iterdir()):
                        if not target.is_file() or target.suffix.lower() not in IMAGE_EXTENSIONS:
                            continue
                        matches = POINT_PATTERN.findall(target.stem)
                        if len(matches) == 1:
                            safe = contained_path(root, str(target))
                            mapping.setdefault(matches[0].replace(' ', ''), []).append(safe.relative_to(root).as_posix())
                    if mapping:
                        row['RT영상'] = mapping
                        row['필름영상화'] = images.relative_to(root).as_posix()
                        row['자동연결영상'] = True
            elif row['보고서종류'] == 'PAUT' and not row.get('데이터'):
                data = directory / '데이터'
                if data.is_dir() and any(p.is_file() for p in data.rglob('*')):
                    row['데이터'] = contained_path(root, str(data)).relative_to(root).as_posix()
    page_cache = {}
    for row in result:
        if not row.get('보고서 매수') and row.get('보고서'):
            try:
                rep_path = contained_path(root, row['보고서'])
                if rep_path and rep_path.is_file():
                    if rep_path not in page_cache:
                        page_cache[rep_path] = count_report_pages(rep_path)
                    if page_cache[rep_path]:
                        row['보고서 매수'] = page_cache[rep_path]
            except Exception:
                pass
        if row.get('보고서종류') == 'RT' and not row.get('결함종류'):
            if row.get('보고서'):
                try:
                    rep_path = contained_path(root, row['보고서'])
                    if rep_path and rep_path.is_file():
                        from site_apps.central.src.contract_links import extract_rt_defects_from_workbook
                        defects = extract_rt_defects_from_workbook(rep_path)
                        joint_key = text(row.get('용접부번호')).casefold()
                        line_key = text(row.get('라인번호')).casefold()
                        df = defects.get((line_key, joint_key)) or defects.get(('', joint_key))
                        if df:
                            row['결함종류'] = df
                except Exception:
                    pass
            if not row.get('결함종류'):
                from site_apps.central.src.contract_links import parse_defects
                raw_sec = row.get('구간정보') or (row.get('원본값') or {}).get('구간정보')
                df = parse_defects(raw_sec)
                if df:
                    row['결함종류'] = df
    return result


def point_images(row, root):
    """Use explicit per-point paths or scan only the record's assigned image folder.

    Never search all RT images by point number: every weld can have a 1-2 image.
    Repeated point images are retained as distinct links rather than picked once.
    """
    if row.get('보고서종류') != 'RT':
        return {}
    mapping = row.get('RT영상') or {}
    if not isinstance(mapping, dict):
        raise ValueError('RT영상은 {"1-2": ["상대경로"]} 형식이어야 합니다.')
    if mapping:
        result = {}
        for label, values in mapping.items():
            label = text(label).replace(' ', '')
            if not re.fullmatch(r'\d+-\d+', label):
                raise ValueError('RT 포인트 표기 확인: ' + label)
            values = [values] if isinstance(values, str) else values
            if not isinstance(values, list) or any(not isinstance(v, str) for v in values):
                raise ValueError('RT 포인트별 영상 경로는 문자열 목록이어야 합니다.')
            result[label] = list(dict.fromkeys(values))
        return result
    folder = contained_path(root, row.get('필름영상화'))
    if not folder or not folder.exists():
        return {}
    if folder.is_dir():
        relative = folder.relative_to(Path(root).resolve())
        joint = text(row.get('용접부번호'))
        line = text(row.get('라인번호'))
        # Shared method/image roots cannot establish ownership of bare 1-2 files.
        # Only a line/weld-specific folder may be scanned without explicit mapping.
        if not joint or not line or line.casefold() not in [p.casefold() for p in relative.parts] or joint.casefold() not in [p.casefold() for p in relative.parts]:
            raise ValueError('RT 자동 검색은 라인번호/용접부번호가 경로에 포함된 전용 폴더만 가능합니다. 포인트별 영상 파일을 직접 지정하세요.')
        # Avoid recursively collecting other welds or repeat inspections below it.
        paths = sorted(folder.iterdir())
    else:
        raise ValueError('RT 영상 파일은 포인트별 영상 연결에서 해당 포인트를 직접 지정하세요.')
    result = {}
    for path in paths:
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        matches = POINT_PATTERN.findall(path.stem)
        # Multiple numeric ranges could be dates or identifiers; require manual mapping.
        if len(matches) == 1:
            point = matches[0].replace(' ', '')
            resolved = contained_path(root, str(path))
            result.setdefault(point, []).append(resolved.relative_to(Path(root).resolve()).as_posix())
    return result


def count_report_pages(file_path):
    p = Path(file_path)
    if not p.is_file():
        return None
    suffix = p.suffix.lower()
    if suffix == '.pdf':
        try:
            import pypdf
            reader = pypdf.PdfReader(str(p))
            return len(reader.pages)
        except Exception:
            return None
    elif suffix in ('.xlsx', '.xls', '.xlsm'):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(p, data_only=True, read_only=True)
            try:
                sheet_count = len(wb.sheetnames)
                for sname in wb.sheetnames:
                    ws = wb[sname]
                    for row in ws.iter_rows(max_row=15, values_only=True):
                        for cell in row:
                            if cell:
                                s = str(cell)
                                m = re.search(r'Page\s*\d+\s*(?:of|/)\s*(\d+)', s, re.I)
                                if m:
                                    return int(m.group(1))
                for sname in wb.sheetnames:
                    ws = wb[sname]
                    for row in ws.iter_rows(max_row=10, values_only=True):
                        row_strs = [str(c).strip() for c in row if c is not None]
                        for idx, val in enumerate(row_strs):
                            if val.lower() == 'of' and idx + 1 < len(row_strs):
                                if row_strs[idx + 1].isdigit():
                                    return int(row_strs[idx + 1])
                return sheet_count
            finally:
                wb.close()
        except Exception:
            return None
    return 1


def issues(row, root):
    from site_apps.central.src.contract_links import allowed
    result = list(row.get('확인사항', []))
    if row.get('보고서종류') not in METHODS:
        result.append('보고서종류 확인: RT/PAUT/MT/PT 중 선택')
    if not row.get('검사결과'):
        result.append('검사결과 미입력')
    if not row.get('보고서번호'):
        result.append('보고서번호 미입력')
    for field in ('공사명', '시공사', '용역사', '도면번호', '라인번호', '용접부번호', '용접부특성', '검사일자'):
        if not text(row.get(field)):
            result.append(field + ' 미입력')
    try:
        dt.date.fromisoformat(text(row.get('검사일자')))
    except ValueError:
        result.append('검사일자 형식 확인')
    if row.get('용접부특성') not in ('O', 'REPAIR'):
        result.append('용접부특성 확인')
    pages = text(row.get('보고서 매수'))
    if not pages.isdigit() or int(pages) < 1:
        result.append('보고서 매수 확인')
    required = ['보고서'] + (['데이터'] if row['보고서종류'] == 'PAUT' else [])
    for field in required:
        try:
            target = contained_path(root, row.get(field))
            if target is None:
                result.append(field + ' 미연결')
            elif not target.exists():
                result.append(field + ' 파일/폴더 없음')
            elif not allowed(root, target, row['보고서종류'], field):
                result.append(field + ' 검사종류 폴더 불일치')
            elif target.is_dir() and not any(p.is_file() for p in target.rglob('*')):
                result.append(field + ' 폴더 비어 있음')
        except ValueError as error:
            result.append(str(error))
    if row['보고서종류'] == 'RT':
        try:
            images = point_images(row, root)
            for point in row.get('촬영포인트', []):
                if not images.get(point):
                    result.append(f'RT {point} 영상 미연결 (중앙지사 촬영포인트)')
            if not images or not any(images.values()):
                result.append('RT 포인트 영상 미연결: 1-2, 2-3 등의 영상 파일을 지정하세요')
            for point, values in images.items():
                if not values:
                    result.append(f'RT {point} 영상 미연결')
                for value in values:
                    target = contained_path(root, value)
                    if target is None or not target.is_file():
                        result.append(f'RT {point} 영상 파일 없음: {value}')
                    elif not allowed(root, target, 'RT', '필름영상화'):
                        result.append(f'RT {point} 영상이 RT 폴더 밖에 있음')
                    elif target.suffix.lower() not in IMAGE_EXTENSIONS:
                        result.append(f'RT {point} 영상 형식 확인: {target.suffix}')
        except ValueError as error:
            result.append(str(error))
    return list(dict.fromkeys(result))


def natural_key(val):
    if not val:
        return ()
    return tuple(int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(val)))


def joint_sort_key(val):
    val = str(val or '').strip().upper()
    if not val:
        return (99, ())
    if val.startswith('HT-S'):
        group = 1
    elif val.startswith('HT-R'):
        group = 2
    elif val.startswith('S'):
        group = 3
    elif val.startswith('R'):
        group = 4
    else:
        group = 5
    return (group, natural_key(val))


def row_sort_key(r):
    method_order = {'RT': 1, 'PAUT': 2, 'MT': 3, 'PT': 4}
    m = method_order.get(str(r.get('보고서종류', '')).upper(), 99)
    rep = natural_key(r.get('보고서번호', ''))
    loc = natural_key(r.get('라인번호', ''))
    joint = joint_sort_key(r.get('용접부번호', ''))
    res = 0 if r.get('검사결과') == '합격' else 1
    date = str(r.get('검사일자', ''))
    return (m, rep, loc, joint, res, date)


def export_workbook(rows, root, destination, unmatched=None, pending=None):
    from site_apps.central.src.contract_links import allowed
    """Native Excel hyperlinks; never generate executable formulas from source text."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
    root, destination = Path(root).resolve(), Path(destination).resolve()
    if destination.parent != root:
        raise ValueError('속성정보 엑셀은 제출 폴더 최상위에 저장하세요.')
    if not rows:
        raise ValueError('생성할 검사 기록이 없습니다.')
    if destination.exists():
        raise FileExistsError('같은 이름의 파일이 있습니다. 다른 이름을 선택하세요.')
    rows = sorted(rows, key=row_sort_key)
    page_cache = {}
    for r in rows:
        joint_no = text(r.get('용접부번호'))
        if re.search(r'\d+R$', joint_no, re.I):
            r['용접부특성'] = 'REPAIR'
        elif not r.get('용접부특성'):
            r['용접부특성'] = 'O'
        if not r.get('보고서 매수') and r.get('보고서'):
            try:
                rep_path = contained_path(root, r['보고서'])
                if rep_path and rep_path.is_file():
                    if rep_path not in page_cache:
                        page_cache[rep_path] = count_report_pages(rep_path)
                    if page_cache[rep_path]:
                        r['보고서 매수'] = page_cache[rep_path]
            except Exception:
                pass
        if r.get('보고서종류') == 'RT' and not r.get('결함종류'):
            if r.get('보고서'):
                try:
                    rep_path = contained_path(root, r['보고서'])
                    if rep_path and rep_path.is_file():
                        from site_apps.central.src.contract_links import extract_rt_defects_from_workbook
                        defects = extract_rt_defects_from_workbook(rep_path)
                        joint_key = text(r.get('용접부번호')).casefold()
                        line_key = text(r.get('라인번호')).casefold()
                        df = defects.get((line_key, joint_key)) or defects.get(('', joint_key))
                        if df:
                            r['결함종류'] = df
                except Exception:
                    pass
            if not r.get('결함종류'):
                from site_apps.central.src.contract_links import parse_defects
                raw_sec = r.get('구간정보') or (r.get('원본값') or {}).get('구간정보')
                df = parse_defects(raw_sec)
                if df:
                    r['결함종류'] = df
    book = Workbook()
    sheet = book.active
    sheet.title = '속성정보'
    mappings = []
    for record in rows:
        try:
            mappings.append(point_images(record, root))
        except ValueError:
            mappings.append({})
    point_counts = {}
    for record, mapping in zip(rows, mappings):
        for point, values in mapping.items():
            point_counts[point] = max(point_counts.get(point, 0), len(values))
        for point in record.get('촬영포인트', []):
            point_counts[point] = max(point_counts.get(point, 0), 1)
    point_columns = [(point, index) for point in sorted(point_counts, key=lambda p: tuple(map(int, p.split('-'))))
                     for index in range(point_counts[point])]
    # 1~12열(FIELDS)은 발주처 공식 규격 유지, 바로 뒤에 보고서번호/검사결과/구간정보를 배치하여 용접부번호와의 거리를 최소화
    headers = (*FIELDS, '보고서번호 (클릭하여 열기)', '검사결과',
               *(f'RT {p}' + (f' ({i + 1})' if i else '') for p, i in point_columns),
               'RT 영상 수', '영상 보관 위치', 'PAUT 데이터 열기', '확인사항', '데이터 출처', '기록ID')
    sheet.append(headers)

    def put(ws, row, col, value):
        cell = ws.cell(row, col, value)
        if isinstance(value, str):
            cell.data_type = 's'  # Source values beginning with '=' remain literal.
        return cell

    def link(ws, row, col, path, label):
        cell = put(ws, row, col, label)
        # Excel's Windows file-link handler expects Unicode file paths, not
        # percent-encoded UTF-8 names (which it can interpret as literal names).
        # Escape only URI delimiters that otherwise change the target meaning.
        cell.hyperlink = path.relative_to(root).as_posix().replace('%', '%25').replace('#', '%23')
        cell.style = 'Hyperlink'

    for index, record in enumerate(rows, 2):
        for col, field in enumerate(FIELDS, 1):
            value = record.get(field, '')
            if field == '검사일자':
                try:
                    value = dt.date.fromisoformat(value)
                except (ValueError, TypeError):
                    pass
            if field == '보고서 매수' and text(value).isdigit():
                value = int(value)
            cell = put(sheet, index, col, value)
            if field == '검사일자':
                cell.number_format = 'yyyy-mm-dd'
        
        # 13: 보고서번호 (링크)
        rep_val = record.get('보고서번호', '')
        put(sheet, index, 13, rep_val)
        try:
            path = contained_path(root, record.get('보고서'))
            if path and path.exists():
                link(sheet, index, 13, path, rep_val or '번호 미입력 · 보고서 열기')
        except ValueError:
            pass

        # 14: 검사결과
        put(sheet, index, 14, record.get('검사결과', ''))

        # 15 ~ (14 + len(point_columns)): 구간정보 (RT 1-2, 2-3...)
        mapping = mappings[index - 2]
        for offset, (point, image_index) in enumerate(point_columns):
            col = 15 + offset
            values = mapping.get(point, [])
            if point in record.get('촬영포인트', []) and not values:
                put(sheet, index, col, '미연결')
            if image_index < len(values):
                try:
                    target = contained_path(root, values[image_index])
                    if target and target.is_file() and allowed(root, target, 'RT', '필름영상화') and target.suffix.lower() in IMAGE_EXTENSIONS:
                        link(sheet, index, col, target, point + (f' ({image_index + 1})' if image_index else ''))
                except ValueError:
                    pass

        # 후순위 관리 열들: RT 영상 수, 영상 보관 위치, PAUT 데이터 열기, 확인사항, 데이터 출처, 기록ID
        base_col = 15 + len(point_columns)
        # RT 영상 수
        put(sheet, index, base_col, sum(len(values) for values in mapping.values()) if record['보고서종류'] == 'RT' else '')
        # 영상 보관 위치
        try:
            film_path = contained_path(root, record.get('필름영상화'))
            if film_path and film_path.exists():
                link(sheet, index, base_col + 1, film_path, '필름영상화 열기')
        except ValueError:
            pass
        # PAUT 데이터 열기
        try:
            paut_path = contained_path(root, record.get('데이터'))
            if paut_path and paut_path.exists():
                link(sheet, index, base_col + 2, paut_path, '데이터 열기')
        except ValueError:
            pass
        # 확인사항
        put(sheet, index, base_col + 3, '\n'.join(issues(record, root)) or '입력 및 연결 확인 완료')
        # 데이터 출처
        put(sheet, index, base_col + 4, record.get('출처', ''))
        # 기록ID
        put(sheet, index, base_col + 5, record['id'])
    for col, options in (('H', 'O,REPAIR'), ('K', ','.join(METHODS))):
        validation = DataValidation(type='list', formula1='"' + options + '"', allow_blank=True)
        validation.showErrorMessage = True
        sheet.add_data_validation(validation)
        validation.add(f'{col}2:{col}{len(rows) + 1}')
    pages = DataValidation(type='whole', operator='greaterThanOrEqual', formula1=1, allow_blank=True)
    pages.showErrorMessage = True
    sheet.add_data_validation(pages)
    pages.add(f'L2:L{len(rows) + 1}')
    overview = book.create_sheet('검사내역', 0)
    # 보고서종류 -> 보고서번호 -> 검사결과 -> 라인번호 -> 용접부번호 -> 구간정보(RT 1-2...) 순으로 배치하여 조인트 번호 바로 옆에 구간정보 연결
    overview_headers = ('보고서종류', '보고서번호 (클릭하여 열기)', '검사결과', '라인번호', '용접부번호',
                        *(f'RT {p}' + (f' ({i + 1})' if i else '') for p, i in point_columns),
                        '검사일자', '용접부특성', '결함종류', 'PAUT 데이터 열기', 'RT 영상 연결현황', '확인사항', '검사차수')
    overview.append(overview_headers)

    for row_index, record in enumerate(rows, 2):
        # 1: 보고서종류
        put(overview, row_index, 1, record.get('보고서종류', ''))
        # 2: 보고서번호 (링크)
        rep_val = record.get('보고서번호', '')
        put(overview, row_index, 2, rep_val)
        try:
            path = contained_path(root, record.get('보고서'))
            if path and path.exists():
                link(overview, row_index, 2, path, rep_val or '번호 미입력 · 보고서 열기')
        except ValueError:
            pass
        # 3: 검사결과
        put(overview, row_index, 3, record.get('검사결과', ''))
        # 4: 라인번호
        put(overview, row_index, 4, record.get('라인번호', ''))
        # 5: 용접부번호
        put(overview, row_index, 5, record.get('용접부번호', ''))

        # 6 ~ (5 + len(point_columns)): 구간정보 (RT 1-2, 2-3...)
        mapping = mappings[row_index - 2]
        for offset, (point, image_index) in enumerate(point_columns):
            target_col = 6 + offset
            source_col = 15 + offset
            source_cell = sheet.cell(row_index, source_col)
            put(overview, row_index, target_col, source_cell.value or '')
            if source_cell.hyperlink:
                cell = put(overview, row_index, target_col, source_cell.value)
                cell.hyperlink = source_cell.hyperlink.target
                cell.style = 'Hyperlink'

        # 후속 정보 열들
        curr_col = 6 + len(point_columns)
        # 검사일자
        put(overview, row_index, curr_col, record.get('검사일자', ''))
        # 용접부특성
        put(overview, row_index, curr_col + 1, record.get('용접부특성', ''))
        # 결함종류
        put(overview, row_index, curr_col + 2, record.get('결함종류', ''))
        # PAUT 데이터 열기
        try:
            paut_path = contained_path(root, record.get('데이터'))
            if paut_path and paut_path.exists():
                link(overview, row_index, curr_col + 3, paut_path, '데이터 열기')
        except ValueError:
            pass
        # RT 영상 연결현황
        expected = record.get('촬영포인트', [])
        if record['보고서종류'] == 'RT':
            def valid_image(value):
                try:
                    target = contained_path(root, value)
                    return bool(target and target.is_file() and allowed(root, target, 'RT', '필름영상화') and target.suffix.lower() in IMAGE_EXTENSIONS)
                except ValueError:
                    return False
            if expected:
                count = sum(any(valid_image(value) for value in mapping.get(point, [])) for point in expected)
                connection_status = f'{count}/{len(expected)} 포인트 연결'
            else:
                count = sum(any(valid_image(value) for value in values) for values in mapping.values())
                connection_status = f'{count}포인트 연결 / 전체 포인트 미확인'
            put(overview, row_index, curr_col + 4, connection_status)
        # 확인사항
        base_col = 15 + len(point_columns)
        put(overview, row_index, curr_col + 5, sheet.cell(row_index, base_col + 3).value)
        # 검사차수
        put(overview, row_index, curr_col + 6, record.get('검사차수', ''))
    if unmatched:
        diagnostics = book.create_sheet('보고서 대조확인')
        keys = ('검사일자', '보고서종류', '라인번호', '용접부번호', '보고서번호', '대조결과', '출처')
        diagnostics.append(keys)
        for index, row in enumerate(unmatched, 2):
            for col, key in enumerate(keys, 1):
                put(diagnostics, index, col, row.get(key, ''))
    if pending:
        recovery = book.create_sheet('연결 복원보류')
        keys = ('검사일자', '보고서종류', '라인번호', '용접부번호', '용접부특성', '보고서번호', '보고서', '필름영상화', '데이터', 'RT영상', 'id')
        recovery.append(keys)
        for index, row in enumerate(pending, 2):
            for col, key in enumerate(keys, 1):
                value = row.get(key, '')
                put(recovery, index, col, json.dumps(value, ensure_ascii=False) if isinstance(value, dict) else value)
    paut_sheet = book.create_sheet('PAUT 데이터 연결')
    paut_sheet.append(('검사일자', '라인번호', 'JOINT', '데이터 파일', '파일 열기'))
    for record in rows:
        for value in record.get('데이터파일', []):
            target = contained_path(root, value)
            if target and target.is_file() and allowed(root, target, 'PAUT', '데이터'):
                paut_sheet.append((record.get('검사일자'), record.get('라인번호'), record.get('용접부번호'), value))
                link(paut_sheet, paut_sheet.max_row, 5, target, '데이터 열기')
    files = book.create_sheet('제출파일목록')
    files.append(('검사종류', '상대경로', '파일 열기'))
    for index, path in enumerate(inventory(root), 2):
        relative = path.relative_to(root)
        put(files, index, 1, relative.parts[0])
        put(files, index, 2, relative.as_posix())
        link(files, index, 3, path, '열기')
    guide = book.create_sheet('작성안내')
    for line in [
        ('항목', '안내'),
        ('폴더 구조', '엑셀을 RT/PAUT/MT/PT 폴더와 같은 최상위 위치에 보관하세요.'),
        ('이동 방법', '메모리카드에 엑셀과 검사 폴더를 함께 복사하고 내부 경로와 파일명을 유지하세요.'),
        ('대조 기준', '검사일자·검사종류·라인번호·용접부번호·용접부특성이 모두 같은 유일한 기록만 병합합니다.'),
        ('기준 목록', '중앙지사 저장 내역만 검사 목록으로 사용합니다. 일치하지 않는 보고서는 보고서 대조확인 시트에 분리합니다.'),
        ('자료 연결', '파일명만으로 확정하지 않습니다. 생성기에서 기록을 선택하고 보고서/영상/데이터를 연결하세요.'),
        ('RT 포인트 영상', 'RT 1-2 / RT 2-3 / RT 3-1 등 포인트 열을 클릭하면 해당 영상 파일이 열립니다. 같은 포인트의 여러 영상도 개별 링크로 표시합니다.'),
        ('RT 영상 수', '연결된 영상 경로 수이며 필수 포인트 수나 검사 완료를 의미하지 않습니다. 보고서의 촬영 포인트와 대조하세요.'),
        ('포인트 연결현황', '중앙지사 구간정보 또는 보고서의 명확한 촬영포인트 목록이 있을 때만 전체 대비 연결 수를 표시합니다. 없으면 전체 포인트 미확인으로 표시합니다.'),
        ('확인사항', '생성 시점의 점검 결과입니다. 엑셀에서 직접 수정해도 자동 갱신되지 않습니다. 생성기에서 재생성하세요.'),
        ('결함종류', '미입력은 무결함을 의미하지 않습니다. 원본 보고서의 코드와 무결함 표기 기준을 확인하세요.'),
        ('원본 대조', '과업내용서 별표1의 누락 항목 및 발주처 지정 항목·파일명은 제출 전에 확인하세요.'),
        ('전용 데이터', 'PAUT 데이터는 해당 프로그램이 설치된 PC에서 열 수 있습니다.'),
        ('생성일시', dt.datetime.now().isoformat(timespec='seconds')),
    ]:
        guide.append(line)
    for ws in book:
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(color='FFFFFF', bold=True)
            cell.fill = PatternFill('solid', fgColor='17365D')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 22
    for cell in guide['B']:
        if cell.row > 1:
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    sheet_issue_col = get_column_letter(18 + len(point_columns))
    sheet_source_col = get_column_letter(19 + len(point_columns))
    sheet.column_dimensions[sheet_issue_col].width = 65
    sheet.column_dimensions[sheet_source_col].width = 40
    files.column_dimensions['B'].width = 90
    guide.column_dimensions['B'].width = 100
    overview.column_dimensions['B'].width = 28
    overview.column_dimensions['D'].width = 38
    overview.column_dimensions[get_column_letter(11 + len(point_columns))].width = 65
    # Atomic creation prevents leaving a half-written deliverable.
    descriptor, temporary_name = tempfile.mkstemp(prefix='.attributes_', suffix='.xlsx', dir=root)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        book.save(temporary)
        check = load_workbook(temporary, read_only=False)
        try:
            if check['속성정보'].max_row != len(rows) + 1:
                raise ValueError('저장 후 행 수 검증 실패')
        finally:
            check.close()
        if destination.exists():
            raise FileExistsError(str(destination))
        temporary.rename(destination)
    finally:
        book.close()
        temporary.unlink(missing_ok=True)
    return destination


def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from site_apps.central.src.data_sync import get_history_path, get_submission_dir

    window = tk.Tk()
    window.title('중앙지사 속성정보 엑셀 생성')
    window.geometry('1250x760')
    state = {'history': [], 'reports': [], 'rows': [], 'unmatched': [], 'cached_keys': set(), 'cached_ids': set()}
    root_var = tk.StringVar(value=os.environ.get('CENTRAL_ATTRIBUTE_ROOT') or get_submission_dir())
    history_var = tk.StringVar(value=os.environ.get('CENTRAL_ATTRIBUTE_HISTORY') or get_history_path())
    only_reported_var = tk.BooleanVar(value=True)
    status_filter_var = tk.StringVar(value='전체 자료')
    status = tk.StringVar(value='중앙지사 이력을 불러온 뒤 보고서 앱에서 저장한 JSON 또는 XLSX를 추가하세요.')
    form = ttk.Frame(window, padding=10)
    form.pack(fill='both', expand=True)

    def run(action):
        try:
            action()
        except Exception as error:
            messagebox.showerror('확인 필요', str(error), parent=window)

    def choose_root():
        path = filedialog.askdirectory(title='메모리카드 또는 제출용 최상위 폴더 선택')
        if path:
            persist()
            root_var.set(path)
            rebuild(preserve=False)
            connect_files()

    def choose_history():
        path = filedialog.askopenfilename(title='중앙지사 daily_work_history.json', filetypes=[('JSON', '*.json')])
        if path:
            history_var.set(path)

    for index, (label, variable, command) in enumerate([
        ('계약 / 제출 폴더', root_var, choose_root), ('중앙지사 이력', history_var, choose_history)
    ]):
        ttk.Label(form, text=label).grid(row=index, column=0, sticky='w')
        ttk.Entry(form, textvariable=variable, state='readonly' if index == 0 else 'normal').grid(row=index, column=1, sticky='ew', padx=6, pady=3)
        ttk.Button(form, text='선택', command=command).grid(row=index, column=2)
    form.columnconfigure(1, weight=1)
    toolbar = ttk.Frame(form)
    toolbar.grid(row=2, column=0, columnspan=3, sticky='ew', pady=8)

    def get_issue_history_keys():
        issue_path = os.environ.get('CENTRAL_ATTRIBUTE_ISSUES')
        if not issue_path:
            from site_apps.central.src.data_sync import get_data_dir
            issue_path = Path(get_data_dir()).parent / 'ndt_report' / 'report_issue_history.json'
        issue_path = Path(issue_path)
        if not issue_path.is_file():
            g_path = Path(r'G:\내 드라이브\07_Antigravity\PMI_Data\ndt_report\report_issue_history.json')
            if g_path.is_file():
                issue_path = g_path
        keys = set()
        if issue_path.is_file():
            try:
                entries = json.loads(issue_path.read_text(encoding='utf-8-sig'))
                for entry in entries:
                    if entry.get('deleted_at') or not entry.get('report_no'):
                        continue
                    mode = str(entry.get('mode') or '').strip().upper()
                    for item in entry.get('items', []):
                        line = str(item.get('Line No.') or item.get('Dwg') or item.get('라인번호') or '').strip()
                        joint = str(item.get('Joint') or item.get('Joint No.') or item.get('용접부번호') or '').strip()
                        date = str(item.get('Date') or item.get('검사일자') or '').strip()
                        keys.add((mode, date, line, joint))
            except Exception:
                pass
        return keys

    issue_keys_cache = [None]
    def has_report_history(row):
        if row.get('보고서번호') or row.get('보고서'):
            return True
        if issue_keys_cache[0] is None:
            issue_keys_cache[0] = get_issue_history_keys()
        k = (
            str(row.get('보고서종류') or '').strip().upper(),
            str(row.get('검사일자') or '').strip(),
            str(row.get('라인번호') or '').strip(),
            str(row.get('용접부번호') or '').strip()
        )
        return k in issue_keys_cache[0]

    def get_row_status(row):
        k = (
            str(row.get('보고서종류') or '').strip().upper(),
            str(row.get('검사일자') or '').strip(),
            str(row.get('라인번호') or '').strip(),
            str(row.get('용접부번호') or '').strip()
        )
        is_existing = (row.get('id') in state.get('cached_ids', set())) or (k in state.get('cached_keys', set()))
        has_report = bool(row.get('보고서') or row.get('보고서번호'))
        
        if is_existing:
            if has_report:
                return '기존 완료', 'existing_done'
            return '기존 미연결', 'unconnected'
        else:
            if has_report:
                return '신규 연결', 'new_connected'
            return '신규 미연결', 'new_unconnected'

    def active_rows():
        rows = state['rows']
        if only_reported_var.get():
            rows = [r for r in rows if has_report_history(r)]
        f = status_filter_var.get()
        if f == '신규 자료만':
            rows = [r for r in rows if '신규' in get_row_status(r)[0]]
        elif f == '기존 완료만':
            rows = [r for r in rows if get_row_status(r)[0] == '기존 완료']
        elif f == '미연결만':
            rows = [r for r in rows if '미연결' in get_row_status(r)[0]]
        return rows

    def refresh():
        tree.delete(*tree.get_children())
        rows_to_show = sorted(active_rows(), key=row_sort_key)
        for row in rows_to_show:
            count = len(issues(row, root_var.get())) if root_var.get() else '폴더 선택'
            st_text, tag = get_row_status(row)
            vals = (
                row.get('검사일자', ''),
                row.get('보고서종류', ''),
                row.get('라인번호', ''),
                row.get('용접부번호', ''),
                st_text,
                row.get('보고서번호', ''),
                row.get('용접부특성', ''),
                row.get('검사결과', ''),
                str(count)
            )
            tree.insert('', 'end', iid=row['id'], values=vals, tags=(tag,))
        pending = pending_saved()
        cnt_exist_done = sum(1 for r in rows_to_show if get_row_status(r)[0] == '기존 완료')
        cnt_new_conn = sum(1 for r in rows_to_show if get_row_status(r)[0] == '신규 연결')
        cnt_unconn = sum(1 for r in rows_to_show if '미연결' in get_row_status(r)[0])
        status_detail = f" [기존완료: {cnt_exist_done}건 | 신규연결: {cnt_new_conn}건 | 미연결: {cnt_unconn}건]"
        filter_msg = f" (필터 적용: {len(rows_to_show)}/{len(state['rows'])}건)" if (len(rows_to_show) != len(state['rows'])) else ""
        status.set(f"중앙지사 검사 {len(rows_to_show)}건{status_detail}{filter_msg} / 보고서 대조확인 {len(state['unmatched'])}건 / 연결 복원보류 {len(pending)}건. 보류 자료는 자동 연결하지 않고 보존합니다.")

    def pending_saved():
        if not root_var.get():
            return []
        path = Path(root_var.get()) / CACHE_NAME
        if not path.exists():
            return []
        read_connections(path)
        with path.open(encoding='utf-8-sig') as stream:
            return json.load(stream).get('unresolved', [])

    def persist():
        if state['rows'] and root_var.get():
            path = Path(root_var.get()) / CACHE_NAME
            if path.exists():
                read_connections(path)
            write_connections(path, state['rows'], root_var.get())

    def rebuild(preserve=True):
        base = copy.deepcopy(state['history'])
        if root_var.get():
            cache = Path(root_var.get()) / CACHE_NAME
            if cache.exists():
                cached_connections = read_connections(cache)
                state['cached_keys'] = {
                    (str(r.get('보고서종류') or '').strip().upper(),
                     str(r.get('검사일자') or '').strip(),
                     str(r.get('라인번호') or '').strip(),
                     str(r.get('용접부번호') or '').strip())
                    for r in cached_connections
                }
                state['cached_ids'] = {r['id'] for r in cached_connections if 'id' in r}
                base = restore_connections(base, cached_connections)
        if preserve:
            base = restore_connections(base, state['rows'])
        for r in base:
            if not r.get('용역사'):
                r['용역사'] = '서울검사'
            if not r.get('시공사') and r.get('원본값', {}).get('업체'):
                r['시공사'] = r['원본값']['업체']
            if not r.get('공사명'):
                r['공사명'] = '2026년 중앙지사 열수송관 비파괴검사 단가계약'
            joint_no = text(r.get('용접부번호'))
            if re.search(r'\d+R$', joint_no, re.I):
                r['용접부특성'] = 'REPAIR'
            elif not r.get('용접부특성') and not any('RT_RE는 불량매수' in s for s in r.get('확인사항', [])):
                r['용접부특성'] = 'O'
            if '확인사항' in r and isinstance(r['확인사항'], list):
                r['확인사항'] = [n for n in r['확인사항'] if '시공사/하도사 구분 확인' not in n]
                if r.get('보고서'):
                    r['확인사항'] = [n for n in r['확인사항'] if '보고서 후보' not in n and '보고서 번호가 여러 개' not in n and '보고서 파일이 여러 개' not in n]
                if r.get('RT영상') or r.get('필름영상화'):
                    r['확인사항'] = [n for n in r['확인사항'] if 'RT 영상 연결 보류' not in n]
                if r.get('데이터'):
                    r['확인사항'] = [n for n in r['확인사항'] if 'PAUT 데이터 연결 보류' not in n]
        unmatched = []
        state['rows'] = merge_rows(base, state['reports'], unmatched)
        state['unmatched'] = unmatched
        refresh()

    def load():
        selected_path = Path(history_var.get())
        with selected_path.open(encoding='utf-8-sig') as stream:
            selected_data = json.load(stream)
        saved_connections = None
        if isinstance(selected_data, dict) and 'version' in selected_data and 'rows' in selected_data:
            saved_connections = read_connections(selected_path)
            selected_path = Path(os.environ.get('CENTRAL_ATTRIBUTE_HISTORY') or get_history_path())
            if selected_path.resolve() == Path(history_var.get()).resolve():
                raise ValueError('중앙지사 이력에서 daily_work_history.json을 선택한 후 연결정보 불러오기를 사용하세요.')
        loaded = load_history(selected_path)
        persist()
        state['history'] = loaded
        history_var.set(str(selected_path))
        rebuild()
        if saved_connections is not None:
            state['rows'] = restore_connections(state['rows'], saved_connections)
            persist()
            refresh()
        if (Path(root_var.get()) / '보고서').is_dir() or (Path(root_var.get()) / 'PAUT DATA').is_dir():
            connect_files()
        else:
            refresh()

    method = tk.StringVar(value='RT')
    def add_report():
        if not state['history']:
            raise ValueError('중앙지사 이력을 먼저 불러오세요. 보고서 데이터는 일치하는 기록만 보완합니다.')
        path = filedialog.askopenfilename(title='보고서 앱 데이터 저장 JSON / 표 형식 XLSX', filetypes=[('보고서 데이터', '*.json *.xlsx')])
        if path:
            loaded = load_report_data(path, method.get())
            existing = {row['id'] for row in state['reports']}
            state['reports'].extend(row for row in loaded if row['id'] not in existing)
            rebuild()
            persist()
            if state['unmatched']:
                messagebox.showinfo('보고서 대조확인', f"{len(state['unmatched'])}건은 중앙지사 기록에 병합하지 않았습니다. 엑셀의 '보고서 대조확인' 시트에서 확인할 수 있습니다.")

    def make_folders():
        if not root_var.get():
            choose_root()
        if not root_var.get():
            return
        rows = active_rows()
        if not rows:
            raise ValueError('생성할 대상 검사 기록이 없습니다.')
        count = create_submission_folders(rows, root_var.get())
        messagebox.showinfo('폴더 생성 완료', f'{count}개 검사 기록의 제출 폴더를 만들었습니다.\n보고서·영상·데이터를 넣은 후 자동 연결을 누르세요. 기존 파일은 변경하지 않습니다.')

    def connect_files():
        if not root_var.get() or not state['rows']:
            return
        state['rows'] = auto_connect(state['rows'], root_var.get())
        persist()
        refresh()

    # 1. 핵심 데이터 로드 및 연결 작업
    ttk.Button(toolbar, text='중앙지사 이력 불러오기', command=lambda: run(load)).pack(side='left')
    ttk.Button(toolbar, text='자료 자동 연결', command=lambda: run(connect_files)).pack(side='left', padx=4)
    ttk.Button(toolbar, text='제출 폴더 틀 생성', command=lambda: run(make_folders)).pack(side='left', padx=4)

    # 2. 구분선
    ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)

    # 3. 데이터 조회 및 필터
    ttk.Checkbutton(toolbar, text='보고서 이력 있는 것만', variable=only_reported_var, command=refresh).pack(side='left', padx=(0, 6))
    ttk.Label(toolbar, text='구분:').pack(side='left', padx=(2, 2))
    status_filter_combo = ttk.Combobox(toolbar, textvariable=status_filter_var, values=['전체 자료', '신규 자료만', '기존 완료만', '미연결만'], state='readonly', width=11)
    status_filter_combo.pack(side='left', padx=(0, 4))
    status_filter_combo.bind('<<ComboboxSelected>>', lambda _: refresh())

    # 4. 구분선
    ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)

    # 5. 외부 보조 데이터 가져오기 (용도 명시)
    ttk.Label(toolbar, text='외부 보조:').pack(side='left', padx=(0, 2))
    ttk.Combobox(toolbar, textvariable=method, values=METHODS, state='readonly', width=5).pack(side='left', padx=2)
    ttk.Button(toolbar, text='보고서 데이터 추가', command=lambda: run(add_report)).pack(side='left', padx=(2, 0))

    ttk.Label(form, text='안내: [자료 자동 연결]을 실행하면 중앙지사 제출 폴더 내 보고서(.xlsx) 및 RT 영상, PAUT 데이터가 100% 자동 매칭됩니다.').grid(row=3, column=0, columnspan=3, sticky='w')
    visible = ('검사일자', '보고서종류', '라인번호', '용접부번호', '구분', '보고서번호', '용접부특성', '검사결과')
    table = ttk.Frame(form)
    table.grid(row=4, column=0, columnspan=3, sticky='nsew', pady=8)
    form.rowconfigure(4, weight=1)
    tree = ttk.Treeview(table, columns=(*visible, '확인 항목 수'), show='headings')
    for field in (*visible, '확인 항목 수'):
        tree.heading(field, text=field)
        width = 160 if field == '보고서번호' else (100 if field in ('검사일자', '보고서종류', '구분', '검사결과', '확인 항목 수') else 145)
        tree.column(field, width=width)
    tree.tag_configure('existing_done', background='#FFFFFF')
    tree.tag_configure('new_connected', background='#E8F8F5')
    tree.tag_configure('new_unconnected', background='#FFF9E6')
    tree.tag_configure('unconnected', background='#FDF2E9')
    scroll = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
    horizontal = ttk.Scrollbar(table, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
    tree.grid(row=0, column=0, sticky='nsew')
    scroll.grid(row=0, column=1, sticky='ns')
    horizontal.grid(row=1, column=0, sticky='ew')
    table.rowconfigure(0, weight=1)
    table.columnconfigure(0, weight=1)

    def edit(_event=None):
        selected = tree.selection()
        if not selected:
            return
        if not root_var.get():
            raise ValueError('제출 폴더를 먼저 선택하세요.')
        selected_set = set(selected)
        target_rows = [row for row in state['rows'] if row['id'] in selected_set]
        if not target_rows:
            return

        is_multi = len(target_rows) > 1
        primary_row = target_rows[0]

        dialog = tk.Toplevel(window)
        title_text = f'속성정보 / 원본 자료 연결 (선택된 {len(target_rows)}건 일괄 편집)' if is_multi else '속성정보 / 원본 자료 연결'
        dialog.title(title_text)
        dialog.transient(window)
        dialog.grab_set()

        all_fields = (*FIELDS, *LINKS, '검사결과', '보고서번호')
        initial_values = {}
        for field in all_fields:
            if is_multi:
                vals = {r.get(field, '') for r in target_rows}
                initial_values[field] = next(iter(vals)) if len(vals) == 1 else ''
            else:
                initial_values[field] = primary_row.get(field, '')

        variables = {field: tk.StringVar(value=initial_values[field]) for field in all_fields}

        curr_row = 0
        if is_multi:
            banner = ttk.Label(
                dialog,
                text=f'선택된 {len(target_rows)}건의 검사 기록을 일괄 편집합니다.\n입력하거나 변경한 항목만 선택된 {len(target_rows)}건 전체에 일괄 적용됩니다. (빈 칸으로 둔 항목은 기존 값 유지)',
                foreground='#0055AA'
            )
            banner.grid(row=curr_row, column=0, columnspan=4, sticky='w', padx=6, pady=(4, 6))
            curr_row += 1

        for field in all_fields:
            variable = variables[field]
            ttk.Label(dialog, text=field).grid(row=curr_row, column=0, sticky='w', padx=6)
            original = primary_row.get('원본값', {}).get(field) if not is_multi else False
            ttk.Entry(dialog, textvariable=variable, width=85, state='readonly' if original else 'normal').grid(row=curr_row, column=1, padx=6, pady=2)
            if field in LINKS:
                def browse(field=field):
                    path = filedialog.askopenfilename(initialdir=root_var.get(), parent=dialog)
                    if path:
                        rel_path = contained_path(root_var.get(), path).relative_to(Path(root_var.get()).resolve()).as_posix()
                        variables[field].set(rel_path)
                        if field == '보고서':
                            stem = Path(path).stem
                            curr_rep = variables['보고서번호'].get().strip()
                            if not curr_rep or curr_rep == initial_values.get('보고서번호', ''):
                                variables['보고서번호'].set(stem)

                def folder(field=field):
                    path = filedialog.askdirectory(initialdir=root_var.get(), parent=dialog)
                    if path:
                        rel_path = contained_path(root_var.get(), path).relative_to(Path(root_var.get()).resolve()).as_posix()
                        variables[field].set(rel_path)

                ttk.Button(dialog, text='파일', command=lambda f=browse: run(f)).grid(row=curr_row, column=2)
                ttk.Button(dialog, text='폴더', command=lambda f=folder: run(f)).grid(row=curr_row, column=3)
            curr_row += 1

        def save():
            for field in LINKS:
                val = variables[field].get().strip()
                if val:
                    contained_path(root_var.get(), val)

            if not is_multi:
                if variables['필름영상화'].get().strip() != primary_row.get('필름영상화', ''):
                    primary_row['자동연결영상'] = False
                primary_row.update({field: var.get().strip() for field, var in variables.items()})
                if '확인사항' in primary_row and isinstance(primary_row['확인사항'], list):
                    if primary_row.get('보고서') or primary_row.get('보고서번호'):
                        primary_row['확인사항'] = [n for n in primary_row['확인사항'] if '발행 보고서 번호가 여러 개입니다' not in n]
            else:
                changed_fields = {}
                for field, var in variables.items():
                    val = var.get().strip()
                    if val != initial_values[field]:
                        changed_fields[field] = val

                for r in target_rows:
                    if '필름영상화' in changed_fields and changed_fields['필름영상화'] != r.get('필름영상화', ''):
                        r['자동연결영상'] = False
                    r.update(changed_fields)
                    if '확인사항' in r and isinstance(r['확인사항'], list):
                        if r.get('보고서') or r.get('보고서번호'):
                            r['확인사항'] = [n for n in r['확인사항'] if '발행 보고서 번호가 여러 개입니다' not in n]

            persist()
            dialog.destroy()
            refresh()

        if not is_multi and primary_row.get('보고서종류') == 'RT':
            def edit_images():
                popup = tk.Toplevel(dialog)
                popup.title('RT 포인트별 영상 연결')
                popup.transient(dialog)
                popup.grab_set()
                ttk.Label(popup, text='한 줄에 포인트=상대경로 입력. 같은 포인트의 영상은 줄을 추가하세요.').pack(padx=8, pady=5)
                editor = tk.Text(popup, width=100, height=15)
                editor.pack(padx=8)
                current = point_images({**primary_row, '필름영상화': variables['필름영상화'].get()}, root_var.get())
                editor.insert('1.0', '\n'.join(f'{p}={v}' for p, values in current.items() for v in values))
                point_var = tk.StringVar(value='1-2')
                controls = ttk.Frame(popup)
                controls.pack(fill='x', padx=8, pady=6)
                ttk.Label(controls, text='포인트').pack(side='left')
                ttk.Entry(controls, textvariable=point_var, width=10).pack(side='left')
                def add_images():
                    if not re.fullmatch(r'\d+-\d+', point_var.get().strip()):
                        raise ValueError('포인트를 1-2 형태로 입력하세요.')
                    paths = filedialog.askopenfilenames(parent=popup, initialdir=root_var.get(), title='이 포인트의 영상 선택')
                    additions = [contained_path(root_var.get(), p).relative_to(Path(root_var.get()).resolve()).as_posix() for p in paths]
                    for value in additions:
                        editor.insert('end', '\n' + point_var.get().strip() + '=' + value)
                def apply_images():
                    mapping = {}
                    for line in editor.get('1.0', 'end').splitlines():
                        if not line.strip():
                            continue
                        point, separator, value = line.partition('=')
                        if not separator:
                            raise ValueError('포인트=상대경로 형식으로 입력하세요.')
                        target = contained_path(root_var.get(), value.strip())
                        from site_apps.central.src.contract_links import allowed
                        if not target or not target.is_file() or not allowed(root_var.get(), target, 'RT', '필름영상화'):
                            raise ValueError('RT 폴더 안의 실제 영상 파일을 선택하세요.')
                        mapping.setdefault(point.strip(), []).append(target.relative_to(Path(root_var.get()).resolve()).as_posix())
                    point_images({**primary_row, 'RT영상': mapping}, root_var.get())
                    primary_row['RT영상'] = mapping
                    primary_row['자동연결영상'] = False
                    persist()
                    popup.destroy()
                    dialog.grab_set()
                ttk.Button(controls, text='영상 파일 추가', command=lambda: run(add_images)).pack(side='left', padx=5)
                ttk.Button(controls, text='포인트 연결 적용', command=lambda: run(apply_images)).pack(side='right')
            ttk.Button(dialog, text='RT 포인트별 영상 연결 (1-2 / 2-3 / 3-1 …)', command=lambda: run(edit_images)).grid(row=curr_row, column=1, pady=4)
            curr_row += 1

        if not is_multi:
            issue_msg = '\n'.join(issues(primary_row, root_var.get()))
        else:
            issue_cnt = sum(1 for r in target_rows if issues(r, root_var.get()))
            issue_msg = f'선택된 {len(target_rows)}건 중 확인 항목이 있는 기록: {issue_cnt}건'

        ttk.Label(dialog, text=issue_msg, wraplength=850).grid(row=curr_row, column=0, columnspan=4, sticky='w', padx=6, pady=4)
        curr_row += 1
        btn_label = f'선택된 {len(target_rows)}건에 일괄 적용' if is_multi else '적용'
        ttk.Button(dialog, text=btn_label, command=lambda: run(save)).grid(row=curr_row, column=1, pady=8)
    tree.bind('<Double-1>', lambda event: run(lambda: edit(event)))

    def open_selected():
        selected = tree.selection()
        if not selected:
            return
        row = next(row for row in state['rows'] if row['id'] == selected[0])
        target = contained_path(root_var.get(), row.get('보고서'))
        if target is None or not target.exists():
            raise ValueError('보고서를 먼저 연결하세요.')
        os.startfile(target)

    def common_fields():
        rows = active_rows()
        if not rows:
            raise ValueError('검사 기록을 먼저 불러오세요.')
        dialog = tk.Toplevel(window)
        dialog.title('공통정보 일괄 입력')
        dialog.transient(window)
        dialog.grab_set()

        default_contract = next((r['공사명'] for r in rows if r.get('공사명')), '2026년 중앙지사 열수송관 비파괴검사 단가계약')
        default_contractor = next((r['시공사'] for r in rows if r.get('시공사')), next((r.get('원본값', {}).get('업체', '') for r in rows if r.get('원본값', {}).get('업체')), ''))
        default_subcontractor = next((r['하도사'] for r in rows if r.get('하도사')), '')
        default_agency = next((r['용역사'] for r in rows if r.get('용역사')), '서울검사')

        defaults = {
            '공사명': default_contract,
            '시공사': default_contractor,
            '하도사': default_subcontractor,
            '용역사': default_agency,
        }

        variables = {}
        for index, field in enumerate(('공사명', '시공사', '하도사', '용역사')):
            variables[field] = tk.StringVar(value=defaults.get(field, ''))
            ttk.Label(dialog, text=field).grid(row=index, column=0, padx=8, pady=4)
            ttk.Entry(dialog, textvariable=variables[field], width=55).grid(row=index, column=1, padx=8)
        ttk.Label(dialog, text='입력한 값은 전체 기록의 빈 칸에만 적용합니다. 기존 값은 유지합니다.').grid(row=4, column=0, columnspan=2, padx=8)
        def apply():
            for row in active_rows():
                for field, variable in variables.items():
                    if not row.get(field) and variable.get().strip():
                        row[field] = variable.get().strip()
            persist()
            dialog.destroy()
            refresh()
        ttk.Button(dialog, text='빈 칸에 적용', command=apply).grid(row=5, column=1, pady=8)

    def save_connections():
        path = filedialog.asksaveasfilename(title='편집 내용 및 연결정보 저장', defaultextension='.json', initialfile='속성정보_연결정보.json', filetypes=[('JSON', '*.json')])
        if path:
            if not state['rows']:
                raise ValueError('저장할 검사 기록이 없습니다.')
            write_connections(path, state['rows'], root_var.get())
            persist()

    def load_connections():
        if not state['history']:
            raise ValueError('중앙지사 이력을 먼저 불러오세요.')
        path = filedialog.askopenfilename(title='저장한 연결정보 불러오기', filetypes=[('JSON', '*.json')])
        if not path:
            return
        rows = read_connections(path)
        state['rows'] = restore_connections(state['rows'], rows)
        persist()
        refresh()

    def export():
        if not root_var.get():
            raise ValueError('제출 폴더를 선택하세요.')
        rows_to_export = active_rows()
        if not rows_to_export:
            raise ValueError('저장할 대상 검사 기록이 없습니다.')
        count = sum(bool(issues(row, root_var.get())) for row in rows_to_export)
        if count and not messagebox.askyesno('검토용 엑셀 생성', f'{count}건에 확인사항이 있습니다. 확인사항을 포함한 검토용 엑셀을 생성할까요?'):
            return
        path = filedialog.asksaveasfilename(initialdir=root_var.get(), initialfile='중앙지사_속성정보.xlsx', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if path:
            persist()
            export_workbook(rows_to_export, root_var.get(), path, state['unmatched'], pending_saved())
            messagebox.showinfo('생성 완료', f'{len(rows_to_export)}건 저장\n확인사항 있는 기록: {count}건\n{path}')

    bottom = ttk.Frame(form)
    bottom.grid(row=5, column=0, columnspan=3, sticky='ew')
    for label, action in [('공통정보 입력', common_fields), ('선택 기록 편집 / 자료 연결', edit), ('선택 보고서 열기', open_selected),
                          ('연결정보 저장', save_connections), ('연결정보 불러오기', load_connections), ('엑셀 생성', export)]:
        ttk.Button(bottom, text=label, command=lambda a=action: run(a)).pack(side='left', padx=3)
    ttk.Label(form, textvariable=status, wraplength=1150).grid(row=6, column=0, columnspan=3, sticky='w', pady=8)
    def close():
        try:
            persist()
        except Exception as error:
            messagebox.showerror('연결정보 저장 실패', str(error), parent=window)
            return
        window.destroy()
    window.protocol('WM_DELETE_WINDOW', close)
    if Path(history_var.get()).is_file():
        window.after(100, lambda: run(load))
    window.mainloop()


if __name__ == '__main__':
    main()
