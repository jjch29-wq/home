"""검사 항목별 담당자와 일일 근무 실적 배분."""


def worker_names(record):
    names = []
    for index in range(1, 11):
        name = str(record.get('User' if index == 1 else f'User{index}', '')).strip()
        if name.casefold() not in ('', 'nan', 'none', 'null') and name not in names:
            names.append(name)
    return ', '.join(names)


def merge_worker_attendance(items):
    """같은 담당자의 검사방법을 합치고 근무시간·OT는 한 번 표시한다."""
    grouped = {}
    for item in items:
        name = str(item.get('names', '')).strip()
        company = str(item.get('company', '')).strip()
        if not name:
            continue
        key = (company, name)
        if key not in grouped:
            grouped[key] = dict(names=name, names_display=name, company=company,
                                method='', ot_hours='', ot_amount='', methods=[])
        row = grouped[key]
        for method in str(item.get('method', '')).split(','):
            method = method.strip()
            if method and method not in row['methods']:
                row['methods'].append(method)
        for field in ('ot_hours', 'ot_amount'):
            value = str(item.get(field, '')).strip()
            if value.lower() not in ('', 'none', 'nan') and row[field] in ('', '0', '0.0'):
                row[field] = value
    for row in grouped.values():
        row['method'] = ', '.join(row.pop('methods'))
    return list(grouped.values())


def include_shared_worker_methods(items, records):
    """한 현장·날짜의 단일 공통 담당자에게 담당자 미배정 항목을 연결한다."""
    result = [dict(item) for item in items]
    workers = {(str(item.get('company', '')).strip(), str(item.get('names', '')).strip())
               for item in items if str(item.get('names', '')).strip()}
    contexts = {(str(row.get('Date', '')).split(' ')[0], str(row.get('Site', '')).strip()) for row in records}
    if len(workers) != 1 or len(contexts) != 1:
        return result
    company, name = next(iter(workers))
    for record in records:
        if worker_names(record):
            continue
        method = str(record.get('검사방법', '')).strip()
        record_company = str(record.get('업체명', '')).strip()
        if method and method.lower() != 'nan' and record_company == company:
            result.append(dict(names=name, company=company, method=method, ot_hours='', ot_amount=''))
    return result


def update_work_details(record, inspection_item, workers):
    """불러온 한 행의 검사품명과 작업자 입력을 함께 수정한다."""
    updates = {'검사품명': inspection_item.strip()}
    for index in range(1, 11):
        suffix = '' if index == 1 else str(index)
        worker = workers.get(index, {})
        for column, field in [('User', 'name'), ('WorkTime', 'time'), ('OT', 'ot'),
                              ('Meal', 'meal'), ('WorkerMethod', 'method')]:
            value = worker.get(field, '')
            if column in ('OT', 'Meal'):
                import math
                value = float(str(value).replace(',', '').strip() or 0)
                if not math.isfinite(value):
                    raise ValueError('OT와 일비는 유효한 숫자로 입력해주세요.')
            updates[column + suffix] = value
    changed = any(str(record.get(key, '')) != str(value) for key, value in updates.items())
    record.update(updates)
    return changed


def apply_work_details(frame, index, record):
    """Excel에서 빈 문자 열이 float64로 읽힌 경우도 안전하게 저장한다."""
    columns = ['검사품명'] + [prefix + ('' if i == 1 else str(i))
        for i in range(1, 11) for prefix in ('User', 'WorkTime', 'OT', 'Meal', 'WorkerMethod')]
    for column in columns:
        numeric = column.startswith(('OT', 'Meal'))
        if column not in frame.columns:
            frame[column] = 0.0 if numeric else ''
        if not numeric:
            frame[column] = frame[column].astype(object)
        frame.at[index, column] = record[column]


def build_worker_assignments(items, workers):
    assignments = []
    counted = set()
    for item in items:
        method = str(item.get('method', '')).strip().upper()
        selected = item.get('workers')
        # 작업자를 등록하기 전에 만든 품목은 저장 시 현재 작업자 기록을 사용한다.
        if selected == [] and not item.get('workers_explicit', True):
            selected = None
        row = {}
        for index in range(1, 11):
            suffix = '' if index == 1 else str(index)
            worker = workers.get(index, {})
            name = worker.get('name', '').strip()
            assigned = bool(name) and (
                index in selected if selected is not None else
                not worker.get('method') or worker['method'].upper() == method
            )
            key = name.casefold()
            first = assigned and key not in counted
            row[f'User{suffix}'] = name if assigned else ''
            row[f'WorkerMethod{suffix}'] = method if assigned else ''
            for field, source in [('WorkTime', 'time'), ('OT', 'ot'), ('Meal', 'meal')]:
                row[f'{field}{suffix}'] = worker.get(source, '') if first else ''
            if assigned:
                counted.add(key)
        assignments.append(row)
    return assignments
