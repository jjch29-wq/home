"""검사 항목별 담당자와 일일 근무 실적 배분."""


def worker_names(record):
    names = []
    for index in range(1, 11):
        name = str(record.get('User' if index == 1 else f'User{index}', '')).strip()
        if name.casefold() not in ('', 'nan', 'none', 'null') and name not in names:
            names.append(name)
    return ', '.join(names)


def build_worker_assignments(items, workers):
    assignments = []
    counted = set()
    for item in items:
        method = str(item.get('method', '')).strip().upper()
        selected = item.get('workers')
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
