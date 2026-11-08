"""Inspection detail lines use original RT points, excluding separate REP records."""
import re
import pandas as pd


def inspection_method_shift(row):
    def text(value):
        return '' if pd.isna(value) else str(value).strip()

    method_text = next((text(row.get(c, '')) for c in ('검사방법', 'TestMethod', 'Method', 'NDT')
                        if text(row.get(c, ''))), '').upper()
    match = re.match(r'^(PAUT|RT|UT|MT|PT|HT|VT|LT|ET)(?=$|[^A-Z])', method_text)
    method = match.group(1) if match else method_text

    def classify(value):
        value = value.upper()
        if '주야간' in value:
            return '주야간'
        if '야간' in value or 'NIGHT' in value or value == 'N':
            return '야간'
        if '일반' in value or '주간' in value or value in ('DAY', 'D'):
            return '주간'
        if '휴일' in value:
            return '휴일'
        return ''

    # The inspection's saved work type defines its quantity/price category.
    # Worker hours may span midnight without making the inspection a night-rate job.
    for column in ('작업형태', '근무구분', 'WorkType', 'Shift'):
        shift = classify(text(row.get(column, '')))
        if shift:
            return method, shift
    shift = classify(method_text)
    if shift:
        return method, shift
    shifts = set()
    for i in range(1, 11):
        time = text(row.get('WorkTime' if i == 1 else f'WorkTime{i}', ''))
        marker = re.match(r'^\((주간|야간|휴일|주야간)\)', time)
        if marker:
            shifts.add(marker.group(1))
    if len(shifts) == 1:
        return method, shifts.pop()
    return method, '주야간' if shifts else '미구분'


def collect_method_totals(records):
    """Keep different shifts, units and prices on separate report rows."""
    groups = {}

    def number(value):
        result = pd.to_numeric(str(value).replace(',', ''), errors='coerce')
        return 0.0 if pd.isna(result) else float(result)

    for _, row in records.iterrows():
        method, shift = inspection_method_shift(row)
        if not method:
            continue
        qty = number(row.get('Usage', row.get('검사량', 0)))
        price = number(row.get('단가', 0))
        travel = number(row.get('출장비', 0))
        total = number(row.get('검사비', 0))
        if not any((qty, price, travel, total)):
            continue
        unit = row.get('Unit', row.get('단위', ''))
        unit = '' if pd.isna(unit) else str(unit).strip()
        key = (method, shift, unit, price)
        item = groups.setdefault(key, dict(unit=unit, price=price, qty=0.0, travel=0.0, total=0.0))
        item['qty'] += qty
        item['travel'] += travel
        item['total'] += total
    result = {}
    order = {'주간': 0, '야간': 1, '휴일': 2, '주야간': 3, '미구분': 4}
    for (method, shift, unit, price), item in sorted(groups.items(), key=lambda v: (
            v[0][0], order[v[0][1]], v[0][2], v[0][3])):
        label = f'{method}({shift})'
        key = label if label not in result else f'{label}#{len(result) + 1}'
        result[key] = dict(item, label=label)
    return result


def collect_inspection_details(records, size_lengths):
    def text(value):
        return '' if pd.isna(value) else str(value).strip()

    def numeric(value):
        result = pd.to_numeric(str(value).replace(',', ''), errors='coerce')
        return 0.0 if pd.isna(result) else float(result)

    pipe_col = next((c for c in ('관경(Inch)', '관경') if c in records.columns), '')
    joint_col = next((c for c in ('조인트수', 'POINT', 'Point') if c in records.columns), '')
    if not pipe_col:
        return []
    groups = {}
    for _, row in records.iterrows():
        pipe = text(row.get(pipe_col, '')).upper()
        method, shift = inspection_method_shift(row)
        if not pipe or (method == 'RT' and text(row.get('검사구분', '')).upper() == 'REP'):
            continue
        detail = groups.setdefault((method, pipe, shift), {'points': 0.0, 'quantity': 0.0, 'reject_films': 0.0})
        detail['points'] += numeric(row.get(joint_col, 0)) if joint_col else 0
        detail['quantity'] += numeric(row.get('Usage', 0))
        if method == 'RT':
            detail['reject_films'] += numeric(row.get('불량수', 0))

    def sort_key(item):
        match = re.search(r'\d+(?:\.\d+)?', item[0][1])
        return (item[0][0], item[0][2] == '야간', item[0][2], -(float(match.group()) if match else -1))

    result = []
    for (method, pipe, shift), detail in sorted(groups.items(), key=sort_key):
        if not detail['points'] and detail['quantity'] > 0 and method in ('MT', 'PT', 'PAUT'):
            length = size_lengths.get(pipe if pipe.endswith('A') else f'{pipe}A', 0)
            if length:
                points = detail['quantity'] / length
                if abs(points - round(points)) < .05:
                    detail['points'] = round(points)
        result.append(dict(detail, method=method, pipe_size=pipe, shift=shift))
    return result
