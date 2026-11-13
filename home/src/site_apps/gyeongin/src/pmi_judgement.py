"""Evaluate selected PMI limits against measured chemistry."""
from decimal import Decimal, InvalidOperation

from site_apps.gyeongin.src.pmi_criteria import normalize_ranges

DEFAULT_TOLERANCE_PERCENT = 10


def judge_measurement(row, criteria, elements, tolerance_percent):
    criterion_id = row.get('criterion_id')
    criterion = next((item for item in criteria if item['id'] == criterion_id), None)
    if criterion is None and not criterion_id and len(criteria) == 1:
        criterion = criteria[0]
    if criterion is None:
        return 'HOLD', '적용할 재질 기준을 선택하세요.'
    try:
        tolerance = Decimal(str(tolerance_percent)) / 100
        if not tolerance.is_finite() or not 0 <= tolerance <= 1:
            raise ValueError
        ranges = normalize_ranges(criterion['ranges'])
    except (ValueError, InvalidOperation, KeyError):
        return 'HOLD', '합격기준 또는 허용오차를 확인하세요.'
    failures = []
    missing = []
    checked = 0
    for symbol in elements:
        limit = ranges.get(symbol)
        if limit in ('-', 'Balance'):
            continue
        if not limit:
            missing.append(symbol)
            continue
        checked += 1
        try:
            value = Decimal(str((row.get('elements') or {}).get(symbol, '')).strip().removesuffix('%'))
            if not value.is_finite() or not 0 <= value <= 100:
                raise InvalidOperation
        except InvalidOperation:
            missing.append(symbol)
            continue
        if '~' in limit:
            lower, upper = map(Decimal, limit.split('~'))
        elif limit.startswith('<='):
            lower, upper = Decimal(0), Decimal(limit[2:])
        else:
            lower, upper = Decimal(limit[2:]), Decimal(100)
        lower = max(Decimal(0), lower * (1 - tolerance))
        upper = min(Decimal(100), upper * (1 + tolerance))
        if not lower <= value <= upper:
            failures.append(f'{symbol} {value} / {lower}~{upper}')
    if failures:
        return 'FAIL', '; '.join(failures)
    if missing:
        return 'HOLD', '측정값·기준 확인: ' + ', '.join(missing)
    if not checked:
        return 'HOLD', '수치로 비교할 기준 원소가 없습니다.'
    return 'PASS', criterion['label']


def report_result(results):
    if 'FAIL' in results:
        return 'FAIL'
    if not results or 'HOLD' in results:
        return 'HOLD'
    return 'PASS'
