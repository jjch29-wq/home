"""RT_OR is all inspected films; RT_RE is rejected films within that total."""
import math
import re
from copy import deepcopy


def films(value):
    try:
        result = float(str(value or 0).replace(',', ''))
    except (ValueError, TypeError):
        raise ValueError('RT 매수는 0 이상의 정수로 입력하세요.')
    if not math.isfinite(result) or result < 0 or not result.is_integer():
        raise ValueError('RT 매수는 0 이상의 정수로 입력하세요.')
    return int(result)


def film_grades(section_info):
    parts = [part.strip() for part in re.split(r'[,;\n]+', str(section_info or '')) if part.strip()]
    result = []
    for part in parts:
        pair = [value.strip() for value in part.split('/')]
        if len(pair) != 2:
            return None
        if pair[1] in ('1', '2', '3', '4'):
            grade, defect = int(pair[1]), pair[0]
        elif pair[0] in ('1', '2', '3', '4'):
            grade, defect = int(pair[0]), pair[1]
        else:
            return None
        if not defect:
            return None
        result.append((grade, defect.upper()))
    return result or None


def rt_result(section_info, expected_total=0):
    parts = [part.strip() for part in re.split(r'[,;\n]+', str(section_info or '')) if part.strip()]
    parsed = [film_grades(part) for part in parts]
    if any(grades and grades[0][0] >= 3 for grades in parsed):
        return '불합격'
    if not parts or any(grades is None for grades in parsed) or len(parts) < expected_total:
        return ''
    return '합격'


def retest_parent(joint):
    match = re.fullmatch(r'(.+\d)R\d*', str(joint or '').strip().upper())
    return match.group(1) if match else ''


def section_slots(section_info):
    return [part.strip() for part in re.split(r'[,;\n]', str(section_info or ''))]


def resolve_rt_row(row, context):
    """Retests inspect only the rejected film positions of their parent joint."""
    result = dict(row)
    parent_joint = retest_parent(row.get('Joint No.', ''))
    if str(row.get('검사방법', '')).strip().upper() != 'RT' or not parent_joint:
        result.pop('원본JointNo', None)
        result.pop('재촬영구간', None)
        return result
    candidates = []
    for candidate in context:
        if str(candidate.get('검사방법', '')).strip().upper() != 'RT':
            continue
        if str(candidate.get('Joint No.', '')).strip().upper() != parent_joint:
            continue
        if any(str(row.get(k, '')).strip() and str(candidate.get(k, '')).strip() != str(row[k]).strip()
               for k in ('업체', '구간', '라인번호')):
            continue
        candidates.append(candidate)
    identities = {tuple(str(candidate.get(k, '')).strip() for k in ('업체', '구간', '라인번호')) for candidate in candidates}
    if len(identities) != 1:
        return result  # No unambiguous parent: retain actual manually entered film counts.
    parent = candidates[-1]
    positions = [index for index, part in enumerate(section_slots(parent.get('구간정보', '')))
                 if (grades := film_grades(part)) and grades[0][0] >= 3]
    if not positions:
        return result
    slots = section_slots(row.get('구간정보', ''))
    # Legacy copied rows contain all original slots. Compact manual retest
    # entries may contain only the new film grades: assign them in target order.
    if len(slots) <= len(positions) and all(part for part in slots):
        expanded = [''] * (max(positions) + 1)
        for index, part in zip(positions, slots):
            expanded[index] = part
        slots = expanded
    selected = [''] * (max(positions) + 1)
    for index in positions:
        selected[index] = slots[index] if index < len(slots) else ''
    result.update({'원본JointNo': parent_joint, '재촬영구간': [i + 1 for i in positions],
                   'RT_OR': str(len(positions)), '구간정보': ','.join(selected).rstrip(',')})
    return result


def effective_section_info(row):
    slots = section_slots(row.get('구간정보', ''))
    positions = row.get('재촬영구간')
    if positions:
        return ','.join(slots[i - 1] for i in positions if 0 < i <= len(slots))
    return row.get('구간정보', '')


def default_rt_copy(row, context):
    """Initialize a newly copied RT row; retain any manually entered readings."""
    result = dict(row)
    if str(row.get('검사방법', '')).strip().upper() != 'RT' or str(row.get('구간정보', '')).strip():
        return result
    result = resolve_rt_row(result, context)
    positions = result.get('재촬영구간')
    if positions:
        slots = [''] * max(positions)
        for position in positions:
            slots[position - 1] = 'N/1'
    else:
        if retest_parent(row.get('Joint No.', '')):
            return result  # Unknown parent: do not guess a retest position.
        size = re.search(r'\d+', str(row.get('관경', '')))
        count = 3 if size and int(size.group()) <= 150 else films(row.get('RT_OR'))
        slots = ['N/1'] * count
    result['구간정보'] = ','.join(slots)
    result['RT_OR'] = str(sum(bool(slot) for slot in slots))
    return result


def widget_rt_values(row):
    values = {key: widget.get() for key, widget in row.items() if hasattr(widget, 'get')}
    values.update(getattr(row.get('구간정보'), '_rt_retake_meta', {}))
    return values


def rt_counts(row):
    grades = film_grades(effective_section_info(row))
    original = row.get('RT_OR')
    total = films(original) if str(original or '').strip() else len(grades or [])
    # A default or previously saved zero must not hide films actually read.
    # Keep a larger explicit total while the remaining grades are being entered.
    total = max(total, len(grades or []))
    rejects = sum(grade >= 3 for grade, _ in grades) if grades is not None else films(row.get('RT_RE', 0))
    if rejects > total:
        raise ValueError('RT_RE(불량매수)는 RT_OR(전체 검사매수)보다 클 수 없습니다.')
    return total, rejects


def sync_rt_widgets(row, from_details=False, all_rows=None):
    if row['검사방법'].get().strip().upper() != 'RT':
        return
    values = widget_rt_values(row)
    if all_rows is not None:
        values = resolve_rt_row(values, [widget_rt_values(item) if hasattr(item.get('검사방법'), 'get') else dict(item) for item in all_rows])
        frame = row['구간정보']
        frame._rt_retake_meta = {key: values[key] for key in ('원본JointNo', '재촬영구간') if key in values}
        if values['구간정보'] != frame.get():
            frame.set(values['구간정보'])
    if row['결과'].get().strip() in ('', '합격', '불합격'):
        try:
            expected = films(values['RT_OR'])
        except ValueError:
            expected = 0
        row['결과'].set(rt_result(effective_section_info(values), expected))
    if from_details and not values['구간정보'].strip():
        values['RT_RE'] = 0
    total, rejects = rt_counts(values)
    for key, value in (('RT_OR', total), ('RT_RE', rejects)):
        row[key].delete(0, 'end')
        row[key].insert(0, str(value))


def corrected_rt_history(history):
    """Recalculate report quantities in memory; never rewrite historical source files."""
    result = deepcopy(history)
    context = []
    for day in sorted(result):
        for index, row in enumerate(result[day].get('ndt_results', [])):
            normalized = resolve_rt_row(row, context + result[day]['ndt_results'])
            result[day]['ndt_results'][index] = normalized
            context.append(normalized)
    previous = {}
    previous_rejected = {}
    for day in sorted(result):
        data = result[day]
        rt_rows = [row for row in data.get('ndt_results', []) if str(row.get('검사방법', '')).strip().upper() == 'RT']
        quantities = data.get('qty_data', {})
        if rt_rows:
            totals = {}
            rejected = {}
            for row in rt_rows:
                total, defects = rt_counts(row)
                size = re.search(r'\d+', str(row.get('관경', '')))
                size = int(size.group()) if size else 0
                shift = str(row.get('근무구분', '') or '')
                if not shift:
                    shift = '야간' if str(row.get('규격', '')).endswith('야간') else ''
                key = 'RT_' + ('80A이하' if size <= 80 else '150A~100A') + ('-야간' if shift == '야간' else '')
                totals[key] = totals.get(key, 0) + total
                rejected[key] = rejected.get(key, 0) + defects
            totals['RT_소계'] = sum(totals.values())
            rejected['RT_소계'] = sum(rejected.values())
            for key, values in quantities.items():
                if key not in totals and not key.startswith(('RT_80A이하', 'RT_150A~100A')):
                    continue
                before = previous.get(key, films(values.get('전일누계', 0)))
                today = totals.get(key, 0)
                values.update({'전일누계': str(before), '금일작업': str(today), '총누계': str(before + today)})
                cumulative_rejected = previous_rejected.get(key, 0) + rejected.get(key, 0)
                cumulative_total = before + today
                values['불량'] = str(cumulative_rejected)
                values['불량률'] = f'{cumulative_rejected / cumulative_total * 100:.1f}' if cumulative_total and cumulative_rejected else ''
                try:
                    expected = films(values.get('예상량', 0))
                except ValueError:
                    expected = 0
                if expected:
                    values['공정률'] = f'{(before + today) / expected * 100:.1f}'
        for key, values in quantities.items():
            if key.startswith('RT_'):
                previous[key] = films(values.get('총누계', 0))
                if not rt_rows:
                    cumulative_rejected = previous_rejected.get(key, 0)
                    values['불량'] = str(cumulative_rejected)
                    values['불량률'] = f'{cumulative_rejected / previous[key] * 100:.1f}' if previous[key] and cumulative_rejected else ''
                previous_rejected[key] = films(values.get('불량', 0))
    return result
