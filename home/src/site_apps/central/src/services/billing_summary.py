"""Preserve billed quantities when daily-report details are unavailable."""
from collections import defaultdict
from decimal import Decimal
import re

from site_apps.central.src.services.billing_totals import grouped_billing_totals
from site_apps.central.src.rt_quantities import corrected_rt_history, rt_counts


def build_company_summaries(records, history, contracts):
    """Use billing quantities as the budget; history supplies attribution only."""
    history = corrected_rt_history(history)
    summaries = {}
    consumed = defaultdict(lambda: Decimal(0))
    contributions = defaultdict(dict)

    def number(value):
        try:
            result = Decimal(str(value or 0).replace(',', ''))
            return result if result.is_finite() else Decimal(0)
        except Exception:
            return Decimal(0)

    def company(value):
        value = str(value or '').strip()
        return value if value and '지역난방공사' not in value else '미지정'

    for record in records:
        remaining = number(record.get('qty')).quantize(Decimal('.0001'))
        if not remaining:
            continue
        method = str(record['ndt_type']).strip().upper()
        material = str(record['material_type']).strip()
        work_time = record.get('work_time', '일반')
        loc = record.get('loc', '')
        location = '플랜트(관리소)' if '관리소' in loc or '플랜트' in record.get('loc_type', loc) else '열배관'
        category = f'{location}_{work_time}_{method}_{material}'
        price = contracts.get(category, {}).get('c_price', 0) or record.get('unit_price', 0)
        if not price:
            price = number(record.get('subtotal')) / remaining
        spec = {'PAUT': '위상배열초음파검사(PAUT)', 'UT': '초음파탐상검사(UT)',
                'RT': '방사선투과검사(RT)', 'MT': '자분탐상검사(MT)',
                'PT': '액체침투탐상검사(PT)'}.get(method, method)
        if method in ('PAUT', 'UT', 'RT'):
            spec += ' ' + material.removeprefix(method + '_')
        if work_time == '야간':
            spec += ' 야간'
        unit = record.get('unit', '')

        def add(comp, section, line, quantity, detail_id=None):
            key = (section, line, spec, unit, price)
            rows = summaries.setdefault(comp, {})
            row = rows.setdefault(key, {'places': None, 'qty': Decimal(0),
                                        'calculated_amt': 0, '_details': set()})
            row['qty'] += quantity
            if detail_id is not None:
                row['_details'].add(detail_id)
                row['places'] = len(row['_details'])
            target = (comp, key)
            contributions[category][target] = contributions[category].get(target, Decimal(0)) + quantity * number(price)

        date = str(record.get('date', ''))[:10]
        for index, detail in enumerate(history.get(date, {}).get('ndt_results', [])):
            if remaining <= 0:
                break
            if record.get('source') == 'work_log' and index not in record.get('history_rows', []):
                continue
            if str(detail.get('검사방법', '')).strip().upper() != method:
                continue
            shift = str(detail.get('근무구분', '')).strip()
            if shift in ('주간', '일반', '야간') and (shift == '야간') != (work_time == '야간'):
                continue
            if method in ('PAUT', 'UT'):
                size = re.search(r'\d+', str(detail.get('관경', '')))
                limits = [int(v) for v in re.findall(r'\d+', material)]
                if not size or not limits:
                    continue
                size = int(size.group())
                matched = (size >= limits[0] if '이상' in material else
                           size <= limits[0] if '이하' in material else
                           size < limits[0] if '미만' in material else
                           min(limits) <= size <= max(limits))
                if not matched:
                    continue
            comp = company(detail.get('업체'))
            explicit_company = company(record.get('company'))
            if explicit_company != '미지정' and comp not in (explicit_company, '미지정'):
                continue
            field = {'PAUT': 'PAUT', 'UT': 'PAUT', 'RT': 'RT_OR', 'MT': 'MT', 'PT': 'PT'}.get(method)
            identity = (date, index)
            detail_qty = rt_counts(detail)[0] if method == 'RT' else detail.get(field)
            available = max(Decimal(0), number(detail_qty) - consumed[identity])
            quantity = min(remaining, available)
            if not quantity:
                continue
            consumed[identity] += quantity
            remaining -= quantity
            joint = str(detail.get('Joint No.', '')).strip()
            detail_id = (date, str(detail.get('구간', '')), str(detail.get('라인번호', '')), joint or index)
            add(comp if comp != '미지정' else explicit_company,
                str(detail.get('구간', '')).strip() or '미지정',
                str(detail.get('라인번호', '')).strip() or '미지정', quantity, detail_id)
        if remaining:
            add(company(record.get('company')), '미지정(작업일보 미연결)', '미지정', remaining)

    budgets = grouped_billing_totals(records, contracts)
    for category, targets in contributions.items():
        allocated = 0
        items = list(targets.items())
        for index, ((comp, key), raw) in enumerate(items):
            amount = budgets[category]['amount'] - allocated if index == len(items) - 1 else int(raw)
            summaries[comp][key]['calculated_amt'] += amount
            allocated += amount
    for rows in summaries.values():
        for row in rows.values():
            row['qty'] = float(row['qty'])
            del row['_details']
    if sum(row['calculated_amt'] for rows in summaries.values() for row in rows.values()) != sum(
            group['amount'] for group in budgets.values()):
        raise ValueError('업체별 요약과 기성내역의 공급가액이 일치하지 않습니다.')
    return summaries


def add_unmatched_record(summary, record, spec, unit, price, has_history_match=False):
    # An existing detail may already have been attributed to another billing
    # row/company. That is not missing history and must not become a fallback.
    if has_history_match:
        return False
    key = ('미지정(작업일보 미연결)', '미지정', spec, unit, price)
    row = summary.setdefault(key, {'places': None, 'qty': 0.0, 'fixed_amt': 0})
    row['qty'] += float(record.get('qty', 0) or 0)
    row['fixed_amt'] += int(record.get('subtotal', 0) or 0)
    return True


def allocate_summary_amounts(summary, ordered_keys):
    groups = defaultdict(list)
    for key in ordered_keys:
        if 'fixed_amt' in summary[key]:
            summary[key]['calculated_amt'] = summary[key]['fixed_amt']
        else:
            groups[(key[2], key[4])].append(key)
    for (_, price), keys in groups.items():
        total = int(sum(summary[key]['qty'] for key in keys) * price)
        allocated = 0
        for index, key in enumerate(keys):
            amount = total - allocated if index == len(keys) - 1 else int(summary[key]['qty'] * price)
            summary[key]['calculated_amt'] = amount
            allocated += amount
