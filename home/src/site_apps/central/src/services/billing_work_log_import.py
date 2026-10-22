"""Create central billing records directly from saved supervisor-log details."""
from datetime import date
from decimal import Decimal, InvalidOperation
import re

from site_apps.central.src.rt_quantities import corrected_rt_history, rt_counts
from site_apps.central.src.services.ndt_calculator import calculate_billing


def build_work_log_billing_records(history, start, end, *, material_costs, labor_costs,
                                   contracts=None, location='열배관', overhead_rate=1.1,
                                   technical_fee_rate=.2, company_filter='전체'):
    start_date = date.fromisoformat(start) if start else date.min
    end_date = date.fromisoformat(end) if end else date.max
    if start_date > end_date:
        raise ValueError('조회 시작일이 종료일보다 늦습니다.')
    history = corrected_rt_history(history)
    groups = {}
    for day in sorted(history):
        selected_date = date.fromisoformat(day)
        if not start_date <= selected_date <= end_date:
            continue
        for index, detail in enumerate(history[day].get('ndt_results', [])):
            method = str(detail.get('검사방법', '')).strip().upper()
            if method not in ('PAUT', 'RT', 'MT', 'PT', 'UT'):
                continue
            company = str(detail.get('업체', '') or '').strip() or '미지정'
            if company_filter and company_filter != '전체' and company != company_filter:
                continue
            try:
                raw = rt_counts(detail)[0] if method == 'RT' else detail.get(method, 0)
                quantity = Decimal(str(raw or 0).replace(',', ''))
                if not quantity.is_finite() or quantity < 0:
                    raise ValueError()
            except (ValueError, InvalidOperation):
                raise ValueError(f'{day} {index + 1}행의 검사수량을 확인해 주세요.')
            if not quantity:
                continue
            spec = str(detail.get('규격', '') or '').strip()
            shift = str(detail.get('근무구분', '') or '').strip()
            if not shift:
                shift = '야간' if spec.endswith('야간') else '주간'
            work_time = '야간' if shift == '야간' else '일반'
            if method == 'RT':
                match = re.search(r'[X×*]\s*(6|12)\s*["”″]?', spec, re.I)
                if not match:
                    raise ValueError(f'{day} {index + 1}행의 RT 필름 규격(6″/12″)을 작업/감독일보에 입력해 주세요.')
                material = f'3 1/3 x {match.group(1)}"'
            elif method in ('PAUT', 'UT'):
                match = re.search(r'\d+', str(detail.get('관경', '')))
                if not match:
                    raise ValueError(f'{day} {index + 1}행의 관경을 확인해 주세요.')
                size = int(match.group())
                material = ('300A 이상' if size >= 300 else '250A' if size == 250 else
                            '200A' if size == 200 else '150A-125A' if 125 <= size <= 150 else
                            '100A 이하' if size <= 100 else '')
                if not material:
                    raise ValueError(f'{day} {index + 1}행의 관경에 해당하는 계약 규격이 없습니다.')
            else:
                material = method
            section = str(detail.get('구간', '') or '').strip()
            line = str(detail.get('라인번호', '') or '').strip()
            key = (day, company, method, material, work_time, section, line)
            group = groups.setdefault(key, {'qty': Decimal(0), 'indices': []})
            group['qty'] += quantity
            group['indices'].append(index)
    records = []
    for (day, company, method, material, work_time, section, line), group in groups.items():
        qty = float(group['qty'].quantize(Decimal('.0001')))
        costs = calculate_billing(quantity=qty, adjusted_quantity=qty,
                                  material_key=f'{method}_{material}', ndt_type=method,
                                  location=location, work_time=work_time,
                                  material_costs=material_costs, labor_costs=labor_costs,
                                  overhead_rate=overhead_rate, technical_fee_rate=technical_fee_rate)
        category = f'{location}_{work_time}_{method}_{material}'
        price = (contracts or {}).get(category, {}).get('c_price', 0) or costs['unit_price']
        if price <= 0:
            raise ValueError(f'{method} {material} {work_time}의 계약 단가를 확인해 주세요.')
        subtotal = int(Decimal(str(qty)) * Decimal(str(price)))
        costs['tech'] += subtotal - costs['subtotal']
        costs.update(unit_price=int(price), subtotal=subtotal,
                     vat=int(Decimal(subtotal) * Decimal('.1')))
        costs['total_amount'] = subtotal + costs['vat']
        records.append(dict(date=day, company=company, actual_company=company,
                            loc=f'[{location}] 중앙지사 / {section} / {line}', loc_type=location,
                            ndt_type=method, material_type=material, work_time=work_time,
                            qty=qty, adjusted_qty=qty, corr=1.0, unit='매' if method == 'RT' else 'M',
                            section=section, line=line, source='work_log',
                            history_rows=group['indices'], **costs))
    return records
