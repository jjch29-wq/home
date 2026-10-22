"""Contract-item totals: sum quantities before truncating to whole won."""
from decimal import Decimal


def grouped_billing_totals(records, contracts):
    groups = {}
    for record in records:
        loc = record.get('loc', '')
        location = ('플랜트(관리소)' if '관리소' in loc or '플랜트' in record.get('loc_type', loc)
                    else '열배관')
        key = f"{location}_{record.get('work_time', '일반')}_{record['ndt_type']}_{record['material_type']}"
        quantity = Decimal(str(record.get('qty', 0) or 0)).quantize(Decimal('.0001'))
        price = contracts.get(key, {}).get('c_price', 0) or record.get('unit_price', 0)
        if not price and quantity:
            price = Decimal(str(record.get('subtotal', 0))) / quantity
        group = groups.setdefault(key, {'qty': Decimal(0), 'raw_amount': Decimal(0)})
        group['qty'] += quantity
        group['raw_amount'] += quantity * Decimal(str(price or 0))
    return {key: {'qty': float(value['qty']), 'amount': int(value['raw_amount'])}
            for key, value in groups.items()}


def grouped_supply_amount(records, contracts):
    return sum(group['amount'] for group in grouped_billing_totals(records, contracts).values())
