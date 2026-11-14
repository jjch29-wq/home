"""Read-only material history using the inventory calculation's contribution rules."""
from copy import copy
import pandas as pd
from site_apps.central.src.models.stock_calculation import build_stock_lookups, stock_breakdown


def material_history(app, mat):
    lookups = build_stock_lookups(app)
    summary = stock_breakdown(app, mat, lookups)
    target = app.normalize_id(mat['MaterialID'])
    name = app.clean_nan(mat.get('품목명', ''))
    combined = (name + app.clean_nan(mat.get('모델명', ''))).replace(' ', '')
    from site_apps.central.src.models.stock_calculation import chemical_key
    chemical = chemical_key(combined)
    rows = []
    is_equipment = not app._is_consumable_material(combined, '')
    def norm(value):
        return ''.join(app.clean_nan(value).upper().split())
    equipment_names = set()
    if is_equipment:
        # Accept exact display names, plus unambiguous legacy names/models.
        display = getattr(app, 'get_material_display_name', None)
        if callable(display):
            equipment_names.add(norm(display(mat['MaterialID'])))
        for field in ('품목명', '모델명'):
            candidate = norm(mat.get(field, ''))
            count = sum(norm(row.get(field, '')) == candidate for _, row in app.materials_df.iterrows())
            if candidate and count == 1:
                equipment_names.add(candidate)
        equipment_names.discard('')

    def add(entry, kind, change, note=''):
        date = pd.to_datetime(entry.get('Date', ''), errors='coerce')
        rows.append(dict(date='' if pd.isna(date) else date.strftime('%Y-%m-%d'),
                         kind=kind, site=app.clean_nan(entry.get('Site', '')),
                         increase=max(change, 0), decrease=max(-change, 0),
                         note=note or app.clean_nan(entry.get('Note', ''))))

    for _, entry in app.transactions_df.iterrows():
        if not target or app.normalize_id(entry.get('MaterialID')) != target:
            continue
        note = app.clean_nan(entry.get('Note', ''))
        kind = str(entry.get('Type', ''))
        if '자동 차감' in note:
            add(entry, '자동 차감 (중복 제외)', 0, note)
        elif kind in ('IN', 'OUT'):
            qty = float(entry.get('Quantity', 0))
            # Match the inventory's absolute aggregate OUT, including reversals.
            change = qty if kind == 'IN' else -qty * (1 if lookups[1].get(target, 0) >= 0 else -1)
            add(entry, '입고' if kind == 'IN' else '출고', change)

    daily = getattr(app, 'daily_usage_df', pd.DataFrame())
    for position, (_, entry) in enumerate(daily.iterrows()):
        if is_equipment:
            from site_apps.central.src.aux_equipment import names
            used_names = names(entry.get('장비명', '')) + names(entry.get('보조장비', ''))
            id_match = bool(target and app.normalize_id(entry.get('MaterialID')) == target)
            if id_match or any(norm(value) in equipment_names for value in used_names):
                info = ['재고 차감 없음']
                for field in ('검사방법', '장비명', '보조장비', 'User', 'WorkTime', 'Note'):
                    value = app.clean_nan(entry.get(field, ''))
                    if value:
                        info.append(value)
                add(entry, '장비 사용', 0, ' · '.join(info))
            continue
        single = copy(app)
        single.transactions_df = app.transactions_df.iloc[:0].copy()
        single.daily_usage_df = daily.iloc[[position]].copy()
        one = build_stock_lookups(single)
        if chemical:
            qty = stock_breakdown(single, mat, one)['ndt_usage']
        else:
            qty = one[2].get(target, 0) + (one[3].get(name, 0) if name else 0)
        if qty:
            add(entry, '현장 사용', -qty)
    rows.sort(key=lambda row: row['date'], reverse=True)
    return summary, rows
