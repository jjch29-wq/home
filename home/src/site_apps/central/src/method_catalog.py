"""Central-only, read-only method/role rules for daily entry suggestions.

Unknown items are opt-in; these rules never rewrite inventory or saved records.
"""
import re


def current_method(app):
    field = getattr(app, 'cb_daily_test_method', None)
    return field.get().strip().upper() if field is not None else ''


def text(value):
    result = str(value).strip()
    return '' if result.lower() in ('nan', 'none') else result


def classify(label, category=''):
    """Return (methods, role); empty methods means classification needs review."""
    s = text(label).upper()
    cat = text(category).upper()
    has = lambda pattern: re.search(pattern, s, re.I) is not None
    if has(r'MT약품|자분|백색.?페인트'):
        return {'MT'}, 'material'
    if has(r'PT약품|침투제|현상제|세척제'):
        return {'PT'}, 'material'
    if cat == 'FILM' or has(r'CARESTREAM|AGFA|필름|\bFILM\b'):
        return {'RT'}, 'material'
    if has(r'BLACK\s*LIGHT|UV.?METER|자외선|조도계|조도측정|조도.?자외선'):
        return {'MT', 'PT'}, 'aux'
    if has(r'SURVEY|DOSIMETER|선량계|COLLIMATOR|콜리메이터|R/CONTROL|GUIDE.?TUBE|가이드.?튜브'):
        return {'RT'}, 'aux'
    if has(r'SCANNER|스캐너|COBRA|MICROBE|BTG006'):
        return {'PAUT'}, 'aux'
    if has(r'PROBE|프로브|탐촉자|WEDGE|웨지|시험편|TEST.?BLOCK|리프팅|LIFTING|팔각|게이지'):
        if cat.startswith('MT') or has(r'자분|팔각|리프팅|LIFTING'):
            return {'MT'}, 'aux'
        if cat.startswith('PT'):
            return {'PT'}, 'aux'
        if has(r'PAUT|PHASED|\bA(?:2|10|12|15|16|17|27|31|32)\b'):
            return {'PAUT'}, 'aux'
        if has(r'시험편|TEST.?BLOCK'):
            return {'UT', 'PAUT'}, 'aux'
        if cat.startswith('UT') or has(r'\bUT\b'):
            return {'UT'}, 'aux'
        return set(), 'aux'
    if has(r'PAUT|OMNI.?SCAN|OMNI2|\bMX2\b|\bX3\b'):
        return {'PAUT'}, 'main'
    if has(r'YOKE|요크|MY-?2|AC POWER SUPPLY') or cat == 'MT장비':
        return {'MT'}, 'main'
    if has(r'SE-?75|IR-?192|X.?RAY|방사선.*장치|방사선투과검사장비'):
        return {'RT'}, 'main'
    if has(r'SITESCAN|MASTERSCAN|초음파.*탐상|두께측정'):
        return {'UT'}, 'main'
    for method in ('RT', 'MT', 'PT', 'UT'):
        if cat == method + '장비':
            return {method}, 'main'
        if cat == method + '자재':
            return {method}, 'material'
    return set(), ''


def row_classification(row):
    return classify(' '.join(text(row.get(k, '')) for k in ('품목명', '모델명')),
                    row.get('품목군코드', ''))


def row_allowed(app, row, role='material', include_unknown=False):
    method = current_method(app)
    if not method:
        return True
    methods, item_role = row_classification(row)
    return (method in methods and role == item_role) or (not methods and include_unknown)


def material_candidates(app):
    inventory = getattr(app, 'materials_df', None)
    if inventory is None:
        return []
    result = []
    for _, row in inventory.iterrows():
        if str(row.get('Active', 1)).strip().lower() in ('0', '0.0', 'false'):
            continue
        if row_allowed(app, row):
            result.append(app.get_material_display_name(row['MaterialID']))
    return sorted(set(filter(None, result)))
