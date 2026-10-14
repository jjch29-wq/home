"""검사방법별 장비 선택. 장비명 내부의 쉼표는 유지한다."""

import re

_METHOD_PREFIX = r'(?:PAUT|PMI|UT|MT|PT|RT|ETC|R-ray)(?=$|[^A-Za-z])'


def equipment_method(name):
    match = re.match(_METHOD_PREFIX, name.strip(), re.IGNORECASE)
    if not match:
        return ''
    prefix = match.group().upper()
    return 'RT' if prefix == 'R-RAY' else prefix


def filter_equipment_text(text, method):
    method = method.strip().upper()
    if method == 'PT':
        return ''
    parts = re.split(r',\s*(?=' + _METHOD_PREFIX + ')', text.strip(), flags=re.IGNORECASE)
    return ', '.join(part.strip() for part in parts if part.strip()
                     and equipment_method(part) in ('', method))
