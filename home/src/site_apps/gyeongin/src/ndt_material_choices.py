"""Editable work-log material choices from Gyeongin report evidence."""
import json
import re
from functools import lru_cache
from pathlib import Path


@lru_cache(maxsize=1)
def report_material_presets():
    path = Path(__file__).resolve().parent.parent / 'resources' / 'ndt_material_presets.json'
    return json.loads(path.read_text(encoding='utf-8'))


def material_choices(company='', method='', history=(), presets=None):
    data = presets if presets is not None else report_material_presets()
    method = str(method).strip().upper()
    company = str(company).strip()
    preferred = data.get('companies', {}).get(company, {}).get(method, [])
    method_values = data.get('methods', {}).get(method, [])
    all_values = [value for values in data.get('methods', {}).values() for value in values]
    return list(dict.fromkeys(str(value).strip()
                              for value in [*preferred, *history, *method_values, *all_values]
                              if str(value or '').strip()))


def filtered_material_choices(choices, query):
    def search_text(value):
        return re.sub(r'[\s\-_/]+', '', str(value)).upper()
    needle = search_text(query)
    if not needle or query in choices:
        return list(choices)
    # Each comma/semicolon-separated part remains editable as entered.
    tail = re.split(r'[,;]', str(query))[-1].strip()
    needle = search_text(tail)
    return [value for value in choices if needle in search_text(value)]
