"""Inspection detail lines use original RT points, excluding separate REP records."""
import re
import pandas as pd


def collect_inspection_details(records, size_lengths):
    def text(value):
        return '' if pd.isna(value) else str(value).strip()

    def numeric(value):
        result = pd.to_numeric(str(value).replace(',', ''), errors='coerce')
        return 0.0 if pd.isna(result) else float(result)

    pipe_col = next((c for c in ('관경(Inch)', '관경') if c in records.columns), '')
    method_col = next((c for c in ('검사방법', 'Method', 'NDT') if c in records.columns), '')
    joint_col = next((c for c in ('조인트수', 'POINT', 'Point') if c in records.columns), '')
    if not pipe_col:
        return []
    groups = {}
    for _, row in records.iterrows():
        pipe = text(row.get(pipe_col, '')).upper()
        method = text(row.get(method_col, '')).upper() if method_col else ''
        if not pipe or (method == 'RT' and text(row.get('검사구분', '')).upper() == 'REP'):
            continue
        detail = groups.setdefault((method, pipe), {'points': 0.0, 'quantity': 0.0, 'reject_films': 0.0})
        detail['points'] += numeric(row.get(joint_col, 0)) if joint_col else 0
        detail['quantity'] += numeric(row.get('Usage', 0))
        if method == 'RT':
            detail['reject_films'] += numeric(row.get('불량수', 0))

    def sort_key(item):
        match = re.search(r'\d+(?:\.\d+)?', item[0][1])
        return (item[0][0], -(float(match.group()) if match else -1))

    result = []
    for (method, pipe), detail in sorted(groups.items(), key=sort_key):
        if not detail['points'] and detail['quantity'] > 0 and method in ('MT', 'PT', 'PAUT'):
            length = size_lengths.get(pipe if pipe.endswith('A') else f'{pipe}A', 0)
            if length:
                points = detail['quantity'] / length
                if abs(points - round(points)) < .05:
                    detail['points'] = round(points)
        result.append(dict(detail, method=method, pipe_size=pipe))
    return result
