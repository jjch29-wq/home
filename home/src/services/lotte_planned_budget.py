"""Reviewed Lotte pre-cost workbook baseline (2026-09-05).

Keep the source's unrounded numeric values; round only for presentation.
"""

import json
from pathlib import Path


def load_planned_budget():
    path = Path(__file__).with_name('lotte_planned_budget.json')
    cells = json.loads(path.read_text(encoding='utf-8'))['cells']

    def value(column, row):
        return cells.get(f'{column}{row}', '')

    labor = {
        value('C', row): dict(personnel=str(value('D', row)),
                              period=str(value('F', row)),
                              unit_price=str(value('H', row)))
        for row in (*range(16, 24), *range(26, 29))
    }
    material = [dict(item=value('A', row), spec=value('C', row),
                     qty=str(value('F', row)), unit=value('G', row),
                     price=str(value('H', row))) for row in range(34, 44)]
    expense = {
        'site_expense': [dict(cat=value('A', row), cont=value('C', row),
                              ppl=value('E', row), qty=str(value('F', row)),
                              unit=value('G', row), price=str(value('H', row)))
                         for row in range(49, 53)],
        'rental': [],
        'outsource': [dict(cat=value('A', 69), work=value('C', 69),
                           count=str(value('G', 69)).replace('매', ''),
                           price=str(value('H', 69)))],
        'depreciation': [dict(item=value('A', row), spec=value('C', row),
                              life=str(value('E', row)), qty=str(value('F', row)),
                              days=str(value('G', row)), rate=str(value('H', row)))
                         for row in range(82, 88)],
    }
    return dict(labor=labor, material=material, expense=expense,
                revenue=value('A', 11), period=value('A', 5))
