"""Restore the district-heating template's indication columns per data row."""
from copy import copy
from openpyxl.styles import Side


def prepare_indication_row(ws, row, has_indication):
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == merged.max_row == row and merged.min_col >= 14 and merged.max_col <= 24:
            ws.unmerge_cells(str(merged))
    if has_indication:
        for first, last in ((14, 15), (16, 17), (18, 19), (20, 21), (23, 24)):
            ws.merge_cells(start_row=row, end_row=row, start_column=first, end_column=last)
    else:
        ws.merge_cells(start_row=row, end_row=row, start_column=14, end_column=24)
    if not has_indication:
        return
    groups = [(14, 15), (16, 17), (18, 19), (20, 21), (22, 22), (23, 24)]
    dotted = Side(style='dotted', color='000000')
    for first, last in groups:
        for column in range(first, last + 1):
            cell = ws.cell(row, column)
            border = copy(cell.border)
            if column != 14:
                border.left = dotted if column == first else Side()
            if column != 24:
                border.right = dotted if column == last else Side()
            cell.border = border
