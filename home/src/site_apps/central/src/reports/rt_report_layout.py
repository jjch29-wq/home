"""Format the RT cover's bilingual equipment headings."""
from copy import copy
from datetime import date
import re


def header_stamp_anchor(ws):
    """Align the seal's right edge with the company/title boundary."""
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU

    boundary_column = next(
        (cell.column for row in ws.iter_rows(min_row=1, max_row=4)
         for cell in row
         if 'RADIOGRAPHIC' in str(cell.value or '').upper()
         or '방사선' in str(cell.value or '')),
        8,
    )
    widths = []
    for column in range(1, boundary_column):
        width = ws.sheet_format.defaultColWidth or 8.43
        for dim in ws.column_dimensions.values():
            if dim.min and dim.min <= column <= (dim.max or dim.min):
                width = 0 if dim.hidden else dim.width
                break
        widths.append(int(width * 7 + 5) if width else 0)
    height = sum((ws.row_dimensions[row].height or
                  ws.sheet_format.defaultRowHeight or 15) * 4 / 3
                 for row in range(1, 5))
    size = min(72, height - 6)
    left = max(0, sum(widths) - size)
    column = 0
    while column < len(widths) - 1 and left >= widths[column]:
        left -= widths[column]
        column += 1
    return OneCellAnchor(
        _from=AnchorMarker(col=column, colOff=pixels_to_EMU(left), row=0,
                           rowOff=pixels_to_EMU(max(0, (height - size) / 2))),
        ext=XDRPositiveSize2D(pixels_to_EMU(size), pixels_to_EMU(size)))


def report_date_from_exam_period(value):
    """Resolve a full or abbreviated period end, preserving unrecognized input."""
    text = str(value or '').strip()
    match = re.fullmatch(
        r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*[~～∼–—]\s*'
        r'(?:(\d{4})[-/.])?(?:(\d{1,2})[-/.])?(\d{1,2})', text)
    if not match:
        return text
    year, month, _, end_year, end_month, end_day = match.groups()
    try:
        end = date(int(end_year or year), int(end_month or month), int(end_day))
    except ValueError:
        return text
    return f'{end.year} 년   {end.month:02d} 월   {end.day:02d} 일'


def fit_film_headings(ws):
    labels = {
        'FILM TYPE': 'FILM TYPE\n필 름 종 류',
        'FILM SIZE': 'FILM SIZE\n필 름 크 기',
        'SCREEN': 'SCREEN\n스 크 린',
    }
    for row in ws.iter_rows(min_row=7, max_row=15):
        for cell in row:
            text = str(cell.value or '').strip()
            for english, label in labels.items():
                if not text.startswith(english):
                    continue
                cell.value = label
                alignment = copy(cell.alignment)
                alignment.horizontal = 'left'
                alignment.vertical = 'center'
                alignment.wrap_text = True
                alignment.shrink_to_fit = False
                cell.alignment = alignment
                break
