"""Gyeongin certificate photo selection and existing-form photo appendices."""
from copy import copy, deepcopy
from io import BytesIO
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Font
from PIL import Image as PillowImage, ImageTk


def batch_photos(history, batch, history_dir, all_rows):
    """Return selectable same-method photos, auto-including unambiguous links."""
    result = []
    ids = {row.get('photo_row_id') for row in batch['rows'] if row.get('photo_row_id')}
    dates = {batch['date']} | {row.get('_source_date') for row in batch['rows'] if row.get('_source_date')}
    for date in sorted(dates):
        for photo in history.get(date, {}).get('process_photos', []):
            if photo.get('process') != batch['method']:
                continue
            if photo.get('company') and photo['company'] != batch['company']:
                continue
            linked = photo.get('row_id') in ids
            if not photo.get('row_id') and photo.get('joint_no'):
                matches = [row for row in all_rows if row.get('검사방법') == photo['process']
                           and row.get('Joint No.', '') == photo['joint_no']
                           and row.get('라인번호', '') == photo.get('line_no', '')
                           and row.get('구간', '') == photo.get('section', '')]
                linked = len(matches) == 1 and matches[0] in batch['rows']
            if photo.get('row_id') and not linked:
                continue
            item = deepcopy(photo)
            item['file_path'] = str((Path(history_dir) / photo['file_path']).resolve())
            item['include'] = linked and photo.get('include', True)
            result.append(item)
    return result


def edit_certificate_photos(parent, photos, on_apply):
    items = deepcopy(photos)
    win = tk.Toplevel(parent)
    win.title('성적서 사진 선택·편집')
    win.geometry('980x570')
    win.transient(parent)
    win.grab_set()
    body = ttk.Frame(win, padding=8)
    body.pack(fill='both', expand=True)
    columns = ('include', 'location', 'description', 'file')
    tree = ttk.Treeview(body, columns=columns, show='headings', selectmode='extended')
    for column, label in zip(columns, ('출력', '대상', '설명', '파일')):
        tree.heading(column, text=label)
        tree.column(column, width=55 if column == 'include' else 155)
    tree.pack(side='left', fill='both', expand=True)
    preview = ttk.Label(body, text='사진을 선택하세요.', width=35, anchor='center')
    preview.pack(side='right', fill='both', padx=8)
    status = tk.StringVar()

    def refresh(selection=()):
        tree.delete(*tree.get_children())
        for i, photo in enumerate(items):
            tree.insert('', 'end', iid=str(i), values=('✓' if photo.get('include', True) else '',
                        photo.get('location', ''), photo.get('description', ''), Path(photo['file_path']).name))
        tree.selection_set(selection)
        status.set(f'등록 {len(items)}장 / 출력 {sum(p.get("include", True) for p in items)}장 · 페이지당 4장')

    def show_preview(_event=None):
        selected = tree.selection()
        if not selected:
            return
        try:
            with PillowImage.open(items[int(selected[0])]['file_path']) as source:
                image = source.copy()
            image.thumbnail((320, 400))
            preview.image = ImageTk.PhotoImage(image, master=win)
            preview.configure(image=preview.image, text='')
        except (OSError, ValueError) as exc:
            preview.configure(image='', text=f'사진 확인 필요\n{exc}')

    def toggle():
        selected = tree.selection()
        for iid in selected:
            items[int(iid)]['include'] = not items[int(iid)].get('include', True)
        refresh(selected)

    def describe():
        selected = tree.selection()
        if not selected:
            return
        value = simpledialog.askstring('사진 설명', '선택 사진에 적용할 설명',
                                       initialvalue=items[int(selected[0])].get('description', ''), parent=win)
        if value is not None:
            for iid in selected:
                items[int(iid)]['description'] = value.strip()
            refresh(selected)

    def move(direction):
        selected = {int(i) for i in tree.selection()}
        for i in sorted(selected, reverse=direction > 0):
            target = i + direction
            if 0 <= target < len(items) and target not in selected:
                items[i], items[target] = items[target], items[i]
                selected.remove(i)
                selected.add(target)
        refresh(tuple(str(i) for i in sorted(selected)))

    def apply():
        try:
            if on_apply(items) is not False:
                win.destroy()
        except (OSError, ValueError) as exc:
            messagebox.showerror('사진 저장 실패', str(exc), parent=win)

    tree.bind('<<TreeviewSelect>>', show_preview)
    tree.bind('<Double-1>', lambda _event: toggle())
    buttons = ttk.Frame(win, padding=8)
    buttons.pack(fill='x')
    for label, command in [('출력 포함/제외', toggle), ('설명 수정', describe),
                           ('위로', lambda: move(-1)), ('아래로', lambda: move(1))]:
        ttk.Button(buttons, text=label, command=command).pack(side='left', padx=3)
    ttk.Button(buttons, text='적용', command=apply).pack(side='right')
    ttk.Button(buttons, text='취소', command=win.destroy).pack(side='right', padx=3)
    ttk.Label(win, textvariable=status, padding=8).pack(anchor='w')
    refresh()
    parent.wait_window(win)


def append_photo_pages(wb, batch, template_dir):
    photos = [p for p in batch.get('certificate_photos', []) if p.get('include', True)]
    if not batch.get('include_photos', True) or not photos or batch['method'] not in ('MT', 'PT'):
        return
    for photo in photos:
        with PillowImage.open(photo['file_path']) as image:
            image.verify()
    template = openpyxl.load_workbook(Path(template_dir) / f'Template_Cert_{batch["method"]}_Photos.xlsx')
    source = template.active
    original_count = len(wb.worksheets)
    page_count = (len(photos) + 3) // 4
    total = original_count + page_count
    for page_index in range(page_count):
        ws = wb.create_sheet(f'{batch["method"]}(사진 {page_index + 1})')
        for cells in source:
            for cell in cells:
                if isinstance(cell, MergedCell):
                    continue
                target = ws.cell(cell.row, cell.column, cell.value)
                for attribute in ('font', 'fill', 'border', 'alignment', 'protection'):
                    setattr(target, attribute, copy(getattr(cell, attribute)))
                target.number_format = cell.number_format
        for merged in source.merged_cells:
            ws.merge_cells(str(merged))
        for key, dimension in source.row_dimensions.items():
            ws.row_dimensions[key] = copy(dimension)
        for key, dimension in source.column_dimensions.items():
            ws.column_dimensions[key] = copy(dimension)
        for attribute in ('page_setup', 'page_margins', 'print_options', 'sheet_properties', 'HeaderFooter'):
            setattr(ws, attribute, deepcopy(getattr(source, attribute)))
        ws.print_area = source.print_area
        ws.sheet_view.showGridLines = False
        for original in source._images:
            payload = original._data()
            original.ref = BytesIO(payload)
            image = Image(BytesIO(payload))
            image.anchor = deepcopy(original.anchor)
            ws.add_image(image)
        last_col = 24 if batch['method'] == 'MT' else 18
        customer_col = 19 if batch['method'] == 'MT' else 15
        ws.cell(4, customer_col, batch.get('customer', ''))
        ws.cell(6, customer_col, batch.get('report_no', ''))
        if batch['method'] == 'MT':
            ws['V3'] = str(original_count + page_index + 1)
            ws['X3'] = str(total)
        else:
            ws['Q3'] = f'{original_count + page_index + 1} of'
            ws['R3'] = str(total)
        half = last_col // 2
        widths = [int((ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width or 8.43) * 7 + 5)
                  for c in range(1, last_col + 1)]
        for index, photo in enumerate(photos[page_index * 4:page_index * 4 + 4]):
            col = 1 if index % 2 == 0 else half + 1
            row = 10 if index < 2 else 23
            width = sum(widths[col - 1:col - 1 + half]) - 16
            height = sum((ws.row_dimensions[r].height or 15) * 96 / 72 for r in range(row, row + 11)) - 12
            image = Image(photo['file_path'])
            ratio = min(width / image.width, height / image.height)
            image.width, image.height = image.width * ratio, image.height * ratio
            image.anchor = OneCellAnchor(
                _from=AnchorMarker(col=col - 1, row=row - 1,
                                   colOff=int((width - image.width + 16) / 2 * 9525),
                                   rowOff=int((height - image.height + 12) / 2 * 9525)),
                ext=XDRPositiveSize2D(int(image.width * 9525), int(image.height * 9525)))
            ws.add_image(image)
            ws.merge_cells(start_row=row + 11, end_row=row + 11, start_column=col, end_column=col + half - 1)
            caption = ws.cell(row + 11, col, photo.get('description') or photo.get('location', ''))
            caption.font = Font(name='바탕', size=9)
            caption.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for ws in wb.worksheets[:original_count]:
        if batch['method'] == 'MT':
            ws['X3'] = str(total)
        else:
            for row in ws.iter_rows(min_row=3, max_row=3):
                for cell in row:
                    if isinstance(cell.value, str) and 'of' in cell.value:
                        import re
                        cell.value = re.sub(r'(of\s*)\d+', lambda match: match[1] + str(total), cell.value)
            if ws['Q3'].value and str(ws['Q3'].value).strip().endswith('of'):
                ws['R3'] = str(total)
    template.close()
