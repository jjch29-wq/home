"""PMI measurement selection and transactional editing for Gyeongin reports."""
from copy import deepcopy
from decimal import Decimal, InvalidOperation
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from site_apps.gyeongin.src.cert_exporter import choose_pmi_element_columns, read_pmi_measurements
from site_apps.gyeongin.src.pmi_criteria import read_state
from site_apps.gyeongin.src.pmi_judgement import DEFAULT_TOLERANCE_PERCENT, judge_measurement


def measurement_value(value):
    text = str(value).strip().upper().removesuffix('%').strip()
    if text in ('', 'ND'):
        return text
    try:
        number = Decimal(text)
    except InvalidOperation:
        raise ValueError('성분값은 0~100의 숫자, ND 또는 빈칸으로 입력하세요.') from None
    if not number.is_finite() or not 0 <= number <= 100:
        raise ValueError('성분값은 0~100 범위로 입력하세요.')
    return format(number, 'f')


def measurement_identity(row):
    if row.get('source_file') and row.get('source_row') is not None:
        return (row['source_file'], str(row['source_row']))
    original = row.get('_original', row)
    return (row.get('source_file', ''), row.get('source_row', ''),
            row.get('measurement_date', ''), original.get('tag', ''),
            original.get('component', ''), original.get('alloy', ''),
            tuple(sorted((original.get('elements') or {}).items())))


def read_measurement_files(paths):
    candidates = []
    for path in paths:
        for index, row in enumerate(read_pmi_measurements(path)):
            row.update(source_file=Path(path).name, source_row=index + 1)
            candidates.append(row)
    return candidates


def merge_measurements(rows, candidates):
    known = {measurement_identity(row): row for row in rows}
    added = []
    duplicates = 0
    for row in candidates:
        identity = measurement_identity(row)
        if identity in known:
            existing = known[identity]
            for key in ('reading_no', 'measurement_date'):
                if not existing.get(key) and row.get(key):
                    existing[key] = row[key]
            duplicates += 1
            continue
        copied = deepcopy(row)
        copied.setdefault('_original', deepcopy({
            key: copied.get(key, {} if key == 'elements' else '')
            for key in ('tag', 'component', 'alloy', 'elements')}))
        added.append(str(len(rows)))
        rows.append(copied)
        known[identity] = copied
    return added, duplicates


def import_measurement_files(parent, batch, persist, on_saved):
    """Register selected files without opening the measurement editor."""
    paths = filedialog.askopenfilenames(parent=parent, title='PMI 측정파일 선택',
                                      filetypes=[('Excel/CSV', '*.xlsx *.xls *.csv')])
    if not paths:
        return
    try:
        candidates = read_measurement_files(paths)
    except Exception as exc:
        messagebox.showerror('측정파일', str(exc), parent=parent)
        return
    if not candidates:
        messagebox.showwarning('측정파일', '읽을 수 있는 측정행이 없습니다.', parent=parent)
        return
    rows = deepcopy(batch.get('pmi_measurements') or [])
    merge_measurements(rows, candidates)
    previous = {key: deepcopy(batch[key]) for key in ('pmi_measurements', 'pmi_source') if key in batch}
    batch['pmi_measurements'] = rows
    batch['pmi_source'] = ', '.join(dict.fromkeys(
        row.get('source_file', '') for row in rows if row.get('source_file')))
    if not persist():
        for key in ('pmi_measurements', 'pmi_source'):
            batch.pop(key, None)
        batch.update(previous)
        return
    on_saved()


def open_measurement_editor(parent, batch, state_path, persist, on_saved, import_files=False):
    rows = deepcopy(batch.get('pmi_measurements') or [])
    elements = choose_pmi_element_columns(rows, batch.get('pmi_elements'))
    criteria = batch.get('pmi_criteria') or []
    tolerance = batch.get('pmi_tolerance_percent', DEFAULT_TOLERANCE_PERCENT)
    criterion_labels = {item['id']: item['label'] for item in criteria}
    for row in rows:
        row.setdefault('_original', deepcopy({k: row.get(k, {} if k == 'elements' else '')
                                             for k in ('tag', 'component', 'alloy', 'elements')}))
    win = tk.Toplevel(parent)
    win.title('PMI 측정자료 선택·편집')
    win.geometry('1100x600')
    win.transient(parent)
    win.grab_set()
    toolbar = ttk.Frame(win, padding=8)
    toolbar.pack(fill='x')
    query = tk.StringVar()
    status = tk.StringVar()
    if len(criteria) > 1:
        criteria_bar = ttk.Frame(win, padding=(8, 2))
        criteria_bar.pack(fill='x')
        ttk.Label(criteria_bar, text='선택 행 적용기준').pack(side='left')
        criterion_choice = tk.StringVar(value=criteria[0]['label'])
        ttk.Combobox(criteria_bar, textvariable=criterion_choice, state='readonly',
                     values=list(criterion_labels.values()), width=55).pack(side='left', padx=4)
        ttk.Button(criteria_bar, text='선택 행에 적용',
                   command=lambda: finish_cell() and apply_selected_criterion()).pack(side='left')
    columns = ('include', 'reading_no', 'date', 'source', 'tag', 'component', 'alloy', 'criterion', 'result') + tuple(elements)
    labels = dict(include='출력', reading_no='Reading No', date='측정일', source='원본 파일', tag='태그',
                  component='구성부품', alloy='장비 합금명', criterion='적용기준', result='판정')
    frame = ttk.Frame(win, padding=8)
    frame.pack(fill='both', expand=True)
    tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
    for column in columns:
        tree.heading(column, text=labels.get(column, column))
        tree.column(column, width=140 if column == 'source' else 85, stretch=True)
    scroll = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    scroll.pack(side='right', fill='y')
    horizontal = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
    tree.configure(xscrollcommand=horizontal.set)
    horizontal.pack(side='bottom', fill='x')
    tree.pack(fill='both', expand=True)
    tree.tag_configure('modified', foreground='#1d4ed8')
    pending = {}

    def cancel_cell(_event=None):
        entry = pending.pop('entry', None)
        pending.clear()
        if entry is not None:
            entry.destroy()
        return 'break'

    def finish_cell():
        if not pending:
            return True
        entry = pending['entry']
        row = pending['row']
        field = pending['field']
        value = entry.get().strip()
        if value == pending['initial']:
            cancel_cell()
            return True
        try:
            if field == 'criterion_id':
                value = next(key for key, label in criterion_labels.items() if label == value)
            if field in elements:
                value = measurement_value(value)
        except (ValueError, StopIteration) as exc:
            messagebox.showerror('입력값 확인', str(exc), parent=win)
            entry.focus_set()
            return False
        if field in elements:
            row.setdefault('elements', {})[field] = value
        else:
            row[field] = value
        iid = pending['iid']
        column = pending['column']
        cancel_cell()
        if tree.exists(iid):
            tree.set(iid, column, criterion_labels.get(value, value) if field == 'criterion_id' else value)
            tree.set(iid, 'result', judge_measurement(row, criteria, elements, tolerance)[0])
            tree.item(iid, tags=('modified',))
        return True

    def edit_cell(event):
        if not finish_cell():
            return 'break'
        iid = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        if not iid or not column_id or tree.identify_region(event.x, event.y) != 'cell':
            return
        column = columns[int(column_id[1:]) - 1]
        row = rows[int(iid)]
        if column == 'include':
            row['include'] = not row.get('include', True)
            refresh()
            return 'break'
        if column in ('source', 'reading_no', 'result'):
            return 'break'
        field = {'date': 'measurement_date', 'criterion': 'criterion_id'}.get(column, column)
        bounds = tree.bbox(iid, column_id)
        if not bounds:
            return 'break'
        value = (row.get('elements') or {}).get(field, '') if field in elements else row.get(field, '')
        if column == 'criterion':
            value = criterion_labels.get(row.get('criterion_id'), '')
            entry = ttk.Combobox(tree, values=list(criterion_labels.values()), state='readonly')
            entry.set(value)
        else:
            entry = ttk.Entry(tree)
            entry.insert(0, str(value))
        pending.update(entry=entry, row=row, field=field, iid=iid, column=column,
                       initial=str(value))
        x, y, width, height = bounds
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()
        entry.select_range(0, 'end')

        def accept(_event=None):
            if finish_cell():
                refresh()
                tree.focus_set()
            return 'break'

        entry.bind('<Return>', accept)
        entry.bind('<Tab>', accept)
        entry.bind('<Escape>', cancel_cell)
        if column == 'criterion':
            entry.bind('<<ComboboxSelected>>', accept)
        return 'break'

    def apply_selected_criterion():
        criterion_id = next(key for key, label in criterion_labels.items() if label == criterion_choice.get())
        for iid in tree.selection():
            rows[int(iid)]['criterion_id'] = criterion_id
        refresh()

    def refresh():
        nonlocal columns
        if not finish_cell():
            return
        elements[:] = choose_pmi_element_columns(rows, batch.get('pmi_elements'))
        columns = ('include', 'reading_no', 'date', 'source', 'tag', 'component', 'alloy', 'criterion', 'result') + tuple(elements)
        tree.configure(columns=columns)
        for column in columns:
            tree.heading(column, text=labels.get(column, column))
            tree.column(column, width={'source': 180, 'date': 165}.get(column, 85), stretch=False)
        selected = tree.selection()
        tree.delete(*tree.get_children())
        term = query.get().strip().casefold()
        for index, row in enumerate(rows):
            values = ('✓' if row.get('include', True) else '', row.get('reading_no', ''), row.get('measurement_date', ''),
                      row.get('source_file', ''), row.get('tag', ''), row.get('component', ''),
                      row.get('alloy', ''),
                      criterion_labels.get(row.get('criterion_id'), criteria[0]['label']
                                           if len(criteria) == 1 and not row.get('criterion_id') else ''),
                      judge_measurement(row, criteria, elements, tolerance)[0]) + tuple(
                          (row.get('elements') or {}).get(s, '') for s in elements)
            if term and term not in ' '.join(map(str, values)).casefold():
                continue
            changed = any(row.get(k, {} if k == 'elements' else '') != v
                          for k, v in row.get('_original', {}).items())
            tree.insert('', 'end', iid=str(index), values=values,
                        tags=('modified',) if changed else ())
        tree.selection_set([iid for iid in selected if tree.exists(iid)])
        status.set(f'전체 {len(rows)}건 / 검색 {len(tree.get_children())}건 / 출력 '
                   f'{sum(r.get("include", True) for r in rows)}건 · 파란 행은 수정된 자료입니다.')

    def add_candidates(candidates):
        if not finish_cell():
            return
        added, duplicates = merge_measurements(rows, candidates)
        query.set('')
        refresh()
        tree.selection_set(added)
        if added:
            tree.see(added[0])
        if duplicates:
            messagebox.showinfo('중복 확인', f'{len(added)}건 추가, 동일 측정자료 {duplicates}건 제외', parent=win)

    def load_files():
        paths = filedialog.askopenfilenames(parent=win, title='여러 날짜의 PMI 측정파일 선택',
                                          filetypes=[('Excel/CSV', '*.xlsx *.xls *.csv')])
        try:
            candidates = read_measurement_files(paths)
        except Exception as exc:
            messagebox.showerror('측정파일', str(exc), parent=win)
            return
        if candidates:
            add_candidates(candidates)
        elif paths:
            messagebox.showwarning('측정파일', '읽을 수 있는 측정행이 없습니다.', parent=win)

    def load_history():
        try:
            candidates = []
            for report in read_state(state_path)['reports'].values():
                if report.get('project_name', '') != batch.get('project_name', ''):
                    continue
                if report.get('_measurement_company', batch.get('company', '')) != batch.get('company', ''):
                    continue
                for row in report.get('pmi_measurements', []):
                    copied = deepcopy(row)
                    copied.setdefault('source_file', report.get('pmi_source', '저장 자료'))
                    if not copied.get('measurement_date'):
                        copied['measurement_date'] = report.get('_measurement_report_date', '')
                    candidates.append(copied)
            if candidates:
                add_candidates(candidates)
            else:
                messagebox.showinfo('저장 자료', '같은 현장의 저장된 측정자료가 없습니다.', parent=win)
        except (ValueError, OSError) as exc:
            messagebox.showerror('저장 자료', str(exc), parent=win)

    def edit(_event=None):
        selected = tree.selection()
        if len(selected) != 1:
            return
        row = rows[int(selected[0])]
        editor = tk.Toplevel(win)
        editor.title('선택 측정행 편집')
        editor.transient(win)
        editor.grab_set()
        fields = ('measurement_date', 'tag', 'component', 'alloy') + tuple(elements)
        variables = {}
        for index, field in enumerate(fields):
            field_row, field_group = index % 12, (index // 12) * 2
            ttk.Label(editor, text={'measurement_date': '측정일'}.get(field, labels.get(field, field))).grid(row=field_row, column=field_group, padx=8, pady=3)
            value = (row.get('elements') or {}).get(field, '') if field in elements else row.get(field, '')
            variables[field] = tk.StringVar(value=value)
            ttk.Entry(editor, textvariable=variables[field], width=30).grid(row=field_row, column=field_group + 1, padx=8)

        def apply():
            try:
                values = {s: measurement_value(variables[s].get()) for s in elements}
            except ValueError as exc:
                messagebox.showerror('입력값 확인', str(exc), parent=editor)
                return
            for key in fields[:4]:
                row[key] = variables[key].get().strip()
            row.setdefault('elements', {}).update(values)
            editor.destroy()
            win.grab_set()
            refresh()

        ttk.Button(editor, text='적용', command=apply).grid(row=min(len(fields), 12), column=1, pady=8)

    def change_selected(action):
        for iid in tree.selection():
            row = rows[int(iid)]
            if action == 'restore':
                row.update(deepcopy(row.get('_original', {})))
            else:
                row['include'] = action
        refresh()

    def bulk_alloy():
        selected = tree.selection()
        if not selected:
            messagebox.showinfo('합금명 일괄 변경', '변경할 행을 선택하세요.', parent=win)
            return
        value = simpledialog.askstring('합금명 일괄 변경', '선택 행에 적용할 합금명', parent=win)
        if value is not None:
            for iid in selected:
                rows[int(iid)]['alloy'] = value.strip()
            refresh()

    def move(direction):
        if query.get().strip():
            messagebox.showinfo('순서 변경', '검색어를 지운 후 순서를 변경하세요.', parent=win)
            return
        selected = {int(i) for i in tree.selection()}
        order = sorted(selected, reverse=direction > 0)
        for index in order:
            target = index + direction
            if 0 <= target < len(rows) and target not in selected:
                rows[index], rows[target] = rows[target], rows[index]
                selected.remove(index)
                selected.add(target)
        refresh()
        tree.selection_set([str(i) for i in sorted(selected)])

    def add_row():
        rows.append(dict(tag='', component='', alloy='', elements={}, include=True,
                         _original=dict(tag='', component='', alloy='', elements={})))
        refresh()
        tree.selection_set(str(len(rows) - 1))
        edit()

    def delete_rows():
        for index in sorted((int(i) for i in tree.selection()), reverse=True):
            del rows[index]
        refresh()

    def paste():
        dialog = tk.Toplevel(win)
        dialog.title('엑셀 측정행 붙여넣기')
        dialog.transient(win)
        dialog.grab_set()
        ttk.Label(dialog, text='열 순서: 태그, 구성부품, 합금명, ' + ', '.join(elements)).pack(padx=8, pady=8)
        text = tk.Text(dialog, width=90, height=12)
        text.pack(padx=8)

        def apply_paste():
            candidates = []
            try:
                for line in text.get('1.0', 'end').strip().splitlines():
                    values = line.split('\t')
                    if len(values) != 3 + len(elements):
                        raise ValueError(f'각 행에 {3 + len(elements)}개 열이 필요합니다.')
                    candidates.append(dict(tag=values[0].strip(), component=values[1].strip(),
                                           alloy=values[2].strip(), elements={s: measurement_value(v)
                                           for s, v in zip(elements, values[3:])},
                                           source_file='엑셀 붙여넣기'))
            except ValueError as exc:
                messagebox.showerror('붙여넣기', str(exc), parent=dialog)
                return
            dialog.destroy()
            win.grab_set()
            if candidates:
                add_candidates(candidates)

        ttk.Button(dialog, text='편집 표에 추가', command=apply_paste).pack(pady=8)

    def save():
        if not finish_cell():
            return
        old = {key: deepcopy(batch.get(key)) for key in ('pmi_measurements', 'pmi_source')}
        batch['pmi_measurements'] = deepcopy(rows)
        batch['pmi_source'] = ', '.join(dict.fromkeys(r.get('source_file', '') for r in rows if r.get('source_file')))
        if not persist():
            for key, value in old.items():
                if value is None:
                    batch.pop(key, None)
                else:
                    batch[key] = value
            return
        on_saved()
        win.destroy()

    ttk.Button(toolbar, text='파일 추가', command=load_files).pack(side='left')
    ttk.Button(toolbar, text='현장 저장자료 불러오기', command=load_history).pack(side='left', padx=4)
    ttk.Button(toolbar, text='엑셀 붙여넣기', command=paste).pack(side='left', padx=4)
    ttk.Label(toolbar, text='검색').pack(side='left', padx=4)
    ttk.Entry(toolbar, textvariable=query).pack(side='left', fill='x', expand=True)
    query.trace_add('write', lambda *_: refresh())
    tree.bind('<Double-1>', edit_cell)
    tree.bind('<Button-1>', lambda _event: None if finish_cell() else 'break')
    def show_verdict_reason(_event=None):
        selected = tree.selection()
        if len(selected) == 1:
            verdict, reason = judge_measurement(rows[int(selected[0])], criteria, elements, tolerance)
            status.set(f'{verdict} · {reason}')

    tree.bind('<<TreeviewSelect>>', show_verdict_reason)
    buttons = ttk.Frame(win, padding=8)
    buttons.pack(fill='x')
    for label, command in [('행 추가', add_row), ('선택 행 편집', edit),
                           ('행 삭제', delete_rows),
                           ('합금명 일괄 변경', bulk_alloy),
                           ('출력 포함', lambda: change_selected(True)),
                           ('출력 제외', lambda: change_selected(False)),
                           ('원래 값 복원', lambda: change_selected('restore')),
                           ('위로', lambda: move(-1)), ('아래로', lambda: move(1))]:
        ttk.Button(buttons, text=label,
                   command=lambda action=command: finish_cell() and action()).pack(side='left', padx=2)
    ttk.Button(buttons, text='저장', command=save).pack(side='right')
    ttk.Button(buttons, text='취소', command=win.destroy).pack(side='right', padx=4)
    ttk.Label(win, text='셀 더블클릭으로 수정 · Enter 적용 · Esc 취소 · 출력 칸 더블클릭으로 포함/제외',
              padding=(8, 2)).pack(anchor='w')
    ttk.Label(win, text=f'선택한 출력 원소 기준으로 판정 · 허용오차 ±{tolerance}% (하한·상한 대비 상대오차)',
              padding=(8, 2)).pack(anchor='w')
    ttk.Label(win, textvariable=status, padding=8).pack(anchor='w')
    refresh()
    if import_files:
        load_files()
