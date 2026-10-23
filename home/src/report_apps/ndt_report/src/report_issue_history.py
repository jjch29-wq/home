"""Persistent snapshots of successfully generated reports."""
import datetime
import json
import os
import re
from pathlib import Path
import tempfile
import uuid


DEFAULT_PATH = Path(__file__).resolve().parents[1] / 'data' / 'report_issue_history.json'


def report_output_path(directory, report_no, extension, fallback='Report'):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '-', str(report_no or '').strip()).rstrip(' .')
    name = name[:180].rstrip(' .') or fallback
    if name.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *[f'COM{i}' for i in range(1, 10)], *[f'LPT{i}' for i in range(1, 10)]}:
        name = '_' + name
    directory = Path(directory)
    candidate = directory / (name + extension)
    number = 2
    while candidate.exists():
        candidate = directory / f'{name} ({number}){extension}'
        number += 1
    return str(candidate)


def load_history(path=DEFAULT_PATH):
    path = Path(path)
    if not path.exists():
        return []
    records = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(records, list) or any(
        not isinstance(r, dict) or not isinstance(r.get('items'), list) for r in records
    ):
        raise ValueError('발행 이력 파일 형식이 올바르지 않습니다.')
    return records


def restore_report_numbers(mode, items, path=DEFAULT_PATH):
    """Restore the latest active issue number for the same source inspection."""
    def identity(item):
        if item.get('_source_id'):
            return ('source', item.get('_source_site', ''), item['_source_id'])
        line = item.get('Line No.') or item.get('Dwg') or item.get('ISO')
        joint = item.get('Joint No.') or item.get('Joint')
        if not line or not joint or not item.get('Date'):
            return None
        return ('inspection', item.get('_source_site', ''), item.get('_source_path', ''),
                item['Date'], line, joint)
    numbers = {}
    for entry in load_history(path):
        if entry.get('deleted_at') or entry.get('mode') != mode or not entry.get('report_no'):
            continue
        for item in entry['items']:
            key = identity(item)
            if key is not None:
                numbers[key] = entry['report_no']
    for item in items:
        key = identity(item)
        if key in numbers:
            item['Report No'] = numbers[key]


def record_issue(mode, report_no, output_path, items, path=DEFAULT_PATH):
    """Record only an existing output; keep a separate snapshot on reissue."""
    output = Path(output_path).resolve()
    if not output.is_file():
        raise FileNotFoundError(output)
    records = load_history(path)
    snapshot = json.loads(json.dumps(items, ensure_ascii=False, default=str))
    entry = {
        'id': uuid.uuid4().hex,
        'issued_at': datetime.datetime.now().astimezone().isoformat(timespec='seconds'),
        'mode': mode, 'report_no': str(report_no or '').strip(),
        'output_path': str(output), 'items': snapshot,
    }
    records.append(entry)
    save_history(records, path)
    return entry


def save_history(records, path=DEFAULT_PATH):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent,
                                         suffix='.tmp', delete=False) as stream:
            temporary = stream.name
            json.dump(records, stream, ensure_ascii=False, indent=2)
        os.replace(temporary, path)
    finally:
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)


def edit_issue(issue_id, item_index=None, fields=None, item_fields=None,
               delete=False, path=DEFAULT_PATH, expected=None):
    """Edit one issue atomically, retaining previous values for recovery."""
    records = load_history(path)
    entry = next((r for r in records if r['id'] == issue_id), None)
    if entry is None or entry.get('deleted_at'):
        raise ValueError('해당 발행 이력을 찾을 수 없습니다. 새로고침해주세요.')
    if expected is not None and entry != expected:
        raise ValueError('다른 작업에서 이력이 변경되었습니다. 새로고침 후 다시 편집해주세요.')
    if item_index is not None and not 0 <= item_index < len(entry['items']):
        raise ValueError('검사 항목을 찾을 수 없습니다.')
    before = json.loads(json.dumps({k: v for k, v in entry.items() if k != 'revisions'}))
    now = datetime.datetime.now().astimezone().isoformat(timespec='seconds')
    if delete:
        if item_index is None:
            entry['deleted_at'] = now
        else:
            del entry['items'][item_index]
            if not entry['items']:
                entry['deleted_at'] = now
    else:
        allowed = {'report_no', 'mode', 'issued_at', 'output_path'}
        if set(fields or {}) - allowed:
            raise ValueError('수정할 수 없는 보고서 필드입니다.')
        entry.update({k: str(v).strip() for k, v in (fields or {}).items()})
        if item_fields:
            if item_index is None:
                raise ValueError('수정할 검사 항목을 선택해주세요.')
            if any(k.startswith('_') for k in item_fields):
                raise ValueError('원본 출처 정보는 수정할 수 없습니다.')
            entry['items'][item_index].update(item_fields)
    entry.setdefault('revisions', []).append({'edited_at': now, 'before': before})
    entry['edited_at'] = now
    save_history(records, path)
    return entry


def item_values(entry, item):
    def first(*keys):
        return next((str(item[k]) for k in keys if item.get(k) not in (None, '')), '')
    return (
        entry['issued_at'], entry.get('report_no') or '(번호 미입력)', entry['mode'],
        first('_source_site'), first('Date'), first('Dwg', 'Line No.', 'ISO', '라인번호'),
        first('Joint', 'Joint No.', 'Joint No'), first('Sec', '구간'),
        first('Size', 'NPS', '관경'), first('Result', 'Evaluation', '결과'),
        entry['output_path'],
    )


def delete_selections(selections, whole_issue=False, path=DEFAULT_PATH):
    """Validate all selections before atomically deleting any of them."""
    records = load_history(path)
    grouped = {}
    for expected, index in selections:
        entry = next((r for r in records if r['id'] == expected['id']), None)
        if entry is None or entry != expected or entry.get('deleted_at'):
            raise ValueError('이력이 변경되었습니다. 새로고침 후 다시 선택해주세요.')
        if not 0 <= index < len(entry['items']):
            raise ValueError('검사 항목을 찾을 수 없습니다.')
        grouped.setdefault(entry['id'], (entry, set()))[1].add(index)
    now = datetime.datetime.now().astimezone().isoformat(timespec='seconds')
    for entry, indices in grouped.values():
        before = json.loads(json.dumps({k: v for k, v in entry.items() if k != 'revisions'}))
        if whole_issue:
            entry['deleted_at'] = now
        else:
            entry['items'] = [item for i, item in enumerate(entry['items']) if i not in indices]
            if not entry['items']:
                entry['deleted_at'] = now
        entry.setdefault('revisions', []).append({'edited_at': now, 'before': before})
        entry['edited_at'] = now
    save_history(records, path)


def show_history(parent, path=DEFAULT_PATH):
    import tkinter as tk
    from tkinter import ttk, messagebox
    window = tk.Toplevel(parent)
    window.title('보고서 발행 이력')
    window.geometry('1250x600')
    window.transient(parent)
    bar = ttk.Frame(window, padding=10)
    bar.pack(fill='x')
    ttk.Label(bar, text='보고서번호 / 검사일 / 라인 / Joint 검색').pack(side='left')
    query = tk.StringVar()
    ttk.Entry(bar, textvariable=query, width=40).pack(side='left', padx=8)
    ttk.Label(window, text='번호·파일 경로는 같은 발행 건 전체에 적용됩니다. 편집은 이력에만 반영되며 Excel 파일과 원본 일보는 변경되지 않습니다.').pack(anchor='w', padx=10)
    frame = ttk.Frame(window, padding=10)
    frame.pack(fill='both', expand=True)
    columns = ('생성일시', '보고서번호', '검사방법', '사업장', '검사일', '라인/도면',
               'Joint', '구간', '관경', '결과', '생성 파일')
    tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
    for column in columns:
        tree.heading(column, text=column)
        tree.column(column, width=150 if column in ('보고서번호', '라인/도면') else 110)
    vertical = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    horizontal = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    tree.grid(row=0, column=0, sticky='nsew')
    vertical.grid(row=0, column=1, sticky='ns')
    horizontal.grid(row=1, column=0, sticky='ew')
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    status = tk.StringVar()
    ttk.Label(window, textvariable=status).pack(anchor='w', padx=10)
    visible = {}

    def refresh(*_args):
        try:
            records = load_history(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror('발행 이력 읽기 실패', str(exc), parent=window)
            return
        tree.delete(*tree.get_children())
        visible.clear()
        term = query.get().strip().casefold()
        issues = set()
        for entry in reversed(records):
            if entry.get('deleted_at'):
                continue
            for item_index, item in enumerate(entry['items']):
                values = item_values(entry, item)
                if term and term not in ' '.join(values).casefold():
                    continue
                iid = str(len(visible))
                tree.insert('', 'end', iid=iid, values=values)
                visible[iid] = (entry, item_index)
                issues.add(entry['id'])
        status.set(f'발행 {len(issues)}회 · 검사 항목 {len(visible)}건')

    def open_output(_event=None):
        selection = tree.selection()
        if not selection:
            return
        try:
            os.startfile(visible[selection[0]][0]['output_path'])
        except OSError as exc:
            messagebox.showerror('파일 열기 실패', str(exc), parent=window)

    def selected():
        selection = tree.selection()
        if not selection:
            messagebox.showinfo('항목 선택', '편집할 검사 항목을 선택해주세요.', parent=window)
            return None
        if len(selection) != 1:
            messagebox.showinfo('단일 항목 선택', '개별 편집은 한 항목을 선택해주세요. 삭제는 여러 항목을 선택할 수 있습니다.', parent=window)
            return None
        return visible[selection[0]]

    def edit_selected(_event=None):
        choice = selected()
        if choice is None:
            return
        entry, index = choice
        item = entry['items'][index]
        editor = tk.Toplevel(window)
        editor.title('발행 이력 편집')
        editor.transient(window)
        editor.grab_set()
        fields = {}
        item_fields = {}
        def add(label, key, value, target):
            row = len(fields) + len(item_fields)
            ttk.Label(editor, text=label).grid(row=row, column=0, sticky='w', padx=10, pady=4)
            var = tk.StringVar(value=value)
            ttk.Entry(editor, textvariable=var, width=65).grid(row=row, column=1, padx=10, pady=4)
            target[key] = var
        add('보고서번호 (발행 건 전체)', 'report_no', entry.get('report_no', ''), fields)
        add('생성 파일 (발행 건 전체)', 'output_path', entry['output_path'], fields)
        for label, aliases in (
            ('검사일', ('Date',)), ('라인/도면', ('Dwg', 'Line No.', 'ISO', '라인번호')),
            ('Joint', ('Joint', 'Joint No.', 'Joint No')), ('구간', ('Sec', '구간')),
            ('관경', ('Size', 'NPS', '관경')), ('결과', ('Result', 'Evaluation', '결과')),
        ):
            key = next((k for k in aliases if item.get(k) not in (None, '')), aliases[0])
            add(label, key, item.get(key, ''), item_fields)
        def save():
            try:
                edit_issue(entry['id'], index,
                           {k: v.get() for k, v in fields.items()},
                           {k: v.get().strip() for k, v in item_fields.items()},
                           path=path, expected=entry)
            except (OSError, ValueError) as exc:
                messagebox.showerror('이력 저장 실패', str(exc), parent=editor)
                return
            editor.destroy()
            refresh()
        ttk.Button(editor, text='저장', command=save).grid(row=9, column=1, sticky='e', padx=10, pady=10)
        ttk.Button(editor, text='취소', command=editor.destroy).grid(row=9, column=0, padx=10, pady=10)

    def delete_selected(whole_issue=False):
        choices = [visible[iid] for iid in tree.selection()]
        if not choices:
            messagebox.showinfo('항목 선택', '삭제할 항목을 선택해주세요.', parent=window)
            return
        issues = {entry['id']: entry for entry, _ in choices}
        scope = (f"선택한 발행 {len(issues)}건의 전체 {sum(len(e['items']) for e in issues.values())}개 항목"
                 if whole_issue else f'선택한 검사 항목 {len(choices)}건')
        if not messagebox.askyesno('발행 이력 삭제', f'{scope}을 이력에서 삭제할까요?\n생성된 보고서 파일은 유지됩니다.', parent=window):
            return
        try:
            delete_selections(choices, whole_issue=whole_issue, path=path)
        except (OSError, ValueError) as exc:
            messagebox.showerror('이력 삭제 실패', str(exc), parent=window)
            return
        refresh()

    actions = ttk.Frame(window, padding=10)
    actions.pack(fill='x')
    ttk.Label(actions, text='Ctrl/Shift: 복수 선택').pack(side='right')
    ttk.Button(actions, text='전체 선택', command=lambda: tree.selection_set(tree.get_children())).pack(side='left', padx=3)
    ttk.Button(actions, text='선택 항목 편집', command=edit_selected).pack(side='left', padx=3)
    ttk.Button(actions, text='선택 항목 삭제', command=delete_selected).pack(side='left', padx=3)
    ttk.Button(actions, text='보고서 이력 전체 삭제', command=lambda: delete_selected(True)).pack(side='left', padx=3)
    ttk.Button(bar, text='새로고침', command=refresh).pack(side='left')
    ttk.Button(bar, text='선택 보고서 열기', command=open_output).pack(side='right')
    query.trace_add('write', refresh)
    tree.bind('<Double-1>', edit_selected)
    refresh()
