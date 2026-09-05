"""Read-only site history selection for the NDT report application."""
import json
from pathlib import Path
import re
import hashlib

SITES = {
    '한국지역난방 중앙지사': ('central', 'daily_work_history.json'),
    '가스공사 가산~가평': ('kogas', 'kogas_daily_work_history.json'),
    '롯데건설 바이오로직스': ('lotte', 'daily_work_history.json'),
}


def history_path(site, source_dir=None):
    folder, filename = SITES[site]
    base = Path(source_dir) if source_dir is not None else Path(__file__).resolve().parent
    return base / 'site_apps' / folder / 'src' / filename


def read_results(site, mode, source_dir=None):
    """Validate before returning any rows; never fall back to shared history."""
    method = 'RT' if mode == 'KOGAS' else mode
    if method not in {'RT', 'PT', 'MT', 'PAUT', 'PMI'}:
        raise ValueError(f'지원하지 않는 검사방법: {mode}')
    with history_path(site, source_dir).open('r', encoding='utf-8-sig') as stream:
        history = json.load(stream)
    if not isinstance(history, dict):
        raise ValueError('일보 파일의 최상위 데이터는 날짜별 객체여야 합니다.')
    rows = []
    for date, day in history.items():
        if not isinstance(day, dict) or not isinstance(day.get('ndt_results', []), list):
            raise ValueError(f'{date}: 일보 또는 ndt_results 형식이 올바르지 않습니다.')
        for record_index, record in enumerate(day.get('ndt_results', [])):
            if not isinstance(record, dict) or any(isinstance(v, (dict, list)) for v in record.values()):
                raise ValueError(f'{date}: NDT 항목 형식이 올바르지 않습니다.')
            clean = {k: str(v).strip() for k, v in record.items() if v is not None and str(v).strip()}
            if not clean:
                continue
            methods = re.findall(r'[A-Z]+', clean.get('검사방법', '').upper())
            if method not in methods:
                continue
            clean['Date'] = date
            identity = json.dumps([site, date, record_index, record], sort_keys=True, ensure_ascii=False)
            clean['_source_id'] = hashlib.sha256(identity.encode('utf-8')).hexdigest()
            clean['_source_site'] = site
            clean['_source_path'] = str(history_path(site, source_dir))
            clean['_source_location'] = clean.get('현장명') or day.get('현장명') or day.get('site') or ''
            rows.append(clean)
    return rows


def select_history_rows(parent, mode):
    """Modal preview. Cancel and source changes cannot modify the report."""
    import tkinter as tk
    from tkinter import ttk, messagebox

    dialog = tk.Toplevel(parent)
    dialog.title(f'일보 NDT 불러오기 — {mode}')
    dialog.geometry('1080x540')
    dialog.transient(parent)
    result = []
    rows = []
    site_var = tk.StringVar(value=next(iter(SITES)))
    path_var = tk.StringVar()
    status_var = tk.StringVar()
    top = ttk.Frame(dialog, padding=10)
    top.pack(fill='x')
    ttk.Label(top, text='사업장').pack(side='left')
    combo = ttk.Combobox(top, textvariable=site_var, values=list(SITES), state='readonly', width=30)
    combo.pack(side='left', padx=10)
    ttk.Label(dialog, textvariable=path_var, wraplength=1030).pack(fill='x', padx=10)
    ttk.Label(dialog, text='가져올 행을 선택하세요. Ctrl/Shift로 여러 행을 선택할 수 있습니다.').pack(anchor='w', padx=10, pady=6)
    frame = ttk.Frame(dialog)
    frame.pack(fill='both', expand=True, padx=10)
    columns = ('사업장', '작업일', '현장명', '검사방법', '라인번호', 'Joint No.', '구간', '결과')
    tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
    for column in columns:
        tree.heading(column, text=column)
        tree.column(column, width=125, minwidth=70)
    scroll = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    scroll.pack(side='right', fill='y')
    tree.pack(fill='both', expand=True)
    ttk.Label(dialog, textvariable=status_var).pack(anchor='w', padx=10, pady=6)
    buttons = ttk.Frame(dialog, padding=10)
    buttons.pack(fill='x')

    def selection_changed(_event=None):
        count = len(tree.selection())
        import_button.configure(state='normal' if count else 'disabled')

    def reload_rows(_event=None):
        # Clear both the data snapshot and selection before attempting a new read.
        rows.clear()
        tree.delete(*tree.get_children())
        import_button.configure(state='disabled')
        site = site_var.get()
        path_var.set(str(history_path(site)))
        status_var.set('불러오는 중…')
        try:
            rows.extend(read_results(site, mode))
        except (OSError, ValueError) as exc:
            status_var.set('불러오기 실패 — 파일과 데이터 형식을 확인하세요.')
            messagebox.showerror('일보 불러오기 실패', f'{site}\n{history_path(site)}\n\n{exc}', parent=dialog)
            return
        for index, row in enumerate(rows):
            tree.insert('', 'end', iid=str(index), values=(site, row['Date'], row['_source_location'] or '(미기록)',
                row.get('검사방법', ''), row.get('라인번호', ''), row.get('Joint No.', ''), row.get('구간', ''), row.get('결과', '')))
        status_var.set(f'{site} · {mode} 항목 {len(rows)}건 (원본 읽기 전용)')

    def accept():
        selected = tree.selection()
        if not selected:
            return
        if not messagebox.askyesno('선택 항목 가져오기',
            f'{site_var.get()}의 {len(selected)}건을 현재 {mode} 보고서에 추가할까요?\n기존 보고서 행은 유지됩니다.', parent=dialog):
            return
        result.extend(dict(rows[int(index)]) for index in selected)
        dialog.destroy()

    ttk.Button(buttons, text='전체 선택', command=lambda: tree.selection_set(tree.get_children())).pack(side='left')
    ttk.Button(buttons, text='새로고침', command=reload_rows).pack(side='left', padx=5)
    ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='right')
    import_button = ttk.Button(buttons, text='선택 항목 가져오기', command=accept, state='disabled')
    import_button.pack(side='right', padx=5)
    combo.bind('<<ComboboxSelected>>', reload_rows)
    tree.bind('<<TreeviewSelect>>', selection_changed)
    dialog.bind('<Escape>', lambda _event: dialog.destroy())
    dialog.grab_set()
    reload_rows()
    parent.wait_window(dialog)
    return result
