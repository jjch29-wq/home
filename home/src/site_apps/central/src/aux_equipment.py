"""Optional auxiliary equipment selection for central daily inspection records."""
import tkinter as tk
from tkinter import ttk


def names(value):
    if value is None or str(value).strip().lower() in ('', 'nan', 'none'):
        return []
    return list(dict.fromkeys(n.strip() for n in str(value).split(';') if n.strip()))


def value(app):
    field = getattr(app, 'daily_aux_equipment_var', None)
    return '; '.join(names(field.get())) if field is not None else ''


def restore(app, text=''):
    field = getattr(app, 'daily_aux_equipment_var', None)
    if field is not None:
        field.set('; '.join(names(text)))


def report_names(records):
    result = []
    for column in ('장비명', '보조장비'):
        if column in records.columns:
            for text in records[column].dropna():
                result.extend(names(text))
    return list(dict.fromkeys(result))


def add_selector(parent, app, variable, width=25):
    frame = ttk.Frame(parent)
    ttk.Entry(frame, textvariable=variable, width=width).pack(side='left', fill='x', expand=True)
    ttk.Button(frame, text='선택', width=5,
               command=lambda: select(frame, app, variable)).pack(side='left', padx=2)
    return frame


def select(parent, app, variable):
    dialog = tk.Toplevel(parent)
    dialog.title('보조장비 선택 (여러 개 선택 가능)')
    dialog.geometry('540x420')
    dialog.transient(parent.winfo_toplevel())
    selected = names(variable.get())
    candidates = list(selected)
    for item in getattr(app, 'equipments', []):
        label = item.get('name', '') if isinstance(item, dict) else str(item)
        candidates.extend(names(label))
    history = getattr(app, 'daily_usage_df', None)
    if history is not None and '보조장비' in history.columns:
        for item in history['보조장비'].dropna():
            candidates.extend(names(item))
    candidates = sorted(set(candidates))
    ttk.Label(dialog, text='클릭하여 선택/해제 · 직접 추가한 항목도 함께 저장됩니다.').pack(padx=10, pady=8)
    body = ttk.Frame(dialog)
    body.pack(fill='both', expand=True, padx=10)
    scrollbar = ttk.Scrollbar(body)
    scrollbar.pack(side='right', fill='y')
    listing = tk.Listbox(body, selectmode='multiple', exportselection=False,
                         yscrollcommand=scrollbar.set)
    listing.pack(fill='both', expand=True)
    scrollbar.configure(command=listing.yview)
    for i, item in enumerate(candidates):
        listing.insert('end', item)
        if item in selected:
            listing.selection_set(i)
    row = ttk.Frame(dialog)
    row.pack(fill='x', padx=10, pady=8)
    entry = ttk.Entry(row)
    entry.pack(side='left', fill='x', expand=True)

    def add():
        for item in names(entry.get()):
            if item not in candidates:
                candidates.append(item)
                listing.insert('end', item)
            listing.selection_set(candidates.index(item))
        entry.delete(0, 'end')

    ttk.Button(row, text='추가', command=add).pack(side='left', padx=4)
    entry.bind('<Return>', lambda event: add())

    def apply():
        add()
        variable.set('; '.join(listing.get(i) for i in listing.curselection()))
        dialog.destroy()

    buttons = ttk.Frame(dialog)
    buttons.pack(pady=(0, 10))
    ttk.Button(buttons, text='적용', command=apply).pack(side='left', padx=4)
    ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left', padx=4)
    dialog.grab_set()
