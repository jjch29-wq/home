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
    if hasattr(app, 'open_list_management_dialog'):
        tk.Button(frame, text="⚙️", font=('Malgun Gothic', 8), bd=0,
                  bg=getattr(app, 'theme_bg', '#f0f0f0'), fg='blue', cursor='hand2',
                  command=lambda: app.open_list_management_dialog('aux_equipments')).pack(side='left', padx=(0, 2))
    return frame


def select(parent, app, variable):
    dialog = tk.Toplevel(parent)
    dialog.title('보조장비 선택 (여러 개 선택 가능)')
    dialog.geometry('560x460')
    dialog.transient(parent.winfo_toplevel())
    selected = names(variable.get())
    candidates = list(selected)

    # 1. Load from app.aux_equipments
    aux_list = getattr(app, 'aux_equipments', None)
    if aux_list is None:
        aux_list = ['JIREH MICROBE BTG006', 'COBRA Manual Scanner']
        if hasattr(app, 'aux_equipments'):
            app.aux_equipments = aux_list
    for item in aux_list:
        candidates.extend(names(item))

    # 2. Load from app.equipments
    for item in getattr(app, 'equipments', []):
        label = item.get('name', '') if isinstance(item, dict) else str(item)
        candidates.extend(names(label))

    # 3. Load from historical records
    history = getattr(app, 'daily_usage_df', None)
    if history is not None and '보조장비' in history.columns:
        for item in history['보조장비'].dropna():
            candidates.extend(names(item))

    # Always ensure default registered aux equipment is present
    for req in ['JIREH MICROBE BTG006', 'COBRA Manual Scanner']:
        if req not in candidates:
            candidates.append(req)

    candidates = sorted(set(candidates))

    def persist_items(items):
        """Save auxiliary equipment list permanently to app and config file"""
        changed = False
        if not hasattr(app, 'aux_equipments') or app.aux_equipments is None:
            app.aux_equipments = []
        for it in items:
            clean_it = str(it).strip()
            if clean_it and clean_it not in app.aux_equipments:
                app.aux_equipments.append(clean_it)
                changed = True
        if changed and hasattr(app, 'save_tab_config'):
            try:
                app.save_tab_config(force=True)
            except Exception as ex:
                print(f"Error persisting aux equipments: {ex}")

    ttk.Label(dialog, text='클릭하여 선택/해제 · 직접 추가/삭제한 항목도 영구 보관됩니다.').pack(padx=10, pady=8)
    body = ttk.Frame(dialog)
    body.pack(fill='both', expand=True, padx=10)
    scrollbar = ttk.Scrollbar(body)
    scrollbar.pack(side='right', fill='y')
    listing = tk.Listbox(body, selectmode='multiple', exportselection=False,
                         yscrollcommand=scrollbar.set)
    listing.pack(fill='both', expand=True)
    scrollbar.configure(command=listing.yview)

    def refresh_listing():
        listing.delete(0, 'end')
        for i, item in enumerate(candidates):
            listing.insert('end', item)
            if item in selected:
                listing.selection_set(i)

    refresh_listing()

    row = ttk.Frame(dialog)
    row.pack(fill='x', padx=10, pady=8)
    entry = ttk.Entry(row)
    entry.pack(side='left', fill='x', expand=True)

    def add():
        new_items = names(entry.get())
        if not new_items:
            return
        for item in new_items:
            if item not in candidates:
                candidates.append(item)
                listing.insert('end', item)
            try:
                idx = [listing.get(i) for i in range(listing.size())].index(item)
                listing.selection_set(idx)
            except ValueError:
                pass
        persist_items(new_items)
        entry.delete(0, 'end')

    def delete_selected():
        sel_indices = list(listing.curselection())
        if not sel_indices:
            return
        to_remove = [listing.get(i) for i in sel_indices]
        changed = False
        if hasattr(app, 'aux_equipments') and app.aux_equipments:
            for r in to_remove:
                if r in app.aux_equipments:
                    app.aux_equipments.remove(r)
                    changed = True
        for r in to_remove:
            if r in candidates:
                candidates.remove(r)
            if r in selected:
                selected.remove(r)
        if changed and hasattr(app, 'save_tab_config'):
            try:
                app.save_tab_config(force=True)
            except Exception as ex:
                print(f"Error saving tab config on delete: {ex}")
        refresh_listing()

    ttk.Button(row, text='추가', command=add).pack(side='left', padx=4)
    ttk.Button(row, text='선택 삭제', command=delete_selected).pack(side='left', padx=4)
    entry.bind('<Return>', lambda event: add())

    def apply():
        add()
        chosen = [listing.get(i) for i in listing.curselection()]
        variable.set('; '.join(chosen))
        persist_items(chosen)
        dialog.destroy()

    buttons = ttk.Frame(dialog)
    buttons.pack(pady=(0, 10))
    ttk.Button(buttons, text='적용', command=apply).pack(side='left', padx=4)
    ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='left', padx=4)
    dialog.grab_set()
