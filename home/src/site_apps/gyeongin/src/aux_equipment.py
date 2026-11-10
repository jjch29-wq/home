"""Optional auxiliary equipment selection for central/gyeongin daily inspection records."""
import tkinter as tk
from tkinter import ttk

METHOD_AUX_MAP = {
    'PAUT': [
        'PAUT SCANNER (MANUAL)',
        'MICROBE (스캐너 Body-137331, Encoder-136861)',
        'COBRA (스캐너 139310)',
        'PAUT PROBE 5L64-A2',
        'PAUT PROBE 2.25L64-A2',
        'PAUT PROBE 7.5CCEV35-16-A15',
        'PAUT PROBE 5L32-A31',
        'PAUT WEDGE SA2-N55S-IHC',
        'PAUT WEDGE SA2-OL-IHC',
        'PAUT WEDGE SA15-N60S-IHC',
        'PAUT WEDGE SA10-N55S',
        '대비시험편 STB-A1'
    ],
    'UT': [
        'UT 사각 PROBE 70˚ (2MHz 10*10)',
        'UT 사각 PROBE 60˚ (2MHz 10*10)',
        'UT 사각 PROBE 45˚ (2MHz 10*10)',
        'UT 사각 PROBE 70˚ (4MHz 10*10)',
        'UT 수직 PROBE (2MHz Φ20)',
        'UT 수직 PROBE (4MHz Φ10)',
        '대비시험편 STB-A1',
        '대비시험편 STB-A2',
        '대비시험편 RB-4',
        '접촉매질(Couplant)'
    ],
    'RT': [
        'ADR, SURVEY, 콜리메타',
        'ADR',
        'SURVEY',
        '콜리메타 (120˚)',
        '콜리메타 (360˚)',
        'R/CONTROL (25ft)',
        'SURVEY METER',
        'ALARM DOSIMETER (ADR)'
    ],
    'MT': [
        '자분살포기',
        'Black Light (자외선등)',
        '조도계',
        '자계강도계',
        'Pie Gauge'
    ],
    'PT': [
        '조도계',
        '온습도계'
    ]
}


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
    dialog.title('보조장비 및 프로브 선택 (다중 체크 가능)')
    dialog.geometry('600x530')
    dialog.transient(parent.winfo_toplevel())

    # Selected set (track items currently chosen)
    selected_set = set(names(variable.get()))

    # Determine current active method
    curr_method = '전체'
    method_var = getattr(app, 'daily_method_var', None) or getattr(app, 'method_var', None)
    if method_var and hasattr(method_var, 'get'):
        val = str(method_var.get()).strip().upper()
        if val in METHOD_AUX_MAP:
            curr_method = val

    # Collect all candidate items
    all_candidates = []
    # 1. From app.aux_equipments
    aux_list = getattr(app, 'aux_equipments', None)
    if aux_list:
        for it in aux_list:
            all_candidates.extend(names(it))

    # 2. From METHOD_AUX_MAP
    for m_list in METHOD_AUX_MAP.values():
        all_candidates.extend(m_list)

    # 3. From app.equipments
    for it in getattr(app, 'equipments', []):
        lbl = it.get('name', '') if isinstance(it, dict) else str(it)
        all_candidates.extend(names(lbl))

    # 4. From history
    history = getattr(app, 'daily_usage_df', None)
    if history is not None and '보조장비' in history.columns:
        for it in history['보조장비'].dropna():
            all_candidates.extend(names(it))

    # Preserve all existing selected items in candidates
    all_candidates.extend(selected_set)

    # Clean and deduplicate candidates
    all_candidates = list(dict.fromkeys(c.strip() for c in all_candidates if c and c.strip()))

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

    # Top info & Selected Count banner
    top_frame = ttk.Frame(dialog)
    top_frame.pack(fill='x', padx=12, pady=(10, 4))
    
    lbl_title = ttk.Label(top_frame, text='✔ 클릭하여 다중 선택 · 검사방법별 필터 및 실시간 검색 지원', font=('Malgun Gothic', 9, 'bold'))
    lbl_title.pack(side='left')

    lbl_count = ttk.Label(top_frame, text=f'선택: {len(selected_set)}개', font=('Malgun Gothic', 9, 'bold'), foreground='#0066cc')
    lbl_count.pack(side='right')

    # Method Filter Tabs
    tab_frame = ttk.Frame(dialog)
    tab_frame.pack(fill='x', padx=12, pady=4)
    
    active_method_filter = tk.StringVar(value=curr_method)

    def on_filter_change():
        apply_filter()

    method_buttons = ['전체', 'PAUT', 'UT', 'RT', 'MT', 'PT']
    for m in method_buttons:
        rb = ttk.Radiobutton(tab_frame, text=m, value=m, variable=active_method_filter, command=on_filter_change)
        rb.pack(side='left', padx=3)

    # Search Bar
    search_frame = ttk.Frame(dialog)
    search_frame.pack(fill='x', padx=12, pady=4)
    ttk.Label(search_frame, text='🔍 검색:').pack(side='left', padx=(0, 4))
    search_var = tk.StringVar()
    search_entry = ttk.Entry(search_frame, textvariable=search_var)
    search_entry.pack(side='left', fill='x', expand=True)

    # Listbox Body
    body = ttk.Frame(dialog)
    body.pack(fill='both', expand=True, padx=12, pady=4)
    scrollbar = ttk.Scrollbar(body)
    scrollbar.pack(side='right', fill='y')
    listing = tk.Listbox(body, selectmode='multiple', exportselection=False,
                         font=('Malgun Gothic', 10), yscrollcommand=scrollbar.set)
    listing.pack(fill='both', expand=True)
    scrollbar.configure(command=listing.yview)

    # Current displayed items in listing
    displayed_items = []
    updating_selection = False

    def apply_filter(*args):
        nonlocal displayed_items, updating_selection
        updating_selection = True
        keyword = search_var.get().strip().lower()
        method_filter = active_method_filter.get()

        if method_filter == '전체':
            allowed_set = None
        else:
            allowed_set = set(METHOD_AUX_MAP.get(method_filter, []))

        filtered = []
        for it in all_candidates:
            # Method filter check
            if allowed_set is not None:
                # If item is in allowed_set or already selected or matches method prefix
                matches_method = (it in allowed_set) or (method_filter in it) or (it in selected_set)
                if not matches_method:
                    continue

            # Keyword filter check
            if keyword and keyword not in it.lower():
                continue

            filtered.append(it)

        displayed_items = filtered
        listing.delete(0, 'end')
        for i, item in enumerate(displayed_items):
            listing.insert('end', item)
            if item in selected_set:
                listing.selection_set(i)

        lbl_count.config(text=f'선택: {len(selected_set)}개')
        updating_selection = False

    search_var.trace_add('write', apply_filter)

    def on_listbox_select(event):
        nonlocal updating_selection
        if updating_selection:
            return
        # Get current selection in visible items
        current_visible_selected = {displayed_items[i] for i in listing.curselection()}
        # For all items currently displayed: if visible and selected -> add to selected_set; if visible not selected -> remove
        for item in displayed_items:
            if item in current_visible_selected:
                selected_set.add(item)
            else:
                selected_set.discard(item)
        lbl_count.config(text=f'선택: {len(selected_set)}개')

    listing.bind('<<ListboxSelect>>', on_listbox_select)

    # Initial populate
    apply_filter()

    # Manual Add & Delete row
    row = ttk.Frame(dialog)
    row.pack(fill='x', padx=12, pady=6)
    entry = ttk.Entry(row)
    entry.pack(side='left', fill='x', expand=True)

    def add_custom():
        new_items = names(entry.get())
        if not new_items:
            return
        for item in new_items:
            if item not in all_candidates:
                all_candidates.append(item)
            selected_set.add(item)
        persist_items(new_items)
        entry.delete(0, 'end')
        apply_filter()

    def delete_selected():
        sel_indices = list(listing.curselection())
        if not sel_indices:
            return
        to_remove = [displayed_items[i] for i in sel_indices]
        changed = False
        if hasattr(app, 'aux_equipments') and app.aux_equipments:
            for r in to_remove:
                if r in app.aux_equipments:
                    app.aux_equipments.remove(r)
                    changed = True
        for r in to_remove:
            if r in all_candidates:
                all_candidates.remove(r)
            selected_set.discard(r)
        if changed and hasattr(app, 'save_tab_config'):
            try:
                app.save_tab_config(force=True)
            except Exception as ex:
                print(f"Error saving tab config on delete: {ex}")
        apply_filter()

    def clear_all():
        selected_set.clear()
        apply_filter()

    ttk.Button(row, text='직접 추가', command=add_custom).pack(side='left', padx=3)
    ttk.Button(row, text='선택 삭제', command=delete_selected).pack(side='left', padx=3)
    ttk.Button(row, text='선택 초기화', command=clear_all).pack(side='left', padx=3)
    entry.bind('<Return>', lambda event: add_custom())

    def apply():
        add_custom()
        # Sort or maintain selected items order
        chosen = [it for it in all_candidates if it in selected_set]
        variable.set('; '.join(chosen))
        persist_items(chosen)
        dialog.destroy()

    buttons = ttk.Frame(dialog)
    buttons.pack(fill='x', padx=12, pady=(0, 10))
    ttk.Button(buttons, text='선택 완료 (적용)', command=apply).pack(side='right', padx=4)
    ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='right', padx=4)

    dialog.grab_set()
