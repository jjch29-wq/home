"""Optional auxiliary equipment selection for central daily inspection records."""
import tkinter as tk
from tkinter import ttk
import re
import unicodedata


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


def _label(item):
    if isinstance(item, dict):
        return str(item.get('name', '') or '').strip()
    return str(item or '').strip()


def _clean_equipment_label(text):
    text = unicodedata.normalize('NFKC', str(text)).strip()
    text = re.sub(r'\[(?:자동등록|수정등록)\]', '', text)
    text = re.sub(r'(?i)(PAUT|UT|RT|MT|PT)\s*장비', r'\1', text)
    text = re.sub(r'(?i)\b(\w+)(?:\s+\1)+\b', r'\1', text)
    return re.sub(r'\s+', ' ', text).strip()


def _equipment_key(text):
    # Keep serial identifiers intact, including their digit order. Only the
    # display order of name/model tokens and repeated words is irrelevant.
    text = _clean_equipment_label(text).casefold()
    tokens = re.findall(r'[a-z]+[a-z0-9]*(?:[-_]\d+)+|[a-z]+\d+[a-z0-9]*|\d+(?:[-_]\d+)+|\d+|[a-z]+|[가-힣]+', text)
    return tuple(sorted(set(re.sub(r'[-_]', '', token) for token in tokens)))


def equipment_catalog(app, mode='aux', selected='', include_unknown=False):
    """Use only this site's master, inventory and history; retain custom entries."""
    items = names(selected)
    attribute = 'equipments' if mode == 'main' else 'aux_equipments'
    for item in getattr(app, attribute, []) or []:
        items.extend(names(_label(item)))
    if mode == 'main' and hasattr(app, '_get_equipment_candidates'):
        for item in app._get_equipment_candidates(include_all=False):
            items.extend(names(_label(item)))
    history = getattr(app, 'daily_usage_df', None)
    column = '장비명' if mode == 'main' else '보조장비'
    if history is not None and column in history.columns:
        for item in history[column].dropna():
            items.extend(names(item))
    inventory = getattr(app, 'materials_df', None)
    inventory_aliases = {}
    classifications = {}
    if inventory is not None:
        for _, row in inventory.iterrows():
            active = str(row.get('Active', 1)).strip().lower()
            if active in ('0', '0.0', 'false'):
                continue
            name = str(row.get('품목명', '') or '').strip()
            if not names(name):
                continue
            if hasattr(app, '_is_consumable_material') and app._is_consumable_material(name, ''):
                continue
            parts = []
            for key in ('모델명', '품목명'):
                part = str(row.get(key, '')).strip()
                if names(part) and part not in parts:
                    parts.append(part)
            serial = str(row.get('SN', '')).strip()
            if names(serial):
                parts.append(f'({serial})')
            if parts:
                label = ' '.join(parts)
                items.append(label)
                from site_apps.central.src.method_catalog import row_classification
                classifications[_equipment_key(label)] = row_classification(row)
                # A short name identifies a unit only when the active inventory
                # has exactly one matching model/S/N. Ambiguous names stay apart.
                inventory_aliases.setdefault(_equipment_key(name), set()).add(_equipment_key(label))
    groups = {}
    raw_keys = {}
    for item in items:
        if item == '전체' or not names(item):
            continue
        key = _equipment_key(item)
        raw_keys[item] = key
        groups.setdefault(key, []).append(_clean_equipment_label(item))
    redirects = {key: next(iter(targets)) for key, targets in inventory_aliases.items()
                 if len(targets) == 1 and next(iter(targets)) != key}
    def resolve(key):
        visited = set()
        while key in redirects and key not in visited:
            visited.add(key)
            key = redirects[key]
        return key
    canonical = {key: min(labels, key=lambda text: (len(text), text.casefold()))
                 for key, labels in groups.items()}
    aliases = {item: canonical[resolve(key)] for item, key in raw_keys.items()}
    candidates = sorted(set(aliases.values()), key=str.casefold)
    from site_apps.central.src.method_catalog import current_method, classify
    method = current_method(app)
    if method:
        def allowed(label):
            methods, role = classifications.get(_equipment_key(label), classify(label))
            return (method in methods and role == mode) or (not methods and include_unknown)
        candidates = [item for item in candidates if allowed(item)]
    return candidates, aliases


def equipment_candidates(app, mode='aux', selected=''):
    return equipment_catalog(app, mode, selected)[0]


def set_selection(target, selected):
    """Normalize without collapsing distinct S/Ns; support StringVar and Entry."""
    text = '; '.join(dict.fromkeys(item for value in selected for item in names(value)))
    if hasattr(target, 'set'):
        target.set(text)
    else:
        target.delete(0, 'end')
        target.insert(0, text)
    return text


def _bind_count_badge(target, label):
    def refresh(*_):
        label.configure(text=f'{len(names(target.get()))}개 선택')
    if isinstance(target, tk.Variable):
        target.trace_add('write', refresh)
    else:
        variable = tk.StringVar(master=target, value=target.get())
        target.configure(textvariable=variable)
        variable.trace_add('write', refresh)
        target._selection_variable = variable
    refresh()


def bind_main_selector(parent, app, target):
    ttk.Button(parent, text='복수 선택', width=9,
               command=lambda: select_main(parent, app, target)).pack(side='left', padx=3)
    badge = ttk.Label(parent)
    badge.pack(side='left', padx=3)
    _bind_count_badge(target, badge)


def add_selector(parent, app, variable, width=25):
    frame = ttk.Frame(parent)
    ttk.Entry(frame, textvariable=variable, width=width).pack(side='left', fill='x', expand=True)
    ttk.Button(frame, text='복수 선택', width=9,
               command=lambda: select(frame, app, variable)).pack(side='left', padx=3)
    badge = ttk.Label(frame)
    badge.pack(side='left', padx=3)
    _bind_count_badge(variable, badge)
    if hasattr(app, 'open_list_management_dialog'):
        tk.Button(frame, text='⚙️', font=('Malgun Gothic', 8), bd=0,
                  bg=getattr(app, 'theme_bg', '#f0f0f0'), fg='blue', cursor='hand2',
                  command=lambda: app.open_list_management_dialog('aux_equipments')).pack(side='left', padx=2)
    return frame


def select(parent, app, variable):
    return _open_multi_select_dialog(parent, app, variable, 'aux')


def select_main(parent, app, target):
    return _open_multi_select_dialog(parent, app, target, 'main')


def _open_multi_select_dialog(parent, app, target, mode):
    title = '주 장비' if mode == 'main' else '보조장비 · 프로브 · 시험편'
    dialog = tk.Toplevel(parent)
    dialog.title(f'{title} 복수 선택')
    owner = parent.winfo_toplevel()
    x = max(0, owner.winfo_rootx() + (owner.winfo_width() - 820) // 2)
    y = max(0, owner.winfo_rooty() + (owner.winfo_height() - 560) // 2)
    dialog.geometry(f'820x560+{x}+{y}')
    dialog.minsize(650, 420)
    dialog.transient(parent.winfo_toplevel())
    candidates, aliases = equipment_catalog(app, mode, target.get())
    selected = dict.fromkeys(aliases.get(item, item) for item in names(target.get()))
    method = getattr(app, 'cb_daily_test_method', None)
    method = method.get().strip() if method is not None else ''
    ttk.Label(dialog, text=f'{title} · 현재 검사방법: {method or "미선택"}',
              font=('Malgun Gothic', 11, 'bold')).pack(anchor='w', padx=12, pady=(12, 4))
    ttk.Label(dialog, text='체크하여 복수 선택 · 모델/S/N 확인 · 검색을 바꿔도 선택 유지').pack(anchor='w', padx=12)
    search_row = ttk.Frame(dialog)
    search_row.pack(fill='x', padx=12, pady=8)
    ttk.Label(search_row, text='검색:').pack(side='left')
    query = tk.StringVar(master=dialog)
    include_unknown = tk.BooleanVar(master=dialog, value=False)
    search = ttk.Entry(search_row, textvariable=query)
    search.pack(side='left', fill='x', expand=True, padx=6)
    # Reserve controls before the expanding result list, including at high DPI.
    buttons = ttk.Frame(dialog)
    buttons.pack(side='bottom', fill='x', padx=12, pady=10)
    direct = ttk.Frame(dialog)
    direct.pack(side='bottom', fill='x', padx=12, pady=5)
    count = ttk.Label(dialog)
    count.pack(side='bottom', anchor='w', padx=12, pady=5)
    body = ttk.Frame(dialog)
    body.pack(fill='both', expand=True, padx=12)
    tree = ttk.Treeview(body, columns=('checked', 'name'), show='headings', selectmode='browse')
    tree.heading('checked', text='선택')
    tree.heading('name', text='장비명 · 모델 · S/N')
    tree.column('checked', width=55, stretch=False, anchor='center')
    tree.column('name', width=650)
    scrollbar = ttk.Scrollbar(body, command=tree.yview)
    scrollbar.pack(side='right', fill='y')
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(fill='both', expand=True)
    visible = {}

    def render(*_):
        nonlocal candidates
        candidates, _ = equipment_catalog(app, mode, target.get(), include_unknown.get())
        for row in tree.get_children():
            tree.delete(row)
        visible.clear()
        needle = query.get().strip().casefold()
        for index, item in enumerate(candidates):
            if needle not in item.casefold():
                continue
            key = str(index)
            visible[key] = item
            tree.insert('', 'end', iid=key, values=('☑' if item in selected else '☐', item))
        hidden = sum(item not in candidates for item in selected)
        count.configure(text=f'{len(selected)}개 선택 · 현재 목록 {len(visible)}개 · 목록 밖 기존 선택 {hidden}개 유지')

    def toggle(event=None):
        row = tree.identify_row(event.y) if event is not None and hasattr(event, 'y') else next(iter(tree.selection()), '')
        if row not in visible:
            return
        item = visible[row]
        if item in selected:
            selected.pop(item)
        else:
            selected[item] = None
        tree.set(row, 'checked', '☑' if item in selected else '☐')
        count.configure(text=f'{len(selected)}개 선택 · 현재 목록 {len(visible)}개')
        return 'break'

    def clear():
        selected.clear()
        render()

    def select_visible():
        selected.update(dict.fromkeys(visible.values()))
        render()

    ttk.Label(direct, text='직접 입력:').pack(side='left')
    entry = ttk.Entry(direct)
    entry.pack(side='left', fill='x', expand=True, padx=6)

    def add():
        for item in names(entry.get()):
            key = _equipment_key(item)
            item = aliases.get(item, next((candidate for candidate in candidates
                                           if _equipment_key(candidate) == key), _clean_equipment_label(item)))
            if item not in candidates:
                candidates.append(item)
            selected[item] = None
        entry.delete(0, 'end')
        query.set('')
        render()

    ttk.Button(direct, text='선택에 추가', command=add).pack(side='right')
    ttk.Checkbutton(search_row, text='미분류 포함', variable=include_unknown,
                    command=render).pack(side='right')
    entry.bind('<Return>', lambda event: add())
    ttk.Button(buttons, text='검색 결과 전체 선택', command=select_visible).pack(side='left')
    ttk.Button(buttons, text='선택 비우기', command=clear).pack(side='left', padx=5)

    def apply():
        if entry.get().strip():
            add()
        set_selection(target, selected)
        attribute = 'equipments' if mode == 'main' else 'aux_equipments'
        master = getattr(app, attribute, None)
        changed = False
        if isinstance(master, list):
            known = {name for item in master for name in names(_label(item))}
            for item in selected:
                if item not in known:
                    master.append(item)
                    known.add(item)
                    changed = True
        if changed and hasattr(app, 'save_tab_config'):
            app.save_tab_config(force=True)
        dialog.destroy()

    ttk.Button(buttons, text='적용', command=apply).pack(side='right', padx=4)
    ttk.Button(buttons, text='취소', command=dialog.destroy).pack(side='right', padx=4)
    tree.bind('<ButtonRelease-1>', toggle)
    tree.bind('<space>', lambda event: toggle())
    query.trace_add('write', render)
    render()
    dialog.grab_set()
    search.focus_set()
    return dialog
