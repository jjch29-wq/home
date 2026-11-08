"""Material history viewer; no workbook or configuration writes."""
import tkinter as tk
from tkinter import ttk, messagebox
from site_apps.central.src.models.stock_history import material_history


def show_material_history(app):
    selected = app.stock_tree.selection()
    if len(selected) != 1:
        messagebox.showinfo('품목 이력', '이력을 볼 품목 하나를 선택해주세요.', parent=app.root)
        return
    values = app.stock_tree.item(selected[0], 'values')
    target = app.normalize_id(values[-1])
    matches = app.materials_df[app.materials_df['MaterialID'].apply(app.normalize_id) == target]
    if matches.empty:
        messagebox.showwarning('품목 이력', '품목 정보를 찾을 수 없습니다.', parent=app.root)
        return
    mat = matches.iloc[0]
    summary, rows = material_history(app, mat)
    window = tk.Toplevel(app.root)
    window.title('품목 이력 · ' + app.clean_nan(mat.get('품목명', '')))
    window.geometry('1000x560')
    window.minsize(750, 350)
    window.transient(app.root)
    unit = app.clean_nan(mat.get('관리단위', ''))
    text = (f"기초수량 {summary['stored_qty']:g}  |  입고 {summary['in_qty']:g}  |  "
            f"총감소 {summary['total_used']:g}  |  현재고 {summary['current_stock']:g} {unit}")
    ttk.Label(window, text=text, padding=10).pack(fill='x')
    ttk.Label(window, text='장비 사용은 재고를 차감하지 않습니다. 자동 차감 기록은 현장 사용과 중복 집계하지 않습니다.').pack(fill='x', padx=10)
    frame = ttk.Frame(window, padding=10)
    frame.pack(fill='both', expand=True)
    columns = ('date', 'kind', 'site', 'increase', 'decrease', 'note')
    tree = ttk.Treeview(frame, columns=columns, show='headings')
    for col, title, width in zip(columns, ('날짜', '구분', '현장', '증가', '감소', '비고'), (105, 175, 120, 75, 75, 350)):
        tree.heading(col, text=title)
        tree.column(col, width=width, anchor='w' if col == 'note' else 'center', stretch=col == 'note')
    scroll = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    horizontal = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
    tree.grid(row=0, column=0, sticky='nsew')
    scroll.grid(row=0, column=1, sticky='ns')
    horizontal.grid(row=1, column=0, sticky='ew')
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    for row in rows:
        tree.insert('', 'end', values=[row[c] if c not in ('increase', 'decrease') else
                    (f"{row[c]:g}" if row[c] else '') for c in columns])
    ttk.Label(window, text=f'이력 {len(rows)}건' if rows else '연결된 입출고·사용 기록이 없습니다. 기초수량만 반영됩니다.').pack(anchor='w', padx=10)
    ttk.Button(window, text='닫기', command=window.destroy).pack(pady=8)
    window.bind('<Escape>', lambda event: window.destroy())
