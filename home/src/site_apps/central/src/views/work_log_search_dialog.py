# History search and copy dialog for central-site daily work logs.
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry

from site_apps.central.src.services.section_line_helper import (
    extract_company_section_line_mapping,
    get_sections_for_company,
    get_lines_for_company_and_section,
)


def open_history_search_dialog(self):
    """저장된 전체 작업내역에서 일자·업체·라인번호별로 검색하고,
    원하는 행을 선택하여 해당 일자의 결과서로 이동해 수정하거나 현재 일보로 복사해옵니다."""
    if hasattr(self, '_history_search_window') and self._history_search_window is not None:
        try:
            if self._history_search_window.winfo_exists():
                self._history_search_window.deiconify()
                self._history_search_window.lift()
                self._history_search_window.focus_force()
                return
        except Exception:
            pass

    history = self.load_history()
    history_mapping = extract_company_section_line_mapping(history)
    all_records = []
    company_set = set()
    line_set = set()
    method_set = set()
    section_set = set()

    for date_key in sorted(history.keys(), reverse=True):
        day_data = history[date_key]
        for row_idx, r in enumerate(day_data.get('ndt_results', [])):
            if not any(str(r.get(k, '')).strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법', '구간')):
                continue
            co = str(r.get('업체', '')).strip()
            ln = str(r.get('라인번호', '')).strip()
            mth = str(r.get('검사방법', '')).strip().upper()
            sec = str(r.get('구간', '')).strip()
            if co: company_set.add(co)
            if ln: line_set.add(ln)
            if mth: method_set.add(mth)
            if sec: section_set.add(sec)
            all_records.append({
                'date': str(date_key),
                'row_idx': row_idx,
                'company': co,
                'work_category': str(r.get('공사구분', '')).strip(),
                'method': mth,
                'section': str(r.get('구간', '')).strip(),
                'work_type': str(r.get('작업구분', '')).strip(),
                'hot_loc': str(r.get('핫태핑개소', '')).strip(),
                'hot_deg': str(r.get('핫태핑차수', '')).strip(),
                'hot_cnt': str(r.get('핫태핑개소수', '')).strip(),
                'line_no': ln,
                'joint_no': str(r.get('Joint No.', '')).strip(),
                'pipe': str(r.get('관경', '')).strip(),
                'thick': str(r.get('두께', '')).strip(),
                'welder': str(r.get('용접사', '')).strip(),
                'joint_info': str(r.get('구간정보', '')).strip(),
                'result': str(r.get('결과', '')).strip(),
                'spec': str(r.get('규격', '')).strip(),
                'shift': str(r.get('근무구분', '')).strip(),
                'raw': dict(r),
            })

    window = tk.Toplevel(self)
    self._history_search_window = window

    def _on_window_close():
        self._history_search_window = None
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", _on_window_close)
    window.title("작업내역 검색 및 결과서 불러오기 (일자 · 업체 · 라인번호)")
    window.geometry("1120x660")
    window.minsize(880, 500)
    window.resizable(True, True)
    window.transient(self.winfo_toplevel())

    # 화면 최대화 / 원래 크기 복원 토글 함수
    def toggle_maximize(_event=None):
        try:
            is_zoomed = (window.state() == 'zoomed')
            if is_zoomed:
                window.state('normal')
                btn_text = "⛶ 화면 최대화"
            else:
                window.state('zoomed')
                btn_text = "🗗 원래 크기로"
            if 'top_max_btn' in locals() and top_max_btn.winfo_exists():
                top_max_btn.config(text=btn_text)
        except Exception:
            pass

    window.bind("<F11>", toggle_maximize)

    # 1. 상단 필터 바
    filter_box = ttk.LabelFrame(window, text="검색 및 필터 조건", padding=8)
    filter_box.pack(fill="x", padx=10, pady=(10, 6))

    # 일자 필터
    row1 = ttk.Frame(filter_box)
    row1.pack(fill="x", pady=2)

    ttk.Label(row1, text="검사일자:").pack(side="left")
    all_dates_var = tk.BooleanVar(value=True)

    try:
        from datetime import timedelta
        today_dt = datetime.strptime(self._selected_date_key(), '%Y-%m-%d').date()
    except Exception:
        today_dt = datetime.now().date()
    month_ago_dt = today_dt - timedelta(days=30)

    start_date_entry = DateEntry(row1, width=11, background='darkblue', foreground='white',
                                 date_pattern='yyyy-mm-dd', year=month_ago_dt.year,
                                 month=month_ago_dt.month, day=month_ago_dt.day)
    start_date_entry.pack(side="left", padx=4)

    ttk.Label(row1, text="~").pack(side="left")
    end_date_entry = DateEntry(row1, width=11, background='darkblue', foreground='white',
                               date_pattern='yyyy-mm-dd', year=today_dt.year,
                               month=today_dt.month, day=today_dt.day)
    end_date_entry.pack(side="left", padx=4)

    ttk.Checkbutton(row1, text="전체 기간(제한 없음)", variable=all_dates_var, command=lambda: refresh_rows()).pack(side="left", padx=(6, 16))

    ttk.Label(row1, text="업체:").pack(side="left")
    company_var = tk.StringVar(value="(전체)")
    company_combo = ttk.Combobox(row1, textvariable=company_var, values=['(전체)'] + sorted(company_set), width=13, state='readonly')
    company_combo.pack(side="left", padx=(4, 10))

    ttk.Label(row1, text="구간:").pack(side="left")
    section_var = tk.StringVar(value="(전체)")
    section_combo = ttk.Combobox(row1, textvariable=section_var, values=['(전체)'] + sorted(section_set), width=11, state='readonly')
    section_combo.pack(side="left", padx=(4, 10))

    ttk.Label(row1, text="방법:").pack(side="left")
    method_var = tk.StringVar(value="(전체)")
    method_combo = ttk.Combobox(row1, textvariable=method_var, values=['(전체)'] + sorted(method_set), width=8, state='readonly')
    method_combo.pack(side="left", padx=(4, 12))

    # 라인번호 및 키워드 검색
    row2 = ttk.Frame(filter_box)
    row2.pack(fill="x", pady=(4, 2))

    # 우측 끝에 최대화 버튼을 가장 먼저 배치하여 글자 잘림 방지 및 고정 공간 확보
    top_max_btn = ttk.Button(row2, text="⛶ 화면 최대화", command=toggle_maximize, width=13)
    top_max_btn.pack(side="right", padx=(8, 2))

    ttk.Label(row2, text="라인번호:").pack(side="left")

    # 라인번호 복수 선택 상태 및 컨테이너
    selected_lines = set()
    current_line_pool = sorted(line_set)

    line_frame = ttk.Frame(row2)
    line_frame.pack(side="left", padx=(4, 12))

    line_var = tk.StringVar(value="(전체)")
    line_entry = ttk.Entry(line_frame, textvariable=line_var, width=22, state='readonly', cursor='hand2')
    line_entry.pack(side="left")

    def update_line_display():
        if not selected_lines:
            line_var.set("(전체)")
        elif len(selected_lines) == 1:
            line_var.set(next(iter(selected_lines)))
        else:
            sample = sorted(selected_lines)
            disp = f"{len(selected_lines)}건 선택: {sample[0]}"
            if len(sample) > 1:
                disp += f", {sample[1]}"
            if len(sample) > 2:
                disp += "..."
            line_var.set(disp)

    def clear_line_selection():
        selected_lines.clear()
        update_line_display()
        refresh_rows()

    def open_line_multiselect_popup():
        if hasattr(window, '_line_popup') and window._line_popup is not None:
            try:
                if window._line_popup.winfo_exists():
                    window._line_popup.lift()
                    window._line_popup.focus_force()
                    return
            except Exception:
                pass

        popup = tk.Toplevel(window)
        window._line_popup = popup
        popup.title("라인번호 복수 선택")
        popup.transient(window)
        popup.resizable(True, True)

        try:
            popup.update_idletasks()
            rx = line_frame.winfo_rootx()
            ry = line_frame.winfo_rooty() + line_frame.winfo_height() + 4
            sw = popup.winfo_screenwidth()
            sh = popup.winfo_screenheight()
            pw, ph = 380, 430
            if rx + pw > sw:
                rx = max(10, sw - pw - 20)
            if ry + ph > sh:
                ry = max(10, line_frame.winfo_rooty() - ph - 10)
            popup.geometry(f"{pw}x{ph}+{rx}+{ry}")
        except Exception:
            popup.geometry("380x430")

        top_p = ttk.Frame(popup, padding=(10, 8, 10, 4))
        top_p.pack(fill="x")
        ttk.Label(top_p, text="검색할 라인번호를 체크하세요 (복수 선택 가능)", font=("맑은 고딕", 9, "bold")).pack(anchor="w")

        search_box = ttk.Frame(top_p)
        search_box.pack(fill="x", pady=(6, 4))
        ttk.Label(search_box, text="🔍 목록 필터:").pack(side="left")
        p_filter_var = tk.StringVar()
        p_filter_entry = ttk.Entry(search_box, textvariable=p_filter_var)
        p_filter_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))

        btn_box = ttk.Frame(top_p)
        btn_box.pack(fill="x", pady=(2, 4))

        items_container = ttk.Frame(popup, padding=(10, 0, 10, 0))
        items_container.pack(fill="both", expand=True)

        canvas = tk.Canvas(items_container, highlightthickness=1, relief="sunken", bg="#ffffff")
        v_scroll = ttk.Scrollbar(items_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_canvas_resize(e):
            canvas.itemconfig(inner_id, width=e.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        def _on_inner_resize(_e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_inner_resize)

        item_vars = {}
        for line in current_line_pool:
            item_vars[line] = tk.BooleanVar(value=(line in selected_lines))

        bottom_p = ttk.Frame(popup, padding=(10, 8, 10, 8))
        bottom_p.pack(fill="x")

        status_lbl = ttk.Label(bottom_p, text="")
        status_lbl.pack(side="left")

        def _update_status():
            cnt = sum(1 for v in item_vars.values() if v.get())
            if cnt == 0:
                status_lbl.config(text="선택: 전체 (0건)")
            else:
                status_lbl.config(text=f"선택: {cnt}건")

        def _render_list():
            for c in inner.winfo_children():
                c.destroy()
            q = p_filter_var.get().strip().lower()
            vis = [l for l in current_line_pool if (not q or q in l.lower())]
            if not vis:
                ttk.Label(inner, text="일치하는 라인번호가 없습니다.", foreground="#888888").pack(anchor="w", padx=8, pady=8)
                return
            for line in vis:
                row_item = ttk.Frame(inner)
                row_item.pack(fill="x", pady=1, padx=4)
                cb = ttk.Checkbutton(row_item, text=line, variable=item_vars[line], command=_update_status)
                cb.pack(side="left", fill="x", expand=True, anchor="w")

        _render_list()
        _update_status()
        p_filter_var.trace_add("write", lambda *_a: _render_list())

        def _select_all():
            q = p_filter_var.get().strip().lower()
            for line in current_line_pool:
                if not q or q in line.lower():
                    item_vars[line].set(True)
            _update_status()

        def _deselect_all():
            for line in current_line_pool:
                item_vars[line].set(False)
            _update_status()

        ttk.Button(btn_box, text="☑ 전체 선택", width=10, command=_select_all).pack(side="left", padx=(0, 4))
        ttk.Button(btn_box, text="☐ 전체 해제", width=10, command=_deselect_all).pack(side="left", padx=4)

        def _apply_and_close():
            selected_lines.clear()
            for line, var in item_vars.items():
                if var.get():
                    selected_lines.add(line)
            update_line_display()
            popup.destroy()
            refresh_rows()

        ttk.Button(bottom_p, text="닫기", width=8, command=popup.destroy).pack(side="right", padx=(4, 0))
        ttk.Button(bottom_p, text="선택 적용 (확인)", width=14, command=_apply_and_close).pack(side="right", padx=4)

        popup.bind("<Return>", lambda _e: _apply_and_close())
        popup.bind("<Escape>", lambda _e: popup.destroy())
        p_filter_entry.focus_set()

    line_btn = ttk.Button(line_frame, text="선택 ▾", width=6, command=open_line_multiselect_popup)
    line_btn.pack(side="left", padx=(2, 0))

    line_clear_btn = ttk.Button(line_frame, text="✕", width=2, command=clear_line_selection)
    line_clear_btn.pack(side="left", padx=(2, 0))

    line_entry.bind('<Button-1>', lambda _e: open_line_multiselect_popup())

    # 호환성을 위한 참조 유지
    window._selected_lines = selected_lines
    window._open_line_multiselect_popup = open_line_multiselect_popup

    ttk.Label(row2, text="통합 검색어:").pack(side="left")
    search_var = tk.StringVar()
    search_entry = ttk.Entry(row2, textvariable=search_var, width=16)
    search_entry.pack(side="left", padx=(4, 6))
    ttk.Label(row2, text="(Joint, 구간 등)", foreground="#666666").pack(side="left", padx=(0, 10))

    ttk.Button(row2, text="🔍 조회", command=lambda: refresh_rows()).pack(side="left", padx=3)

    def reset_filters():
        all_dates_var.set(True)
        company_var.set("(전체)")
        section_var.set("(전체)")
        section_combo['values'] = ['(전체)'] + sorted(section_set)
        method_var.set("(전체)")
        selected_lines.clear()
        current_line_pool.clear()
        current_line_pool.extend(sorted(line_set))
        update_line_display()
        search_var.set("")
        refresh_rows()

    ttk.Button(row2, text="초기화", command=reset_filters).pack(side="left", padx=2)

    def _update_search_lines():
        sel_co = company_var.get().strip()
        sel_sec = section_var.get().strip()
        current_line_pool.clear()
        lines = get_lines_for_company_and_section(history_mapping, sel_co, sel_sec)
        current_line_pool.extend(lines)
        selected_lines.intersection_update(current_line_pool)
        update_line_display()
        refresh_rows()

    def on_company_change(*_args):
        sel_co = company_var.get().strip()
        sec_opts = ['(전체)'] + get_sections_for_company(history_mapping, sel_co)
        section_combo['values'] = sec_opts
        if section_var.get() not in sec_opts:
            section_var.set('(전체)')
        _update_search_lines()

    def on_section_change(*_args):
        _update_search_lines()

    company_combo.bind('<<ComboboxSelected>>', on_company_change)
    section_combo.bind('<<ComboboxSelected>>', on_section_change)
    method_combo.bind('<<ComboboxSelected>>', lambda _e: refresh_rows())
    search_entry.bind('<Return>', lambda _e: refresh_rows())

    # 2. 중앙 테이블
    table_frame = ttk.Frame(window, padding=(10, 4, 10, 4))
    table_frame.pack(fill="both", expand=True)

    cols = ('date', 'company', 'work_category', 'method', 'section', 'line_no', 'joint_no', 'pipe', 'thick', 'welder', 'result', 'spec')
    tree = ttk.Treeview(table_frame, columns=cols, show='headings', selectmode='extended')

    headers = {
        'date': ('검사일자', 95, 'center'),
        'company': ('업체', 120, 'w'),
        'work_category': ('공사구분', 85, 'center'),
        'method': ('방법', 65, 'center'),
        'section': ('구간', 85, 'center'),
        'line_no': ('라인번호', 240, 'w'),
        'joint_no': ('Joint No.', 90, 'center'),
        'pipe': ('관경', 70, 'center'),
        'thick': ('두께', 60, 'center'),
        'welder': ('용접사', 90, 'center'),
        'result': ('결과', 65, 'center'),
        'spec': ('규격/근무', 95, 'center'),
    }
    for col_id, (header_text, col_width, align) in headers.items():
        tree.heading(col_id, text=header_text)
        tree.column(col_id, width=col_width, minwidth=40, anchor=align, stretch=False)

    v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

    tree.grid(row=0, column=0, sticky="nsew")
    v_scroll.grid(row=0, column=1, sticky="ns")
    h_scroll.grid(row=1, column=0, sticky="ew")
    table_frame.columnconfigure(0, weight=1)
    table_frame.rowconfigure(0, weight=1)

    # 3. 하단 상태 및 액션 버튼 바
    status_var = tk.StringVar()
    ttk.Label(window, textvariable=status_var, padding=(12, 2, 12, 4), foreground="#1d4ed8").pack(fill="x")

    bottom_frame = ttk.Frame(window, padding=(10, 4, 10, 8))
    bottom_frame.pack(fill="x")

    # 1단: 가이드 문구 + 이동/복사 시 창 닫기 체크박스
    guide_row = ttk.Frame(bottom_frame)
    guide_row.pack(fill="x", pady=(0, 6))

    ttk.Label(
        guide_row,
        text="💡 항목 더블클릭 또는 [이동하여 수정]을 누르면 해당 일자의 결과서로 이동해 편집할 수 있습니다.",
        foreground="#555555"
    ).pack(side="left")

    close_on_move_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(guide_row, text="이동/복사 시 창 닫기", variable=close_on_move_var).pack(side="right")

    # 2단: 액션 버튼들(우측)
    action_row = ttk.Frame(bottom_frame)
    action_row.pack(fill="x")

    filtered_records = []

    def refresh_rows(*_args):
        nonlocal filtered_records
        tree.delete(*tree.get_children())

        sel_co = company_var.get().strip()
        sel_sec = section_var.get().strip()
        sel_mth = method_var.get().strip()
        sel_line = line_var.get().strip()
        query = search_var.get().strip().lower()
        all_dates = all_dates_var.get()

        try:
            start_str = start_date_entry.get_date().strftime('%Y-%m-%d')
            end_str = end_date_entry.get_date().strftime('%Y-%m-%d')
        except Exception:
            start_str, end_str = "0000-00-00", "9999-99-99"

        filtered = []
        for r in all_records:
            if not all_dates:
                if not (start_str <= r['date'] <= end_str):
                    continue
            if sel_co and sel_co != '(전체)' and r['company'] != sel_co:
                continue
            if sel_sec and sel_sec != '(전체)' and r['section'] != sel_sec:
                continue
            if sel_mth and sel_mth != '(전체)' and r['method'] != sel_mth:
                continue
            if selected_lines:
                if r['line_no'] not in selected_lines:
                    continue
            elif sel_line and sel_line != '(전체)':
                if sel_line.lower() not in r['line_no'].lower():
                    continue
            if query:
                search_corpus = f"{r['date']} {r['company']} {r['line_no']} {r['joint_no']} {r['section']} {r['welder']} {r['joint_info']}".lower()
                if query not in search_corpus:
                    continue
            filtered.append(r)

        filtered_records = filtered
        for idx, r in enumerate(filtered):
            spec_display = f"{r['spec']} ({r['shift']})" if r['shift'] else r['spec']
            tree.insert(
                '', 'end', iid=str(idx),
                values=(r['date'], r['company'], r.get('work_category', ''), r['method'], r['section'], r['line_no'],
                        r['joint_no'], r['pipe'], r['thick'], r['welder'], r['result'], spec_display)
            )

        status_var.set(f"총 {len(filtered)}건의 작업내역이 조회되었습니다. (선택 후 버튼 클릭)")

    def move_to_selected(_event=None):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("선택 확인", "이동할 작업내역을 목록에서 먼저 선택하세요.", parent=window)
            return
        idx = int(sel[0])
        record = filtered_records[idx]
        target_date = record['date']
        target_line = record['line_no']
        target_joint = record['joint_no']

        current_date = self._selected_date_key()
        if target_date != current_date:
            baseline = getattr(self, '_log_baseline', None)
            if baseline is not None and self._capture_log_input() != baseline:
                if not messagebox.askyesno(
                    '일보 날짜 이동',
                    f'현재 {current_date} 일보에 저장하지 않은 작성 내용이 있습니다.\n'
                    f'작성 내용을 버리고 {target_date} 일보로 이동할까요?',
                    parent=window
                ):
                    return
            try:
                self.date_entry.set_date(datetime.strptime(target_date, '%Y-%m-%d').date())
            except Exception:
                pass
            self.on_date_change()

        target_row_idx = None
        for r_i, row in enumerate(self.ndt_grid_entries):
            if (row['라인번호'].get().strip() == target_line and
                row['Joint No.'].get().strip() == target_joint and
                row['업체'].get().strip() == record['company']):
                target_row_idx = r_i
                break
        if target_row_idx is None:
            for r_i, row in enumerate(self.ndt_grid_entries):
                if (row['라인번호'].get().strip() == target_line and
                    row['Joint No.'].get().strip() == target_joint):
                    target_row_idx = r_i
                    break

        if target_row_idx is not None:
            self._clear_ndt_selection()
            self._select_ndt_row(self.ndt_grid_entries[target_row_idx])
            self._focus_ndt_joint(target_row_idx)

        if close_on_move_var.get():
            _on_window_close()
            if target_row_idx is not None:
                messagebox.showinfo(
                    "작업내역 이동 완료",
                    f"[{target_date}] {record['company']} - {target_line} (Joint {target_joint})\n"
                    f"해당 행({target_row_idx + 1}행)을 선택했습니다. 수정 후 상단 [💾 자동 집계/저장]을 눌러 주세요.",
                    parent=self
                )
        else:
            if target_row_idx is not None:
                status_var.set(
                    f"✅ [{target_date}] {record['company']} · {target_line} · Joint {target_joint} "
                    f"-> 결과서 {target_row_idx + 1}행 선택 및 포커스 완료! 메인 화면에서 바로 수정 후 [💾 자동 집계/저장]하세요."
                )
            else:
                status_var.set(
                    f"⚠️ [{target_date}] 일보로 이동했으나 결과서에서 일치하는 행을 찾지 못했습니다."
                )

    def copy_to_current():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("선택 확인", "가져올 작업내역을 1개 이상 선택하세요.", parent=window)
            return
        items_to_copy = [filtered_records[int(i)] for i in sel]

        empty_indices = [
            idx for idx, row in enumerate(self.ndt_grid_entries)
            if not any(row[k].get().strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법'))
        ]

        needed = len(items_to_copy) - len(empty_indices)
        if needed > 0:
            self._append_ndt_rows(((needed + 9) // 10) * 10)
            empty_indices = [
                idx for idx, row in enumerate(self.ndt_grid_entries)
                if not any(row[k].get().strip() for k in ('업체', '라인번호', 'Joint No.', '검사방법'))
            ]

        first_filled_idx = None
        for item, row_idx in zip(items_to_copy, empty_indices):
            row = self.ndt_grid_entries[row_idx]
            if first_filled_idx is None:
                first_filled_idx = row_idx
            field_map = {
                '업체': item['company'],
                '공사구분': item.get('work_category', ''),
                '검사방법': item['method'],
                '구간': item['section'],
                '작업구분': item['work_type'],
                '핫태핑개소': item['hot_loc'],
                '핫태핑차수': item['hot_deg'],
                '핫태핑개소수': item['hot_cnt'],
                '라인번호': item['line_no'],
                'Joint No.': item['joint_no'],
                '관경': item['pipe'],
                '두께': item['thick'],
                '용접사': item['welder'],
                '구간정보': item['joint_info'],
                '결과': item['result'] or '합격',
                '규격': item['spec'],
                '근무구분': item['shift'] or '주간',
            }
            for k, v in field_map.items():
                if k in row:
                    if hasattr(row[k], 'set'):
                        row[k].set(v)
                    elif hasattr(row[k], 'delete'):
                        row[k].delete(0, tk.END)
                        row[k].insert(0, v)
            self._sync_ndt_spec(row)

        if first_filled_idx is not None:
            self._clear_ndt_selection()
            self._select_ndt_row(self.ndt_grid_entries[first_filled_idx])
            self._focus_ndt_joint(first_filled_idx)

        if close_on_move_var.get():
            _on_window_close()
            messagebox.showinfo(
                "가져오기 완료",
                f"{len(items_to_copy)}건의 작업내역을 현재 일보({self._selected_date_key()}) 결과서로 가져왔습니다.\n"
                "수정 또는 확인 후 상단 [💾 자동 집계/저장]을 눌러 주세요.",
                parent=self
            )
        else:
            status_var.set(
                f"✅ {len(items_to_copy)}건의 작업내역을 현재 일보({self._selected_date_key()}) 결과서로 복사했습니다. "
                f"메인 화면에서 확인 후 [💾 자동 집계/저장]하세요."
            )

    tree.bind('<Double-1>', move_to_selected)
    tree.bind('<Return>', move_to_selected)

    ttk.Button(action_row, text="닫기", command=_on_window_close, width=10).pack(side="right", padx=(4, 0))
    ttk.Button(action_row, text="📋 현재 일보로 행 복사하여 가져오기", command=copy_to_current).pack(side="right", padx=4)
    ttk.Button(action_row, text="📌 해당 일자로 이동하여 수정", command=move_to_selected).pack(side="right", padx=4)

    refresh_rows()
    search_entry.focus_set()
