# -*- coding: utf-8 -*-
"""
중앙지사 비파괴검사 보고서 테이블 조작, 컬럼 정렬/관리, 일괄수정 컨트롤러 Mixin
"""

import json
import logging
import os
import re
import sys
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

logger = logging.getLogger(__name__)

class ReportTableControllerMixin:
    """테이블 데이터 파퓰레이션, 정렬, 컬럼 설정, 일괄 변경 및 저장 컨트롤러 Mixin"""
    def manage_columns(self, mode="PMI"):
        """프로그램의 각 모드별 미리보기 컬럼 보이기/숨기기 관리"""
        if mode == "RT":
            tree, idx_map, data, keys_attr = self.rt_preview_tree, self.rt_item_idx_map, self.rt_extracted_data, "rt_column_keys"
        elif mode == "KOGAS":
            tree, idx_map, data, keys_attr = self.kogas_preview_tree, self.kogas_item_idx_map, self.kogas_extracted_data, "kogas_column_keys"
        elif mode == "PT":
            tree, idx_map, data, keys_attr = self.pt_preview_tree, self.pt_item_idx_map, self.pt_extracted_data, "pt_column_keys"
        elif mode == "MT":
            tree, idx_map, data, keys_attr = self.mt_preview_tree, self.mt_item_idx_map, self.mt_extracted_data, "mt_column_keys"
        elif mode == "PAUT":
            tree, idx_map, data, keys_attr = self.paut_preview_tree, self.paut_item_idx_map, self.paut_extracted_data, "paut_column_keys"
        else:
            tree, idx_map, data, keys_attr = self.preview_tree, self.item_idx_map, self.extracted_data, "column_keys"

        dialog = tk.Toplevel(self.root)
        dialog.title(f"{mode} 컬럼 관리")
        dialog.geometry("420x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text=f"{mode} 표시할 컬럼을 클릭 순서대로 정렬 및 선택하세요:", font=("Malgun Gothic", 10, "bold")).pack(pady=(0, 5), anchor='w')
        ttk.Label(main_frame, text="* 체크 시 클릭한 순서대로 번호(순서)가 부여됩니다.", foreground="gray", font=("Malgun Gothic", 8)).pack(pady=(0, 10), anchor='w')
        
        # 엑셀 데이터 또는 기존 정의에서 모든 키 추출
        all_keys = []
        if data:
            all_keys = list(data[0].keys())
        else:
            # 데이터가 없을 경우 현재 정의된 키들을 기본으로 함
            all_keys = list(getattr(self, keys_attr))
            
        # 내부 관리용 키 제외
        exclude = ["selected", "date_filtered", "line_filtered", "size_filtered", "order_index", "visual_group_joint", "is_merged_iso", "is_merged_joint"]
        display_keys = [k for k in all_keys if k not in exclude]
        if mode == "PAUT" and "TCG MAX(%)" not in display_keys:
            display_keys.append("TCG MAX(%)")
        
        # 스크롤 가능한 영역
        scroll_frame = ttk.Frame(main_frame)
        scroll_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mouse wheel support for popup
        def _on_popup_mw(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_popup_mw)
        
        # RT용 레이블 매핑 정의 (표준화된 ID 대응)
        label_map = {}
        if mode == "RT" or mode == "KOGAS":
            label_map = {
                "Dwg": "Drawing No.", "Joint": "Film Ident. No.", "Loc": "Film Location",
                "Acc": "Acc", "Rej": "Rej", "Deg": "Deg", "Welder": "Welder No",
                "T": "T", "Mat": "Mat", "No": "No", "Date": "Date", "Sec": "SEC", "Remarks": "Remarks",
                "Size": "구경",
                "D1": "① Crack", "D2": "② IP", "D3": "③ LF", "D4": "④ Slag", "D5": "⑤ Por",
                "D6": "⑥ U/C", "D7": "⑦ RUC", "D8": "⑧ BT", "D9": "⑨ TI", "D10": "⑩ CP",
                "D11": "⑪ RC", "D12": "⑫ Mis", "D13": "⑬ EP", "D14": "⑭ SD", "D15": "⑮ Oth"
            }
        elif mode == "PAUT":
            label_map = {"ISO": "ISO/DWG", "Location": "Loc"}

        vars_dict = {}
        current_keys = list(getattr(self, keys_attr))
        
        # 클릭 순서를 추적할 리스트 초기화 (현재 활성화된 키들의 순서 반영)
        click_ordered_keys = [k for k in current_keys if k in display_keys]
        ordered_keys_for_ui = list(display_keys)
            
        def _on_check_toggled(key, var):
            if var.get():
                if key not in click_ordered_keys:
                    click_ordered_keys.append(key)
            else:
                if key in click_ordered_keys:
                    click_ordered_keys.remove(key)
            _refresh_list()

        def _refresh_list():
            for widget in scrollable_frame.winfo_children():
                widget.destroy()
            
            for i, k in enumerate(ordered_keys_for_ui):
                f = tk.Frame(scrollable_frame, background="#ffffff")
                f.pack(fill='x', pady=2)
                
                v = vars_dict.get(k)
                if v is None:
                    v = tk.BooleanVar(value=(k in current_keys))
                    vars_dict[k] = v
                
                display_name = label_map.get(k, k)
                if v.get() and k in click_ordered_keys:
                    order_num = click_ordered_keys.index(k) + 1
                    label_text = f"[{order_num}] {display_name}"
                else:
                    label_text = f"[ ] {display_name}"
                
                cb = ttk.Checkbutton(f, text=label_text, variable=v, command=lambda _k=k, _v=v: _on_check_toggled(_k, _v))
                cb.pack(side='left', padx=5)
                
                # Up/Down Buttons as supplementary tool inside order list
                btn_f = tk.Frame(f, background="#ffffff")
                btn_f.pack(side='right', padx=5)
                
                if k in click_ordered_keys:
                    o_idx = click_ordered_keys.index(k)
                    if o_idx > 0:
                        ttk.Button(btn_f, text="▲", width=2, command=lambda _k=k: _move_in_order(_k, -1)).pack(side='left')
                    if o_idx < len(click_ordered_keys) - 1:
                        ttk.Button(btn_f, text="▼", width=2, command=lambda _k=k: _move_in_order(_k, 1)).pack(side='left')

        def _move_in_order(key, delta):
            if key in click_ordered_keys:
                idx = click_ordered_keys.index(key)
                target = idx + delta
                if 0 <= target < len(click_ordered_keys):
                    click_ordered_keys[idx], click_ordered_keys[target] = click_ordered_keys[target], click_ordered_keys[idx]
                    _refresh_list()

        _refresh_list()

        # 도구 버튼 영역 (전체 선택/전체 해제/컬럼 추가/컬럼 삭제)
        tool_btn_frame = ttk.Frame(main_frame)
        tool_btn_frame.pack(fill='x', pady=5)

        def _select_all():
            for k in ordered_keys_for_ui:
                if k not in vars_dict:
                    vars_dict[k] = tk.BooleanVar(value=True)
                else:
                    vars_dict[k].set(True)
                if k not in click_ordered_keys:
                    click_ordered_keys.append(k)
            _refresh_list()

        def _clear_all():
            for k in ordered_keys_for_ui:
                if k in vars_dict:
                    vars_dict[k].set(False)
            click_ordered_keys.clear()
            _refresh_list()

        ttk.Button(tool_btn_frame, text="모두 선택", command=_select_all).pack(side='left', padx=2, expand=True, fill='x')
        ttk.Button(tool_btn_frame, text="모두 해제", command=_clear_all).pack(side='left', padx=2, expand=True, fill='x')

        # 컬럼 관리 버튼 영역 (추가/삭제)
        mgr_btn_frame = ttk.Frame(main_frame)
        mgr_btn_frame.pack(fill='x', pady=5)
        
        from tkinter import simpledialog

        def _add_col():
            name = simpledialog.askstring("새 컬럼 추가", "추가할 컬럼 이름을 입력하세요:", parent=dialog)
            if name:
                name = name.strip()
                if not name: return
                if any(k.lower() == name.lower() for k in all_keys):
                    messagebox.showerror("오류", "이미 존재하는 컬럼 이름입니다.", parent=dialog)
                    return
                # 데이터 구조에 반영
                if data:
                    for item in data:
                        if name not in item: item[name] = ""
                # 현재 키 리스트에도 즉시 반영 (데이터가 없을 때를 위해)
                curr = getattr(self, keys_attr)
                if name not in curr: curr.append(name)

                self.log(f"➕ {mode} 새 컬럼 추가됨: {name}")
                dialog.destroy()
                self.manage_columns(mode)

        def _del_col():
            name = simpledialog.askstring("컬럼 삭제", "삭제할 컬럼 이름을 정확히 입력하세요:", parent=dialog)
            if name:
                name = name.strip()
                if not name: return
                essential = ["V", "No", "selected", "ISO", "Joint", "ISO/DWG", "Dwg"]
                if any(k.lower() == name.lower() for k in essential):
                    messagebox.showerror("오류", f"'{name}' 컬럼은 필수 항목이므로 삭제할 수 없습니다.", parent=dialog)
                    return
                
                target_key = next((k for k in display_keys if k.lower() == name.lower()), None)
                if not target_key:
                    messagebox.showerror("오류", f"'{name}' 컬럼을 찾을 수 없습니다.", parent=dialog)
                    return
                
                if messagebox.askyesno("삭제 확인", f"'{target_key}' 컬럼을 데이터에서 완전히 삭제하시겠습니까?", parent=dialog):
                    if data:
                        for item in data:
                            if target_key in item: del item[target_key]
                    curr = getattr(self, keys_attr)
                    if target_key in curr: curr.remove(target_key)

                    self.log(f"🗑️ {mode} 컬럼 삭제됨: {target_key}")
                    dialog.destroy()
                    self.manage_columns(mode)

        ttk.Button(mgr_btn_frame, text="컬럼 추가", command=_add_col).pack(side='left', padx=2, expand=True, fill='x')
        ttk.Button(mgr_btn_frame, text="컬럼 삭제", command=_del_col).pack(side='left', padx=2, expand=True, fill='x')
            
        def _apply():
            selected_keys = []
            
            if "_status" in vars_dict and vars_dict["_status"].get():
                selected_keys.append("_status")
            elif "ST" in vars_dict and vars_dict["ST"].get():
                selected_keys.append("ST")
                
            selected_keys.append("selected") # "V" column is required
            
            if "No" in vars_dict and vars_dict["No"].get():
                selected_keys.append("No")
            elif "No" not in selected_keys:
                selected_keys.append("No") # Fallback safety

            # [FIX] 사용자가 클릭한 순서(click_ordered_keys)를 기준으로 컬럼 수집
            for k in click_ordered_keys:
                if k not in ["ST", "_status", "No", "selected"] and vars_dict[k].get():
                    selected_keys.append(k)
            
            # click_ordered_keys에 없지만 체크되어 있는 예외적인 항목 처리 (안전장치)
            for k, v in vars_dict.items():
                if k not in ["ST", "_status", "No", "selected"] and v.get() and k not in selected_keys:
                    selected_keys.append(k)
            
            setattr(self, keys_attr, selected_keys)
            
            # [FIX] RT/KOGAS 모드는 전용 UI 갱신 함수 호출로 처리 (ID 정합성 보장)
            if mode in ["RT", "KOGAS"]:
                self._update_rt_preview_columns(mode=mode)
                self._refresh_rt_config_tabs() # [NEW] 세부 설정 탭도 함께 새로고침
            else:
                # 타 모드(PMI 등)용 기존 로직
                header_mapping = {}
                current_cols = []
                old_display_cols = list(tree["displaycolumns"])
                if old_display_cols == ["#all"] or not old_display_cols:
                    old_display_cols = list(tree["columns"])
                
                for k in selected_keys:
                    col_id = k
                    if k == "selected": col_id = "V"
                    elif k == "_status": col_id = "ST"
                    elif mode == "PMI":
                        if k == "Dwg": col_id = "Dwg"
                        elif k == "Joint": col_id = "Joint"
                    
                    current_cols.append(col_id)
                    header_mapping[col_id] = label_map.get(k, col_id)

                tree["columns"] = tuple(current_cols)
                tree["displaycolumns"] = tuple(current_cols)
                
                for col in tree["columns"]:
                    w = 80
                    tree.heading(col, text=header_mapping.get(col, col), anchor='center', command=lambda _c=col: self.sort_by_column(_c, mode=mode))
                    tree.column(col, width=w, anchor='center', stretch=False)
            
            self.populate_preview(data if data else [], switch_tab=False, mode=mode)
            self.log(f"⚙️ {mode} 미리보기 컬럼 설정 변경 완료")
            dialog.destroy()
            
        ttk.Button(main_frame, text="적용 (Apply)", command=_apply).pack(pady=5)




    def save_preview_data(self, mode="PMI"):
        """현재 추출/편집된 데이터를 JSON 파일로 저장"""
        if mode == "RT": data = self.rt_extracted_data
        elif mode == "KOGAS": data = self.kogas_extracted_data
        elif mode == "PT": data = self.pt_extracted_data
        elif mode == "MT": data = self.mt_extracted_data
        elif mode == "PAUT": data = self.paut_extracted_data
        else: data = self.extracted_data
        if not data:
            messagebox.showwarning("알림", f"{mode} 저장할 데이터가 없습니다.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            initialfile=f"{mode}_Data_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            title=f"{mode} 미리보기 데이터 저장"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                self.log(f"💾 {mode} 데이터 저장 완료: {os.path.basename(file_path)}")
                messagebox.showinfo("완료", "데이터가 성공적으로 저장되었습니다.")
            except Exception as e:
                self.log(f"❌ {mode} 저장 오류: {e}")
                messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {e}")



    def load_preview_data(self, mode="PMI"):
        """저장된 JSON 데이터를 불러와서 목록에 채움"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json")],
            title=f"{mode} 미리보기 데이터 불러오기"
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                    
                if not isinstance(loaded_data, list):
                    raise ValueError("올바른 데이터 형식이 아닙니다.")
                
                if mode == "RT": data = self.rt_extracted_data
                elif mode == "KOGAS": data = self.kogas_extracted_data
                elif mode == "PT": data = self.pt_extracted_data
                elif mode == "MT": data = self.mt_extracted_data
                elif mode == "PAUT": data = self.paut_extracted_data
                else: data = self.extracted_data
                
                if data and not messagebox.askyesno("확인", f"현재 작업 중인 {mode} 데이터가 있습니다. 불러온 데이터로 덮어쓰시겠습니까?\n(아니오를 선택하면 기존 데이터에 추가됩니다.)"):
                    data.extend(loaded_data)
                else:
                    if mode == "RT": self.rt_extracted_data = loaded_data
                    elif mode == "KOGAS": self.kogas_extracted_data = loaded_data
                    elif mode == "PT": self.pt_extracted_data = loaded_data
                    elif mode == "MT": self.mt_extracted_data = loaded_data
                    elif mode == "PAUT": self.paut_extracted_data = loaded_data
                    else: self.extracted_data = loaded_data
                    data = loaded_data
            
                # [FIX] 필수 속성 보정
                for i, item in enumerate(data):
                    if 'date_filtered' not in item: item['date_filtered'] = True
                    if 'selected' not in item: item['selected'] = True
                    if 'order_index' not in item: item['order_index'] = i
                    
                self.update_date_listbox(mode)
                getattr(self, '_last_auto_material_sig', {}).pop(mode, None)
                self.populate_preview(data, mode=mode)
                self.log(f"📂 {mode} 데이터 불러오기 완료: {os.path.basename(file_path)}")
                messagebox.showinfo("완료", f"데이터를 성공적으로 불러왔습니다.\n(총 {len(data)} 건)")
            except Exception as e:
                self.log(f"❌ {mode} 불러오기 오류: {e}")
                messagebox.showerror("오류", f"파일을 불러오는 중 오류가 발생했습니다: {e}")



    def sort_by_column(self, col, mode="PMI"):
        """특정 컬럼을 기준으로 데이터 전체 정렬 (문자열/숫자 구분 및 헤더 표시)"""
        # 0. 데이터 대상 선정
        data_targets = {
            "PMI": self.extracted_data,
            "RT": self.rt_extracted_data,
            "KOGAS": self.kogas_extracted_data,
            "PT": self.pt_extracted_data,
            "MT": getattr(self, 'mt_extracted_data', []),
            "PAUT": self.paut_extracted_data
        }
        target_list = data_targets.get(mode)
        if not target_list: return
        
        # 1. 정렬 기준 컬럼 매핑 (모든 모드 통합)
        key_map = {
            "selected": "selected", "V": "selected", "No": "No", "Date": "Date", 
            "Dwg": "Dwg", "Joint": "Joint", "Loc": "Loc", "Ni": "Ni", "Cr": "Cr", "Mo": "Mo", "Grade": "Grade",
            "NPS": "NPS", "Thk.": "Thk.", "Material": "Material", "WType": "WType", "Result": "Result",
            "ISO/DWG": "ISO", "t": "t", "h": "h", "l": "l", "d": "d", "Location": "Location", "Nature": "Nature",
            "ST": "_status", "ISO Drawing No.": "Dwg", "Joint No": "Joint", "Test Location": "Loc", "Weld Type": "WType",
            "Drawing No.": "Dwg", "Joint No.": "Joint"
        }
        data_key = key_map.get(col, col)
        # PAUT 데이터는 공통 검사 데이터와 달리 화면 컬럼명을 그대로 키로 사용한다.
        # 따라서 Joint No.를 공통 키인 "Joint"로 매핑하면 모든 값이 빈 문자열로
        # 평가되어 헤더를 클릭해도 실제 정렬이 일어나지 않는다.
        if mode == "PAUT":
            paut_key_map = {
                "ISO/DWG": "Line No.",
                "ISO Drawing No.": "Line No.",
                "Drawing No.": "Line No.",
                "Joint": "Joint No.",
                "Joint No": "Joint No.",
                "Joint No.": "Joint No.",
            }
            data_key = paut_key_map.get(col, data_key)
        if not data_key: return

        # 2. 정렬 방향 결정 (기존 방향과 같으면 토글)
        attr_col = f"{mode.lower()}_sort_col"
        attr_rev = f"{mode.lower()}_sort_rev"
        
        if getattr(self, attr_col, None) == col:
            setattr(self, attr_rev, not getattr(self, attr_rev, False))
        else:
            setattr(self, attr_col, col)
            setattr(self, attr_rev, False)

        sort_rev = getattr(self, attr_rev, False)

        # 3. 데이터 정렬 (지능형 복합 자연어 정렬)
        def get_natural_key(val):
            if val is None: return [(2, "")]
            s_val = str(val).strip().upper()
            if not s_val: return [(2, "")]
            # 숫자와 문자열 세그먼트를 분리하여 (0, 숫자) 또는 (1, 문자열) 튜플 리스트 반환
            return [(0, int(c)) if c.isdigit() else (1, c) 
                    for c in re.split(r'(\d+)', s_val) if c]

        def get_value(item, key):
            if key == "_status" and mode == "PMI":
                is_def = False
                if hasattr(self, 'element_filters'):
                    for f in self.element_filters:
                        fk = f['key'].get().strip()
                        if not fk: continue
                        fm = self.to_float(f['min'].get())
                        fx = self.to_float(f['max'].get())
                        v = self.to_float(self._get_val_ci(item, fk))
                        if (fm > 0 and v < fm) or (fx > 0 and v > fx):
                            is_def = True; break
                return 0 if is_def else 1
            # '_current' 가 넘어오면 클릭된 컬럼(data_key) 값을 반환
            return item.get(data_key if key=="_current" else key, "")

        # 4. 계층적 정렬 수행 (클릭컬럼 -> Dwg -> Joint 순서로 Tie-break)
        try:
            k_dwg = "Line No." if mode == "PAUT" else "Dwg"
            k_joint = "Joint No." if mode == "PAUT" else "Joint"
            
            sort_key_func = lambda x: (
                get_natural_key(get_value(x, "_current")),
                get_natural_key(get_value(x, k_dwg)),
                get_natural_key(get_value(x, k_joint)),
                x.get('order_index', 0)
            )

            target_list.sort(key=sort_key_func, reverse=sort_rev)
        except Exception as e:
            self.log(f"⚠️ {mode} 정렬 오류: {e}")

        # 5. 헤더 텍스트 갱신 (화살표)
        tree_map = {
            "PMI": self.preview_tree, "RT": self.rt_preview_tree,
            "KOGAS": self.kogas_preview_tree,
            "PT": self.pt_preview_tree, "MT": getattr(self, 'mt_preview_tree', None), "PAUT": self.paut_preview_tree
        }
        target_tree = tree_map.get(mode)
        if target_tree:
            for c in target_tree["columns"]:
                current_text = target_tree.heading(c).get("text", c)
                clean_text = current_text.replace("▲ ", "").replace("▼ ", "")
                prefix = ""
                if c == col:
                    prefix = "▼ " if sort_rev else "▲ "
                target_tree.heading(c, text=prefix + clean_text)

        # 6. 화면 갱신
        self.populate_preview(target_list, switch_tab=False, mode=mode)
        self.log(f"📊 {mode}: '{col}' 기준 {'내림차순' if sort_rev else '오름차순'} 정렬 완료")




    def populate_preview(self, data_list, switch_tab=True, mode="PMI"):
        """추출된 데이터를 미리보기 표에 채움 (필터 반영 및 그룹 색상 적용)"""
        if mode == "PAUT":
            from site_apps.central.src.reports.paut_fields import normalize_tcg
            for item in data_list:
                normalize_tcg(item)
        if mode in ('PAUT', 'MT', 'RT', 'KOGAS'):
            from site_apps.central.src.reports.report_issue_history import restore_report_numbers
            try:
                restore_report_numbers(mode, data_list)
            except (OSError, ValueError) as exc:
                self.log(f'보고서번호 복원 실패: {exc}')
        if mode in ('PAUT', 'MT', 'PT', 'RT', 'KOGAS') and data_list and hasattr(self, '_sync_gapji_material_from_items'):
            self._sync_gapji_material_from_items(mode, data_list)
        if mode == "RT":
            tree = self.rt_preview_tree
            self.rt_item_idx_map = []
            idx_map = self.rt_item_idx_map
        elif mode == "PT":
            tree = self.pt_preview_tree
            self.pt_item_idx_map = []
            idx_map = self.pt_item_idx_map
        elif mode == "MT":
            tree = self.mt_preview_tree
            self.mt_item_idx_map = []
            idx_map = self.mt_item_idx_map
        elif mode == "PAUT":
            tree = self.paut_preview_tree
            self.paut_item_idx_map = []
            idx_map = self.paut_item_idx_map
        elif mode == "KOGAS":
            tree = self.kogas_preview_tree
            self.kogas_item_idx_map = []
            idx_map = self.kogas_item_idx_map
        else:
            tree = self.preview_tree
            self.item_idx_map = []
            idx_map = self.item_idx_map
            # [NEW] Update Loc List
            self.update_pmi_loc_listbox()
            
        # [NEW] Deficiency Detection & Auto-Reset (PMI Only)
            # [REMOVED global check, moved to local loop below]

        for item in tree.get_children():
            tree.delete(item)
        
        if mode == 'MT':
            self._refresh_mt_filter_headers()
        elif mode in ('RT', 'KOGAS'):
            self._refresh_rt_filter_headers(mode)
        filter_enabled = self.show_selected_only.get()
        hidden_by_def_count = 0
        total_matched_search = 0
        local_pmi_deficient_count = 0 # Count within Date/Loc filter
        
        active_locs = set()
        if mode == "PMI" and hasattr(self, 'pmi_loc_listbox'):
            for i in range(self.pmi_loc_listbox.size()):
                val = self.pmi_loc_listbox.get(i)
                if val.startswith("[v] "): active_locs.add(val.replace("[v] ", "").lower())
        
        last_iso = None
        last_joint = None
        current_tag = "group_even"
        
        for idx, item in enumerate(data_list):
            if not item.get('date_filtered', True):
                continue
            if mode == "PAUT" and not (
                    item.get('line_filtered', True)
                    and item.get('size_filtered', True)):
                continue
            
            if mode in ('MT', 'RT', 'KOGAS') and not item.get('drawing_filtered', True):
                continue
            is_selected = item.get('selected', True)
            if filter_enabled and not is_selected:
                continue
            
            # [NEW] PMI Preview Filters (Multi-Select Test Location & Deficiency)
            selected_locs_names = []
            if mode == "PMI":
                # item Date/Loc filter already passed (Date was line 2608, Loc is here)
                loc_val = str(item.get('Loc', '')).lower()
                if active_locs and loc_val not in active_locs:
                    continue
                
                # Check for deficiency highlighting (AND local count)
                is_deficient = False
                if hasattr(self, 'element_filters'):
                    for f_item in self.element_filters:
                        f_key = f_item['key'].get().strip()
                        if not f_key: continue
                        f_min = round(self.to_float(f_item['min'].get()), 2)
                        f_max = round(self.to_float(f_item['max'].get()), 2)
                        val = round(self.to_float(self._get_val_ci(item, f_key)), 2)
                        if (f_min > 0 and val < f_min) or (f_max > 0 and val > f_max):
                            is_deficient = True; break
                
                if is_deficient: local_pmi_deficient_count += 1
                total_matched_search += 1
                selected_locs_names.append(item.get('Loc', '')) 
                
                if self.pmi_show_deficiency_only.get() and not is_deficient:
                    hidden_by_def_count += 1
                    continue
            else:
                is_deficient = False

            idx_map.append(idx)
            v_mark = "●" if is_selected else "○"
            st_mark = "⚠️" if is_deficient else "✅"
            
            # ISO/DWG 번호가 바뀌면 배경색 태그 교체 (PAUT 'ISO' 대응)
            curr_iso = item.get('Dwg', item.get('ISO', ''))
            norm_iso = self.normalize_iso(curr_iso)
            
            # [NEW] 시각적 병합 (visual_group_joint) 값이 존재하면 해당 값으로 그룹화 계산
            curr_joint = item.get('visual_group_joint', item.get('Joint', ''))
            
            # [FIX] Joint-Aware: Joint가 바뀌었는지도 함께 체크
            is_new_iso = (last_iso is None or self.normalize_iso(last_iso) != norm_iso)
            is_new_joint = (last_joint is None or last_joint != curr_joint)
            
            if is_new_iso:
                current_tag = "group_odd" if current_tag == "group_even" else "group_even"

            # [REFINED] 하이브리드 표시 로직: 블록 시작점 정렬이거나 "새로운 그룹(ISO or Joint)" 시작점일 때 표시
            display_count = len(idx_map) - 1 
            is_block_start = (display_count % 3 == 0)
            
            # ISO나 Joint가 바뀌었거나 3줄 단위 시작이면 표시
            is_show = is_new_iso or is_new_joint or is_block_start

            display_iso = curr_iso if is_show else ""
            display_joint = item.get('Joint', '') if is_show else ""
            
            # [FIX] 명시적으로 병합된 경우라도 '새 그룹 시작'이면 표시 강제
            if item.get('is_merged_iso') and not (is_new_iso or is_new_joint): display_iso = ""
            if item.get('is_merged_joint') and not (is_new_iso or is_new_joint): display_joint = ""

            last_iso = curr_iso
            last_joint = curr_joint
            
            row_vals = []
            if mode == "RT" or mode == "KOGAS":
                keys = self.kogas_column_keys if mode == "KOGAS" else self.rt_column_keys
                for k in keys:
                    if k == "selected": row_vals.append(v_mark)
                    elif k == "Acc":
                        val = item.get("Acc", "")
                        if not val:
                            res = str(item.get("Result", "")).upper()
                            val = "OK" if "ACC" in res or "OK" in res else ""
                        row_vals.append(val)
                    elif k == "Rej":
                        val = item.get("Rej", "")
                        if not val:
                            res = str(item.get("Result", "")).upper()
                            val = "NG" if "REJ" in res or "NG" in res or "RE" in res else ""
                        row_vals.append(val)
                    elif k == "Deg":
                        row_vals.append(str(item.get("Deg", "")).strip())
                    elif mode == "KOGAS" and k in ["Dwg", "Mat", "Welder"]:
                        main_val = str(item.get(k, "")).strip()
                        sub_val = str(item.get(f"{k}_Sub", "")).strip()
                        if sub_val and sub_val != main_val:
                            row_vals.append(f"{main_val} / {sub_val}")
                        else:
                            row_vals.append(main_val)
                    else: 
                        val = str(item.get(k, "")).strip()
                        if (k in ["Dwg", "ISO"] and item.get('is_merged_iso')) or (k == "Joint" and item.get('is_merged_joint')): val = ""
                        row_vals.append(val)
            elif mode == "PT":
                for k in self.pt_column_keys:
                    if k == "selected": row_vals.append(v_mark)
                    else: 
                        val = str(item.get(k, "")).strip()
                        row_vals.append(val)
            elif mode == "MT":
                for k in self.mt_column_keys:
                    if k == "selected": row_vals.append(v_mark)
                    else:
                        val = str(item.get(k, "")).strip()
                        row_vals.append(val)
            elif mode == "PAUT":
                # PAUT use paut_column_keys
                for k in self.paut_column_keys:
                    if k == "selected": row_vals.append(v_mark)
                    elif k == "No": row_vals.append(len(idx_map))
                    else: 
                        val = str(item.get(k, "")).strip()
                        if (k in ["Dwg", "ISO"] and item.get('is_merged_iso')) or (k == "Joint" and item.get('is_merged_joint')): val = ""
                        row_vals.append(val)
            else: # PMI
                for k in self.column_keys:
                    if k == "_status": row_vals.append(st_mark)
                    elif k == "selected": row_vals.append(v_mark)
                    elif k in ["Ni", "Cr", "Mo", "Mn"]:
                        row_vals.append(f"{self.to_float(item.get(k, 0)):.2f}")
                    else:
                        val = str(item.get(k, "")).strip()
                        if (k in ["Dwg", "ISO"] and item.get('is_merged_iso')) or (k == "Joint" and item.get('is_merged_joint')): val = ""
                        row_vals.append(val)

            row_tags = [str(idx), current_tag]
            if is_deficient:
                row_tags.append("deficient")
            tree.insert("", "end", values=tuple(row_vals), tags=tuple(row_tags))
            
        # [REACTIVE AUTO-RESET] Trigger if current view is clean but filters are active
        if mode == "PMI":
            has_pmi_filter = self.pmi_show_deficiency_only.get()
            if not has_pmi_filter and hasattr(self, 'element_filters'):
                for f in self.element_filters:
                    if f['min'].get().strip() or f['max'].get().strip():
                        has_pmi_filter = True; break
            
            if local_pmi_deficient_count == 0 and has_pmi_filter:
                self.reset_pmi_filters(refresh=True, clear_search=False)
                return

        if mode == "PMI" and hidden_by_def_count > 0:
            sel_str = ", ".join(list(set(selected_locs_names))) if selected_locs_names else "전체"
            self.log(f"ℹ️ '{sel_str}' 결과 중 {hidden_by_def_count}건이 '함량 미달만 보기' 옵션으로 인해 숨겨져 있습니다.")

            
        # [REMOVED] Legacy tab switching logic (Preview is now permanently visible in splitscreen)
        # if switch_tab:
        #     if mode == "RT":
        #         self.rt_tab_notebook.select(self.rt_tab_preview)
        #     elif mode == "PT":
        #         self.pt_tab_notebook.select(self.pt_tab_preview)
        #     elif mode == "PAUT":
        #         pass 
        #     else:
        #         self.tab_notebook.select(self.tab_preview)


    def show_bulk_update_dialog(self, mode="PMI"):
        """선택된 항목들을 필터 조건에 따라 일괄 변경하는 다이얼로그 표시"""
        target_data = self._get_mode_info(mode)[2]
        
        if not target_data:
            messagebox.showwarning("알림", f"{mode} 데이터가 없습니다.")
            return

        # 체크(●)된 항목의 인덱스 추출
        selected_indices = [idx for idx, item in enumerate(target_data) if item.get('selected', True)]
        if not selected_indices:
            messagebox.showwarning("항목 미선택", "일괄 변경할 항목을 체크(선택)해주세요.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"{mode} 선택 항목 일괄 변경")
        dialog.geometry("450x570")
        dialog.configure(background="#f9fafb")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry(f"+{self.root.winfo_x() + 100}+{self.root.winfo_y() + 50}")

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="1. 대상 필터링 (선택 사항)", font=("Malgun Gothic", 10, "bold")).pack(anchor='w', pady=(0, 10))
        
        filter_frame = ttk.Frame(main_frame)
        filter_frame.pack(fill='x', pady=(0, 10))
        
        # 모드별 필터 가능 열 설정
        if mode == "RT":
            filter_cols = ["전체(필터 없음)", "Drawing No.", "Film Ident. No.", "Film Location", "Welder No", "Remarks"]
            key_map = {"Drawing No.": "Dwg", "Film Ident. No.": "Joint", "Film Location": "Loc", "Welder No": "Welder", "Remarks": "Remarks"}
        elif mode == "PT":
            filter_cols = ["전체(필터 없음)", "ISO Drawing No.", "Joint", "NPS", "Thk.", "Material", "Welder"]
            key_map = {"ISO Drawing No.": "Dwg", "Joint": "Joint", "NPS": "NPS", "Thk.": "Thk.", "Material": "Material", "Welder": "Welder"}
        elif mode == 'MT':
            filter_cols = ['전체(필터 없음)', 'Dwg', 'Joint', 'Material', 'TestItem', 'Result', 'Welder']
            key_map = {key: key for key in filter_cols[1:]}
        else: # PMI
            filter_cols = ["전체(필터 없음)", "ISO/DWG", "Joint No", "Test Location", "Ni", "Cr", "Mo", "Grade"]
            key_map = {"ISO/DWG": "Dwg", "Joint No": "Joint", "Test Location": "Loc", "Ni": "Ni", "Cr": "Cr", "Mo": "Mo", "Grade": "Grade"}

        ttk.Label(filter_frame, text="필터 열:").grid(row=0, column=0, sticky='w')
        col_var = tk.StringVar(value="전체(필터 없음)")
        col_combo = ttk.Combobox(filter_frame, textvariable=col_var, state='readonly', width=20, values=filter_cols)
        col_combo.grid(row=0, column=1, padx=5, sticky='w')
        
        ttk.Label(filter_frame, text="필터 값:").grid(row=1, column=0, sticky='w', pady=5)
        val_var = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=val_var, width=23).grid(row=1, column=1, padx=5, sticky='w', pady=5)
        
        # [NEW] Filter Integration Checkboxes
        only_def_var = tk.BooleanVar(value=False)
        only_visible_var = tk.BooleanVar(value=True) # Default to True for convenience
        
        filter_options_frame = ttk.Frame(main_frame)
        filter_options_frame.pack(fill='x', pady=(0, 10))
        
        if mode == "PMI":
            tk.Checkbutton(filter_options_frame, text="⚠️ 함량 미달인 항목만 변경 적용", variable=only_def_var, 
                           background="#f9fafb", font=("Malgun Gothic", 9)).pack(anchor='w')
        
        tk.Checkbutton(filter_options_frame, text="🔍 현재 화면에 보이는(필터링된) 항목만 변경", variable=only_visible_var, 
                       background="#f9fafb", font=("Malgun Gothic", 9)).pack(anchor='w')
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=5)

        ttk.Label(main_frame, text="2. 변경할 값 입력 (입력된 항목만 반영)", font=("Malgun Gothic", 10, "bold")).pack(anchor='w', pady=(0, 10))
        
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill='both', expand=True)
        
        # 모드별 입력 필드 구성
        if mode == "RT":
            labels = [("Dwg:", "Dwg"), ("Film Ident:", "Joint"), ("Loc:", "Loc"), ("Grade:", "Grade"), ("Welder:", "Welder"), ("Remarks:", "Remarks"), ("Date:", "Date")]
        elif mode == "PT":
            labels = [("Dwg:", "Dwg"), ("Joint:", "Joint"), ("NPS:", "NPS"), ("Thk:", "Thk."), ("Material:", "Material"), ("Welder:", "Welder"), ("Weld Type:", "WType"), ("Result:", "Result"), ("Date:", "Date")]
        elif mode == 'MT':
            labels = [(key + ':', key) for key in ('Dwg', 'Joint', 'Material', 'TestItem', 'Result', 'Welder', 'Remarks', 'Date')]
        else: # PMI
            labels = [("Ni (%):", "Ni"), ("Cr (%):", "Cr"), ("Mo (%):", "Mo"), ("Location:", "Loc"), ("Grade:", "Grade"), ("Date:", "Date")]

        entries = {}
        for i, (lbl, key) in enumerate(labels):
            ttk.Label(fields_frame, text=lbl).grid(row=i, column=0, sticky='e', pady=3, padx=5)
            entries[key] = ttk.Entry(fields_frame, width=20)
            entries[key].grid(row=i, column=1, sticky='w', pady=3)

        def apply_action():
            filter_col = col_var.get()
            filter_val = val_var.get().strip().lower()
            target_key = key_map.get(filter_col)
            
            update_spec = {}
            for key, entry in entries.items():
                val = entry.get().strip()
                if val:
                    if mode == "PMI" and key not in ["No", "Joint", "Loc", "Grade", "Date", "Dwg", "_status", "ST", "V", "selected", "order_index"]:
                        if val.startswith(('+', '-')):
                            try: update_spec[key] = ('relative', float(val))
                            except: update_spec[key] = ('absolute', self.to_float(val))
                        else:
                            update_spec[key] = ('absolute', self.to_float(val))
                    else:
                        update_spec[key] = ('absolute', val)
            
            if not update_spec:
                messagebox.showwarning("입력 부족", "변경할 값을 하나 이상 입력해주세요.")
                return

            count = 0
            for idx in selected_indices:
                item = target_data[idx]
                match = True
                if target_key:
                    curr_val = str(item.get(target_key, "")).strip().lower()
                    if filter_val not in curr_val: match = False
                
                # [NEW] Check for deficiency if requested
                if match and mode == "PMI" and only_def_var.get():
                    is_deficient = False
                    for f_item in self.element_filters:
                        f_key = f_item['key'].get().strip()
                        if not f_key: continue
                        f_min = round(self.to_float(f_item['min'].get()), 2)
                        f_max = round(self.to_float(f_item['max'].get()), 2)
                        val = round(self.to_float(self._get_val_ci(item, f_key)), 2)
                        if (f_min > 0 and val < f_min) or (f_max > 0 and val > f_max):
                            is_deficient = True; break
                    if not is_deficient: match = False
                
                # [NEW] Check for sidebar visibility (Date & Loc)
                if match and only_visible_var.get():
                    if mode in ('MT', 'RT', 'KOGAS') and not item.get('drawing_filtered', True):
                        match = False
                    if not item.get('date_filtered', True):
                        match = False
                    if match and mode == "PAUT" and not (
                            item.get('line_filtered', True)
                            and item.get('size_filtered', True)):
                        match = False
                    
                    if match and mode == "PMI" and hasattr(self, 'pmi_loc_listbox'):
                        # Get active sidebar locs
                        active_locs = set()
                        for i in range(self.pmi_loc_listbox.size()):
                            v = self.pmi_loc_listbox.get(i)
                            if v.startswith("[v] "): active_locs.add(v.replace("[v] ", "").lower())
                        
                        if active_locs:
                            loc_val = str(item.get('Loc', '')).lower()
                            if loc_val not in active_locs:
                                match = False

                if match:
                    for k, spec in update_spec.items():
                        up_mode, up_val = spec
                        if up_mode == 'relative':
                            item[k] = self.to_float(item.get(k, 0)) + up_val
                        else:
                            item[k] = up_val
                    
                    # [NEW] Re-calculate Grade ONLY if NOT manually provided
                    if "Grade" not in update_spec:
                        new_grade = self.check_material_grade(item)
                        if new_grade:
                            item['Grade'] = new_grade
                    
                    count += 1
            
            self.populate_preview(target_data, switch_tab=False, mode=mode)
            messagebox.showinfo("일괄 변경 완료", f"총 {count}개의 항목이 업데이트되었습니다.")
            dialog.destroy()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(20, 0))
        ttk.Button(btn_frame, text="적용 (Apply)", command=apply_action).pack(side='right', padx=5)
        ttk.Button(btn_frame, text="취소 (Cancel)", command=dialog.destroy).pack(side='right')


    def import_request_mapping(self, mode="PMI"):
        """외부 검사 요청서/도면 리스트 엑셀 파일을 읽어와 현재 테이블의 Drawing No, Joint No, Location을 자동 매핑/업데이트"""
        mode_info = self._get_mode_info(mode)
        if not mode_info: return
        tree, idx_map, target_data, _ = mode_info

        if not target_data:
            messagebox.showwarning("알림", f"{mode} 데이터가 없습니다. 먼저 PMI 데이터를 불러와주세요.")
            return

        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="검사 요청서 / 도면 리스트 엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx;*.xlsm;*.xls"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일을 열 수 없습니다:\n{e}")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("검사 요청서(Drawing / Joint / Location) 매핑")
        dialog.geometry("520x460")
        dialog.configure(background="#f9fafb")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry(f"+{self.root.winfo_x() + 80}+{self.root.winfo_y() + 60}")

        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)

        ttk.Label(main_frame, text="📋 검사 요청서 데이터 매핑", font=("Malgun Gothic", 12, "bold")).pack(anchor='w', pady=(0, 5))
        ttk.Label(main_frame, text=f"파일: {os.path.basename(file_path)}", font=("Malgun Gothic", 9), foreground="#6b7280").pack(anchor='w', pady=(0, 15))

        # 시트 선택
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.pack(fill='x', pady=5)
        ttk.Label(sheet_frame, text="시트 선택:", width=12).pack(side='left')
        sheet_var = tk.StringVar(value=wb.sheetnames[0])
        sheet_combo = ttk.Combobox(sheet_frame, textvariable=sheet_var, state='readonly', values=wb.sheetnames, width=28)
        sheet_combo.pack(side='left', padx=5)

        # 시작 행 선택
        start_row_frame = ttk.Frame(main_frame)
        start_row_frame.pack(fill='x', pady=5)
        ttk.Label(start_row_frame, text="데이터 시작 행:", width=12).pack(side='left')
        start_row_var = tk.IntVar(value=2)
        start_row_spin = ttk.Spinbox(start_row_frame, from_=1, to=100, textvariable=start_row_var, width=10)
        start_row_spin.pack(side='left', padx=5)
        ttk.Label(start_row_frame, text="(헤더 제목 다음 행)", foreground="#6b7280").pack(side='left')

        # 컬럼 선택 프레임
        cols_frame = ttk.LabelFrame(main_frame, text=" 컬럼 매핑 설정 ", padding=10)
        cols_frame.pack(fill='x', pady=10)

        col_options_var = {"list": []}

        def get_col_header_list(ws_cur, header_r=1):
            opts = ["(미사용 / 매핑 안함)"]
            max_c = min(ws_cur.max_column or 26, 30)
            for c in range(1, max_c + 1):
                col_letter = openpyxl.utils.get_column_letter(c)
                v = ws_cur.cell(row=header_r, column=c).value
                header_text = str(v).strip().replace('\n', ' ') if v is not None else ""
                label = f"{col_letter}열: {header_text}" if header_text else f"{col_letter}열"
                opts.append(label)
            return opts

        dwg_var = tk.StringVar()
        joint_var = tk.StringVar()
        loc_var = tk.StringVar()
        no_var = tk.StringVar()

        def update_sheet_columns(*args):
            s_name = sheet_var.get()
            if s_name not in wb.sheetnames: return
            ws_cur = wb[s_name]

            best_hr = 1
            best_cnt = 0
            for r in range(1, min(6, (ws_cur.max_row or 1) + 1)):
                cnt = sum(1 for c in range(1, min(ws_cur.max_column or 26, 26) + 1) if ws_cur.cell(r, c).value is not None)
                if cnt > best_cnt:
                    best_cnt = cnt
                    best_hr = r

            start_row_var.set(best_hr + 1)
            opts = get_col_header_list(ws_cur, header_r=best_hr)
            col_options_var["list"] = opts

            for cb in (dwg_combo, joint_combo, loc_combo, no_combo):
                cb['values'] = opts

            def match_col(keywords):
                for opt in opts:
                    opt_l = opt.lower()
                    for kw in keywords:
                        if kw in opt_l:
                            return opt
                return opts[0]

            dwg_var.set(match_col(['도면', 'drawing', 'dwg', 'iso', 'line', '라인']))
            joint_var.set(match_col(['joint', '조인트', 'j/n', 'weld', '용접']))
            loc_var.set(match_col(['location', 'loc', '위치', '부위', '지점', 'point']))
            no_var.set(match_col(['순번', '번호', 'no', 'seq']))

        for r_i, (lbl_txt, v_target) in enumerate([
            ("Drawing No. (도면):", dwg_var),
            ("Joint No. (조인트):", joint_var),
            ("Location (검사부위):", loc_var),
            ("No (순번/비교용):", no_var),
        ]):
            f_row = ttk.Frame(cols_frame)
            f_row.pack(fill='x', pady=3)
            ttk.Label(f_row, text=lbl_txt, width=18).pack(side='left')
            combo = ttk.Combobox(f_row, textvariable=v_target, state='readonly', width=30)
            combo.pack(side='left', padx=5)
            if r_i == 0: dwg_combo = combo
            elif r_i == 1: joint_combo = combo
            elif r_i == 2: loc_combo = combo
            elif r_i == 3: no_combo = combo

        sheet_combo.bind("<<ComboboxSelected>>", update_sheet_columns)
        update_sheet_columns()

        match_mode_var = tk.StringVar(value="ORDER")
        mode_frame = ttk.Frame(main_frame)
        mode_frame.pack(fill='x', pady=5)
        ttk.Label(mode_frame, text="매핑 방식:", width=12).pack(side='left')
        ttk.Radiobutton(mode_frame, text="행 순서대로 1:1 매핑 (추천)", variable=match_mode_var, value="ORDER").pack(side='left', padx=5)
        ttk.Radiobutton(mode_frame, text="No 번호 일치 매핑", variable=match_mode_var, value="NO").pack(side='left', padx=5)

        def apply_mapping():
            s_name = sheet_var.get()
            if s_name not in wb.sheetnames: return
            ws_cur = wb[s_name]

            def parse_col_letter(opt_str):
                if not opt_str or "(미사용" in opt_str: return None
                return opt_str.split("열")[0].strip()

            c_dwg = parse_col_letter(dwg_var.get())
            c_joint = parse_col_letter(joint_var.get())
            c_loc = parse_col_letter(loc_var.get())
            c_no = parse_col_letter(no_var.get())

            if not any([c_dwg, c_joint, c_loc]):
                messagebox.showwarning("알림", "최소 하나의 매핑 열(Drawing, Joint, Location 중)을 선택해주세요.")
                return

            col_dwg_idx = openpyxl.utils.column_index_from_string(c_dwg) if c_dwg else None
            col_joint_idx = openpyxl.utils.column_index_from_string(c_joint) if c_joint else None
            col_loc_idx = openpyxl.utils.column_index_from_string(c_loc) if c_loc else None
            col_no_idx = openpyxl.utils.column_index_from_string(c_no) if c_no else None

            start_r = max(1, start_row_var.get())
            max_r = ws_cur.max_row or start_r

            req_items = []
            for r in range(start_r, max_r + 1):
                d_val = ws_cur.cell(row=r, column=col_dwg_idx).value if col_dwg_idx else None
                j_val = ws_cur.cell(row=r, column=col_joint_idx).value if col_joint_idx else None
                l_val = ws_cur.cell(row=r, column=col_loc_idx).value if col_loc_idx else None
                n_val = ws_cur.cell(row=r, column=col_no_idx).value if col_no_idx else None

                if d_val is None and j_val is None and l_val is None and n_val is None:
                    continue

                req_items.append({
                    'Dwg': str(d_val).strip() if d_val is not None else '',
                    'Joint': str(j_val).strip() if j_val is not None else '',
                    'Loc': str(l_val).strip() if l_val is not None else '',
                    'No': str(n_val).strip() if n_val is not None else '',
                })

            if not req_items:
                messagebox.showwarning("알림", "선택한 엑셀 시트에서 유효한 데이터를 찾지 못했습니다.")
                return

            m_mode = match_mode_var.get()
            updated_count = 0

            if m_mode == "NO" and col_no_idx:
                req_map = {item['No']: item for item in req_items if item['No']}
                for item in target_data:
                    cur_no = str(item.get('No', '')).strip()
                    if cur_no in req_map:
                        matched = req_map[cur_no]
                        if col_dwg_idx and matched['Dwg']: item['Dwg'] = matched['Dwg']
                        if col_joint_idx and matched['Joint']: item['Joint'] = matched['Joint']
                        if col_loc_idx and matched['Loc']: item['Loc'] = matched['Loc']
                        updated_count += 1
            else:
                for i, req_row in enumerate(req_items):
                    if i >= len(target_data): break
                    item = target_data[i]
                    if col_dwg_idx and req_row['Dwg']: item['Dwg'] = req_row['Dwg']
                    if col_joint_idx and req_row['Joint']: item['Joint'] = req_row['Joint']
                    if col_loc_idx and req_row['Loc']: item['Loc'] = req_row['Loc']
                    updated_count += 1

            self.populate_preview(target_data, switch_tab=False, mode=mode)
            self.log(f"📥 {mode} 요청서 매핑 완료: 총 {updated_count}개 항목 업데이트됨 (파일: {os.path.basename(file_path)})")
            messagebox.showinfo("매핑 완료", f"총 {updated_count}개 항목의 도면/Joint/위치 정보가 요청서로부터 성공적으로 업데이트되었습니다.")
            dialog.destroy()

        btn_box = ttk.Frame(main_frame)
        btn_box.pack(fill='x', pady=(15, 0))
        ttk.Button(btn_box, text="매핑 적용 (Apply)", command=apply_mapping).pack(side='right', padx=5)
        ttk.Button(btn_box, text="취소 (Cancel)", command=dialog.destroy).pack(side='right')

    # --- Integrated Verification Logic ---

