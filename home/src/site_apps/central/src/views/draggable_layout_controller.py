# -*- coding: utf-8 -*-
import os
import re
import sys
import json
import math
import time
import shutil
import datetime
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
import numpy as np
from tkcalendar import DateEntry, Calendar
from site_apps.central.src.utils.helpers import NAN_PATTERN, DOT_ZERO_PATTERN, MARKER_PATTERN
from site_apps.central.src.views.components import *


def create_draggable_container_impl(self, parent, label_text, widget_class, config_key, manage_list_key=None, grid_info=None, **widget_kwargs):
    """Create a draggable container with a label and a widget, styled as a visible box"""
    box_border_color = (
        getattr(self, 'style', None).lookup('TLabelframe', 'bordercolor') if hasattr(self, 'style') else ''
    ) or (
        getattr(self, 'style', None).lookup('TFrame', 'bordercolor') if hasattr(self, 'style') else ''
    ) or '#c0c0c0'

    container = tk.Frame(
        parent,
        relief="solid",
        borderwidth=1,
        highlightthickness=1,
        highlightbackground=box_border_color,
        bg=self.theme_bg
    )

    # Header container for label and buttons
    hdr = ttk.Frame(container) 
    hdr.pack(side='top', fill='x', padx=0, pady=0)

    # Drag handle icon
    lbl_drag = ttk.Label(hdr, text="✥", font=('Arial', 9), cursor='fleur')
    lbl_drag.pack(side='left', padx=1)
    self.make_header_draggable(lbl_drag, container)

    # Label
    lbl = ttk.Label(hdr, text=label_text, font=('Malgun Gothic', 9, 'bold'))
    lbl.pack(side='left', padx=1)
    self.make_header_draggable(lbl, container)

    # Icons container on the right
    btn_box = ttk.Frame(hdr)
    btn_box.pack(side='right', padx=1)

    # Rename icon
    btn_rename = ttk.Label(btn_box, text="✏️", font=('Arial', 8), cursor='hand2')
    btn_rename.pack(side='left', padx=1)
    btn_rename.bind('<Button-1>', lambda e: self.rename_widget_label(config_key))

    # Clone icon
    btn_clone = ttk.Label(btn_box, text="📋", font=('Arial', 8), cursor='hand2')
    btn_clone.pack(side='left', padx=1)
    btn_clone.bind('<Button-1>', lambda e: self.clone_widget(config_key))

    # Delete icon (X)
    btn_del = ttk.Label(btn_box, text="❌", font=('Arial', 8), cursor='hand2')
    btn_del.pack(side='left', padx=1)
    btn_del.bind('<Button-1>', lambda e: self.remove_box(config_key))

    # Internal Content Area
    content_area = ttk.Frame(container, padding=(1, 0))
    content_area.pack(side='top', fill='both', expand=True)

    # Widget
    widget = widget_class(content_area, **widget_kwargs)
    widget.pack(side='left', fill='both', expand=True)

    # Manage List Icon (Gear) - if it's a list-based widget
    if manage_list_key:
        btn_manage = ttk.Label(btn_box, text="⚙️", font=('Arial', 8), cursor='hand2')
        btn_manage.pack(side='left', padx=1)
        # [FIX] Bind AFTER widget is defined
        btn_manage.bind('<Button-1>', lambda e, cb=widget: self.open_list_management_dialog(manage_list_key, target_cb=cb))

    # If the widget is a basic tk widget (Text, Canvas), set its background
    if hasattr(widget, 'config') and 'bg' in widget.keys():
        try: widget.config(bg=self.theme_bg)
        except: pass

    # NEW: Handle grid placement first if grid_info provided
    if grid_info:
        container.grid(**grid_info)

    # Track for layout reset/config
    container._config_key = config_key
    container._label_widget = lbl
    container._widget = widget
    container._widget_class = widget_class
    container._widget_kwargs = widget_kwargs
    container._manage_list_key = manage_list_key

    # Register and make draggable (this now captures the CORRECT grid info)
    self.draggable_items[config_key] = container
    self.make_draggable(container, config_key)

    return container, widget


def bulk_register_materials_to_stock_impl(self, material_list):
    """Register items from the preferred list to the main Material inventory if they don't exist or are inactive"""
    print(f"DEBUG: Starting bulk registration for {len(material_list)} items...")
    if not material_list:
        messagebox.showwarning("경고", "등록할 품목이 목록에 없습니다.")
        return

    added_count = 0
    reactivated_count = 0
    existing_indices = {} # Name.lower() -> index in materials_df
    max_id = 10000

    # Ensure materials_df is loaded
    if self.materials_df is None:
        self.materials_df = pd.DataFrame(columns=['MaterialID', '품목명', 'Active'])
    old_materials = self.materials_df.copy(deep=True)

    if not self.materials_df.empty:
        # Collect existing names and their positions
        # [STABILITY] Build index using (Name, Model) to allow parent items even if sub-items exist
        existing_keys = {} # (name_lower, model_lower) -> index
        for idx, mat in self.materials_df.iterrows():
            nm = str(mat.get('품목명', '')).strip().lower()
            md = str(mat.get('모델명', '')).strip().lower()
            if nm:
                existing_keys[(nm, md)] = idx

        try:
            # Calculate next MaterialID
            if 'MaterialID' in self.materials_df.columns:
                valid_ids = pd.to_numeric(self.materials_df['MaterialID'], errors='coerce').dropna()
                if not valid_ids.empty:
                    max_id = int(valid_ids.max())
        except Exception as e:
            print(f"DEBUG: Error calculating max_id: {e}")
            max_id = 10000

    new_rows = []
    for mat in material_list:
        mat_clean = str(mat).strip()
        if not mat_clean: continue

        mat_lower = mat_clean.lower()
        # Check for parent item (empty model)
        if (mat_lower, '') in existing_keys:
            idx = existing_keys[(mat_lower, '')]
            # Check if it's inactive
            curr_active = self.materials_df.at[idx, 'Active']
            if str(curr_active) == '0' or curr_active == 0:
                # Reactivate!
                self.materials_df.at[idx, 'Active'] = 1
                reactivated_count += 1
                print(f"DEBUG: Reactivated item: {mat_clean}")
            else:
                print(f"DEBUG: Skipping '{mat_clean}' (Already Active)")
        else:
            # Add as New
            max_id += 1
            new_row = {
                'MaterialID': max_id,
                '품목명': mat_clean,
                'SN': '',
                '모델명': '',
                '관리단위': '매',
                '수량': 0,
                '재고하한': 0,
                'Active': 1,
                '구분': '소모품'
            }
            for col in self.materials_df.columns:
                if col not in new_row:
                    new_row[col] = ''

            new_rows.append(new_row)
            added_count += 1
            print(f"DEBUG: Prepared new item: {mat_clean} (ID: {max_id})")

    if new_rows or reactivated_count > 0:
        try:
            if new_rows:
                new_df = pd.DataFrame(new_rows)
                new_df = new_df.reindex(columns=self.materials_df.columns, fill_value='')
                self.materials_df = pd.concat([self.materials_df, new_df], ignore_index=True)

            # Save and Refresh
            if not self.save_data(require_primary=True):
                raise RuntimeError('원본 재고 파일에 일괄 등록 내역을 저장하지 못했습니다.')
            self.update_stock_view()
            self.update_material_combo()

            msg = ""
            if added_count > 0: msg += f"{added_count}개의 품목 신규 등록 완료\n"
            if reactivated_count > 0: msg += f"{reactivated_count}개의 품목 재활성화(복구) 완료"

            messagebox.showinfo("등록 완료", msg)
        except Exception as e:
            self.materials_df = old_materials
            print(f"DEBUG: Error during bulk registration: {e}")
            messagebox.showerror("오류", f"일괄 등록 중 오류가 발생해 원본 상태를 복원했습니다:\n{e}")
    else:
        messagebox.showinfo("정보", "이미 모든 품목이 재고 현황에 활성화되어 있습니다.")


def refresh_ui_for_list_change_impl(self, config_key):
    """Update all related UI elements after a list (sites, users, etc) has changed"""
    # Dictionary mapping config keys to their current values
    list_map = {
        'sites': self.sites,
        'users': self.users,
        'warehouses': self.warehouses,
        'equipments': self.equipments,
        'aux_equipments': getattr(self, 'aux_equipments', []),
        'worktimes': self.worktimes,
        'vehicles': self.vehicles,
        'companies': getattr(self, 'companies', []),
        'materials': getattr(self, 'carestream_films', []),
        'daily_units': getattr(self, 'daily_units', []),
        'test_items': getattr(self, 'test_items', []),
        'applied_codes': getattr(self, 'applied_codes', []),
        'co_code': getattr(self, 'co_code_list', []),
        '회사코드': getattr(self, 'co_code_list', [])
    }

    if config_key not in list_map:
        return

    current_vals = list_map[config_key]
    # [NEW] Ensure uniqueness and strip whitespace before sorting
    unique_vals = sorted(list(set([str(v).strip() for v in current_vals if v])))
    sorted_vals = unique_vals

    # [FIX] Update the list in-place to ensure all UI dialogs holding references see the changes
    if isinstance(current_vals, list):
        current_vals[:] = sorted_vals

    # 1. Update standard widgets
    if config_key == 'sites':
        if hasattr(self, 'cb_daily_site'): self.cb_daily_site['values'] = sorted_vals
        if hasattr(self, 'cb_trans_site'): self.cb_trans_site['values'] = sorted_vals
        if hasattr(self, 'cb_daily_filter_site'):
            self.cb_daily_filter_site['values'] = ['전체'] + sorted_vals
        if hasattr(self, 'cb_budget_site'): self.cb_budget_site['values'] = sorted_vals
        if hasattr(self, 'cb_budget_view_site'): self.cb_budget_view_site['values'] = sorted_vals
    elif config_key == 'companies':
        # [NEW] Update all company comboboxes in daily usage
        if hasattr(self, 'cb_daily_company'): self.cb_daily_company['values'] = sorted_vals
        if hasattr(self, 'cb_daily_filter_company'):
            self.cb_daily_filter_company['values'] = ['전체'] + sorted_vals
    elif config_key == 'daily_units':
        print(f"DEBUG: Refreshing All Unit Comboboxes with: {sorted_vals}")
        # [FIX] daily_units is already updated in-place above via current_vals[:] = sorted_vals
        # 1. Re-merge to include DB units and update all UI automatically
        self.update_registration_combos()
        # update_registration_combos() handles cb_daily_unit and cb_unit automatically
    elif config_key == 'users':
        # Updated to match current attribute names if needed, 
        # but usually it's cb_daily_user for 1, and cb_daily_user{i} for 2-10
        for i in range(1, 11):
            attr = 'cb_daily_user' if i == 1 else f'cb_daily_user{i}'
            if hasattr(self, attr):
                widget = getattr(self, attr)
                if isinstance(widget, WorkerCompositeWidget):
                     widget.cb_name['values'] = [''] + sorted_vals
                elif hasattr(widget, 'configure'):
                    try: widget['values'] = [''] + sorted_vals
                    except: pass

        if hasattr(self, 'ent_user'): self.ent_user['values'] = [''] + sorted_vals
    elif config_key == 'worktimes':
        # Update all worktime fields (1-10)
        for i in range(1, 11):
            attr = f'ent_worktime{i}'
            if hasattr(self, attr):
                widget = getattr(self, attr)
                if hasattr(widget, 'configure'):
                    try: 
                        widget['values'] = [''] + sorted_vals
                    except: pass
            grp_attr = f'worker_group{i}'
            if hasattr(self, grp_attr):
                getattr(self, grp_attr).update_time_list(sorted_vals)
    elif config_key == 'equipments' or config_key == 'materials':
        # Both affect the unified material/equipment suggestion lists
        self.update_material_combo()
    elif config_key == 'test_items' or config_key == 'applied_codes':
        # These are handled by dynamic lambdas in _bind_combobox_word_suggest
        pass 
    elif config_key == 'warehouses':
        if hasattr(self, 'cb_warehouse'): self.cb_warehouse['values'] = sorted_vals
    elif config_key == 'vehicles':
        # 1. Update fixed panel
        if hasattr(self, 'fixed_vehicle_widget'):
            self.fixed_vehicle_widget.update_vehicle_list(sorted_vals)
        # 2. Update floating panels
        if hasattr(self, 'draggable_items'):
            for key, cont in self.draggable_items.items():
                w = getattr(cont, '_widget', None)
                if w and type(w).__name__ == 'VehicleInspectionWidget':
                    w.update_vehicle_list(sorted_vals)
        # (Legacy compatibility)
        if hasattr(self, 'vehicle_boxes'):
            for box in getattr(self, 'vehicle_boxes', []):
                box.update_vehicle_list(sorted_vals)
        if hasattr(self, 'vehicle_inspections'):
            for key, widget_instance in getattr(self, 'vehicle_inspections', {}).items():
                widget_instance.update_vehicle_list(sorted_vals)
    elif config_key == 'co_code' or config_key == '회사코드':
        # [NEW] Update company code combobox in daily usage
        if hasattr(self, 'cb_daily_co_code'): self.cb_daily_co_code['values'] = sorted_vals
        if hasattr(self, 'cb_co_code'): self.cb_co_code['values'] = sorted_vals
        if hasattr(self, 'ndt_company_entries'):
            for entries in self.ndt_company_entries:
                cb_ndt_co = entries.get('_company')
                if cb_ndt_co is not None:
                    try: cb_ndt_co['values'] = sorted_vals
                    except Exception: pass

    # Always save configuration after any list change
    self.save_tab_config()


    # 2. Update ALL draggable widgets (clones) that depend on this list
    for key, container in self.draggable_items.items():
        # Heuristic: if manage_list_key is missing but label suggests it's a worker/user/site
        m_key = getattr(container, '_manage_list_key', None)
        if not m_key:
            if hasattr(container, '_label_widget'):
                lbl_text = container._label_widget.cget('text').lower()
                if config_key == 'users' and any(x in lbl_text for x in ['작업자', '담당자', 'user', 'worker']):
                    m_key = 'users'
                    container._manage_list_key = 'users'
                elif config_key == 'sites' and any(x in lbl_text for x in ['현장', 'site']):
                    m_key = 'sites'
                    container._manage_list_key = 'sites'
                elif config_key == 'equipments' and any(x in lbl_text for x in ['장비', 'equip']):
                    m_key = 'equipments'
                    container._manage_list_key = 'equipments'
                elif config_key == 'vehicles' and any(x in lbl_text for x in ['차량', 'vehicle']):
                    m_key = 'vehicles'
                    container._manage_list_key = 'vehicles'

        if m_key == config_key:
            if hasattr(container, '_widget') and hasattr(container._widget, 'config'):
                try:
                    container._widget['values'] = sorted_vals
                except:
                    pass


def remove_box_impl(self, key):
    """Intelligently remove a box: hide standard ones, destroy custom ones"""
    panel = getattr(self, 'vehicle_panel', None)
    editor = getattr(self, 'vehicle_inspections', {}).get(key)
    if panel is not None and editor in panel.editors:
        panel.remove(editor)
        return
    is_dynamic = key.startswith('memo_') or key.startswith('clone_') or key.startswith('checklist_') or key.startswith('vehicle_inspection_')

    # Allow deletion of dynamic widgets even if layout is locked
    if self.layout_locked and not is_dynamic: return

    if is_dynamic:
        # Permanent deletion for custom/dynamic items
        self.destroy_custom_widget(key)
    else:
        # Hiding for standard items (can be restored via Reset All)
        widget = self.draggable_items.get(key)
        if widget:
            # We reuse the existing hide_widget logic
            self.hide_widget(None, widget=widget)


def destroy_custom_widget_impl(self, key):
    """Destroy and remove from config any dynamic widget (clone or memo)"""
    # Allow destruction even if layout_locked (since it's a content management action)

    widget = self.draggable_items.get(key)
    if widget:
        widget.destroy()
        if key in self.draggable_items:
            del self.draggable_items[key]
        if key in self.memos:
            del self.memos[key]
        if key in self.checklists:
            del self.checklists[key]
        if key in self.vehicle_inspections:
            widget_to_remove = self.vehicle_inspections[key]
            if widget_to_remove in self.vehicle_boxes:
                self.vehicle_boxes.remove(widget_to_remove)
            del self.vehicle_inspections[key]


        # Clean from in-memory config to prevent resurrection on next save_tab_config call
        if hasattr(self, 'tab_config') and 'draggable_geometries' in self.tab_config:
            if key in self.tab_config['draggable_geometries']:
                del self.tab_config['draggable_geometries'][key]

        # Clean from config immediately to prevent resurrection on next load
        try:
            import json
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                if 'draggable_geometries' in config and key in config['draggable_geometries']:
                    del config['draggable_geometries'][key]

                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
        except:
            pass

        self.save_tab_config()


def rename_widget_label_impl(self, key):
    """Show a simple dialog to rename a widget's label"""
    if self.layout_locked: return

    widget = self.draggable_items.get(key)
    if not widget: return

    current_text = ""
    if hasattr(widget, '_label_widget'):
        current_text = widget._label_widget.cget('text')
    elif key in self.memos:
        current_text = self.memos[key]['title_entry'].get()
    elif key in self.checklists:
        current_text = self.checklists[key]['title_entry'].get()

    new_name = simpledialog.askstring("이름 변경", "새 이름을 입력하세요:", initialvalue=current_text)
    if new_name is not None:
        if hasattr(widget, '_label_widget'):
            widget._label_widget.config(text=new_name)
        elif key in self.memos:
            self.memos[key]['title_entry'].delete(0, 'end')
            self.memos[key]['title_entry'].insert(0, new_name)
        elif key in self.checklists:
            self.checklists[key]['title_entry'].delete(0, 'end')
            self.checklists[key]['title_entry'].insert(0, new_name)
        self.save_tab_config()


def clone_widget_impl(self, key):
    """Create a clone of an existing widget as a new custom box"""
    if self.layout_locked: return

    orig = self.draggable_items.get(key)
    if not orig: return

    import time
    new_key = f"clone_{int(time.time() * 1000)}"

    label_text = ""
    if hasattr(orig, '_label_widget'):
        label_text = orig._label_widget.cget('text')

    if hasattr(orig, '_widget_class'):
        # It's a container created via create_draggable_container
        cont, w = self.create_draggable_container(
            self.entry_inner_frame, 
            label_text, 
            orig._widget_class, 
            new_key, 
            manage_list_key=getattr(orig, '_manage_list_key', None), # Pass manage_list_key
            **orig._widget_kwargs
        )

        # Copy value from original widget
        if hasattr(orig, '_widget'):
            try:
                if hasattr(orig._widget, 'get'):
                    current_val = str(orig._widget.get()) # Ensure string

                    # Try generic Entry-like setting (works for Entry and Combobox text area)
                    if hasattr(w, 'delete') and hasattr(w, 'insert'):
                        try:
                            w.delete(0, 'end')
                            w.insert(0, current_val)
                        except:
                            # Readonly comboboxes might fail delete/insert
                            pass

                    # Try specific set method (Combobox, Scale, etc)
                    if hasattr(w, 'set'):
                        w.set(current_val)
                elif hasattr(orig._widget, 'get_data') and hasattr(w, 'set_data'):
                    # For complex widgets like VehicleInspectionWidget
                    w.set_data(orig._widget.get_data())
            except Exception as e:
                print(f"Failed to copy value: {e}")


        cont.place(x=50, y=50) # Start position
        self.save_tab_config()

    elif key in self.memos:
        # It's a memo
        content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
        title = self.memos[key]['title_entry'].get()
        self.add_new_memo(initial_text=content, initial_title=title, key=new_key)
        self.save_tab_config()
    elif key in self.checklists:
        # It's a checklist
        self.duplicate_checklist(key)
        self.save_tab_config()


def _bind_recursive_impl(self, widget, target_container):
    """Recursively bind drag events to widget and its children"""
    from functools import partial

    # Bind events to this widget, targeting the container
    # Note: We use add=True to avoid overwriting existing bindings if possible, 
    # but for drag we usually want exclusive control or at least priority.
    # Here we just bind standard.

    # We need to capture the target_container in the callback
    widget.bind("<Button-3>", partial(self.on_drag_start, widget=target_container))
    widget.bind("<Shift-Button-3>", partial(self.on_resize_start, widget=target_container))
    widget.bind("<B3-Motion>", partial(self.on_mouse_motion, widget=target_container))
    widget.bind("<Double-Button-3>", partial(self.reset_widget_position, widget=target_container))
    widget.bind("<Control-Button-3>", partial(self.hide_widget, widget=target_container))
    widget.bind("<ButtonRelease-3>", partial(self.on_drag_stop, widget=target_container))

    # Recurse for children
    try:
        for child in widget.winfo_children():
            # Toplevel(DateEntry 달력 팝업 등) 내부는 절대 순회하지 않음
            # → winfo_children() 재귀만으로도 Toplevel이 화면에 나타날 수 있음
            if child.winfo_class() == 'Toplevel':
                continue
            self._bind_recursive(child, target_container)
    except:
        pass


def hide_widget_impl(self, event, widget=None):
    """Hide a widget (Ctrl + Right Click)"""
    if self.layout_locked:
        return "break"

    if widget is None:
        widget = event.widget

    # [STABILITY FIX] Prevent hiding core widgets
    if hasattr(widget, '_config_key') and widget._config_key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
        messagebox.showwarning("보호됨", f"'{widget._config_key}' 항목은 화면의 필수 구성 요소이므로 숨길 수 없습니다.")
        return "break"

    # Remove from place layout
    if widget.winfo_manager() == 'place':
        widget.place_forget()
    elif widget.winfo_manager() == 'grid':
        widget.grid_forget()

    # Remove placeholder if exists
    self._remove_placeholder(widget)

    # Mark as hidden in config
    if hasattr(widget, '_config_key') and widget._config_key:
        try:
            import json
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {}

            if 'draggable_geometries' not in config:
                config['draggable_geometries'] = {}

            if widget._config_key not in config['draggable_geometries']:
                config['draggable_geometries'][widget._config_key] = {}

            config['draggable_geometries'][widget._config_key]['hidden'] = True

            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save hide status: {e}")
    return "break"


def make_draggable_impl(self, widget, config_key=None):
    """Make a widget draggable and resizable with right mouse button"""
    widget._config_key = config_key
    # Save original grid info for reset/placeholder
    # We need to do this immediately while it's still in the grid
    widget._original_grid_info = widget.grid_info()

    # Recursively bind to ensure clicking anywhere works
    self._bind_recursive(widget, widget)


def make_header_draggable_impl(self, widget, target_container):
    """Make a specific widget (header/label) draggable with Left Mouse Button targeting a container"""
    from functools import partial
    widget.bind("<Button-1>", partial(self.on_drag_start, widget=target_container))
    widget.bind("<B1-Motion>", partial(self.on_mouse_motion, widget=target_container))
    widget.bind("<ButtonRelease-1>", partial(self.on_drag_stop, widget=target_container))


def reset_widget_position_impl(self, event, widget=None):
    """Reset widget to original grid position"""
    if self.layout_locked:
        return "break"

    if widget is None:
        widget = event.widget

    # Remove from place layout
    widget.place_forget()

    # Remove placeholder if exists
    self._remove_placeholder(widget)

    # Restore to grid
    if hasattr(widget, '_original_grid_info'):
        widget.grid(**widget._original_grid_info)

    # Reset size variables if any
    if hasattr(widget, '_start_width'): del widget._start_width
    if hasattr(widget, '_start_height'): del widget._start_height

    # Remove from config
    if hasattr(widget, '_config_key') and widget._config_key:
        try:
            import json
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            if 'draggable_geometries' in config and widget._config_key in config['draggable_geometries']:
                del config['draggable_geometries'][widget._config_key]

            # Backward compatibility: also check top level
            if widget._config_key in config:
                del config[widget._config_key]

            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)

        except Exception as e:
            print(f"Failed to reset config: {e}")


def reset_all_widgets_layout_impl(self):
    """Reset all widgets to their original grid slots, unhide them, and reset paned window sashes"""
    if messagebox.askyesno("초기화", "모든 항목의 위치와 크기를 초기화하고 숨겨진 항목을 다시 표시하시겠습니까?\n(창 분할 위치도 초기화됩니다.)"):
        # 1. Reset draggable widgets
        for key, widget in list(self.draggable_items.items()):
            # Explicitly unhide if it was hidden
            if widget.winfo_manager() == '':
                # For grid, we just grid it back
                if hasattr(widget, '_original_grid_info'):
                     widget.grid(**widget._original_grid_info)

            self.reset_widget_position(None, widget=widget)

        # 1.05 [NEW] Use custom defaults if they exist
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                custom_defaults = config.get('custom_default_geometries', {})
                if custom_defaults:
                    # Clear existing geometries to ensure fresh start
                    if 'draggable_geometries' not in config: config['draggable_geometries'] = {}

                    for key, geo in custom_defaults.items():
                        if key in self.draggable_items:
                            w = self.draggable_items[key]
                            w.place(x=geo['x'], y=geo['y'], width=geo['width'], height=geo['height'])
                            # Also update config so it persists as the new current layout
                            config['draggable_geometries'][key] = geo

                    # Save the updated config back to disk and memory
                    with open(self.config_path, 'w', encoding='utf-8') as f:
                        json.dump(config, f, ensure_ascii=False, indent=2)
                    self.tab_config = config
            except Exception as e:
                print(f"Error applying custom defaults: {e}")

        # 1.1 Restore parent propagation safely and reset size
        if hasattr(self, 'entry_inner_frame'):
            self.entry_inner_frame.pack_propagate(True)
            self.entry_inner_frame.grid_propagate(True)

            # Perform a layout update pass
            self.root.update_idletasks()

            # Refresh scrollregion if in canvas
            if hasattr(self, 'entry_canvas'):
                self.entry_canvas.configure(scrollregion=self.entry_canvas.bbox("all"))

        # 2. Reset sash positions (splitters)
        try:
            if hasattr(self, 'daily_usage_paned'):
                # Give more space to the entry form by default (500px)
                getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, 500) 
            if hasattr(self, 'daily_history_paned'):
                total_w = self.daily_history_paned.winfo_width()
                if total_w > 100:
                    self.daily_history_paned.sashpos(0, int(total_w * 0.7))
                else:
                    self.daily_history_paned.sashpos(0, 800)
        except:
            pass

        self.save_tab_config()
        messagebox.showinfo("완료", "레이아웃과 창 분할 위치가 초기화되었습니다.")


def save_current_layout_as_default_impl(self):
    """Save the current layout as the custom 'reset' default"""
    try:
        current_geos = {}
        for key, widget in self.draggable_items.items():
            # [FIX] Capture all managed widgets (grid or place) 
            # This ensures the snapshot is complete even if some items weren't moved.
            manager = widget.winfo_manager()
            if manager in ['place', 'grid']:
                current_geos[key] = {
                    'x': widget.winfo_x(),
                    'y': widget.winfo_y(),
                    'width': widget.winfo_width(),
                    'height': widget.winfo_height(),
                    'hidden': False
                }

        if not current_geos:
            messagebox.showwarning("실패", "수동으로 배치된(드래그된) 항목이 없어 기본값으로 저장할 수 없습니다.")
            return

        if os.path.exists(self.config_path):
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = {}

        config['custom_default_geometries'] = current_geos

        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

        messagebox.showinfo("완료", "현재 레이아웃이 '기본 배치'로 저장되었습니다.\n앞으로 [초기화] 시 이 배치로 되돌아갑니다.")
    except Exception as e:
        messagebox.showerror("오류", f"기본 배치 저장 중 오류 발생: {e}")


def toggle_layout_lock_impl(self):
    """Toggle layout locking state"""
    self.layout_locked = not self.layout_locked
    if hasattr(self, 'btn_lock_layout'):
        if self.layout_locked:
            self.btn_lock_layout.config(text="🔒 배치 고정됨")
            self.style.configure("Lock.TButton", foreground="black")
        else:
            self.btn_lock_layout.config(text="🔓 배치 수정 중")
            self.style.configure("Lock.TButton", foreground="red")

    # Save layout lock state to config immediately
    if not hasattr(self, 'tab_config'):
        self.tab_config = {}
    self.tab_config['layout_locked'] = self.layout_locked

    # [FIX] Unlocking is no longer destructive. 
    # It simply enables movement without resetting positions.
    if self.layout_locked:
        print("Layout locked - current positions preserved")
    else:
        print("Layout unlocked - movement enabled")

    # Force save the state immediately to file
    self.save_tab_config(force=True)

    print(f"Layout lock {'enabled' if self.layout_locked else 'disabled'}")


def add_new_memo_impl(self, initial_text="", initial_title="메모", key=None):
    """Create a new movable/editable/copyable memo box with editable title"""
    import time
    if key is None:
        key = f"memo_{int(time.time() * 1000)}"

    # Create container

    memo_container = ttk.LabelFrame(self.entry_inner_frame)

    # Controls frame (top of memo)
    ctrl_frame = ttk.Frame(memo_container)
    ctrl_frame.pack(fill='x', side='top')

    # Editable Title Entry
    title_entry = ttk.Entry(ctrl_frame, font=('Arial', 9, 'bold'), width=15)
    title_entry.pack(side='left', padx=2)
    title_entry.insert(0, initial_title)
    title_entry.bind('<FocusOut>', lambda e: self.save_tab_config())

    btn_copy = ttk.Button(ctrl_frame, text="📋", width=3, command=lambda: self.duplicate_memo(key))
    btn_copy.pack(side='right')

    btn_del = ttk.Button(ctrl_frame, text="❌", width=3, command=lambda: self.remove_box(key))
    btn_del.pack(side='right')

    # Text area
    text_area = tk.Text(memo_container, wrap='word', height=5, width=30, font=('Arial', 10), bg=self.theme_bg, highlightthickness=0)
    text_area.pack(fill='both', expand=True, padx=2, pady=2)
    text_area.insert('1.0', initial_text)

    # Bind text change to auto-save
    text_area.bind('<FocusOut>', lambda e: self.save_tab_config())

    # Make draggable (Right Click)
    self.make_draggable(memo_container, key)

    # Make header draggable (Left Click)
    self.make_header_draggable(ctrl_frame, memo_container)
    self.draggable_items[key] = memo_container
    self.memos[key] = {
        'container': memo_container, 
        'text_widget': text_area,
        'title_entry': title_entry
    }

    # Initial placement if not loaded from config (will be handled by load_tab_config if exists)
    if key not in getattr(self, '_loading_memos', []):
        memo_container.place(x=1450, y=50) # Default start position

    return memo_container


def duplicate_memo_impl(self, key):
    """Duplicate an existing memo with its title and content"""
    if self.layout_locked: return
    if key in self.memos:
        content = self.memos[key]['text_widget'].get('1.0', 'end-1c')
        title = self.memos[key]['title_entry'].get()
        self.add_new_memo(initial_text=content, initial_title=title)
        self.save_tab_config()


def add_new_checklist_impl(self, initial_data=None, initial_title="체크리스트", key=None):
    """Create a new movable/editable checklist box"""
    import time
    if key is None:
        key = f"checklist_{int(time.time() * 1000)}"

    # Create container
    check_container = ttk.LabelFrame(self.entry_inner_frame)

    # Controls frame (top of checklist)
    ctrl_frame = ttk.Frame(check_container)
    ctrl_frame.pack(fill='x', side='top')

    # Editable Title Entry
    title_entry = ttk.Entry(ctrl_frame, font=('Arial', 9, 'bold'), width=15)
    title_entry.pack(side='left', padx=2)
    title_entry.insert(0, initial_title)
    title_entry.bind('<FocusOut>', lambda e: self.save_tab_config())

    btn_copy = ttk.Button(ctrl_frame, text="📋", width=3, command=lambda: self.duplicate_checklist(key))
    btn_copy.pack(side='right')

    btn_del = ttk.Button(ctrl_frame, text="❌", width=3, command=lambda: self.remove_box(key))
    btn_del.pack(side='right')

    # Add Item Area
    add_frame = ttk.Frame(check_container)
    add_frame.pack(fill='x', padx=2, pady=2)

    new_item_var = tk.StringVar()
    entry_new = ttk.Entry(add_frame, textvariable=new_item_var, width=20)
    entry_new.pack(side='left', fill='x', expand=True)

    def add_item(event=None):
         text = new_item_var.get().strip()
         if text:
             self.add_checklist_item(item_frame, text, False, key)
             new_item_var.set("")
             self.save_tab_config()

    def _adjust_parent_height(self, parent, force=False):
        """Adjust parent frame height with performance check"""
        try:
            # Only update idletasks if forced or we're not in the middle of a high-speed interaction
            # This is the single biggest cause of UI stutter.
            if force:
                parent.update_idletasks()
        except:
            pass

    entry_new.bind('<Return>', add_item)
    btn_add = ttk.Button(add_frame, text="➕", width=3, command=add_item)
    btn_add.pack(side='right')

    # Scrollable Frame for Items
    canvas_frame = ttk.Frame(check_container)
    canvas_frame.pack(fill='both', expand=True, padx=2, pady=2)

    canvas = tk.Canvas(canvas_frame, height=100, width=200, bg=self.theme_bg, highlightthickness=0) # Match theme
    scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
    item_frame = ttk.Frame(canvas)

    item_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=item_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Scroll binding is handled globally in MaterialManager.__init__

    # Make draggable
    self.make_draggable(check_container, key)
    self.draggable_items[key] = check_container
    self.checklists[key] = {
        'container': check_container,
        'title_entry': title_entry,
        'item_frame': item_frame,
        'items': [] # List of item widgets/vars will be managed dynamically via children of item_frame
    }

    # Add initial items if provided
    if initial_data:
        for item in initial_data:
            self.add_checklist_item(item_frame, item.get('text', ''), item.get('checked', False), key)

    # Initial placement
    if key not in getattr(self, '_loading_memos', []): # Reuse loading flag or logic
        check_container.place(x=1450, y=50)

    return check_container


def add_vehicle_inspection_box_impl(self, initial_data=None, key=None):
    """Add a vehicle tab; keep the legacy API for record/session loaders."""
    panel = getattr(self, 'vehicle_panel', None)
    if panel is not None:
        return panel.add(initial_data=initial_data, key=key)
    import datetime
    if key is None:
        key = f"vehicle_inspection_{int(datetime.datetime.now().timestamp() * 1000)}"

    # Create draggable container
    container, widget = self.create_draggable_container(
        self.entry_inner_frame, "차량 점검", VehicleInspectionWidget, key,
        manage_list_key='vehicles',
        theme_bg=getattr(self, 'theme_bg', '#f0f0f0'),
        vehicle_list=getattr(self, 'vehicles', []),
        on_save=self.save_single_vehicle_data
    )

    # Track specifically for data gathering
    self.vehicle_inspections[key] = widget
    if widget not in self.vehicle_boxes:
        self.vehicle_boxes.append(widget)

    # Load initial data if provided
    if initial_data:
        widget.set_data(initial_data)

    # Start in the visible left side, including when the input is scrolled.
    if key not in getattr(self, '_loading_memos', []):
        offset = 20 * min(len(self.vehicle_inspections) - 1, 4)
        canvas = self.entry_canvas
        container.place(x=int(canvas.canvasx(12)) + offset,
                        y=int(canvas.canvasy(12)) + offset,
                        width=600, height=300)
        container.lift()

    return widget


def add_checklist_item_impl(self, parent_frame, text, checked, checklist_key):
    """Add a single item row to the checklist"""
    row_frame = ttk.Frame(parent_frame)
    row_frame.pack(fill='x', pady=1)

    var = tk.BooleanVar(value=checked)
    cb = ttk.Checkbutton(row_frame, variable=var, command=lambda: self.save_tab_config())
    cb.pack(side='left')

    entry = ttk.Entry(row_frame)
    entry.insert(0, text)
    entry.pack(side='left', fill='x', expand=True)
    entry.bind('<FocusOut>', lambda e: self.save_tab_config())

    def delete_this_item():
        row_frame.destroy()
        self.save_tab_config()

    btn_del_item = ttk.Label(row_frame, text="❌", font=('Arial', 8), cursor='hand2', foreground='gray')
    btn_del_item.pack(side='right', padx=2)
    btn_del_item.bind('<Button-1>', lambda e: delete_this_item())

    # Store refs in widget for retrieval during save
    row_frame._checklist_data = {'var': var, 'entry': entry}


def duplicate_checklist_impl(self, key):
    """Duplicate an existing checklist"""
    if self.layout_locked: return
    if key in self.checklists:
        # get current data
        original_items = []
        item_frame = self.checklists[key]['item_frame']
        for child in item_frame.winfo_children():
            if hasattr(child, '_checklist_data'):
                data = child._checklist_data
                original_items.append({
                    'text': data['entry'].get(),
                    'checked': data['var'].get()
                })

        title = self.checklists[key]['title_entry'].get()
        self.add_new_checklist(initial_data=original_items, initial_title=title)
        self.save_tab_config()


def _update_widget_position_impl(self, event, widget):
    """Internal helper to calculate and set widget position/size during interaction"""
    dx = event.x_root - widget._drag_start_root_x
    dy = event.y_root - widget._drag_start_root_y
    parent = widget.master
    parent_w = parent.winfo_width()
    parent_h = parent.winfo_height()

    if widget._interaction_mode == 'move':
        new_x = widget._drag_start_pos_x + dx
        new_y = widget._drag_start_pos_y + dy
        widget_w = widget.winfo_width()
        widget_h = widget.winfo_height()

        # Clamp
        if new_x < 0: new_x = 0
        elif new_x + widget_w > parent_w: new_x = max(0, parent_w - widget_w)

        # Loosen Y clamping: allow moving slightly below parent to trigger growth
        if new_y < 0: new_y = 0
        # Instead of strict clamp at parent_h, allow a small overflow to trigger resize logic
        elif new_y + widget_h > parent_h + 100: 
            new_y = parent_h + 100 - widget_h


        if widget.winfo_manager() != 'place':
            # Save original grid info and explicitly un-grid
            if not hasattr(widget, '_original_grid_info') and widget.grid_info():
                widget._original_grid_info = widget.grid_info()
            self._ensure_placeholder(widget)
            widget.grid_forget() # Explicitly un-grid before placing
            widget.place(width=widget_w, height=widget_h)

        widget.place(x=new_x, y=new_y)
        widget.lift()

        # Auto-update scroll region if we're dragging near bounds
        self._ensure_canvas_scroll_region()

    elif widget._interaction_mode == 'resize':
        new_width = max(50, widget._start_width + dx)
        new_height = max(20, widget._start_height + dy)
        current_x = widget.winfo_x()
        current_y = widget.winfo_y()

        # Clamp
        if current_x + new_width > parent_w: new_width = max(50, parent_w - current_x)
        if current_y + new_height > parent_h: new_height = max(20, parent_h - current_y)

        if widget.winfo_manager() != 'place':
            # Save original grid info and explicitly un-grid
            if not hasattr(widget, '_original_grid_info') and widget.grid_info():
                widget._original_grid_info = widget.grid_info()
            self._ensure_placeholder(widget)
            widget.grid_forget() # Explicitly un-grid before placing
            widget.place(x=current_x, y=current_y)

        widget.place(width=new_width, height=new_height)
        widget.lift()

    return "break"


def _apply_push_down_logic_impl(self, dragged_widget):
    """Recursively push widgets down if they overlap with the dragged widget"""
    try:
        # Current bounding box of the dragging widget
        x1 = dragged_widget.winfo_x()
        y1 = dragged_widget.winfo_y()
        w1 = dragged_widget.winfo_width()
        h1 = dragged_widget.winfo_height()

        padding = 10

        # Parent of the widgets
        parent = dragged_widget.master

        # OPTIMIZATION: Early filter candidates by parent to avoid massive iterations
        candidates = [w for w in self.draggable_items.values() if w.master == parent]

        for other in candidates:
            if other == dragged_widget:
                continue

            # [STABILITY FIX] Skip if the 'other' widget's BOTTOM is ABOVE the dragged widget's TOP.
            # Strictly ignore anything physically above the current interaction zone.
            if (other.winfo_y() + other.winfo_height()) <= y1:
                continue

            # Skip if not in the same parent
            if other.master != parent:
                continue

            # Skip if hidden
            if not other.winfo_ismapped():
                continue

            # Get current pos/size of the 'other' widget
            x2 = other.winfo_x()
            y2 = other.winfo_y()
            w2 = other.winfo_width()
            h2 = other.winfo_height()

            # 1. Horizontal Overlap Check
            if (x1 < x2 + w2) and (x1 + w1 > x2):
                # 2. Vertical Collision Check
                # If dragged_widget is hitting 'other' from the top
                # We check if the bottom of dragged_widget is below the top of other
                if y1 < y2 and (y1 + h1) > y2:
                    # New target Y for 'other'
                    new_y2 = y1 + h1 + padding

                    # Only move if it's actually pushing it DOWN
                    if y2 < new_y2:
                        if other.winfo_manager() == 'grid':
                            # Convert grid to place
                            self._ensure_placeholder(other)
                            other.lift()
                            other.place(x=x2, y=new_y2, width=w2, height=h2)
                        else:
                            # Just update place
                            other.place(y=new_y2)

                        # 3. Recursive Push
                        self._apply_push_down_logic(other)
    except Exception as e:
        # Silent fail to avoid interrupting drag
        print(f"Error in push-down logic: {e}")


def _adjust_parent_height_impl(self, parent, force=False):
    """Adjust parent frame height with performance check"""
    try:
        # Only update idletasks if forced or we're not in the middle of a high-speed interaction
        # This is the single biggest cause of UI stutter.
        if force:
            parent.update_idletasks()
        # 2. Start with the bounding box of all GRIDDED items
        try:
            # grid_bbox returns (x, y, width, height) of the grid
            bbox = parent.grid_bbox()
            required_h = bbox[1] + bbox[3] if bbox[3] > 0 else 0
            required_w = bbox[0] + bbox[2] if bbox[2] > 0 else 0
        except:
            required_h = 0
            required_w = 0

        # 3. Handle PLACE items as well (dragged/custom items)
        for child in parent.winfo_children():
            try:
                manager = child.winfo_manager()
                if manager == 'place':
                    info = child.place_info()
                    # Get relative or absolute y + height
                    y = child.winfo_y()
                    h = child.winfo_height()
                    x = child.winfo_x()
                    w = child.winfo_width()
                    required_h = max(required_h, y + h)
                    required_w = max(required_w, x + w)
            except:
                continue

        # Add some padding
        new_height = required_h + 30

        # [STABILITY FIX] Guard against collapse but allow growth
        # Width is now dictated by canvas parent in responsive mode.
        if new_height < 500: new_height = 800

        # Only resize if the required dimensions are significantly different
        current_h = parent.winfo_height()
        if abs(current_h - new_height) > 10:
             parent.config(height=new_height)

             # [RECURSIVE FIX] If this parent has a master that is part of the entry system,
             # adjust its height too. 
             if parent.master and parent.master != self.entry_inner_frame.master: # Don't go above canvas
                 self._adjust_parent_height(parent.master, force=force)

    except Exception as e:
        print(f"Error adjusting parent height: {e}")


def _ensure_placeholder_impl(self, widget, width=None, height=None):
    """Ensure a placeholder exists in the grid where the widget used to be"""
    # Map widget to a placeholder attribute name dynamically
    # We can use the widget id or a dictionary, but simpler to attach to widget
    if not hasattr(widget, '_placeholder'):
        # Use provided dims or current widget dims
        w = width if width is not None else widget.winfo_width()
        h = height if height is not None else widget.winfo_height()

        # Create a frame to hold the space
        widget._placeholder = ttk.Frame(widget.master, width=w, height=h)

        # Grid it at the original position
        if hasattr(widget, '_original_grid_info'):
            widget._placeholder.grid(**widget._original_grid_info)
            # Ensure the placeholder doesn't shrink and holds its size
            widget._placeholder.grid_propagate(False)


def _remove_placeholder_impl(self, widget):
    """Remove placeholder for a widget"""
    if hasattr(widget, '_placeholder'):
        widget._placeholder.destroy()
        del widget._placeholder

