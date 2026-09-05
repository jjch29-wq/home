import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import math
from datetime import datetime
import json
import traceback
from tkcalendar import DateEntry
from site_apps.lotte.src.utils.helpers import *

def _on_daily_usage_sash_changed_impl(self, event=None):
    """Handle sash position change to save ratio"""
    try:
        # Skip saving if locked
        if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
            return

        if hasattr(self, 'daily_usage_paned'):
            self.daily_usage_paned.update_idletasks()
            total_h = self.daily_usage_paned.winfo_height()
            if total_h > 0:
                sash_pos = 500
                ratio = sash_pos / total_h
                
                if not hasattr(self, 'tab_config'):
                    self.tab_config = {}
                
                self.tab_config['daily_usage_sash_ratio'] = ratio
                self.tab_config['daily_usage_sash_pos'] = sash_pos
                
                self.save_tab_config()
                print(f"Sash ratio saved: {ratio:.3f}")
    except Exception as e:
        print(f"Error saving sash position: {e}")


def on_material_selected_impl(self, event=None):
    """Update the model listbox based on selected material"""
    selection = self.cb_material.get()
    if not selection:
        return
        
    # Clear existing models
    self.list_models.delete(0, tk.END)
    
    # Extract pure material name
    mat_name = selection
    if " - " in mat_name:
        mat_name = mat_name.split(" - ")[0]
    if " (SN: " in mat_name:
        mat_name = mat_name.split(" (SN: ")[0]
    
    pure_mat_name = mat_name
    
    # Find unique models for this material in materials_df
    if not self.materials_df.empty:
        relevant_mats = self.materials_df[self.materials_df['품목명'] == pure_mat_name]
        if not relevant_mats.empty:
            unique_models = relevant_mats['모델명'].dropna().unique()
            unique_models = sorted([str(m).strip() for m in unique_models if str(m).strip()])
            
            for model in unique_models:
                self.list_models.insert(tk.END, model)
        # If no models found, add a placeholder
            if not unique_models:
                self.list_models.insert(tk.END, "(등록된 모델명 없음)")


def _on_trans_site_return_impl(self, event):
    self.auto_save_to_list(event, self.cb_trans_site, self.sites, 'sites')
    self.cb_warehouse.focus_set()


def _on_warehouse_return_impl(self, event):
    self.auto_save_to_list(event, self.cb_warehouse, self.warehouses, 'warehouses')
    self.ent_user.focus_set()


def _on_user_return_impl(self, event):
    self.auto_save_to_list(event, self.ent_user, self.users, 'users')
    self.ent_note.focus_set()


def on_monthly_usage_select_impl(self, event):
    """Update site and worker summaries when a row is selected in monthly usage tree"""
    selection = self.monthly_usage_tree.selection()
    if not selection:
        return
        
    item = selection[0]
    values = self.monthly_usage_tree.item(item, 'values')
    tags = self.monthly_usage_tree.item(item, 'tags')
    
    # If no data stored yet or error
    if not hasattr(self, 'current_monthly_df') or self.current_monthly_df.empty:
        return
        
    # If total row is selected, show total summaries
    if 'total' in tags:
        self._populate_monthly_summary_trees(self.current_monthly_df)
        return
        
    # Extract row info (Year, Month, Site, Material)
    try:
        year = int(values[0])
        month = int(values[1])
        site = str(values[2]).strip()
        # Material is at index 22 (after 연도-월-현장-작업자-작업시간-OT시간-OT금액-OT1...OT10-검사량-단가-출장비-일식-검사비)
        mat_name = str(values[22]).strip()
        
        matching_ids = []
        # Find matching MateriaIDs for this mat_name from master data
        if hasattr(self, 'materials_df') and not self.materials_df.empty:
            matches = self.materials_df[self.materials_df['MaterialName'].astype(str).str.contains(mat_name, case=False, na=False)]
            if not matches.empty:
                matching_ids = matches['MaterialID'].tolist()

        # Filter the current monthly dataset
        mask = (self.current_monthly_df['Year'] == year) & \
               (self.current_monthly_df['Month'] == month) & \
               (self.current_monthly_df['Site'] == site)

        
        # [ROBUST] Material Filter: matches 품목명 or direct MaterialID (for manual entries)
        if not matching_ids:
            # If not in master materials, check if mat_name itself exists as an ID in the data
            if mat_name in self.current_monthly_df['MaterialID'].astype(str).values:
                matching_ids = [mat_name]
            else:
                # Fallback: check case/space insensitive match in data
                m_norm = mat_name.replace(' ', '').upper()
                possible_ids = self.current_monthly_df['MaterialID'].dropna().unique()
                for p_id in possible_ids:
                    if str(p_id).replace(' ', '').upper() == m_norm:
                        matching_ids.append(p_id)
        
        if matching_ids:
            mask = mask & (self.current_monthly_df['MaterialID'].isin(matching_ids))
        
        filtered_subset = self.current_monthly_df[mask]
        self._populate_monthly_summary_trees(filtered_subset)
        
    except Exception as e:
        print(f"DEBUG: Error in on_monthly_usage_select: {e}")


def on_drag_stop_impl(self, event, widget=None):
    """Handle end of dragging or resizing and auto-save"""
    if widget is None:
        widget = event.widget
    if hasattr(widget, '_interaction_mode'):
        mode = getattr(widget, '_interaction_mode')
        del widget._interaction_mode
        
        # Update parent height if something moved or resized
        self._adjust_parent_height(widget.master, force=True)
        
        # Auto-save layout
        self.save_tab_config()


def on_drag_start_impl(self, event, widget=None):
    """Begin dragging widget"""
    if self.layout_locked:
        return "break" # Prevent movement and stop propagation
        
    if widget is None:
        widget = event.widget
    widget._interaction_mode = 'move'
    
    # Save absolute start position of mouse
    widget._drag_start_root_x = event.x_root
    widget._drag_start_root_y = event.y_root
    
    # Save initial widget position relative to parent
    widget._drag_start_pos_x = widget.winfo_x()
    widget._drag_start_pos_y = widget.winfo_y()
    
    # Ensure we have grid info (redundant but safe)
    if not hasattr(widget, '_original_grid_info') and widget.grid_info():
        widget._original_grid_info = widget.grid_info()


def on_resize_start_impl(self, event, widget=None):
    """Begin resizing widget"""
    if self.layout_locked:
        return "break"
        
    if widget is None:
        widget = event.widget
    widget._interaction_mode = 'resize'
    
    # Save absolute start position of mouse
    widget._drag_start_root_x = event.x_root
    widget._drag_start_root_y = event.y_root
    
    # Save initial size
    widget._start_width = widget.winfo_width()
    widget._start_height = widget.winfo_height()
    
    # Ensure we have grid info
    if not hasattr(widget, '_original_grid_info') and widget.grid_info():
        widget._original_grid_info = widget.grid_info()
    return "break"


def on_mouse_motion_impl(self, event, widget=None):
    """Handle dragging or resizing motion with performance throttling"""
    if self.layout_locked:
        return "break"
        
    if widget is None:
        widget = event.widget
    
    if not hasattr(widget, '_interaction_mode'):
        return
    
    # PERFORMANCE THROTTLE: Limit updates to ~60fps (16ms)
    curr_time = time.time()
    if curr_time - self._last_motion_time < 0.016:
        # Still update the physical position of the widget being interacted with
        # or the user will feel lag in the initial drag/resize itself.
        self._update_widget_position(event, widget)
        return
    
    self._last_motion_time = curr_time
    
    # Apply positioning only (No collision, No auto-resize)
    self._update_widget_position(event, widget)


def on_recent_record_click_impl(self, event):
    """최근 기록 테이블의 항목을 클릭했을 때 상단 입력 폼에 해당 데이터를 로드"""
    selection = self.tv_recent.selection()
    if not selection: return
    item = self.tv_recent.item(selection[0])
    values = item.get('values')
    if not values: return
    
    record_id = values[0]
    
    if hasattr(self, 'daily_usage_df') and not self.daily_usage_df.empty:
        try:
            record_idx = int(record_id)
            if record_idx in self.daily_usage_df.index:
                record = self.daily_usage_df.loc[record_idx].to_dict()
                # Use existing method to populate the form
                self.load_daily_usage_to_form(record)
                print(f"DEBUG: Loaded recent record {record_idx} to form.")
        except Exception as e:
            print(f"DEBUG: Error loading recent record: {e}")


def _on_daily_usage_select_impl(self, event):
    """[NEW] Update Note Detail Area and load record to form when a row is selected in Site tab"""
    if not hasattr(self, 'daily_usage_tree'): return
    
    selection = self.daily_usage_tree.selection()
    if not selection:
        if hasattr(self, 'txt_daily_note_detail'):
            self.txt_daily_note_detail.config(state='normal')
            self.txt_daily_note_detail.delete('1.0', tk.END)
            self.txt_daily_note_detail.config(state='disabled')
        return
        
    item = selection[0]
    tags = self.daily_usage_tree.item(item, 'tags')
    if tags and tags[0].isdigit():
        idx = int(tags[0])
        if idx in self.daily_usage_df.index:
            row_data = self.daily_usage_df.loc[idx].to_dict()
            self.load_daily_usage_to_form(row_data)

    # Update note detail area as before
    if hasattr(self, 'txt_daily_note_detail'):
        values = self.daily_usage_tree.item(item, 'values')
        if values:
            try:
                cols = self.daily_usage_tree['columns']
                if '비고' in cols:
                    note_idx = list(cols).index('비고')
                    if note_idx < len(values):
                        note_text = values[note_idx]
                        self.txt_daily_note_detail.config(state='normal')
                        self.txt_daily_note_detail.delete('1.0', tk.END)
                        self.txt_daily_note_detail.insert(tk.END, note_text)
                        self.txt_daily_note_detail.config(state='disabled')
            except Exception as e:
                print(f"Detail view error: {e}")


def on_budget_tree_select_impl(self, event):
    """No-op - 하단 목록 제거됨"""
    pass


def _on_daily_usage_resize_impl(self, event):
    """Handle window resize to maintain sash ratio or absolute position if locked"""
    try:
        if not hasattr(self, 'daily_usage_paned'): return
        
        # If locked, maintain absolute position from top
        if hasattr(self, 'daily_usage_sash_locked') and self.daily_usage_sash_locked:
            self._restore_locked_position()
            return

        # Otherwise maintain ratio
        if hasattr(self, 'tab_config') and 'daily_usage_sash_ratio' in self.tab_config:
            ratio = self.tab_config['daily_usage_sash_ratio']
            total_h = self.daily_usage_paned.winfo_height()
            
            if total_h > 200:
                new_pos = int(total_h * ratio)
                min_pos, max_pos = 50, total_h - 50
                new_pos = max(min_pos, min(new_pos, max_pos))
                
            getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, new_pos)
    except Exception as e:
        print(f"Error handling resize: {e}")


def _on_main_window_resize_impl(self, event):
    """Handle main window resize to maintain all sash ratios"""
    try:
        # Only process for actual window resize, not widget events
        if event.widget == self.root:
            # Check if daily usage tab exists and has saved ratio
            if hasattr(self, 'daily_usage_paned') and hasattr(self, 'tab_config') and 'daily_usage_sash_ratio' in self.tab_config:
                self.root.after(100, self._ensure_daily_usage_sash_visibility)
            
            # Check if inout tab exists and has saved ratio
            if hasattr(self, 'inout_paned') and hasattr(self, 'tab_config') and 'inout_sash_ratio' in self.tab_config:
                self.root.after(100, self._ensure_inout_sash_visibility)
                
    except Exception as e:
        print(f"Error handling main window resize: {e}")


def on_daily_usage_tree_select_impl(self, event):
    """No-op: details panel removed"""
    pass


def on_tab_drag_start_impl(self, event):
    """Start tab dragging by identifying the tab under the cursor"""
    try:
        # Check if click is on a tab tab
        clicked_tab = self.notebook.identify(event.x, event.y)
        if clicked_tab == "label":
            # Find which index this is
            self._drag_start_index = self.notebook.index(f"@{event.x},{event.y}")
            self._current_drag_index = self._drag_start_index
        else:
            self._drag_start_index = None
    except:
        self._drag_start_index = None


def on_tab_drag_impl(self, event):
    """Handle visual swapping of tabs during drag"""
    if not hasattr(self, '_drag_start_index') or self._drag_start_index is None:
        return

    try:
        # Find the index of the tab currently under the cursor
        target_index = self.notebook.index(f"@{event.x},{event.y}")
        
        if target_index != self._current_drag_index:
            # Get the widget of the tab we are dragging
            tab_widget = self.notebook.tabs()[self._current_drag_index]
            tab_text = self.notebook.tab(self._current_drag_index, "text")
            
            # Use insert to move tab widget. 
            # Note: insert(pos, widget) handles the reordering logic in Notebook
            self.notebook.insert(target_index, tab_widget, text=tab_text)
            
            self._current_drag_index = target_index
            self.notebook.select(target_index)
    except:
        pass


def on_tab_drag_end_impl(self, event):
    """Finalize tab order and save configuration"""
    if hasattr(self, '_drag_start_index') and self._drag_start_index is not None:
        if self._current_drag_index != self._drag_start_index:
            print(f"Tab reordered: {self._drag_start_index} -> {self._current_drag_index}")
            # [FIX] Force save on manual drag end to ensure order is persisted
            self.save_tab_config(force=True)
        
        self._drag_start_index = None
        self._current_drag_index = None


def on_tab_changed_impl(self, event=None):
    """Handle tab selection change event"""
    try:
        # 1. Save configuration when tab changes (respects is_ready via save_tab_config)
        self.save_tab_config()

        # 1-1. 탭 이동 시 실제 데이터(Excel)도 자동 저장
        if getattr(self, 'is_ready', False):
            self.save_data()
        
        # 2. Handle specific tab UI adjustments
        current_tab = self.notebook.select()
        if not current_tab:
            return
            
        # Convert widget path to index if needed
        try:
            current_tab_idx = self.notebook.index("current")
            tab_text = self.notebook.tab(current_tab_idx, "text")
        except:
            tab_text = ""

        # Check for Daily Usage tab
        if (hasattr(self, 'tab_daily_usage') and str(current_tab) == str(self.tab_daily_usage)) or \
           tab_text == '현장별 일일 사용량 기입':
            
            print("Daily usage tab selected - ensuring visibility")
            # Force multiple updates when tab is selected
            self.refresh_inquiry_filters()
            self.update_daily_usage_view() # [NEW] Trigger auto-hiding logic
            self.root.after(50, self._ensure_daily_usage_sash_visibility)
            self.root.after(200, self._ensure_daily_usage_sash_visibility)
            self.root.after(400, self._ensure_canvas_scroll_region)
            
            # Also ensure the inner history sash is visible
            self.root.after(150, self._ensure_sash_visible)
        
        elif tab_text == '월별 집계':
            print("Monthly usage tab selected - refreshing view")
            self.refresh_inquiry_filters()
            self.update_monthly_usage_view()
        elif tab_text == '공사실행예산서':
            print("Construction budget tab selected - refreshing view")
            self.update_budget_site_view()
        elif tab_text == '입출고 관리':
            print("In/Out Management tab selected - refreshing history")
            self.refresh_inout_history()
    except Exception as e:
        print(f"Error in tab change handler: {e}")


def on_closing_impl(self):
    """Handle window closing event"""
    self.save_tab_config(force=True)
    if hasattr(self, 'ndt_calculator') and hasattr(self.ndt_calculator, 'save_ui_state'):
        self.ndt_calculator.save_ui_state()
    self.is_ready = False  # 프로그램 종료 중 발생하는 UI 이벤트가 설정을 덮어쓰는 것 방지
    self.root.destroy()


