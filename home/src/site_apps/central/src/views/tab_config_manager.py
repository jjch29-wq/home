# -*- coding: utf-8 -*-
"""
중앙지사 자재관리 탭 설정 및 탭 네비게이션 관리 Mixin
"""

import json
import logging
import os
import re
import sys
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

logger = logging.getLogger(__name__)

class TabConfigMixin:
    """Notebook 탭 설정 저장/복원 및 순서/단축키 관리 Mixin"""
    def setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts for navigation"""
        # Ctrl+Tab to switch between notebook tabs (forward)
        self.root.bind('<Control-Tab>', self.next_tab)
        # Ctrl+Shift+Tab to switch between notebook tabs (backward)
        self.root.bind('<Control-Shift-Tab>', self.prev_tab)
        
        # Alt+숫자 for direct tab access
        self.root.bind('<Alt-Key-1>', lambda e: self.notebook.select(0))
        self.root.bind('<Alt-Key-2>', lambda e: self.notebook.select(1))
        self.root.bind('<Alt-Key-3>', lambda e: self.notebook.select(2))
        self.root.bind('<Alt-Key-4>', lambda e: self.notebook.select(3))
        self.root.bind('<Alt-Key-5>', lambda e: self.notebook.select(4))
        self.root.bind('<Alt-Key-6>', lambda e: self.notebook.select(5))
        
        # Right-click on notebook for tab reordering
        # Tab interactions removed here, kept in create_widgets
    
    def next_tab(self, event=None):
        """Switch to next tab"""
        current = self.notebook.index(self.notebook.select())
        total = self.notebook.index('end')
        next_tab = (current + 1) % total
        self.notebook.select(next_tab)
        return 'break'
    
    def prev_tab(self, event=None):
        """Switch to previous tab"""
        current = self.notebook.index(self.notebook.select())
        total = self.notebook.index('end')
        prev_tab = (current - 1) % total
        self.notebook.select(prev_tab)
        return 'break'
    
    def show_tab_context_menu(self, event):
        """Show context menu for tab options"""
        try:
            clicked_tab = self.notebook.index(f"@{event.x},{event.y}")
            self.notebook.select(clicked_tab)
            
            context_menu = tk.Menu(self.root, tearoff=0)
            context_menu.add_command(label="🔄 기본 탭 순서로 초기화", 
                                    command=self.reset_tab_order_to_default)
            context_menu.post(event.x_root, event.y_root)
        except:
            pass
    
    def reset_tab_order_to_default(self):
        """Reset notebook tabs to standard default order"""
        default_order = [
            '현재 재고 현황',
            '데이터 가져오기/내보내기',
            '입출고 관리',
            '현장별 일일 사용량 기입',
            '현장 일일기록 조회 및 관리',
            '월별 집계',
            '작업/감독일보',
            '기성 정산 (NDT)',
            '공사실행예산서'
        ]
        try:
            tab_map = {}
            for tab_id in self.notebook.tabs():
                tab_map[self.notebook.tab(tab_id, "text")] = tab_id
                
            for i, tab_text in enumerate(default_order):
                if tab_text in tab_map:
                    tab = tab_map[tab_text]
                    self.notebook.insert(i, tab, text=tab_text)
                    
            if hasattr(self, 'tab_config'):
                self.tab_config['tab_order'] = default_order
            self.save_tab_config(force=True)
            messagebox.showinfo("완료", "탭 순서가 기본 상태로 복원되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"탭 순서 복원 실패: {e}")
    
    def move_tab_left(self, tab_index):
        """Move tab one position to the left"""
        if tab_index > 0:
            # Get tab info
            tab = self.notebook.tabs()[tab_index]
            text = self.notebook.tab(tab_index, "text")
            
            # Remove and reinsert at new position
            self.notebook.insert(tab_index - 1, tab, text=text)
            self.notebook.select(tab_index - 1)
            
            # Save configuration immediately
            self.save_tab_config()
    
    def move_tab_right(self, tab_index):
        """Move tab one position to the right"""
        if tab_index < self.notebook.index('end') - 1:
            # Get tab info
            tab = self.notebook.tabs()[tab_index]
            text = self.notebook.tab(tab_index, "text")
            
            # Remove and reinsert at new position
            self.notebook.insert(tab_index + 2, tab, text=text)
            self.notebook.select(tab_index + 1)
            
            # Save configuration immediately
            self.save_tab_config()

    def on_tab_drag_start(self, *args, **kwargs):
        pass

    def on_tab_drag(self, *args, **kwargs):
        pass

    def on_tab_drag_end(self, *args, **kwargs):
        pass
    
    def save_tab_config(self, force=False):
        """Save current tab configuration to memory and disk"""
        try:
            # [STABILITY] Skip saving if app is not fully ready (prevents overwriting saved tab index during init)
            if not getattr(self, 'is_ready', False) and not force:
                return

            if getattr(self, '_tab_only_config_save', False) and not force:
                if not hasattr(self, 'tab_config') or not isinstance(self.tab_config, dict):
                    self.tab_config = {}
                try:
                    current_tab_widget = self.notebook.select()
                    self.tab_config['selected_tab'] = self.notebook.index(current_tab_widget)
                    self.tab_config['selected_tab_text'] = self.notebook.tab(current_tab_widget, "text")
                except Exception:
                    pass
                final_config = dict(self.tab_config)
                if os.path.exists(self.config_path):
                    try:
                        with open(self.config_path, 'r', encoding='utf-8') as f:
                            loaded_disk = json.load(f)
                        if isinstance(loaded_disk, dict):
                            loaded_disk['selected_tab'] = self.tab_config.get('selected_tab', 0)
                            loaded_disk['selected_tab_text'] = self.tab_config.get('selected_tab_text', '')
                            final_config = loaded_disk
                    except (OSError, json.JSONDecodeError):
                        pass
                import tempfile
                cfg_dir = os.path.dirname(os.path.abspath(self.config_path))
                os.makedirs(cfg_dir, exist_ok=True)
                temp_cfg = None
                try:
                    with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=cfg_dir, prefix='.tab_config-', suffix='.json', delete=False) as tf:
                        temp_cfg = tf.name
                        json.dump(final_config, tf, ensure_ascii=False, indent=2)
                    os.replace(temp_cfg, self.config_path)
                    temp_cfg = None
                finally:
                    if temp_cfg and os.path.exists(temp_cfg):
                        try:
                            os.remove(temp_cfg)
                        except OSError:
                            pass
                self.tab_config = final_config
                return

            # Force update to ensure all coordinates and sizes are accurate
            try:
                self.root.update_idletasks()
            except:
                pass

            # Initialize or keep existing tab_config
            if not hasattr(self, 'tab_config'):
                self.tab_config = {}
                
            # Update with current core values
            current_tab_idx = 0
            current_tab_text = ""
            try:
                current_tab_widget = self.notebook.select()
                current_tab_idx = self.notebook.index(current_tab_widget)
                current_tab_text = self.notebook.tab(current_tab_widget, "text")
            except:
                current_tab_idx = self.tab_config.get('selected_tab', 0)
                current_tab_text = self.tab_config.get('selected_tab_text', "")

            # Get current tab order
            tab_order = []
            for tab_id in self.notebook.tabs():
                tab_order.append(self.notebook.tab(tab_id, "text"))

            daily_usage_sash_pos = None
            if hasattr(self, 'daily_usage_paned'):
                try: daily_usage_sash_pos = 500
                except: pass

            daily_history_sash_pos = None
            if hasattr(self, 'daily_history_paned'):
                try: daily_history_sash_pos = self.daily_history_paned.sashpos(0)
                except: pass

            bottom_dashboard_sashes = []
            if hasattr(self, 'bottom_dashboard') and isinstance(self.bottom_dashboard, ttk.PanedWindow):
                try:
                    bottom_dashboard_sashes = [self.bottom_dashboard.sashpos(0), self.bottom_dashboard.sashpos(1)]
                except: pass

            # Only user drags update sash preferences; hidden/startup layouts are temporary.
            entry_inner_frame_height = None
            if hasattr(self, 'entry_inner_frame'):
                try: entry_inner_frame_height = self.entry_inner_frame.winfo_height()
                except: pass
                
            window_state = 'normal'
            window_width = 1200
            window_height = 800
            try:
                window_state = self.root.state()
                window_width = self.root.winfo_width()
                window_height = self.root.winfo_height()
            except: pass

            self.tab_config.update({
                'selected_tab': current_tab_idx,
                'selected_tab_text': current_tab_text,
                'tab_order': tab_order,
                'sites': self.sites,
                'daily_units': self.daily_units,
                'users': getattr(self, 'users', []),
                'companies': getattr(self, 'companies', []),
                'warehouses': getattr(self, 'warehouses', []),
                'equipments': getattr(self, 'equipments', []),
                'aux_equipments': getattr(self, 'aux_equipments', ['JIREH MICROBE BTG006', 'COBRA Manual Scanner']),
                'test_items': getattr(self, 'test_items', []),
                'applied_codes': getattr(self, 'applied_codes', []),
                'co_code_list': getattr(self, 'co_code_list', []),
                'vehicles': getattr(self, 'vehicles', []),
                'materials': getattr(self, 'carestream_films', []),
                'worktimes': getattr(self, 'worktimes', []),
                'ot_times': getattr(self, 'ot_times', []),
                'layout_locked': getattr(self, 'layout_locked', False),
                'resolution_locked': getattr(self, 'resolution_locked', False),
                'locked_width': getattr(self, 'locked_width', 1200),
                'locked_height': getattr(self, 'locked_height', 800),
                'daily_usage_sash_locked': getattr(self, 'daily_usage_sash_locked', False),
                'daily_usage_sash_pos': daily_usage_sash_pos,
                'daily_history_sash_pos': daily_history_sash_pos,
                'bottom_dashboard_sashes': bottom_dashboard_sashes,
                'monthly_sash_positions': self.tab_config.get('monthly_sash_positions', [125, 262]),
                'monthly_sash_ratios': self.tab_config.get('monthly_sash_ratios', [0.282, 0.590]),
                'entry_inner_frame_height': entry_inner_frame_height,
                'history_visible_cols': getattr(self, 'manual_visible_cols', []),
                'monthly_visible_cols': getattr(self, 'monthly_visible_cols', []),
                'budget_view_visible_cols': getattr(self, 'budget_view_visible_cols', []),
                'budget_view_heading_aliases': getattr(self, 'budget_view_heading_aliases', {}),
                'budget_view_custom_columns': getattr(self, 'budget_view_custom_columns', []),
                'window_state': window_state,
                'window_width': window_width,
                'window_height': window_height
            })
            
            # Save current stock column widths
            if hasattr(self, 'stock_tree'):
                self.tab_config['stock_col_widths'] = {}
                for col in self.stock_tree['columns']:
                    self.tab_config['stock_col_widths'][col] = self.stock_tree.column(col, 'width')

            # Save in/out history column widths
            self.tab_config['inout_col_widths'] = {}
            if hasattr(self, 'inout_tree'):
                for col in self.inout_tree['columns']:
                    self.tab_config['inout_col_widths'][col] = self.inout_tree.column(col, 'width')

            # Save recent entries column widths
            self.tab_config['tv_recent_col_widths'] = {}
            if hasattr(self, 'tv_recent'):
                for col in self.tv_recent['columns']:
                    self.tab_config['tv_recent_col_widths'][col] = self.tv_recent.column(col, 'width')

            # Save daily usage column widths
            self.tab_config['daily_usage_col_widths'] = {}
            if hasattr(self, 'daily_usage_tree'):
                for col in self.daily_usage_tree['columns']:
                    self.tab_config['daily_usage_col_widths'][col] = self.daily_usage_tree.column(col, 'width')
            
            # Save monthly usage column widths
            self.tab_config['monthly_usage_col_widths'] = {}
            if hasattr(self, 'monthly_usage_tree'):
                for col in self.monthly_usage_tree['columns']:
                    self.tab_config['monthly_usage_col_widths'][col] = self.monthly_usage_tree.column(col, 'width')
            
            # [NEW] Save summary table column widths
            self.tab_config['site_summary_col_widths'] = {}
            if hasattr(self, 'site_summary_tree'):
                for col in self.site_summary_tree['columns']:
                    self.tab_config['site_summary_col_widths'][col] = self.site_summary_tree.column(col, 'width')
            
            self.tab_config['worker_summary_col_widths'] = {}
            if hasattr(self, 'worker_summary_tree'):
                for col in self.worker_summary_tree['columns']:
                    self.tab_config['worker_summary_col_widths'][col] = self.worker_summary_tree.column(col, 'width')
            
            # [NEW] Save report column widths
            self.tab_config['report_col_widths'] = {}
            if hasattr(self, 'report_tree'):
                for col in self.report_tree['columns']:
                    self.tab_config['report_col_widths'][col] = self.report_tree.column(col, 'width')

            # Save budget site performance column widths
            self.tab_config['budget_view_col_widths'] = {}
            if hasattr(self, 'budget_view_tree'):
                for col in self.budget_view_tree['columns']:
                    self.tab_config['budget_view_col_widths'][col] = self.budget_view_tree.column(col, 'width')

            # Save Daily Work Log pane layout & NDT result tree column widths
            work_log = getattr(self, 'tab_daily_work_log', None)
            if work_log is not None and hasattr(work_log, 'sync_layout_to_config'):
                work_log.sync_layout_to_config()
            
            # Save tab order (항상 표준 기본 순서 고정)
            default_tab_order = [
                '현재 재고 현황',
                '데이터 가져오기/내보내기',
                '입출고 관리',
                '현장별 일일 사용량 기입',
                '현장 일일기록 조회 및 관리',
                '월별 집계',
                '작업/감독일보',
                '기성 정산 (NDT)',
                '공사실행예산서'
            ]
            self.tab_config['tab_order'] = default_tab_order
                
            # [NEW] Save equipment list
            self.tab_config['equipments'] = self.equipments
            self.tab_config['aux_equipments'] = getattr(self, 'aux_equipments', ['JIREH MICROBE BTG006', 'COBRA Manual Scanner'])
            
            # Save Draggable items
            # Prepare in-memory geometries for merging (capture current states)
            current_geometries = {}
            
            for key, widget in self.draggable_items.items():
                try:
                    # [LAYOUT FIX] 핵심 박스는 좌표 저장 제외
                    # 재시작 시 기본 grid 배치로 복원하여 겹침/잘림 방지
                    if key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
                        continue

                    manager = widget.winfo_manager()
                    # Capture everything that is visible and managed
                    if manager in ['place', 'grid']:
                        current_geometries[key] = {
                            'x': widget.winfo_x(),
                            'y': widget.winfo_y(),
                            'width': widget.winfo_width(),
                            'height': widget.winfo_height(),
                            'hidden': False
                        }
                        
                        if hasattr(widget, '_label_widget'):
                            current_geometries[key]['custom_label'] = widget._label_widget.cget('text')
                        if hasattr(widget, '_manage_list_key') and widget._manage_list_key:
                            current_geometries[key]['manage_list_key'] = widget._manage_list_key
                        if key.startswith('clone_'):
                            current_geometries[key]['is_clone'] = True
                            current_geometries[key]['widget_class_name'] = widget._widget_class.__name__
                            saved_kwargs = widget._widget_kwargs.copy()
                            if 'values' in saved_kwargs: del saved_kwargs['values']
                            # Remove un-serializable callable objects (like on_save callbacks)
                            keys_to_del = [k for k, v in saved_kwargs.items() if callable(v)]
                            for k in keys_to_del: del saved_kwargs[k]
                            current_geometries[key]['widget_kwargs'] = saved_kwargs
                        if key.startswith('memo_'):
                            current_geometries[key]['text'] = self.memos[key]['text_widget'].get('1.0', 'end-1c')
                            current_geometries[key]['memo_title'] = self.memos[key]['title_entry'].get()
                        if key in getattr(self, 'checklists', {}):
                            current_geometries[key]['checklist_title'] = self.checklists[key]['title_entry'].get()
                            items_data = []
                            for child in self.checklists[key]['item_frame'].winfo_children():
                                if hasattr(child, '_checklist_data'):
                                    items_data.append({
                                        'text': child._checklist_data['entry'].get(),
                                        'checked': child._checklist_data['var'].get()
                                    })
                            current_geometries[key]['checklist_items'] = items_data
                        if key in getattr(self, 'vehicle_inspections', {}):
                            current_geometries[key]['vehicle_data'] = self.vehicle_inspections[key].get_data()
                    elif manager == '': # Hidden or not managed
                        # Only mark as hidden if it was previously in place/grid (to avoid junk)
                        if key in self.tab_config.get('draggable_geometries', {}):
                            current_geometries[key] = {'hidden': True}
                except Exception as e:
                    print(f"Skipping save for draggable widget {key}: {e}")

            # [DEFINITIVE FIX] Deep merge current_geometries into disk config
            final_config = dict(getattr(self, 'tab_config', {}) or {})
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, 'r', encoding='utf-8') as f:
                        loaded_disk = json.load(f)
                    if isinstance(loaded_disk, dict):
                        final_config = loaded_disk
                except (OSError, json.JSONDecodeError) as read_err:
                    logger.warning("Existing tab config could not be read before save (%s: %s); preserving in-memory config", type(read_err).__name__, read_err)
            
            # Use disk's geometries as base, then update with WHAT WE JUST CAPTURED in memory
            disk_geos = dict(final_config.get('draggable_geometries', {}) or {})

            # [LAYOUT FIX] 과거에 저장된 핵심 박스 좌표도 제거
            for core_key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
                if core_key in disk_geos:
                    del disk_geos[core_key]

            disk_geos.update(current_geometries)
            
            # Apply other tab state updates
            final_config.update(self.tab_config)
            final_config['draggable_geometries'] = disk_geos
            # [NEW] Save hidden_sites to config
            final_config['hidden_sites'] = list(getattr(self, 'hidden_sites', []))
            
            import tempfile
            cfg_dir = os.path.dirname(os.path.abspath(self.config_path))
            os.makedirs(cfg_dir, exist_ok=True)
            temp_cfg = None
            try:
                with tempfile.NamedTemporaryFile('w', encoding='utf-8', dir=cfg_dir, prefix='.tab_config-', suffix='.json', delete=False) as tf:
                    temp_cfg = tf.name
                    json.dump(final_config, tf, ensure_ascii=False, indent=2)
                os.replace(temp_cfg, self.config_path)
                temp_cfg = None
            finally:
                if temp_cfg and os.path.exists(temp_cfg):
                    try:
                        os.remove(temp_cfg)
                    except OSError:
                        pass
            
            # Update in-memory copy as well
            self.tab_config = final_config
            
            print(f"Configuration saved. Locked: {self.tab_config.get('layout_locked')}, Tab: {self.tab_config.get('selected_tab')}")
        except Exception as e:
            logger.error("Failed to save tab config: %s", e, exc_info=True)
            print(f"Failed to save tab config: {e}")

    
    def load_tab_config(self):
        """Load and restore tab configuration"""
        try:
            if not hasattr(self, 'tab_config'):
                self.tab_config = {}
                
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.tab_config = json.load(f)
                
                config = self.tab_config
                
                # Restore lists - Respecting user deletions
                if 'daily_units' in config:
                    loaded = config['daily_units']
                    if loaded: # If user has a saved list (even if they deleted items)
                        self.daily_units = list(dict.fromkeys(loaded))
                    else: # Fallback only if the list is literally empty
                        self.daily_units = ['EA', 'CAN', 'SET', 'KG', 'M', '매', 'I/D', 'P,M,I/D', 'M,I/D', 'Point', 'Meter', 'Inch', 'Dia']
                
                # Apply high-level lock states immediately to internal variables
                self.layout_locked = config.get('layout_locked', False)
                self.resolution_locked = config.get('resolution_locked', False)
                self.daily_usage_sash_locked = config.get('daily_usage_sash_locked', False)
                self.locked_width = config.get('locked_width', 1200)
                self.locked_height = config.get('locked_height', 800)

                # Get current tab order
                current_order = []
                for tab in self.notebook.tabs():
                    tab_index = self.notebook.index(tab)
                    tab_text = self.notebook.tab(tab_index, "text")
                    current_order.append((tab_text, tab))
                
                # Create mapping from text to tab widget
                tab_map = {text: tab for text, tab in current_order}
                
                # Restore tab order to standard default order (항상 표준 순서 강제 복원)
                default_order = [
                    '현재 재고 현황',
                    '데이터 가져오기/내보내기',
                    '입출고 관리',
                    '현장별 일일 사용량 기입',
                    '현장 일일기록 조회 및 관리',
                    '월별 집계',
                    '작업/감독일보',
                    '기성 정산 (NDT)',
                    '공사실행예산서'
                ]
                for i, tab_text in enumerate(default_order):
                    if tab_text in tab_map:
                        tab = tab_map[tab_text]
                        current_pos = self.notebook.index(tab)
                        if current_pos != i:
                            self.notebook.insert(i, tab, text=tab_text)
                self.tab_config['tab_order'] = default_order
                
                # 마지막으로 보던 탭 복원 (설정 없으면 기본값: '현장별 일일 사용량 기입')
                selected_text = config.get('selected_tab_text')
                selected_idx = config.get('selected_tab')
                if not selected_text and isinstance(selected_idx, int) and 0 <= selected_idx < self.notebook.index('end'):
                    try:
                        selected_text = self.notebook.tab(selected_idx, "text")
                    except: pass
                if not selected_text:
                    selected_text = "현장별 일일 사용량 기입"
                
                tab_restored = False
                if selected_text:
                    # [ROBUST] Find tab by name instead of index
                    for i in range(self.notebook.index('end')):
                        if self.notebook.tab(i, "text") == selected_text:
                            try:
                                self.notebook.select(i)
                                tab_restored = True
                                break
                            except: pass
                
                if not tab_restored and isinstance(selected_idx, int) and 0 <= selected_idx < self.notebook.index('end'):
                    try:
                        self.notebook.select(selected_idx)
                        tab_restored = True
                    except: pass
                
                # Force update after selection so the tab is rendered and children computed
                self.root.update_idletasks()
                
                # Restore sites list - use in-place update
                self.sites[:] = config.get('sites', [])
                # If sites list is empty, try to populate from current daily_usage_df
                if not self.sites and not self.daily_usage_df.empty:
                    # [FIX] Collect and normalize sites in-place to maintain widget references and resolve crash
                    self.sites[:] = sorted(list(self.daily_usage_df['Site'].dropna().astype(str).apply(self.normalize_site_name).unique()))
                # [NEW] Restore hidden sites list and ensure requested defaults are present
                saved_hidden = config.get('hidden_sites', [])
                for s in ["초안", "롯데현장"]:
                    if s not in saved_hidden: saved_hidden.append(s)
                self.hidden_sites[:] = saved_hidden
                
                # Restore users list - use in-place update to preserve references
                self.users[:] = config.get('users', [])
                
                # [MIGRATION] Load and immediately migrate worktimes to ensure 익일 marker
                raw_worktimes = config.get('worktimes', [])
                self.worktimes = self._migrate_worktimes(raw_worktimes)
                
                # Restore vehicles, equipments, and preferred materials
                raw_vehicles = config.get('vehicles', [])
                updated_vehicles = []
                from site_apps.central.src.vehicle_catalog import vehicle_label
                for v in raw_vehicles:
                    updated_vehicles.append(vehicle_label(v))
                self.vehicles[:] = sorted(list(set(updated_vehicles)))
                
                self.equipments[:] = config.get('equipments', [])
                loaded_aux = config.get('aux_equipments', ['JIREH MICROBE BTG006', 'COBRA Manual Scanner'])
                if not loaded_aux:
                    loaded_aux = ['JIREH MICROBE BTG006', 'COBRA Manual Scanner']
                for req in ['JIREH MICROBE BTG006', 'COBRA Manual Scanner']:
                    if req not in loaded_aux:
                        loaded_aux.append(req)
                self.aux_equipments = list(dict.fromkeys(loaded_aux))
                self.test_items[:] = config.get('test_items', [])
                self.applied_codes[:] = config.get('applied_codes', [])
                if 'co_code_list' in config and isinstance(config.get('co_code_list'), list):
                    if not hasattr(self, 'co_code_list') or not isinstance(self.co_code_list, list):
                        self.co_code_list = []
                    self.co_code_list[:] = sorted(set(self.co_code_list) | {str(v).strip() for v in config.get('co_code_list', []) if v and str(v).strip()})
                if 'materials' in config:
                    self.carestream_films[:] = config['materials']
                
                # [NEW] Restore companies list
                self.companies[:] = config.get('companies', [])
                if hasattr(self, 'cb_daily_company'):
                    self.cb_daily_company['values'] = sorted(self.companies)

                
                # Refresh all WorkerDataGroup dropdowns with the migrated list
                for i in range(1, 11):
                    group_attr = f'worker_group{i}'
                    if hasattr(self, group_attr):
                        getattr(self, group_attr).update_time_list(self.worktimes)

                self.ot_times = config.get('ot_times', [])

                # [CLEANUP] Clean up users list from shift markers
                import re
                cleaned_users = []
                for user in self.users:
                    if re.match(r"^\((주간|야간|휴일)\)$", user.strip()):
                        continue
                    match = re.match(r"\((주간|야간|휴일)\)\s*(.*)", user.strip())
                    if match:
                        actual_name = match.group(2).strip()
                        if actual_name and actual_name not in cleaned_users:
                            cleaned_users.append(actual_name)
                    elif user.strip() and user.strip() not in cleaned_users:
                        cleaned_users.append(user.strip())
                
                if len(cleaned_users) != len(self.users):
                    self.users[:] = cleaned_users
                    # Don't save yet, let the normal flow handle it
                
                # Restore history visibility if saved
                # [FIX] Mismatch between 'history_visible_cols' and 'daily_usage_visible_cols'
                self.manual_visible_cols = config.get('daily_usage_visible_cols', config.get('history_visible_cols', []))
                
                # [MIGRATION] If 'OT합계' is in saved cols, replace with 'OT시간', 'OT금액'
                if 'OT합계' in self.manual_visible_cols:
                    index = self.manual_visible_cols.index('OT합계')
                    self.manual_visible_cols.pop(index)
                    if 'OT시간' not in self.manual_visible_cols:
                        self.manual_visible_cols.insert(index, 'OT시간')
                    if 'OT금액' not in self.manual_visible_cols:
                        self.manual_visible_cols.insert(index + 1, 'OT금액')
                
                # [MIGRATION] Translate English names to Korean
                translations = {
                    'Date': '날짜', 'Site': '현장', 'User': '작업자', 'WorkTime': '작업시간',
                    'Equipment': '장비명', 'Method': '검사방법', 'Remark': '비고', 
                    'Note': '비고', 'EntryTime': '입력시간', 'MaterialName': '품목명',
                    '제경비율': '제경비', '기술료율': '기술료',
                    '차량번호': '차량번호', '주행거리': '주행거리', '차량점검': '차량점검', '차량비고': '차량비고'
                }
                for i, col in enumerate(self.manual_visible_cols):
                    if col in translations:
                        self.manual_visible_cols[i] = translations[col]
                
                # [CLEANUP] Remove noisy NDT calculation breakdown columns and noise columns from display list
                for nc in ['작업형태', '조건1', '조건2', '보정계수', '제경비', '기술료', '환산물량', '재료비', '인건비', '적용코드', '현장']:
                    while nc in self.manual_visible_cols:
                        self.manual_visible_cols.remove(nc)

                # Ensure preferred columns in clean order
                preferred_order = ['날짜', '구분', '검사방법', '품목명', '수량', '단위', '단가', '검사비', '작업자', '작업시간', 'OT금액', '장비명', '업체명', '비고']
                if not self.manual_visible_cols or any(ot in self.manual_visible_cols for ot in ['OT1', 'OT2', 'OT3']):
                    self.manual_visible_cols = list(preferred_order)
                else:
                    if '품목명' not in self.manual_visible_cols:
                        if '검사방법' in self.manual_visible_cols:
                            idx = self.manual_visible_cols.index('검사방법')
                            self.manual_visible_cols.insert(idx + 1, '품목명')
                        elif '수량' in self.manual_visible_cols:
                            idx = self.manual_visible_cols.index('수량')
                            self.manual_visible_cols.insert(idx, '품목명')
                        else:
                            self.manual_visible_cols.append('품목명')
                    for pc in preferred_order:
                        if pc not in self.manual_visible_cols:
                            self.manual_visible_cols.append(pc)
                
                if hasattr(self, 'daily_usage_tree'):
                    try:
                        self.daily_usage_tree['displaycolumns'] = [c for c in self.manual_visible_cols if c in self.daily_usage_tree['columns']]
                    except: pass
                
                # Restore monthly history visibility if saved
                self.monthly_visible_cols = config.get('monthly_visible_cols', [])
                if self.monthly_visible_cols and hasattr(self, 'monthly_usage_tree'):
                    try:
                        self.monthly_usage_tree['displaycolumns'] = [c for c in self.monthly_visible_cols if c in self.monthly_usage_tree['columns']]
                    except: pass

                # Restore budget site performance view visibility if saved (exclude billing breakdown details by default)
                raw_budget_cols = config.get('budget_view_visible_cols', [])
                self.budget_view_visible_cols = [c for c in raw_budget_cols if c not in {'청구인건비', '청구재료비', '제경비', '기술료'}]
                self.budget_view_heading_aliases = config.get('budget_view_heading_aliases', {})
                self.budget_view_custom_columns = config.get('budget_view_custom_columns', [])
                if hasattr(self, 'budget_view_tree'):
                    try:
                        self._refresh_budget_view_tree_columns(reload_data=False)
                    except:
                        pass
                if self.budget_view_visible_cols and hasattr(self, 'budget_view_tree'):
                    try:
                        visible = [c for c in self.budget_view_visible_cols if c in self.budget_view_tree['columns']]
                        self.budget_view_tree['displaycolumns'] = visible if visible else self.budget_view_tree['columns']
                    except:
                        pass
                
                # Restore warehouses list
                self.warehouses[:] = config.get('warehouses', [])
                if not self.warehouses and not self.materials_df.empty:
                    self.warehouses = sorted(self.materials_df['창고'].dropna().unique().tolist())
                    self.warehouses = [str(w).strip() for w in self.warehouses if str(w).strip()]
                
                self.equipments[:] = config.get('equipments', [])
                if not self.equipments and not self.daily_usage_df.empty and '장비명' in self.daily_usage_df.columns:
                    self.equipments = sorted(self.daily_usage_df['장비명'].dropna().unique().tolist())
                    self.equipments = [str(e).strip() for e in self.equipments if str(e).strip()]
                
                # Vehicle list
                from site_apps.central.src.vehicle_catalog import restore_vehicle_catalog
                self.vehicles[:] = restore_vehicle_catalog(config)
                # [FIX] Update combo box values after loading lists from config
                # Update sites combo boxes
                if hasattr(self, 'cb_trans_site'): 
                    self.cb_trans_site['values'] = sorted(self.sites)
                # Update warehouses combo boxes
                if hasattr(self, 'cb_warehouse'): 
                    self.cb_warehouse['values'] = sorted(self.warehouses)
                # Update users combo boxes
                if hasattr(self, 'ent_user'): 
                    self.ent_user['values'] = sorted(self.users)
                # Update vehicles for VehicleInspectionWidgets
                if hasattr(self, 'fixed_vehicle_widget'):
                    self.fixed_vehicle_widget.update_vehicle_list(sorted(self.vehicles))
                if hasattr(self, 'vehicle_boxes'):
                    for w in self.vehicle_boxes:
                        if hasattr(w, 'update_vehicle_list'):
                            w.update_vehicle_list(sorted(self.vehicles))
                if hasattr(self, 'vehicle_inspections'):
                    for w in self.vehicle_inspections.values():
                        if hasattr(w, 'update_vehicle_list'):
                            w.update_vehicle_list(sorted(self.vehicles))
                
                # Restore stock column widths
                stock_col_widths = config.get('stock_col_widths', {})
                if stock_col_widths and hasattr(self, 'stock_tree'):
                    for col, width in stock_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.stock_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass

                # Restore in/out history column widths
                inout_col_widths = config.get('inout_col_widths', {})
                if inout_col_widths and hasattr(self, 'inout_tree'):
                    for col, width in inout_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.inout_tree.column(col, width=w, minwidth=50, stretch=False)
                        except:
                            pass
                
                # Restore recent entries column widths
                tv_recent_col_widths = config.get('tv_recent_col_widths', {})
                if tv_recent_col_widths and hasattr(self, 'tv_recent'):
                    for col, width in tv_recent_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.tv_recent.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass

                # Restore daily usage column widths
                daily_usage_col_widths = config.get('daily_usage_col_widths', {})
                if daily_usage_col_widths and hasattr(self, 'daily_usage_tree'):
                    for col, width in daily_usage_col_widths.items():
                        try:
                            # Only apply if width is reasonable (e.g. > 10)
                            w = int(width)
                            
                            # All columns are now user-resizable
                            if w > 10:
                                # Enforce minimums for high-precision cols
                                if col == '날짜': w = max(w, 160)
                                elif col == '입력시간': w = max(w, 300)
                                self.daily_usage_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass

                # Restore monthly usage column widths
                monthly_usage_col_widths = config.get('monthly_usage_col_widths', {})
                if monthly_usage_col_widths and hasattr(self, 'monthly_usage_tree'):
                    for col, width in monthly_usage_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.monthly_usage_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass
                
                # [NEW] Restore site summary column widths
                site_summary_col_widths = config.get('site_summary_col_widths', {})
                if site_summary_col_widths and hasattr(self, 'site_summary_tree'):
                    for col, width in site_summary_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.site_summary_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass

                # Restore budget site performance column widths
                budget_view_col_widths = config.get('budget_view_col_widths', {})
                if budget_view_col_widths and hasattr(self, 'budget_view_tree'):
                    for col, width in budget_view_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.budget_view_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass

                # [NEW] Restore worker summary column widths
                worker_summary_col_widths = config.get('worker_summary_col_widths', {})
                if worker_summary_col_widths and hasattr(self, 'worker_summary_tree'):
                    for col, width in worker_summary_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.worker_summary_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass
                # [NEW] Restore report column widths
                report_col_widths = config.get('report_col_widths', {})
                if report_col_widths and hasattr(self, 'report_tree'):
                    for col, width in report_col_widths.items():
                        try:
                            w = int(width)
                            if w > 10:
                                self.report_tree.column(col, width=w, minwidth=20, stretch=False)
                        except:
                            pass
                
                if hasattr(self, '_loading_memos'):
                    del self._loading_memos
                    
                # Restore layout lock state
                if hasattr(self, 'btn_lock_layout'):
                    if self.layout_locked:
                        self.btn_lock_layout.config(text="🔒 배치 고정됨")
                        self.style.configure("Lock.TButton", foreground="black")
                    else:
                        self.btn_lock_layout.config(text="🔓 배치 수정 중")
                        self.style.configure("Lock.TButton", foreground="red")

                if config.get('daily_usage_sash_locked', False):
                    self.daily_usage_sash_locked = True
                    
                    if hasattr(self, 'btn_sash_lock'):
                        self.btn_sash_lock.config(text="🔒 경계 잠금됨")
                    
                    print("LOADED: Restored sash lock state")

                # Tab selection handled already at line 4531
                
                # Recreate Memos and Clones first (these must exist before they can be placed)
                self._loading_memos = []
                # Map class names to actual classes for recreation
                from site_apps.central.src.views.components import VehicleInspectionWidget
                class_map = {'Entry': ttk.Entry, 'Combobox': ttk.Combobox, 'VehicleInspectionWidget': VehicleInspectionWidget}
                
                draggable_geos = config.get('draggable_geometries', {})
                for key, geo in draggable_geos.items():
                    if key.startswith('clone_'):
                        self._loading_memos.append(key)
                        cls_name = geo.get('widget_class_name', 'Entry')
                        cls = class_map.get(cls_name, ttk.Entry)
                        label = geo.get('custom_label', "복제항목")
                        kwargs = geo.get('widget_kwargs', {})
                        m_list_key = geo.get('manage_list_key')
                        cont, w = self.create_draggable_container(self.entry_inner_frame, label, cls, key, manage_list_key=m_list_key, **kwargs)
                    # [UI REVISION] Legacy floating boxes (memo, checklist, vehicle) are no longer recreated 
                    # because they are now permanently embedded in the bottom dashboard.

                # 2. DELAYED RESTORATION: Restore complex states after UI mapped
                def delayed_restore():
                    self.root.update_idletasks()
                    
                    # 1. Restore window state and size
                    try:
                        if config.get('window_state'):
                            if config['window_state'] == 'zoomed':
                                self.root.state('zoomed')
                            else:
                                try:
                                    self.root.state(config['window_state'])
                                except: pass
                        
                        w = config.get('window_width')
                        h = config.get('window_height')
                        if w and h:
                            # Only set geometry if not zoomed
                            if self.root.state() != 'zoomed':
                                self.root.geometry(f"{w}x{h}")
                    except: pass

                    # 2. Restore resolution lock if active
                    if self.resolution_locked:
                        self.root.resizable(False, False)
                        if hasattr(self, 'btn_resolution_lock'):
                            self.btn_resolution_lock.config(text="🔒 해상도 고정됨")

                    # 3. Restore draggable components regardless of lock state
                    # [FIX] Always restore positions if they exist, so movement is persisted 
                    # even if the app was closed while 'unlocked'.
                    draggable_geos = config.get('draggable_geometries', {})
                    for key, geo in draggable_geos.items():
                        if key in self.draggable_items:
                            widget = self.draggable_items[key]
                            try:
                                # [LAYOUT FIX] 핵심 박스는 저장 좌표를 무시하고 기본 grid 배치 사용
                                if key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
                                    continue

                                if geo.get('hidden'):
                                    # [STABILITY] If it's a core widget, ignore hidden status and reset it
                                    if key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
                                        print(f"LOADER: Recovering hidden core widget: {key}")
                                        self.reset_widget_position(None, widget=widget)
                                        continue
                                        
                                    if widget.winfo_manager() == 'grid':
                                        widget.grid_forget()
                                    elif widget.winfo_manager() == 'place':
                                        widget.place_forget()
                                    continue
                                
                                if geo.get('custom_label') and hasattr(widget, '_label_widget'):
                                    widget._label_widget.config(text=geo['custom_label'])
                                
                                # [FIX] Explicitly un-grid before using place manager
                                if widget.winfo_manager() == 'grid':
                                    if not hasattr(widget, '_original_grid_info'):
                                        widget._original_grid_info = widget.grid_info()
                                    widget.grid_forget()
                                
                                self._ensure_placeholder(widget, width=geo.get('width'), height=geo.get('height'))
                                widget.lift()
                                
                                # [FIX] Enforce minimum dimensions to prevent invisible widgets
                                load_x = geo.get('x', 0)
                                load_y = geo.get('y', 0)
                                load_w = max(100, geo.get('width', 200))
                                load_h = max(50, geo.get('height', 100))

                                # [RECOVERY] 작업자 박스는 화면 밖/과소 복원을 방지
                                if key == 'workers_box_geometry':
                                    try:
                                        p_w = widget.master.winfo_width()
                                        p_h = widget.master.winfo_height()
                                        load_w = max(520, load_w)
                                        load_h = max(360, load_h)
                                        if p_w > 0:
                                            load_x = max(0, min(int(load_x), max(0, p_w - load_w)))
                                        else:
                                            load_x = max(0, int(load_x))
                                        if p_h > 0:
                                            load_y = max(0, min(int(load_y), max(0, p_h - load_h)))
                                        else:
                                            load_y = max(0, int(load_y))
                                    except Exception:
                                        load_x = max(0, int(load_x))
                                        load_y = max(0, int(load_y))
                                
                                widget.place(x=load_x, y=load_y, width=load_w, height=load_h)

                                # [VISUAL FIX] 핵심 박스는 기존 placeholder 잔상 제거 + 최소 필요 높이 보정
                                if key in ('ndt_usage_box_geometry', 'rtk_usage_box_geometry', 'workers_box_geometry'):
                                    try:
                                        self._remove_placeholder(widget)
                                    except Exception:
                                        pass
                                    try:
                                        widget.update_idletasks()
                                        if key == 'workers_box_geometry':
                                            min_w = max(520, widget.winfo_reqwidth() + 8)
                                            min_h = max(360, widget.winfo_reqheight() + 12)
                                        else:
                                            min_w = max(160, widget.winfo_reqwidth() + 4)
                                            min_h = max(90, widget.winfo_reqheight() + 6)
                                        cur_x = int(float(widget.place_info().get('x', load_x)))
                                        cur_y = int(float(widget.place_info().get('y', load_y)))
                                        cur_w = int(float(widget.place_info().get('width', load_w)))
                                        cur_h = int(float(widget.place_info().get('height', load_h)))
                                        if cur_w < min_w or cur_h < min_h:
                                            widget.place(x=cur_x, y=cur_y, width=max(cur_w, min_w), height=max(cur_h, min_h))
                                        if key == 'workers_box_geometry':
                                            widget.lift()
                                    except Exception:
                                        pass
                            except Exception as e:
                                print(f"Error placing widget {key}: {e}")
                                
                    # 3.5 [LAYOUT FIX] 핵심 박스는 항상 기본 grid 위치로 복구
                    for key in getattr(self, 'CORE_DRAGGABLE_KEYS', []):
                        if key not in self.draggable_items:
                            continue
                        widget = self.draggable_items[key]
                        try:
                            self.reset_widget_position(None, widget=widget)
                            self._remove_placeholder(widget)
                            widget.lift()
                        except Exception as e:
                            print(f"LOADER: Failed to reset core widget {key}: {e}")

                    # 3.8 [SAFEGUARD] 저장 버튼 강제 복구 (화면 밖/숨김/겹침 방지)
                    if 'save_btn_geometry' in self.draggable_items:
                        sbox = self.draggable_items['save_btn_geometry']
                        try:
                            sbox.update_idletasks()
                            smgr = sbox.winfo_manager()
                            sw = sbox.winfo_width()
                            sh = sbox.winfo_height()

                            # Hidden/unmanaged or too small => reset to grid default first
                            if smgr == '' or sw < 80 or sh < 20:
                                self.reset_widget_position(None, widget=sbox)
                                smgr = sbox.winfo_manager()

                            # If placed, clamp into visible area and keep above siblings
                            if smgr == 'place':
                                pi = sbox.place_info()
                                sx = int(float(pi.get('x', 10)))
                                sy = int(float(pi.get('y', 340)))
                                sw = int(float(pi.get('width', max(120, sbox.winfo_reqwidth() or 120))))
                                sh = int(float(pi.get('height', max(28, sbox.winfo_reqheight() or 28))))
                                p_w = sbox.master.winfo_width()
                                p_h = sbox.master.winfo_height()
                                sw = max(120, sw)
                                sh = max(28, sh)
                                if p_w > 0:
                                    sx = max(0, min(sx, max(0, p_w - sw)))
                                if p_h > 0:
                                    sy = max(0, min(sy, max(0, p_h - sh)))
                                sbox.place(x=sx, y=sy, width=sw, height=sh)

                            sbox.lift()
                        except Exception:
                            pass

                    # 4. Restore sash positions
                    def apply_sashes():
                        try:
                            # Daily usage sash
                            daily_sash = config.get('daily_usage_sash_pos')
                            if daily_sash is not None and hasattr(self, 'daily_usage_paned'):
                                total_h = self.daily_usage_paned.winfo_height()
                                if total_h > 0:
                                    min_pos = max(int(total_h * 0.1), 460)
                                    max_pos = int(total_h * 0.9)
                                    safe_pos = int(daily_sash)
                                    if safe_pos < min_pos:
                                        safe_pos = min_pos
                                    if safe_pos > max_pos:
                                        safe_pos = max_pos
                                    getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, safe_pos)
                                else:
                                    getattr(self.daily_usage_paned, "sashpos", lambda *args: 500)(0, int(daily_sash))
                            
                            # History sash
                            history_sash = config.get('daily_history_sash_pos')
                            if history_sash is not None and hasattr(self, 'daily_history_paned'):
                                self.daily_history_paned.sashpos(0, int(history_sash))
                                
                            # Bottom dashboard sashes
                            bottom_sashes = config.get('bottom_dashboard_sashes', [])
                            if bottom_sashes and len(bottom_sashes) >= 2 and hasattr(self, 'bottom_dashboard') and isinstance(self.bottom_dashboard, ttk.PanedWindow):
                                try:
                                    # Ensure widget is fully drawn before setting sashes
                                    self.bottom_dashboard.update_idletasks()
                                    total_w = self.bottom_dashboard.winfo_width()
                                    
                                    if total_w > 300: # Make sure window is mapped and has a reasonable width
                                        # Clamp sash 1: at least 100px from left, at most 200px from right edge
                                        s1 = max(100, min(int(bottom_sashes[0]), total_w - 200))
                                        # Clamp sash 2: at least 100px from sash 1, at most 50px from right edge
                                        s2 = max(s1 + 100, min(int(bottom_sashes[1]), total_w - 50))
                                        
                                        self.bottom_dashboard.sashpos(0, s1)
                                        self.bottom_dashboard.sashpos(1, s2)
                                    else:
                                        self.bottom_dashboard.sashpos(0, int(bottom_sashes[0]))
                                        self.bottom_dashboard.sashpos(1, int(bottom_sashes[1]))
                                except: pass
                                
                            # Monthly usage table sashes (월별 집계 3개 테이블 높이 복원)
                            self._ensure_monthly_sash_visibility()

                            # Daily Work Log left/right frame sash & visibility restoration
                            work_log = getattr(self, 'tab_daily_work_log', None)
                            if work_log is not None and hasattr(work_log, 'restore_work_log_layout'):
                                work_log.restore_work_log_layout()

                            # If sash lock is active (check config heavily)
                            config_locked = config.get('daily_usage_sash_locked', False)
                            if config_locked:
                                self.daily_usage_sash_locked = True
                                if hasattr(self, 'btn_sash_lock'):
                                    self.btn_sash_lock.config(text="🔒 경계 고정됨")
                                    self.btn_sash_lock.configure(style="SashLock.TButton")
                                self._start_sash_monitor()
                            elif self.daily_usage_sash_locked: # If self says locked but config didn't (fallback)
                                self._start_sash_monitor()
                        except: pass

                    self.root.after(100, apply_sashes)
                    self.root.after(800, apply_sashes)

                    # 5. Refresh canvas scroll region
                    if hasattr(self, 'entry_inner_frame'):
                        self._ensure_canvas_scroll_region()
                    
                    # 6. Final UI refresh
                    for l_key in ['users', 'sites', 'equipments', 'warehouses', 'worktimes']:
                        self.refresh_ui_for_list_change(l_key)

                    # 7. Mark as ready for future saves AFTER all restoration is done
                    def finalize_loading():
                        try:
                            print("[STARTUP] Finalizing loading...")
                            self.is_ready = True
                            
                            # [FIX] Forcibly show and deiconify window to prevent "closing downwards" or disappearing
                            print("[STARTUP] Ensuring window visibility...")
                            try:
                                self.root.deiconify()
                                self.root.lift()
                                self.root.focus_force()
                            except: pass
                            
                            # [V13_RESET_ON_STARTUP] Ensure Daily Usage form starts BLANK as requested
                            print("[STARTUP] Resetting forms...")
                            try:
                                self.clear_daily_usage_form_all()
                            except: pass
                            
                            # [PERF] 활성 탭만 필요한 뷰를 갱신하고 비활성 탭의 불필요한 중복 렌더링 방지
                            print("[STARTUP] Refreshing views...")
                            try:
                                self._monthly_view_dirty = True
                                active_tab = self.notebook.select()
                                active_title = self.notebook.tab(active_tab, "text") if active_tab else ""
                                if active_title == '현장별 일일 사용량 기입':
                                    self.root.after(80, self.update_daily_usage_view)
                                elif active_title == '월별 집계':
                                    self.root.after(80, self.update_monthly_usage_view)
                                elif active_title == '공사실행예산서':
                                    self.root.after(80, self.update_budget_site_view)
                            except Exception as e:
                                print(f"[STARTUP] View refresh error: {e}")
                                
                            print("APP READY: State restoration complete.")
                        except Exception as startup_err:
                            print(f"[CRITICAL STARTUP ERROR] {startup_err}")
                            messagebox.showerror("Startup Error", f"An error occurred during startup: {startup_err}\n\n{traceback.format_exc()}")
                    
                    print("[STARTUP] Scheduling finalization...")
                    self.root.after(500, finalize_loading) 
                
                # Execute delayed restoration
                try:
                    self.root.after(300, delayed_restore)
                except Exception as outer_err:
                    messagebox.showerror("Initialization Error", f"Critical error during initialization: {outer_err}")

        except Exception as e:
            print(f"Failed to load tab config: {e}")
            # Ensure we eventually become ready even if load failed, to allow new saves
            def _mark_ready():
                try:
                    if hasattr(self, 'root') and self.root.winfo_exists():
                        self.is_ready = True
                except tk.TclError:
                    pass
            try:
                if hasattr(self, 'root') and self.root.winfo_exists():
                    self.root.after(2000, _mark_ready)
            except tk.TclError:
                pass
    
