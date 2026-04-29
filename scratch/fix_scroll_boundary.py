import re

file_path = 'src/Material-Master-Manager-V13.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Make rtk_grid a class attribute
content = content.replace(
    'rtk_grid = ttk.LabelFrame(self.master_form_panel, text="RTK 분류")',
    'self.rtk_grid = ttk.LabelFrame(self.master_form_panel, text="RTK 분류")'
)
content = content.replace(
    'for c in range(6): rtk_grid.columnconfigure(c, weight=1, uniform="ndt_rtk")',
    'for c in range(6): self.rtk_grid.columnconfigure(c, weight=1, uniform="ndt_rtk")'
)
content = content.replace(
    'for i, cat in enumerate(rtk_cats):\n            r = i // 3; col = (i % 3) * 2\n            ttk.Label(rtk_grid, text=f"{cat}:", font=(\'Arial\', 8)).grid(row=r, column=col, padx=1, pady=1, sticky=\'w\')\n            e = ttk.Entry(rtk_grid, width=6)',
    'for i, cat in enumerate(rtk_cats):\n            r = i // 3; col = (i % 3) * 2\n            ttk.Label(self.rtk_grid, text=f"{cat}:", font=(\'Arial\', 8)).grid(row=r, column=col, padx=1, pady=1, sticky=\'w\')\n            e = ttk.Entry(self.rtk_grid, width=6)'
)
content = content.replace(
    'e.grid(row=r, column=col+1, padx=1, pady=1, sticky=\'ew\')',
    'e.grid(row=r, column=col+1, padx=1, pady=1, sticky=\'ew\')'
)

# 2. Update _ensure_canvas_scroll_region to use self.rtk_grid bottom
new_ensure_scroll = """    def _ensure_canvas_scroll_region(self):
        \"\"\"Update canvas scroll region based on content height (stops at RTK bottom)\"\"\"
        try:
            if hasattr(self, 'entry_canvas') and self.entry_canvas:
                self.entry_canvas.update_idletasks()
                
                # Get max Y from all core elements
                max_y = 0
                
                # 1. Use the bottom of master_form_panel which contains form, NDT, RTK, and Workers
                if hasattr(self, 'master_form_panel'):
                    self.master_form_panel.update_idletasks()
                    panel_y = self.master_form_panel.winfo_y()
                    panel_h = self.master_form_panel.winfo_height()
                    max_y = max(max_y, panel_y + panel_h)
                
                # 2. Specifically check RTK bottom if requested by user
                if hasattr(self, 'rtk_grid'):
                    # rtk_grid is inside master_form_panel, so calculate relative to master_form_panel master
                    self.rtk_grid.update_idletasks()
                    rtk_bottom = self.rtk_grid.winfo_y() + self.rtk_grid.winfo_height()
                    # Add master_form_panel offset
                    if hasattr(self, 'master_form_panel'):
                        rtk_bottom += self.master_form_panel.winfo_y()
                    max_y = max(max_y, rtk_bottom)
                
                # 3. Handle draggable items (Memos, Checklists, etc.)
                for key, widget in self.draggable_items.items():
                    try:
                        if widget.winfo_manager() == 'place':
                            info = widget.place_info()
                            y = int(float(info.get('y', 0)))
                            h = int(float(info.get('height', widget.winfo_height())))
                            max_y = max(max_y, y + h)
                    except: pass

                # Final scroll height with minimal buffer
                scroll_h = max_y + 10
                scroll_w = max(1100, self.entry_inner_frame.winfo_width())
                
                self.entry_canvas.configure(scrollregion=(0, 0, scroll_w, scroll_h))
        except Exception as e:
            print(f"DEBUG: Scroll region update error: {e}")"""

# Find and replace the function
content = re.sub(r'def _ensure_canvas_scroll_region\(self\):.*?self\.entry_canvas\.configure\(scrollregion=\(0, 0, scroll_w, scroll_h\)\).*?else:.*?self\.entry_canvas\.configure\(scrollregion=\(0, 0, 2000, 2000\)\).*?except:.*?pass', new_ensure_scroll, content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print('RTK grid logic and scroll region updated.')
