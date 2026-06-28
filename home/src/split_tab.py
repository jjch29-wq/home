import re

with open('Material-Master-Manager-V14_20260627.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Add new tab creation
old_tab_creation = """        # Tab 6: Daily Usage Entry by Site
        self.tab_daily_usage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_daily_usage, text='현장별 일일 사용량 입력')
        self.setup_daily_usage_tab()"""

new_tab_creation = """        # Tab 6: Daily Usage Entry by Site
        self.tab_daily_usage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_daily_usage, text='현장별 일일 사용량 입력')
        self.setup_daily_usage_tab()
        
        # Tab 7: Daily Usage Query
        self.tab_daily_usage_query = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_daily_usage_query, text='현장 일일기록 조회 및 관리')
        self.setup_daily_usage_query_tab()"""

code = code.replace(old_tab_creation, new_tab_creation)

# 2. Change Panedwindow to Frame
old_paned = """        # Create PanedWindow for resizable frames
        self.daily_usage_paned = ttk.Panedwindow(self.tab_daily_usage, orient='vertical')
        self.daily_usage_paned.pack(fill='both', expand=True, padx=5, pady=5)  # Reduced padding
        
        # Save sash position on adjustment and lock it
        self.daily_usage_paned.bind("<ButtonRelease-1>", self._on_daily_usage_sash_changed)
        self.daily_usage_paned.bind("<Configure>", self._on_daily_usage_resize)
        # [FIX] Respect loaded config if available, otherwise default to False
        self.daily_usage_sash_locked = getattr(self, 'daily_usage_sash_locked', False)
        
        # Set initial sash position to ensure visibility (30% for top frame, 70% for bottom)
        self.daily_usage_paned.after(200, self._ensure_daily_usage_sash_visibility)
        self.daily_usage_paned.after(500, self._ensure_daily_usage_sash_visibility)
        self.daily_usage_paned.after(1000, self._ensure_daily_usage_sash_visibility)
        self.daily_usage_paned.after(1200, self._ensure_canvas_scroll_region)

        
        entry_frame = ttk.LabelFrame(self.daily_usage_paned, text="현장별 일일 사용량 기입")
        self.daily_usage_paned.add(entry_frame, weight=1) # Changed from weight=3 to weight=1"""

new_paned = """        # Main Frame (No longer PanedWindow since we separated tabs)
        self.daily_usage_paned = ttk.Frame(self.tab_daily_usage)
        self.daily_usage_paned.pack(fill='both', expand=True, padx=5, pady=5)  # Reduced padding
        
        self.daily_usage_sash_locked = getattr(self, 'daily_usage_sash_locked', False)
        
        entry_frame = ttk.LabelFrame(self.daily_usage_paned, text="현장별 일일 사용량 기입")
        entry_frame.pack(fill='both', expand=True) # Changed from add to pack"""

code = code.replace(old_paned, new_paned)

# 3. Disable Sash Lock Button
old_sash_btn = """        self.btn_sash_lock = ttk.Button(row1, text="🔒 경계 고정됨" if self.daily_usage_sash_locked else "🔓 경계 고정", command=self.toggle_sash_lock)
        self.btn_sash_lock.pack(side='right', padx=5)"""

new_sash_btn = """        # Sash lock button disabled since UI is separated
        # self.btn_sash_lock = ttk.Button(row1, text="🔒 경계 고정됨" if self.daily_usage_sash_locked else "🔓 경계 고정", command=self.toggle_sash_lock)
        # self.btn_sash_lock.pack(side='right', padx=5)"""

code = code.replace(old_sash_btn, new_sash_btn)

# 4. Find the display_frame block and move it to a new function
start_marker = '        display_frame = ttk.LabelFrame(self.daily_usage_paned, text="일일 사용량 기록 조회")'
end_marker = '    def _on_daily_usage_select(self, event):'

start_idx = code.find(start_marker)
end_idx = code.find(end_marker)

if start_idx != -1 and end_idx != -1:
    display_block = code[start_idx:end_idx]
    
    # We replace the old block with nothing
    code = code[:start_idx] + code[end_idx:]
    
    # Now we insert the new method definition
    display_block_modified = display_block.replace(
        'display_frame = ttk.LabelFrame(self.daily_usage_paned, text="일일 사용량 기록 조회")',
        'display_frame = ttk.Frame(self.tab_daily_usage_query)'
    ).replace(
        'self.daily_usage_paned.add(display_frame, weight=1) # Less weight for the list',
        'display_frame.pack(fill="both", expand=True, padx=5, pady=5)'
    )
    
    new_method = f'''
    def setup_daily_usage_query_tab(self):
        """Setup the daily usage query tab"""
{display_block_modified}
'''
    
    # Insert new_method before _on_daily_usage_select
    code = code[:code.find(end_marker)] + new_method + code[code.find(end_marker):]
else:
    print("Could not find display_frame markers!")

with open('Material-Master-Manager-V14_20260627.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Modifications done!")
