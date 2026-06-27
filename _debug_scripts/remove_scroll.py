import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Pattern for PMI
pattern_pmi = re.compile(r"""        # \[LEFT\] Settings & Actions Pane \(Scrollable Wrapper\)
        left_pane_outer = tk.Frame\(self.pmi_paned, background="#f9fafb"\)
        self.pmi_paned.add\(left_pane_outer, stretch="always"\) 

        left_canvas = tk.Canvas\(left_pane_outer, background="#f9fafb", highlightthickness=0\)
        left_vsb = ttk.Scrollbar\(left_pane_outer, orient="vertical", command=left_canvas.yview\)
        left_pane = tk.Frame\(left_canvas, background="#f9fafb", padx=10, pady=10\)
        left_pane.bind\("<Configure>", lambda e: left_canvas.configure\(scrollregion=left_canvas.bbox\("all"\)\)\)
        canvas_window = left_canvas.create_window\(\(0, 0\), window=left_pane, anchor="nw"\)
        
        # \[CRITICAL\] Strictly follow canvas width to prevent "Hiding" \(Clipping\)
        def _on_canvas_configure\(e\):
            self.root.update_idletasks\(\)
            left_canvas.itemconfig\(canvas_window, width=e.width\)
        left_canvas.bind\("<Configure>", _on_canvas_configure\)
        left_canvas.configure\(yscrollcommand=left_vsb.set\)
        
        left_vsb.pack\(side='right', fill='y'\)
        left_canvas.pack\(side='left', fill='both', expand=True\)

        # Mousewheel scroll for settings
        def _left_scroll\(event\):
            left_canvas.yview_scroll\(int\(-1\*\(event.delta/120\)\), "units"\)
            return "break"
        left_canvas.bind\("<MouseWheel>", _left_scroll\)
        left_pane.bind\("<MouseWheel>", _left_scroll\)
        def _bind_left_mousewheel_recursive\(widget\):
            widget.bind\("<MouseWheel>", _left_scroll, add="\+"\)
            for child in widget.winfo_children\(\):
                _bind_left_mousewheel_recursive\(child\)
        left_pane.bind\("<Configure>", lambda e: _bind_left_mousewheel_recursive\(left_pane\), add="\+"\)\n""", re.MULTILINE)

replacement_pmi = """        # [LEFT] Settings & Actions Pane (Simplified Layout)
        left_pane = tk.Frame(self.pmi_paned, background="#f9fafb", padx=10, pady=10)
        self.pmi_paned.add(left_pane, stretch="always")\n"""

# Pattern for RT, PT, PAUT
# Note: They have slightly different bindings, varying prefix (rt_paned, pt_paned, paut_paned)
modes = ["rt", "pt", "paut"]
for m in modes:
    pat = re.compile(rf"""        # \[LEFT\] Settings Pane \(Scrollable Wrapper\)
        left_pane_outer = tk.Frame\(self.{m}_paned, background="#f9fafb"\)
        self.{m}_paned.add\(left_pane_outer, stretch="always"\)\s+left_canvas = tk.Canvas\(left_pane_outer, background="#f9fafb", highlightthickness=0\)
        left_vsb = ttk.Scrollbar\(left_pane_outer, orient="vertical", command=left_canvas.yview\)
        left_pane = tk.Frame\(left_canvas, background="#f9fafb", padx=10, pady=10\)\s+left_pane.bind\("<Configure>", lambda e: left_canvas.configure\(scrollregion=left_canvas.bbox\("all"\)\)\)
        canvas_window = left_canvas.create_window\(\(0, 0\), window=left_pane, anchor="nw"\)
        def _on_{m}_canvas_configure\(e\):
            self.root.update_idletasks\(\)
            left_canvas.itemconfig\(canvas_window, width=e.width\)
        left_canvas.bind\("<Configure>", _on_{m}_canvas_configure\)
        left_canvas.configure\(yscrollcommand=left_vsb.set\)\s+left_vsb.pack\(side='right', fill='y'\)
        left_canvas.pack\(side='left', fill='both', expand=True\)\s+# Mousewheel scroll
        def _{m}_left_scroll\(event\):
            left_canvas.yview_scroll\(int\(-1\*\(event.delta/120\)\), "units"\)
            return "break"
        left_canvas.bind\("<MouseWheel>", _{m}_left_scroll\)
""")
    rep = f"""        # [LEFT] Settings Pane (Simplified Layout)
        left_pane = tk.Frame(self.{m}_paned, background="#f9fafb", padx=10, pady=10)
        self.{m}_paned.add(left_pane, stretch="always")\n"""
    
    content, count = pat.subn(rep, content)
    print(f"{m} replacement count: {count}")

content, count_pmi = pattern_pmi.subn(replacement_pmi, content)
print(f"pmi replacement count: {count_pmi}")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
