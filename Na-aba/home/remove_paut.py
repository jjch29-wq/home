import os
import re

file_path = r"c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

pat = re.compile(r"""        # \[LEFT\] Settings Pane \(Scrollable Wrapper\)
        left_pane_outer = tk.Frame\(self.paut_paned, background="#f9fafb"\)
        self.paut_paned.add\(left_pane_outer, stretch="always"\)\s+left_canvas = tk.Canvas\(left_pane_outer, background="#f9fafb", highlightthickness=0\)
        left_vsb = ttk.Scrollbar\(left_pane_outer, orient="vertical", command=left_canvas.yview\)
        left_pane = tk.Frame\(left_canvas, background="#f9fafb", padx=10, pady=10\)\s+left_pane.bind\("<Configure>", lambda e: left_canvas.configure\(scrollregion=left_canvas.bbox\("all"\)\)\)
        canvas_window = left_canvas.create_window\(\(0, 0\), window=left_pane, anchor="nw"\)
        def _on_paut_canvas_configure\(e\):
            self.root.update_idletasks\(\)
            left_canvas.itemconfig\(canvas_window, width=e.width\)
        left_canvas.bind\("<Configure>", _on_paut_canvas_configure\)
        left_canvas.configure\(yscrollcommand=left_vsb.set\)\s+left_vsb.pack\(side='right', fill='y'\)
        left_canvas.pack\(side='left', fill='both', expand=True\)\s+# Mousewheel scroll
        def _paut_left_scroll\(event\):
            left_canvas.yview_scroll\(int\(-1\*\(event.delta/120\)\), "units"\)
            return "break"
        left_canvas.bind\("<MouseWheel>", _paut_left_scroll\)
""")
rep = """        # [LEFT] Settings Pane (Simplified Layout)
        left_pane = tk.Frame(self.paut_paned, background="#f9fafb", padx=10, pady=10)
        self.paut_paned.add(left_pane, stretch="always")\n"""

content, count = pat.subn(rep, content)
print(f"paut replacement count: {count}")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
