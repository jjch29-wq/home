import re

file_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\src\main.py'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Ensure Canvas window filling is robust
pattern1 = re.compile(r'def _on_canvas_configure\(event\):.*?self\.canvas\.bind\("<Configure>", _on_canvas_configure\)', re.DOTALL)
replacement1 = r'''def _on_canvas_configure(event):
            # Ensure the scrollable_frame window expands to the canvas width minus scrollbar
            # Using width=event.width ensures it fills horizontally
            self.canvas.itemconfigure(self.canvas_frame_window, width=event.width)
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.canvas.bind("<Configure>", _on_canvas_configure)'''

# Fix 2: Ensure mode_notebook and its tabs are fully expanded
# The previous code was:
# self.mode_notebook.pack(fill='both', expand=True, pady=(0, 10))
# This is correct, but let's re-verify and ensure pmi_mode_frame expansion inside setup.

# Fix 3: Force PanedWindow and its children to stretch
content = content.replace('pw.add(right_pane, stretch="always")', 'pw.add(right_pane, stretch="always", minsize=400)')

# Fix 4: Check if any parent of container in _setup_pmi_ui is constrained
# Standardize container packing in _setup_pmi_ui
new_pmi_setup_start = r'''
    def _setup_pmi_ui(self, parent):
        # [FORCE] Ensure parent (pmi_mode_frame) allows expansion
        container = tk.Frame(parent, background="#f9fafb")
        container.pack(fill='both', expand=True)

        # --- Dual Pane Layout (PanedWindow) ---
        # [REFINED] Use sticky for children
        pw = tk.PanedWindow(container, orient='horizontal', background="#d1d5db", 
                            sashwidth=6, sashpad=0, sashrelief='raised', borderwidth=0)
        pw.pack(fill='both', expand=True)
'''

# Apply pattern replacement for _setup_pmi_ui start
pattern2 = re.compile(r'def _setup_pmi_ui\(self, parent\):.*?container = tk\.Frame\(parent, background="#f9fafb"\)\s+container\.pack\(fill=\'both\', expand=True\)\s+# --- Dual Pane Layout \(PanedWindow\) ---.*?pw = tk\.PanedWindow\(container, orient=\'horizontal\', background="#d1d5db",.*?sashwidth=6, sashpad=0, sashrelief=\'raised\', borderwidth=0\)\s+pw\.pack\(fill=\'both\', expand=True\)', re.DOTALL)
content = pattern2.sub(new_pmi_setup_start.strip(), content)

# Fix 5: Ensure Treeview columns use stretch=True in all modes
# RT, PT, PAUT might need this to fill the space
content = content.replace("stretch=False", "stretch=True")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Horizontal expansion fixes applied.")
