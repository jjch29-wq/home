"""Compatibility launcher. Edit site_apps/central/src for this site."""
from pathlib import Path
import sys

_src = str(Path(__file__).resolve().parent)
if _src not in sys.path:
    sys.path.insert(0, _src)

from site_apps.central.src.app import MaterialManager
from site_apps.central.main import CentralMaterialManager

def main():
    import tkinter as tk
    root = tk.Tk()
    app = CentralMaterialManager(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        try:
            root.destroy()
        except Exception:
            pass

if __name__ == "__main__":
    main()
