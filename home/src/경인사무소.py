"""경인사무소 실행기. Edit site_apps/gyeongin for this site."""
from pathlib import Path
import sys

_src = str(Path(__file__).resolve().parent)
if _src not in sys.path:
    sys.path.insert(0, _src)

from site_apps.gyeongin.main import GyeonginMaterialManager

def main():
    import tkinter as tk
    root = tk.Tk()
    app = GyeonginMaterialManager(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        try:
            root.destroy()
        except Exception:
            pass

if __name__ == "__main__":
    main()

