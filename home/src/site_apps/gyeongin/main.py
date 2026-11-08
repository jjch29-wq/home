"""경인사무소 실행기."""
from pathlib import Path
import os
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

# Set Gyeongin-specific environment / data paths
os.environ["PMI_SITE"] = "gyeongin"

from site_apps.gyeongin.src.app import MaterialManager

class GyeonginMaterialManager(MaterialManager):
    def __init__(self, root):
        super().__init__(root)
        self.root.title("자재 및 소모품 관리 시스템 - 경인사무소 (Material Manager)")

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

if __name__ == '__main__':
    main()

