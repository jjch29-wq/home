"""경인사무소 실행기."""
from pathlib import Path
import os
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

# Set Gyeongin-specific environment / data paths
os.environ["PMI_SITE"] = "gyeongin"

from site_apps.central.src.app import MaterialManager

class GyeonginMaterialManager(MaterialManager):
    def __init__(self, root):
        # Override title before super().__init__ finishes
        super().__init__(root)
        self.root.title("자재 및 소모품 관리 시스템 - 경인사무소 (Material Manager)")
        
        # Override DB path to Gyeongin data
        cloud_p = os.path.join(r"G:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data", "Material_Inventory.xlsx")
        local_p = str(Path(__file__).resolve().parent / "data" / "Material_Inventory.xlsx")
        if os.path.exists(cloud_p):
            self.db_path = cloud_p
        elif os.path.exists(local_p):
            self.db_path = local_p

        # Override config path
        cfg_p = str(Path(__file__).resolve().parent / "src" / "Material_Manager_Config.json")
        if os.path.exists(cfg_p):
            self.config_path = cfg_p

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

