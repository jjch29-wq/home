import pandas as pd
import os

class Tester:
    def __init__(self):
        self.carestream_films = [
            "Carestream AA400-3⅓*12\"", "Carestream AA400-3⅓*17\"", 
            "Carestream AA400-4½*12\"", "Carestream AA400-10*12\"",
            "Carestream M100-3⅓*12\"", "Carestream M100-10*12\"",
            "Carestream M100-14*17\"", "Carestream MX125-3⅓*6\"",
            "Carestream MX125-3⅓*12\"", "Carestream MX125-4½*12\"",
            "Carestream MX125-10*12\"", "Carestream MX125-14*17\"",
            "Carestream T200-3⅓*12\"", "Carestream T200-3⅓*17\"",
            "Carestream T200-4½*12\"", "Carestream T200-10*12\"",
            "Carestream T200-14*17\""
        ]

    def is_consumable(self, name):
        if not name or pd.isna(name): return False
        name_upper = str(name).upper().strip()
        if any(k in name_upper for k in ["PAUT", "UT", "MPI", "PMI", "SCANNER", "WEDGE", "PROBE", "CABLE", "장비", "본체"]):
            return False
        rt_keywords = ["필름", "FILM", "CARESTREAM", "FUJIFILM", "AGFA", "KODAK", "AA400", "M100", "MX125", "T200", "HS800"]
        if any(k in name_upper for k in rt_keywords): return True
        ndt_keywords = ["자분", "페인트", "침투제", "세척제", "현상제", "CHEMICAL", "DEVELOPER", "CLEANER", "PENETRANT", "SM-15", "MP-35", "MEGA-CHECK"]
        if any(k in name_upper for k in ndt_keywords): return True
        ndt_list = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        if any(m in name_upper for m in ndt_list): return True
        return False

t = Tester()
file_path = 'data/Material_Inventory.xlsx'
if os.path.exists(file_path):
    excel_file = pd.ExcelFile(file_path)
    sheets = {name: excel_file.parse(name) for name in excel_file.sheet_names}
    if 'Materials' in sheets:
        df = sheets['Materials']
        count = 0
        for idx, row in df.iterrows():
            name = row.get('품목명', '')
            if not t.is_consumable(name):
                if df.at[idx, 'Active'] != 0:
                    df.at[idx, 'Active'] = 0
                    count += 1
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for name, sheet_df in sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
        print(f"Deactivated {count} non-consumable items (Active=0).")
    else:
        print("Materials sheet not found.")
else:
    print(f"File not found: {file_path}")
