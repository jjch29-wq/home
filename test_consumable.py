import pandas as pd

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
        
        if "PAUT" in name_upper: return False
        
        if any(f.upper() in name_upper for f in getattr(self, 'carestream_films', [])): return True
        if "필름" in name_upper or "FILM" in name_upper: return True
        
        ndt_keywords = ["자분", "페인트", "침투제", "세척제", "현상제", "CHEMICAL", "DEVELOPER", "CLEANER", "PENETRANT"]
        if any(k in name_upper for k in ndt_keywords): return True
        
        ndt_list = ["형광자분", "흑색자분", "백색페인트", "침투제", "세척제", "현상제", "형광침투제"]
        if any(m in name_upper for m in ndt_list): return True
        
        return False

t = Tester()
test_names = ["PAUT SCANNER", "Carestream AA400", "흑색자분", "UT Probe", "Film", "자분"]
for name in test_names:
    print(f"{name}: {t.is_consumable(name)}")
