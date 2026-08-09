import os
import json
from datetime import datetime
import win32com.client

class MonthlyReportExporter:
    def __init__(self, history, target_month, template_path, doc_num="01"):
        """
        history: dict of daily logs
        target_month: 'YYYY-MM'
        template_path: path to the original .xls or .xlsx template
        """
        self.history = history
        self.target_month = target_month
        self.template_path = template_path
        self.doc_num = doc_num
        self.year, self.month = target_month.split('-')
        
        # Basic config loading
        self.config = self._load_config()
        self.expected_qtys = self.config.get("CONTRACT_QTY", {})
        
        self.replacements = {}

    def _load_config(self):
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    def _aggregate_data(self):
        # 1. Basic Text
        self.replacements['[[보고서_연월]]'] = f"{self.year}년 {self.month}월"
        self.replacements['[[보고서_월]]'] = f"{self.month}월"
        self.replacements['[[문서번호]]'] = self.doc_num
        self.replacements['[[작성일자]]'] = datetime.now().strftime("%Y. %m. %d.")
        self.replacements['[[계약명]]'] = "2026년 중앙지사 열수송관 비파괴검사용역 단가계약"
        self.replacements['[[지사명]]'] = "중앙지사"
        
        # Automatically replace old hardcoded strings so the user doesn't have to tag them everywhere
        self.replacements['2025년 동탄지사 열수송관 비파괴검사용역 단가계약'] = "2026년 중앙지사 열수송관 비파괴검사용역 단가계약"
        self.replacements['2025년 동탄지사 열배관  비파괴검사용역 단가계약'] = "2026년 중앙지사 열수송관 비파괴검사용역 단가계약"
        self.replacements['2025년 동탄지사 열배관 비파괴검사용역 단가계약'] = "2026년 중앙지사 열수송관 비파괴검사용역 단가계약"
        self.replacements['동 탄 지 사'] = "중 앙 지 사"
        self.replacements['분 당 사 업 소'] = "중 앙 지 사"
        
        # 2. Personnel
        total_p = 0
        for date_key, log in self.history.items():
            if date_key.startswith(self.target_month):
                p_data = log.get('personnel_data', {})
                day_total = sum(int(v) if str(v).isdigit() else 0 for k,v in p_data.items() if '계' not in k)
                total_p = max(total_p, day_total)
        self.replacements['[[당월_최대_인원]]'] = str(total_p)
        
        # 3. NDT Quantities
        stats = {'PAUT': [0,0,0,0], 'RT': [0,0,0,0], 'MT': [0,0,0,0], 'PT': [0,0,0,0]}
        for key, val in self.expected_qtys.items():
            if 'PAUT' in key: stats['PAUT'][0] += int(val)
            elif 'RT' in key: stats['RT'][0] += int(val)
            elif 'MT' in key: stats['MT'][0] += int(val)
            elif 'PT' in key: stats['PT'][0] += int(val)
            
        for date_key, log in self.history.items():
            try:
                dt = datetime.strptime(date_key, "%Y-%m-%d")
            except:
                continue
                
            is_curr_month = date_key.startswith(self.target_month)
            is_prev_month = dt.strftime("%Y-%m") < self.target_month
            
            for ndt in log.get('ndt_results', []):
                for method in ['PAUT', 'RT', 'MT', 'PT']:
                    val = ndt.get(method, '0')
                    if not str(val).isdigit(): val = '0'
                    v = int(val)
                    if v > 0:
                        if is_prev_month: stats[method][1] += v
                        if is_curr_month: stats[method][2] += v
                        stats[method][3] = stats[method][1] + stats[method][2]
                        
        for method in ['PAUT', 'RT', 'MT', 'PT']:
            self.replacements[f'[[{method}_도급물량]]'] = str(stats[method][0])
            self.replacements[f'[[{method}_전월누계]]'] = str(stats[method][1])
            self.replacements[f'[[{method}_당월물량]]'] = str(stats[method][2])
            self.replacements[f'[[{method}_총누계]]'] = str(stats[method][3])
            
        # 4. Defect Rates
        rt_or_curr, rt_re_curr = 0, 0
        rt_or_total, rt_re_total = 0, 0
        for date_key, log in self.history.items():
            is_curr_month = date_key.startswith(self.target_month)
            for ndt in log.get('ndt_results', []):
                v_or = int(ndt.get('RT_OR', 0) if str(ndt.get('RT_OR', '')).isdigit() else 0)
                v_re = int(ndt.get('RT_RE', 0) if str(ndt.get('RT_RE', '')).isdigit() else 0)
                try:
                    dt = datetime.strptime(date_key, "%Y-%m-%d")
                    if dt.strftime("%Y-%m") <= self.target_month:
                        rt_or_total += v_or
                        rt_re_total += v_re
                        if is_curr_month:
                            rt_or_curr += v_or
                            rt_re_curr += v_re
                except:
                    pass
                    
        curr_rate = round((rt_re_curr / rt_or_curr * 100) if rt_or_curr > 0 else 0, 2)
        total_rate = round((rt_re_total / rt_or_total * 100) if rt_or_total > 0 else 0, 2)
        
        self.replacements['[[당월_촬영매수]]'] = str(rt_or_curr)
        self.replacements['[[당월_불량매수]]'] = str(rt_re_curr)
        self.replacements['[[당월_불량률]]'] = f"{curr_rate}%"
        
        self.replacements['[[누계_촬영매수]]'] = str(rt_or_total)
        self.replacements['[[누계_불량매수]]'] = str(rt_re_total)
        self.replacements['[[누계_불량률]]'] = f"{total_rate}%"

    def generate(self, output_path):
        self._aggregate_data()
        
        try:
            import openpyxl
            wb = openpyxl.load_workbook(self.template_path)
            
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            new_val = cell.value
                            for tag, value in self.replacements.items():
                                new_val = new_val.replace(tag, str(value))
                            if new_val != cell.value:
                                cell.value = new_val
                                
            wb.save(output_path)
            wb.close()
            
        except Exception as e:
            print(f"Error generating report: {e}")
            raise