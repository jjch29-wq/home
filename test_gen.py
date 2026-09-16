import sys, os
sys.path.append(r'C:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src')
from ndt import NdtReportGenerator
import openpyxl

class MockApp:
    def __init__(self):
        self.config = {'PT_END_ROW': '37', 'START_ROW': '10', 'DATA_END_ROW': '37'}
        self.project_data = {'Customer': 'Test'}
        self.current_mode = 'PT'
        self.progress = {'value': 0}
        self.mode_notebook = None
    def log(self, m): print(m)
    def update_idletasks(self): pass

app = NdtReportGenerator(None)
app.config = {'PT_END_ROW': '37', 'PT_START_ROW': '10'}
app.project_data = {'Customer': 'Test'}
app.current_mode = 'PT'
app.progress = {'value': 0}
app.mode_notebook = None
app.log = lambda x: print(x)
app.update_idletasks = lambda: None

final_list = []
for i in range(40):
    final_list.append({'Dwg': f'DWG-{i}', 'Joint': f'{i}', 'Result': 'Acc'})

template_path = r'C:\Users\jjch2\Desktop\PT-템플릿.xlsx'
try:
    app._run_pt_process(template_path, final_list)
except Exception as e:
    import traceback
    traceback.print_exc()