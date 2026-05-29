"""
요청서 합치기.py의 실제 merge_logic을 직접 호출하는 테스트
"""
import os, sys, types, importlib.util

# tkinter 완전 모킹
tk_mock = types.ModuleType('tkinter')
ttk_mock = types.ModuleType('tkinter.ttk')
fd_mock = types.ModuleType('tkinter.filedialog')
mb_mock = types.ModuleType('tkinter.messagebox')

class MockVar:
    def __init__(self, val=None): self._val = val
    def get(self): return self._val
    def set(self, v): self._val = v

tk_mock.BooleanVar = lambda value=False: MockVar(value)
tk_mock.StringVar = lambda value='': MockVar(value)
tk_mock.Tk = lambda: types.SimpleNamespace(configure=lambda **k: None, geometry=lambda s: None, title=lambda s: None, mainloop=lambda: None)
tk_mock.Label = lambda *a, **k: None
tk_mock.Button = lambda *a, **k: {'state': ''}
tk_mock.Frame = lambda *a, **k: types.SimpleNamespace(pack=lambda **k: None)
tk_mock.LabelFrame = lambda *a, **k: types.SimpleNamespace(pack=lambda **k: None)
tk_mock.Entry = lambda *a, **k: types.SimpleNamespace(pack=lambda **k: None)
tk_mock.Text = lambda *a, **k: types.SimpleNamespace(insert=lambda *a: None, see=lambda *a: None, delete=lambda *a: None, pack=lambda **k: None)
tk_mock.Checkbutton = lambda *a, **k: types.SimpleNamespace(pack=lambda **k: None)
tk_mock.BOTH = tk_mock.X = tk_mock.LEFT = tk_mock.RIGHT = tk_mock.W = tk_mock.END = tk_mock.FLAT = tk_mock.NORMAL = tk_mock.DISABLED = ''

ttk_mock.Style = lambda: types.SimpleNamespace(theme_use=lambda *a: None, configure=lambda *a, **k: None)
ttk_mock.Frame = lambda *a, **k: types.SimpleNamespace(pack=lambda **k: None)

mb_mock.showinfo = lambda *a, **k: None
mb_mock.showwarning = lambda *a, **k: None
mb_mock.showerror = lambda *a, **k: None
fd_mock.askdirectory = lambda: None
fd_mock.askopenfilenames = lambda **k: []

sys.modules['tkinter'] = tk_mock
sys.modules['tkinter.ttk'] = ttk_mock
sys.modules['tkinter.filedialog'] = fd_mock
sys.modules['tkinter.messagebox'] = mb_mock

# 프로그램 로드
spec = importlib.util.spec_from_file_location('merger', r'c:\Users\-\OneDrive\바탕 화면\home\요청서 합치기.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

FOLDER = r'C:\Users\-\OneDrive\바탕 화면\JHC RT'
KEYWORDS = 'No, Joint, Dwg, THK, Result, Date, Report No, Defect Rev'

# App 생성
class TestApp(mod.ExcelMergerApp):
    def __init__(self):
        self.selected_folder = FOLDER
        self.excel_files = [
            f for f in os.listdir(FOLDER)
            if f.endswith('.xlsx') and not f.startswith('~$') and 'Smart_Merged' not in f and 'BoxLabel' not in f
        ]
        self.keyword_var = MockVar(KEYWORDS)
        self.only_totals_var = MockVar(True)
        self.export_box_label_var = MockVar(True)
        self.btn_merge = {'state': ''}
        self.status_var = MockVar('')
        self.log_text = types.SimpleNamespace(insert=lambda *a: None, see=lambda *a: None, delete=lambda *a: None)
        self.root = types.SimpleNamespace(update_idletasks=lambda: None)

    def add_log(self, msg):
        if any(x in msg for x in ['ERROR', 'error', '오류', '완료', 'RT-0001', 'RT-0022', 'header', 'joint', '라벨']):
            try:
                print(msg.encode('cp949', errors='replace').decode('cp949'))
            except:
                pass

app = TestApp()
print(f"Files: {len(app.excel_files)}")
app.merge_logic()

