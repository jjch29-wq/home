import re
import os

codebook_path = r'c:\Users\jjch2\Desktop\PMI\home\src\코드집_사전.py'
hub_path = r'c:\Users\jjch2\Desktop\PMI\home\src\절차서_수정_통합허브.py'
output_path = r'c:\Users\jjch2\Desktop\PMI\home\src\통합_마스터_허브.py'

with open(codebook_path, 'r', encoding='utf-8') as f:
    codebook_code = f.read()

with open(hub_path, 'r', encoding='utf-8') as f:
    hub_code = f.read()

# Extract everything inside class ProcedureHubApp
hub_methods = re.search(r'class ProcedureHubApp:.*?(?=if __name__ ==)', hub_code, re.DOTALL).group(0)

# Replace tab variables
hub_methods = hub_methods.replace('self.tab1', 'self.tab_batch')
hub_methods = hub_methods.replace('self.tab2', 'self.tab_compare')
hub_methods = hub_methods.replace('create_widgets_tab1', 'create_batch_widgets')
hub_methods = hub_methods.replace('create_widgets_tab2', 'create_compare_widgets')

# Extract methods to inject
methods_to_inject = re.sub(r'class ProcedureHubApp:.*?def create_batch_widgets', '    def create_batch_widgets', hub_methods, flags=re.DOTALL)

# Modify CodebookApp
codebook_code = codebook_code.replace('self.root.title("코드집 사전 / 규격 관리")', 'self.root.title("🔥 PAUT 절차서 통합 허브 (뷰어/코드관리/일괄수정/비교) 🔥")')
codebook_code = codebook_code.replace('self.root.geometry("1200x800")', 'self.root.geometry("1400x900")')

init_addition = """
        self.tab_batch = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_batch, text="✍️ 다중 일괄 변환 및 편집")
        
        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text="🔍 개정 내용 비교 (구버전 vs 신버전)")
        
        self.create_batch_widgets()
        self.create_compare_widgets()
"""

codebook_code = codebook_code.replace('self.create_viewer_widgets()', 'self.create_viewer_widgets()' + init_addition)

# Both apps use entry_find and entry_replace and tree. We need to isolate batch ones.
methods_to_inject = methods_to_inject.replace('self.entry_find', 'self.batch_entry_find')
methods_to_inject = methods_to_inject.replace('self.entry_replace', 'self.batch_entry_replace')
methods_to_inject = methods_to_inject.replace('self.tree', 'self.batch_tree')

methods_to_inject = methods_to_inject.replace('def add_item(self', 'def batch_add_item(self')
methods_to_inject = methods_to_inject.replace('def update_item(self', 'def batch_update_item(self')
methods_to_inject = methods_to_inject.replace('def delete_item(self', 'def batch_delete_item(self')
methods_to_inject = methods_to_inject.replace('def on_tree_select(self', 'def batch_on_tree_select(self')

methods_to_inject = methods_to_inject.replace('command=self.add_item', 'command=self.batch_add_item')
methods_to_inject = methods_to_inject.replace('command=self.update_item', 'command=self.batch_update_item')
methods_to_inject = methods_to_inject.replace('command=self.delete_item', 'command=self.batch_delete_item')
methods_to_inject = methods_to_inject.replace('"<<TreeviewSelect>>", self.on_tree_select', '"<<TreeviewSelect>>", self.batch_on_tree_select')

codebook_code = codebook_code.replace('if __name__ == "__main__":', methods_to_inject + '\nif __name__ == "__main__":')

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(codebook_code)
