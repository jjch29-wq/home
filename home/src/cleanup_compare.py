import re

with open('c:/Users/jjch2/Desktop/PMI/home/src/코드절차서관리.py', encoding='utf-8') as f:
    code = f.read()

# Update title
code = code.replace('self.root.title("코드집 사전 (Codebook Manager)")', 'self.root.title("🔥 통합 절차서 규격 관리 및 일괄 개정 허브 🔥")')

# Remove tab_compare creation
code = re.sub(r'\s*self\.tab_compare = ttk\.Frame\(self\.notebook\)\s*self\.notebook\.add\(self\.tab_compare, text="🔍 개정 문서 비교기"\)\s*', '\n        ', code)

# Remove call to create_compare_widgets
code = code.replace('        self.create_compare_widgets()\n', '')

# Remove methods: create_compare_widgets, browse_comp_file, compare_docs
# Since they are at the very end of the file before if __name__ == '__main__':
code = re.sub(r'\s*def create_compare_widgets\(self\):.*?if __name__ == "__main__":', '\n\nif __name__ == "__main__":', code, flags=re.DOTALL)

with open('c:/Users/jjch2/Desktop/PMI/home/src/코드절차서관리.py', 'w', encoding='utf-8') as f:
    f.write(code)
print('Cleanup done')
