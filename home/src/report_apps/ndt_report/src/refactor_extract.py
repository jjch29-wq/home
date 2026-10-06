import sys

file_path = r'c:\Users\jjch2\Desktop\PMI\home\src\report_apps\ndt_report\src\비파괴검사보고서.py'
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1
pre_loop_end_idx = -1

for i, line in enumerate(lines):
    if 'self.log(f"🔍 {mode} 데이터 추출 시작' in line:
        start_idx = i
        break

for i in range(start_idx, len(lines)):
    if 'if mode == "PMI" and hasattr(self, \'element_filters\') and self.element_filters:' in lines[i]:
        end_idx = i
        break

# Find where auto column addition ends
for i, line in enumerate(lines):
    if 'def _find_col(df, keywords, exclude=None):' in line:
        pre_loop_end_idx = i
        break

print(f"start_idx: {start_idx}, pre_loop_end_idx: {pre_loop_end_idx}, end_idx: {end_idx}")

new_lines = []
# 1. Everything before line 8624 (where `if not target_file:` is)
target_file_check_idx = start_idx - 6
new_lines.extend(lines[:target_file_check_idx])

# 2. Add multiple file check and initialization
new_lines.append('        try: target_files = self.root.tk.splitlist(target_file)\n')
new_lines.append('        except: target_files = [target_file]\n')
new_lines.append('        if not target_files or not target_files[0]:\n')
new_lines.append('            messagebox.showwarning("파일 미선택", f"{mode} 데이터 파일을 선택해주세요.")\n')
new_lines.append('            return False\n')
new_lines.append('        \n')
new_lines.append('        self.progress[\'value\'] = 0\n')
new_lines.append('        all_extracted_data = []\n')

# 3. Add AUTO COLUMN ADDITION block (lines 8664 to 8699 approx)
auto_col_start = -1
for i in range(start_idx, pre_loop_end_idx):
    if '# [AUTO COLUMN ADDITION]' in lines[i]:
        auto_col_start = i
        break

new_lines.extend(lines[auto_col_start:pre_loop_end_idx])

# 4. Add the helper functions definition (they don't need to be indented inside loop)
# lines[pre_loop_end_idx] to lines[8723] where `try:` starts
loop_start = -1
for i in range(pre_loop_end_idx, end_idx):
    if 'try:' in lines[i] and 'target_input = self.sequence_filter.get().strip()' in lines[i+1]:
        loop_start = i
        break

new_lines.extend(lines[pre_loop_end_idx:loop_start])

# 5. Start loop
new_lines.append('        for target_file in target_files:\n')
new_lines.append('            if not target_file: continue\n')

# 6. Add start_idx to auto_col_start (which contains log, date extraction, etc.) indented
for line in lines[start_idx:auto_col_start]:
    new_lines.append('    ' + line if line.strip() else line)

# 7. Add loop_start to end_idx indented
for line in lines[loop_start:end_idx]:
    new_lines.append('    ' + line if line.strip() else line)

# 8. Add the rest of the file
new_lines.extend(lines[end_idx:])

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)
print("Done!")
