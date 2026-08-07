import os

with open('daily_work_log_exporter.py', 'r', encoding='utf-8') as f:
    code = f.read()

old_loop = """        for rng, text in headers_ndt:
            merge_and_set(rng, text, font=self.font_small, fill=self.fill_header)"""
new_loop = """        for rng, text in headers_ndt:
            merge_and_set(rng, text, font=self.font_small, fill=self.fill_header, align=self.align_nowrap)"""

code = code.replace(old_loop, new_loop)

with open('daily_work_log_exporter.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("Updated NDT headers to use nowrap successfully")
