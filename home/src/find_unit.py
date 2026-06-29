import sys
import os

with open(r'c:\Users\-\OneDrive\바탕 화면\home_new\home\src\ndt_billing_tab.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()
    for i, line in enumerate(lines):
        if 'c_price' in line or 'unit_cost' in line or 'self.contract_vars[full_key]' in line:
            if 'self.contract_vars' in line:
                print(f"Line {i+1}: {line.strip()}")
