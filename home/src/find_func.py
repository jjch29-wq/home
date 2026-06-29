with open(r'c:\Users\-\OneDrive\바탕 화면\home_new\home\src\ndt_billing_tab.py', 'r', encoding='utf-8') as f:
    for i, line in enumerate(f.readlines()):
        if 'auto_load_contract_qty' in line:
            print(f'Line {i+1}: {line.rstrip()}')
