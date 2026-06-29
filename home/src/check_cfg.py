import json
with open(r'c:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json', 'r', encoding='utf-8') as f:
    c = json.load(f)
print('LABOR:', c['LABOR_COST']['수송배관(주배관)']['일반']['RT'])
print('MAT:', c['MATERIAL_COST'].get('RT (B필름: 3⅓"x17")', 'missing'))
