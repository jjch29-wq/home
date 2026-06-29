import json
with open(r'c:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json', 'r', encoding='utf-8') as f:
    c = json.load(f)
print('LABOR UT:', c['LABOR_COST']['수송배관(주배관)']['일반']['UT'])
print('LABOR PT:', c['LABOR_COST']['수송배관(주배관)']['일반']['PT'])
