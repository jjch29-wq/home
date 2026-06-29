import json

config_path = r'C:\Users\-\OneDrive\바탕 화면\home_new\home\src\config.json'
with open(config_path, 'r', encoding='utf-8') as f:
    c = json.load(f)

# The total price is calculated as:
# For each location, each shift, each test type (RT, UT, PT):
# Amount = Quantity * (Material_Cost + Labor_Cost_Adjusted_for_Shift)
# Wait! In `ndt_billing_tab.py`, how is the contract amount calculated?
# Ah! Let's check `ndt_billing_tab.py` to see the exact calculation logic for Contract Amount!
# Actually, the user just asked to check "현재 앱 수량 금액이 맞는지" (Check if the app's quantity and amount are correct).
