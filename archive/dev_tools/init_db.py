import pandas as pd
import os

db_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Material_Inventory.xlsx'

if not os.path.exists(db_path):
    with pd.ExcelWriter(db_path, engine='openpyxl') as writer:
        # Materials Sheet with expanded fields
        materials_columns = [
            'Material ID',           # 자재 ID
            'Equipment Code',        # 설비코드
            'Item Name',            # 품목
            'Classification',       # 분류
            'Specification',        # 규격
            'Unit',                 # 단위
            'Supplier',             # 공급업자
            'Manufacturer',         # 제조조
            'Reorder Point',        # 재주문 수준
            'Initial Stock',        # 초기 재고
            'Current Stock'         # 현재 재고
        ]
        pd.DataFrame(columns=materials_columns).to_excel(writer, sheet_name='Materials', index=False)
        
        # Transactions Sheet
        transaction_columns = ['Date', 'Material ID', 'Type', 'Quantity', 'Note', 'User']
        pd.DataFrame(columns=transaction_columns).to_excel(writer, sheet_name='Transactions', index=False)
    print(f"Created initial database at {db_path}")
else:
    print(f"Database already exists at {db_path}")
