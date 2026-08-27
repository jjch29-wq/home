import sys, os
f=open(r'c:\Users\jjch2\Desktop\PMI\home\src\services\data_loader.py', 'r', encoding='utf-8')
content = f.read()
f.close()

old = '''        if not os.path.exists(self.db_path):
            bundled_db = os.path.join(self.bundle_dir, 'Material_Inventory.xlsx')
            print(f"DEBUG: Main DB not found. Trying to restore from bundle: {bundled_db}")
            if os.path.exists(bundled_db):
                import shutil
                try:
                    shutil.copy2(bundled_db, self.db_path)
                    print("DEBUG: Restored DB from bundle.")
                    # Also try to copy config if it exists in bundle but not in app_dir
                    bundled_config = os.path.join(self.bundle_dir, 'Material_Manager_Config.json')'''

new = '''        if not os.path.exists(self.db_path):
            db_basename = os.path.basename(self.db_path)
            bundled_db = os.path.join(self.bundle_dir, db_basename)
            print(f"DEBUG: Main DB not found. Trying to restore from bundle: {bundled_db}")
            if os.path.exists(bundled_db):
                import shutil
                try:
                    shutil.copy2(bundled_db, self.db_path)
                    print("DEBUG: Restored DB from bundle.")
                    # Also try to copy config if it exists in bundle but not in app_dir
                    config_basename = os.path.basename(self.config_path)
                    bundled_config = os.path.join(self.bundle_dir, config_basename)'''

content = content.replace(old, new)
f=open(r'c:\Users\jjch2\Desktop\PMI\home\src\services\data_loader.py', 'w', encoding='utf-8')
f.write(content)
f.close()
print('Done')
