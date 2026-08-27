import sys
import os
import json

sys.path.append(r'c:\Users\jjch2\Desktop\PMI\home\src\services')
from monthly_report_manager import MonthlyReportManager

history_path = r'c:\Users\jjch2\Desktop\PMI\home\src\daily_work_history.json'
with open(history_path, 'r', encoding='utf-8') as f:
    history = json.load(f)

print("Loaded history.")
process_photos = []
target_dates = [d for d in history.keys() if d.startswith('2026-08')]
print(f"Target dates: {target_dates}")

for d in target_dates:
    photos = history[d].get('process_photos', [])
    print(f"Date {d}: {len(photos)} photos")
    process_photos.extend(photos)

print(f"Total photos: {len(process_photos)}")

# Check path resolution
base_dir = os.path.dirname(os.path.abspath(history_path))
valid = []
for p in process_photos:
    process = str(p.get('process', '')).strip().upper()
    if process not in {'PAUT', 'MT', 'RT', 'PT'}:
        print(f"Invalid process: {process}")
        continue
    stored_path = str(p.get('file_path', '') or '')
    image_path = stored_path if os.path.isabs(stored_path) else os.path.abspath(os.path.join(base_dir, stored_path))
    
    if os.path.isfile(image_path):
        valid.append(image_path)
    else:
        print(f"File not found: {image_path}")
        
print(f"Valid photos: {len(valid)}")
