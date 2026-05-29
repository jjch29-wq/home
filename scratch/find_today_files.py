import os
import time

def find_today():
    print("Excel files modified today:")
    for root, dirs, files in os.walk('.'):
        if '.venv' in root or '.git' in root:
            continue
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm'):
                fp = os.path.join(root, file)
                mtime = os.path.getmtime(fp)
                # Check if modified today (2026-05-29)
                mtime_struct = time.localtime(mtime)
                if mtime_struct.tm_year == 2026 and mtime_struct.tm_mon == 5 and mtime_struct.tm_mday == 29:
                    print(f"File: {fp} | Modified: {time.ctime(mtime)} | Size: {os.path.getsize(fp)} bytes")

find_today()
