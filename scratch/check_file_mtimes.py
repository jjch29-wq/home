import os
import time

def show_mtime(fp):
    if os.path.exists(fp):
        mtime = os.path.getmtime(fp)
        print(f"File: {fp} | Modified: {time.ctime(mtime)}")
    else:
        print(f"File not found: {fp}")

show_mtime("요청서 합치기.py")
show_mtime("Na-aba/Final_Smart_Merged_v2.8_220225.xlsx")
show_mtime("Na-aba/Final_Smart_Merged_v2.8_215354.xlsx")
