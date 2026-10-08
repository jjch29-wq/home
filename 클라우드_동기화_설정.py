import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import json

def setup_cloud_sync():
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("클라우드 동기화 설정", "원드라이브(OneDrive)나 구글 드라이브 등 동기화할 폴더를 선택해주세요.\n해당 폴더 내에 'PMI_Data' 폴더가 생성됩니다.")
    
    selected_dir = filedialog.askdirectory(title="클라우드 동기화 폴더 선택")
    
    if not selected_dir:
        messagebox.showwarning("취소", "폴더 선택이 취소되었습니다.")
        return
        
    target_data_dir = os.path.join(selected_dir, "PMI_Data", "central")
    os.makedirs(target_data_dir, exist_ok=True)
    
    # 1. 기존 데이터 복사 (daily_work_history.json)
    base_src_dir = os.path.join(os.path.dirname(__file__), "home", "src", "site_apps", "central", "src")
    src_history = os.path.join(base_src_dir, "daily_work_history.json")
    
    if os.path.exists(src_history):
        shutil.copy2(src_history, os.path.join(target_data_dir, "daily_work_history.json"))
        print("작업일보 데이터 복사 완료.")
        
    # 2. 기존 사진 복사 (data/process_photos)
    src_photos_dir = os.path.join(os.path.dirname(__file__), "home", "src", "site_apps", "central", "data", "process_photos")
    target_photos_dir = os.path.join(target_data_dir, "data", "process_photos")
    
    if os.path.exists(src_photos_dir):
        if not os.path.exists(target_photos_dir):
            os.makedirs(os.path.dirname(target_photos_dir), exist_ok=True)
            shutil.copytree(src_photos_dir, target_photos_dir)
            print("사진 데이터 복사 완료.")
        else:
            print("사진 데이터 폴더가 이미 존재합니다.")
            
    # 3. 설정 파일 저장 (sync_config.json)
    config_path = os.path.join(base_src_dir, "sync_config.json")
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump({"data_dir": target_data_dir}, f, ensure_ascii=False, indent=4)
        
    messagebox.showinfo("설정 완료", f"클라우드 동기화 설정이 완료되었습니다!\n이제 데이터가 다음 위치에 저장됩니다:\n{target_data_dir}\n\n다른 컴퓨터에서도 동일하게 이 스크립트를 실행하고 같은 폴더를 선택해주시면 됩니다.")

if __name__ == "__main__":
    setup_cloud_sync()
