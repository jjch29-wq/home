#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MaterialManager-1 EXE 빌드 스크립트
"""

import os
import sys
import subprocess
import shutil

def build_exe():
    """PyInstaller로 EXE 파일 생성"""
    
    # 현재 디렉토리
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
    
    print("MaterialManager-2 EXE 파일 생성을 시작합니다...")
    
    # PyInstaller 설치 확인
    try:
        import PyInstaller
        print("PyInstaller가 설치되어 있습니다.")
    except ImportError:
        print("PyInstaller를 설치합니다...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("PyInstaller 설치 완료!")
    
    # 빌드 대상 파일
    target_script = "MaterialManager-5.py"
    exe_name = "MaterialManager_v5"
    
    # 빌드 명령어
    build_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",           # 단일 파일로 생성
        "--windowed",          # 콘솔 창 없이
        f"--name={exe_name}",  # 파일 이름
        "--icon=NONE",         # 아이콘 없이
        "--clean",             # 이전 빌드 정리
        "--add-data", f"Material_Inventory.xlsx;.", # 데이터 포함
        "--add-data", f"Material_Manager_Config.json;.", # 이미지/설정 포함
        target_script
    ]
    
    try:
        print("빌드를 시작합니다...")
        subprocess.run(build_cmd, check=True)
        print("빌드 완료!")
        
        # 생성된 파일 확인
        exe_path = os.path.join(current_dir, "dist", f"{exe_name}.exe")
        if os.path.exists(exe_path):
            print(f"EXE 파일 생성 완료: {exe_path}")
            
            # 배포용 ZIP 파일 생성
            print("배포용 ZIP 파일을 생성합니다...")
            zip_filename = f"{exe_name}_Portable_Pack"
            shutil.make_archive(os.path.join(current_dir, zip_filename), 'zip', os.path.join(current_dir, "dist"))
            print(f"배포용 ZIP 파일 생성 완료: {zip_filename}.zip")

            # 바탕화면에 바로가기 대신 배포용 폴더 링크 표시 (선택사항)
            # 여기서는 편의를 위해 dist 폴더의 파일을 복사
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop, f"{exe_name}.exe")
            
            try:
                shutil.copy2(exe_path, shortcut_path)
                print(f"바탕화면에 실행 파일 복사 완료: {shortcut_path}")
            except Exception as e:
                print(f"바탕화면 복사 실패: {e}")
            
            return True
        else:
            print("EXE 파일 생성에 실패했습니다.")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"빌드 오류: {e}")
        return False
    except Exception as e:
        print(f"예상치 못한 오류: {e}")
        return False

def main():
    if build_exe():
        print("\n=== 빌드 성공 ===")
        print("MaterialManager-1.exe 파일이 생성되었습니다.")
        print("바탕화면에서 실행하거나 dist 폴더에서 실행할 수 있습니다.")
        input("엔터를 누르면 종료합니다...")
    else:
        print("\n=== 빌드 실패 ===")
        input("엔터를 누르면 종료합니다...")

if __name__ == "__main__":
    main()
