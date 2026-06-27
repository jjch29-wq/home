import os
import sys
import subprocess
import shutil

def build():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
    
    spec_file = "PMI_Report_V2_Unified.spec"
    
    if not os.path.exists(spec_file):
        print(f"Error: {spec_file} not found.")
        return

    print("--- Starting Build Process ---")
    
    # Ensure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)

    # Run PyInstaller
    print(f"Running PyInstaller for {spec_file}...")
    subprocess.run([sys.executable, "-m", "PyInstaller", "--clean", spec_file], check=True)
    
    exe_name = "PMI_Report_V2_Unified.exe"
    dist_exe = os.path.join(current_dir, "dist", exe_name)
    
    if os.path.exists(dist_exe):
        print(f"\nSUCCESS: Executable created at {dist_exe}")
        
        # Copy to Desktop if requested or just keep in dist
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        try:
            shutil.copy2(dist_exe, os.path.join(desktop, exe_name))
            print(f"Copied to Desktop: {os.path.join(desktop, exe_name)}")
        except Exception as e:
            print(f"Could not copy to Desktop: {e}")
    else:
        print("\nERROR: Build failed. dist/EXE not found.")

if __name__ == "__main__":
    build()
