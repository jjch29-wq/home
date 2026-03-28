
import subprocess
import sys
import os

def build():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
    
    script_name = "PipeAreaCalculator.py"
    exe_name = "PipeAreaCalculator"
    
    print(f"--- Building {exe_name} Standalone EXE ---")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--clean",
        "--name", exe_name,
        script_name
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print("\nSUCCESS: Build complete.")
        
        dist_path = os.path.join(current_dir, "dist", f"{exe_name}.exe")
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        
        if os.path.exists(dist_path):
            import shutil
            shutil.copy2(dist_path, os.path.join(desktop, f"{exe_name}.exe"))
            print(f"Copied to Desktop: {os.path.join(desktop, f'{exe_name}.exe')}")
            
    except subprocess.CalledProcessError as e:
        print(f"\nERROR: Build failed with return code {e.returncode}")

if __name__ == "__main__":
    build()
