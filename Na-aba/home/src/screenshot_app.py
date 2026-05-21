import sys
import os
import importlib.util
import tkinter as tk
from PIL import ImageGrab
import time

script_path = r'c:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI\Na-aba\home\src\Archived-Main-App-20260405-RT-Fix.py'

spec = importlib.util.spec_from_file_location("main_app", script_path)
main_app = importlib.util.module_from_spec(spec)
sys.modules["main_app"] = main_app
spec.loader.exec_module(main_app)

root = tk.Tk()

# Initialize app
app = main_app.PMIReportApp(root)

# Position and size of window (set AFTER app initialization)
root.state('normal')
root.geometry("1200x800+50+50")
root.deiconify()
root.lift()
root.attributes('-topmost', True)
root.focus_force()

# Active wait loop to let window manager map and render the window
for i in range(30):
    # Force select in main notebook
    app.mode_notebook.select(1)
    # Force select in RT notebook
    app.rt_preview_nb.select(1)
    
    root.update()
    time.sleep(0.1)

# Grab screenshot
screenshot = ImageGrab.grab(bbox=(50, 50, 1250, 850))

# Save to artifacts directory
save_dir = r'C:\Users\jjch2\.gemini\antigravity\brain\d385025f-1282-4eab-888e-f4f37927ae88'
if not os.path.exists(save_dir):
    os.makedirs(save_dir)
screenshot.save(os.path.join(save_dir, 'app_screenshot.png'))

print("Screenshot saved successfully!")
root.destroy()
