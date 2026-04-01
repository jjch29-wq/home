import sys
import os

print(f"Python Executable: {sys.executable}")
print(f"Python Version: {sys.version}")
print("System Path:")
for p in sys.path:
    print(p)

try:
    import tkcalendar
    print(f"tkcalendar found at: {tkcalendar.__file__}")
except ImportError as e:
    print(f"Error importing tkcalendar: {e}")
