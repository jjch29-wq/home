import subprocess
import sys

def install(package):
    print(f"Installing {package}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    import tkcalendar
    print("tkcalendar is already installed.")
except ImportError:
    print("tkcalendar not found. Installing...")
    try:
        install("tkcalendar")
        print("Successfully installed tkcalendar.")
    except Exception as e:
        print(f"Failed to install tkcalendar: {e}")

try:
    import babel
    print("babel is already installed.")
except ImportError:
    print("babel not found. Installing...")
    try:
        install("babel")
        print("Successfully installed babel.")
    except Exception as e:
        print(f"Failed to install babel: {e}")

required_packages = ['pandas', 'openpyxl']
for package in required_packages:
    try:
        __import__(package)
        print(f"{package} is already installed.")
    except ImportError:
        print(f"{package} not found. Installing...")
        try:
            install(package)
            print(f"Successfully installed {package}.")
        except Exception as e:
            print(f"Failed to install {package}: {e}")

input("Press Enter to exit...")
