import json
import os
import sys

# Define path to config
documents_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'MaterialManager')
config_path = os.path.join(documents_dir, 'Material_Manager_Config.json')

def check_config():
    if not os.path.exists(config_path):
        # Fallback to current directory for script mode tests
        config_path_local = 'Material_Manager_Config.json'
        if not os.path.exists(config_path_local):
            print("Config file not found.")
            return
        path = config_path_local
    else:
        path = config_path

    try:
        with open(path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            selected_tab = config.get('selected_tab', 'Not Found')
            print(f"Current saved tab index in {path}: {selected_tab}")
            return selected_tab
    except Exception as e:
        print(f"Error reading config: {e}")

if __name__ == "__main__":
    check_config()
