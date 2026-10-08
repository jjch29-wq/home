import os
import json

def get_data_dir():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(base_dir, "sync_config.json")
    if os.path.exists(config_file):
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                config = json.load(f)
                custom_dir = config.get("data_dir")
                if custom_dir and os.path.exists(custom_dir):
                    return custom_dir
        except Exception:
            pass
    return base_dir

def get_history_path():
    return os.path.join(get_data_dir(), 'daily_work_history.json')

def get_process_photos_dir():
    return os.path.join(get_data_dir(), 'data', 'process_photos')
