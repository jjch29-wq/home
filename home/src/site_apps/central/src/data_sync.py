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
                if custom_dir:
                    if os.path.exists(custom_dir):
                        return custom_dir
                    # PC마다 구글 드라이브 드라이브 문자(예: G:, F:)가 다를 경우 자동 탐색
                    _, path_without_drive = os.path.splitdrive(custom_dir)
                    path_without_drive = path_without_drive.lstrip(r'\/')
                    for drive in 'CDEFGHIJK':
                        alt_path = f"{drive}:\\{path_without_drive}"
                        if os.path.exists(alt_path):
                            return alt_path
        except Exception:
            pass
    return base_dir

def get_history_path():
    return os.path.join(get_data_dir(), 'daily_work_history.json')


def get_business_data_dir():
    """Use cloud business files when configured, otherwise the site's data folder."""
    local_dir = os.path.dirname(os.path.abspath(__file__))
    root = get_data_dir()
    if os.path.normcase(os.path.abspath(root)) == os.path.normcase(local_dir):
        return os.path.abspath(os.path.join(local_dir, '..', 'data'))
    return os.path.join(root, 'data')


def get_submission_dir():
    return os.path.join(get_data_dir(), '중앙지사_제출자료')

def get_process_photos_dir():
    return os.path.join(get_data_dir(), 'data', 'process_photos')


def resolve_process_photo_path(history_path, stored_path):
    """Find a registered photo after moving central history to cloud storage."""
    if not stored_path:
        return None
    history_dir = os.path.dirname(os.path.abspath(history_path))
    direct = (stored_path if os.path.isabs(stored_path)
              else os.path.abspath(os.path.join(history_dir, stored_path)))
    if os.path.isfile(direct):
        return direct

    # Match the original UUID filename and date folder only. Never search
    # another site's storage or substitute a similarly named image.
    normalized = str(stored_path).replace('\\', '/')
    parts = normalized.rstrip('/').split('/')
    if len(parts) < 2 or parts[-2] in ('', '.', '..'):
        return None
    date_folder, filename = parts[-2:]
    local_dir = os.path.dirname(os.path.abspath(__file__))
    roots = (
        os.path.join(history_dir, 'data', 'process_photos'),
        os.path.join(local_dir, '..', 'data', 'process_photos'),
        os.path.join(local_dir, 'data', 'process_photos'),
    )
    for root in roots:
        candidate = os.path.abspath(os.path.join(root, date_folder, filename))
        if os.path.isfile(candidate):
            return candidate
    return None
