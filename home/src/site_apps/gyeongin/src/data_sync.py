import json
import os


class SharedDataUnavailableError(RuntimeError):
    """Raised when configured shared business storage is unavailable."""


def _configured_data_dir():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(base_dir, "sync_config.json")
    if not os.path.exists(config_file):
        return None
    try:
        with open(config_file, "r", encoding="utf-8") as stream:
            return str(json.load(stream).get("data_dir") or "").strip() or None
    except (OSError, ValueError, TypeError):
        return None


def _shared_dir_candidates(custom_dir):
    yield custom_dir
    _, relative = os.path.splitdrive(custom_dir)
    relative = relative.lstrip(r'\/')
    for drive in 'CDEFGHIJK':
        candidate = f"{drive}:\\{relative}"
        if os.path.normcase(candidate) != os.path.normcase(custom_dir):
            yield candidate


def get_data_connection_error():
    """Explain a configured-but-unavailable shared data directory."""
    custom_dir = _configured_data_dir()
    if not custom_dir:
        return None
    if any(os.path.isdir(path) for path in _shared_dir_candidates(custom_dir)):
        return None
    return (
        "경인사무소 공유 작업자료 경로에 연결할 수 없습니다.\n"
        f"설정 경로: {custom_dir}\n\n"
        "오래된 로컬 자료에 잘못 저장되는 것을 막기 위해 저장을 중단했습니다. "
        "Google Drive 또는 공유 드라이브 연결을 복구한 뒤 다시 시도해 주세요."
    )


def require_shared_data_available():
    error = get_data_connection_error()
    if error:
        raise SharedDataUnavailableError(error)


def get_data_dir():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    custom_dir = _configured_data_dir()
    if custom_dir:
        for candidate in _shared_dir_candidates(custom_dir):
            if os.path.isdir(candidate):
                return candidate
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
    return os.path.join(get_data_dir(), '경인사무소_제출자료')


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
