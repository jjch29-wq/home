import json
import os
import time
from site_apps.central.src.runtime_profile import is_development


class SharedDataUnavailableError(RuntimeError):
    """Raised when configured shared business storage is unavailable."""


def _configured_data_dir():
    if is_development():
        return None
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(base_dir, "sync_config.json")
    if not os.path.exists(config_file):
        return None
    try:
        with open(config_file, "r", encoding="utf-8") as stream:
            return str(json.load(stream).get("data_dir") or "").strip() or None
    except (OSError, ValueError, TypeError):
        return None


_CACHED_SHARED_DIR = None
_CACHED_CUSTOM_DIR = None


def _shared_dir_candidates(custom_dir):
    norm_custom = os.path.normpath(custom_dir)
    orig_drive, relative = os.path.splitdrive(norm_custom)
    rel_clean = relative.replace('/', '\\').lstrip('\\')

    rel_variants = [rel_clean]
    if rel_clean.startswith('내 드라이브\\'):
        rel_variants.append('My Drive\\' + rel_clean[len('내 드라이브\\'):])
    elif rel_clean.startswith('My Drive\\'):
        rel_variants.append('내 드라이브\\' + rel_clean[len('My Drive\\'):])

    seen = set()

    # 1. 설정된 원래 경로의 드라이브 루트가 실존하면 최우선 탐색
    if orig_drive:
        try:
            if os.path.exists(f"{orig_drive}\\"):
                seen.add(os.path.normcase(norm_custom))
                yield norm_custom
        except Exception:
            pass

    # 2. 회사(F:)와 집(G:)을 최우선 순위로 탐색하되, 활성화된 드라이브 먼저 탐색
    priority_drives = [d for d in 'GFEDCBAHIJKLMNOPQRSTUVWXYZ']
    active_drives = []
    for d in priority_drives:
        try:
            if os.path.exists(f"{d}:\\"):
                active_drives.append(d)
        except Exception:
            continue

    # 활성 드라이브 후보 먼저 산출
    for drive in active_drives:
        for rel in rel_variants:
            candidate = os.path.normpath(f"{drive}:\\{rel}")
            key = os.path.normcase(candidate)
            if key not in seen:
                seen.add(key)
                yield candidate

    # 비활성 드라이브는 후순위로 시도 (일시적 마운트 대기)
    for drive in priority_drives:
        if drive in active_drives:
            continue
        for rel in rel_variants:
            candidate = os.path.normpath(f"{drive}:\\{rel}")
            key = os.path.normcase(candidate)
            if key not in seen:
                seen.add(key)
                yield candidate


def _find_available_shared_dir(custom_dir, retries=3, delay=0.4):
    """Google Drive 가상 드라이브의 일시적 마운트 지연/응답 대기를 고려하여 재시도하며 유효 경로를 찾는다."""
    global _CACHED_SHARED_DIR, _CACHED_CUSTOM_DIR
    if _CACHED_SHARED_DIR and _CACHED_CUSTOM_DIR == custom_dir:
        try:
            if os.path.isdir(_CACHED_SHARED_DIR):
                return _CACHED_SHARED_DIR
        except Exception:
            _CACHED_SHARED_DIR = None
            _CACHED_CUSTOM_DIR = None

    candidates = list(_shared_dir_candidates(custom_dir))
    for attempt in range(retries + 1):
        for path in candidates:
            try:
                if os.path.isdir(path):
                    _CACHED_SHARED_DIR = path
                    _CACHED_CUSTOM_DIR = custom_dir
                    return path
            except Exception:
                continue
        if attempt < retries:
            time.sleep(delay)
    return None


def get_data_connection_error():
    """Explain a configured-but-unavailable shared data directory."""
    custom_dir = _configured_data_dir()
    if not custom_dir:
        return None
    if _find_available_shared_dir(custom_dir) is not None:
        return None
    return (
        "중앙지사 공유 작업자료 경로에 연결할 수 없습니다.\n"
        f"설정 경로: {custom_dir}\n"
        "(회사 컴퓨터 F: 드라이브 / 현재 컴퓨터 G: 드라이브 자동 감지)\n\n"
        "오래된 로컬 자료에 잘못 저장되는 것을 막기 위해 저장을 중단했습니다.\n"
        "Google Drive 데스크톱 앱이 켜져 있는지 확인한 뒤 다시 시도해 주세요."
    )


def require_shared_data_available():
    error = get_data_connection_error()
    if error:
        raise SharedDataUnavailableError(error)


def get_data_dir():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    custom_dir = _configured_data_dir()
    if custom_dir:
        found = _find_available_shared_dir(custom_dir)
        if found:
            return found
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
