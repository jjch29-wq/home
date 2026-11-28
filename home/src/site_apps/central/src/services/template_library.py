"""Per-profile, per-client NDT report & photo-log template library."""
from __future__ import annotations

from datetime import datetime
import json
import os
import shutil
from typing import Any

from site_apps.central.src.site_profile import (
    SiteProfile,
    get_active_profile,
    get_profile_data_dir,
    get_profiles_root,
)

SUPPORTED_TEMPLATE_MODES: tuple[str, ...] = ("PMI", "RT", "PT", "MT", "PAUT", "PHOTO", "STATEMENT")
DEFAULT_CLIENT_KEY = "default"


def _library_paths(profile: SiteProfile | None = None) -> tuple[str, str]:
    active = profile or get_active_profile()
    data_dir = get_profile_data_dir(active)
    if data_dir:
        templates_dir = os.path.join(os.path.dirname(data_dir), "templates")
        registry_path = os.path.join(data_dir, "template_library.json")
    else:
        base_dir = os.path.abspath(os.path.join(get_profiles_root(), "central"))
        templates_dir = os.path.join(base_dir, "templates")
        registry_path = os.path.join(base_dir, "template_library.json")
    os.makedirs(templates_dir, exist_ok=True)
    os.makedirs(os.path.dirname(registry_path), exist_ok=True)
    return registry_path, templates_dir


def _load_registry(profile: SiteProfile | None = None) -> dict[str, Any]:
    registry_path, _ = _library_paths(profile)
    if not os.path.isfile(registry_path):
        return {"version": 1, "entries": {}}
    try:
        with open(registry_path, "r", encoding="utf-8") as stream:
            data = json.load(stream)
        if isinstance(data, dict) and isinstance(data.get("entries"), dict):
            return data
    except (OSError, ValueError, TypeError):
        pass
    return {"version": 1, "entries": {}}


def _save_registry(data: dict[str, Any], profile: SiteProfile | None = None) -> None:
    registry_path, _ = _library_paths(profile)
    tmp_path = registry_path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)
    os.replace(tmp_path, registry_path)


def register_template(
    source_path: str,
    mode: str,
    *,
    client_id: str | None = None,
    description: str = "",
    copy_into_library: bool = True,
    profile: SiteProfile | None = None,
) -> dict[str, Any]:
    mode_up = str(mode or "").strip().upper()
    if mode_up not in SUPPORTED_TEMPLATE_MODES:
        raise ValueError(f"지원하지 않는 템플릿 모드입니다: {mode}")
    if not source_path or not os.path.isfile(source_path):
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {source_path}")

    client_key = str(client_id or "").strip() or DEFAULT_CLIENT_KEY
    _, templates_dir = _library_paths(profile)

    if copy_into_library:
        target_dir = os.path.join(templates_dir, client_key)
        os.makedirs(target_dir, exist_ok=True)
        ext = os.path.splitext(source_path)[1] or ".xlsx"
        base_name = os.path.basename(source_path)
        stored_path = os.path.join(target_dir, f"{mode_up}_{base_name}" if not base_name.upper().startswith(mode_up) else base_name)
        if not stored_path.lower().endswith(ext.lower()):
            stored_path += ext
        if os.path.abspath(source_path) != os.path.abspath(stored_path):
            shutil.copy2(source_path, stored_path)
    else:
        stored_path = os.path.abspath(source_path)

    registry = _load_registry(profile)
    client_bucket = registry["entries"].setdefault(client_key, {})
    entry = {
        "client_id": client_key,
        "mode": mode_up,
        "path": stored_path,
        "filename": os.path.basename(stored_path),
        "description": description.strip(),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    client_bucket[mode_up] = entry
    _save_registry(registry, profile)
    return entry


def resolve_template(
    mode: str,
    *,
    client_id: str | None = None,
    profile: SiteProfile | None = None,
) -> dict[str, Any] | None:
    """Resolve client-specific template first, falling back to profile default template."""
    mode_up = str(mode or "").strip().upper()
    registry = _load_registry(profile)
    entries = registry.get("entries", {})

    candidates = []
    if client_id and str(client_id).strip():
        candidates.append(str(client_id).strip())
    if DEFAULT_CLIENT_KEY not in candidates:
        candidates.append(DEFAULT_CLIENT_KEY)

    for key in candidates:
        entry = entries.get(key, {}).get(mode_up)
        if isinstance(entry, dict) and entry.get("path") and os.path.isfile(entry["path"]):
            return dict(entry)
    return None


def list_templates(
    *,
    client_id: str | None = None,
    profile: SiteProfile | None = None,
) -> list[dict[str, Any]]:
    registry = _load_registry(profile)
    entries = registry.get("entries", {})
    result: list[dict[str, Any]] = []
    if client_id:
        bucket = entries.get(str(client_id).strip(), {})
        for mode in sorted(bucket):
            result.append(dict(bucket[mode]))
        return result
    for key in sorted(entries):
        for mode in sorted(entries[key]):
            result.append(dict(entries[key][mode]))
    return result


def remove_template(
    mode: str,
    *,
    client_id: str | None = None,
    profile: SiteProfile | None = None,
) -> bool:
    mode_up = str(mode or "").strip().upper()
    client_key = str(client_id or "").strip() or DEFAULT_CLIENT_KEY
    registry = _load_registry(profile)
    bucket = registry.get("entries", {}).get(client_key)
    if not isinstance(bucket, dict) or mode_up not in bucket:
        return False
    del bucket[mode_up]
    if not bucket:
        del registry["entries"][client_key]
    _save_registry(registry, profile)
    return True

