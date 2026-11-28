"""Dedicated brand identity and launcher entry point for the NDT-Pro Multi-Site App."""
from __future__ import annotations

import argparse
import os
import sys

APP_BRAND_NAME = "NDT-Pro 통합 자재·비파괴검사 관리"
APP_VERSION = "1.0.0-multi"


def format_window_title(profile) -> str:
    """Return branded window title for multi-site profiles while preserving legacy central title."""
    if getattr(profile, "is_central", False):
        return "자재 및 소모품 관리 시스템 (Material Manager)"
    inv_tag = "공용재고" if getattr(profile, "inventory_mode", "shared") == "shared" else "독립재고"
    display = getattr(profile, "display_name", "") or getattr(profile, "profile_id", "기본")
    return f"[{APP_BRAND_NAME}] {display} ({inv_tag})"


def main(argv: list[str] | None = None) -> int:
    src_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
    if src_root not in sys.path:
        sys.path.insert(0, src_root)

    parser = argparse.ArgumentParser(description=APP_BRAND_NAME)
    parser.add_argument("--profile", dest="profile_id", default=None)
    parser.add_argument("--pick-profile", action="store_true")
    args, _ = parser.parse_known_args(argv)

    from site_apps.central.src.site_profile import pick_profile_dialog, set_active_profile_id

    profile_id = args.profile_id
    if args.pick_profile or not profile_id:
        profile_id = pick_profile_dialog()
        if not profile_id:
            return 0

    set_active_profile_id(profile_id)

    import tkinter as tk
    from site_apps.central.src.app import MaterialManager

    root = tk.Tk()
    MaterialManager(root, profile_id=profile_id)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        try:
            root.destroy()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

