"""Multi-site / multi-contract profile manager for the central NDT & material app.

Supports three business modes on a single shared codebase:
- ``kdhc``: Korea District Heating branch operations (full 10 tabs, hot-tapping, KDHC reports)
- ``annual_contract``: Annual unit-price contract operations (shared company inventory + isolated contract records/billing)
- ``smb_general``: Small/medium business per-job / lump-sum operations (shared company inventory + streamlined tabs + statement export)
"""
from __future__ import annotations

from dataclasses import dataclass, field
import json
import os
import re
from typing import Any

PROFILE_ENV_VAR = "PMI_SITE_PROFILE"
DEFAULT_PROFILE_ID = "central"

ALL_TAB_KEYS: tuple[str, ...] = (
    "stock",
    "io_transfer",
    "inout",
    "daily_usage",
    "daily_history",
    "monthly",
    "work_log",
    "ndt_billing",
    "budget",
    "ndt_report",
    "clients_contracts",
)

CENTRAL_DEFAULT_TAB_KEYS: tuple[str, ...] = (
    "stock",
    "io_transfer",
    "inout",
    "daily_usage",
    "daily_history",
    "monthly",
    "work_log",
    "ndt_billing",
    "budget",
    "ndt_report",
)

SITE_TYPE_DEFAULTS: dict[str, dict[str, Any]] = {
    "kdhc": {
        "enabled_tabs": list(CENTRAL_DEFAULT_TAB_KEYS),
        "enable_hot_tapping": True,
        "inventory_mode": "shared",
    },
    "annual_contract": {
        "enabled_tabs": [
            "stock",
            "io_transfer",
            "inout",
            "daily_usage",
            "daily_history",
            "monthly",
            "work_log",
            "ndt_billing",
            "budget",
            "ndt_report",
            "clients_contracts",
        ],
        "enable_hot_tapping": False,
        "inventory_mode": "shared",
    },
    "smb_general": {
        "enabled_tabs": [
            "stock",
            "io_transfer",
            "inout",
            "daily_usage",
            "daily_history",
            "monthly",
            "budget",
            "ndt_report",
            "clients_contracts",
        ],
        "enable_hot_tapping": False,
        "inventory_mode": "shared",
    },
}

SITE_TYPE_LABELS: dict[str, str] = {
    "kdhc": "지역난방 지사형",
    "annual_contract": "연간 단가계약형",
    "smb_general": "중소업체 건별/일식형",
}

_CENTRAL_DEFAULT_COMPANIES = (
    "GS네오텍",
    "다림건설(주)",
    "한국지역난방공사",
    "한국지역난방공사 중앙지사",
)

_CENTRAL_DEFAULT_WORK_CATEGORIES: dict[str, list[str]] = {
    "다림": ["", "제조", "수탁", "건가"],
    "다림건설": ["", "제조", "수탁", "건가"],
    "다림건설(주)": ["", "제조", "수탁", "건가"],
    "GS네오텍": ["", "건가-추가", "건가-개선"],
    "GS": ["", "건가-추가", "건가-개선"],
    "지에스네오텍": ["", "건가-추가", "건가-개선"],
}

_CENTRAL_DEFAULT_ALL_CATEGORIES = ("", "제조", "수탁", "건가", "건가-추가", "건가-개선")


@dataclass(frozen=True)
class SiteProfile:
    profile_id: str
    display_name: str
    site_type: str = "kdhc"
    inventory_mode: str = "shared"  # "shared" | "isolated"
    window_title: str = ""
    description: str = ""
    enabled_tabs: tuple[str, ...] = ALL_TAB_KEYS
    companies: tuple[str, ...] = ()
    company_work_categories: dict[str, list[str]] = field(default_factory=dict)
    all_work_categories: tuple[str, ...] = ("",)
    enable_hot_tapping: bool = True
    billing_title: str = ""
    billing_overrides: dict[str, Any] = field(default_factory=dict)
    data_dir: str = ""

    @property
    def is_central(self) -> bool:
        return self.profile_id == DEFAULT_PROFILE_ID

    def is_tab_enabled(self, tab_key: str) -> bool:
        return tab_key in self.enabled_tabs

    def status_badge(self) -> str:
        inv_label = "공용재고" if self.inventory_mode == "shared" else "독립재고"
        type_label = SITE_TYPE_LABELS.get(self.site_type, self.site_type)
        return f"프로필: {self.display_name} ({type_label} · {inv_label})"


def get_profiles_root() -> str:
    src_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(src_dir, "..", "profiles"))


def get_active_profile_id() -> str:
    raw = os.environ.get(PROFILE_ENV_VAR, "").strip()
    return raw or DEFAULT_PROFILE_ID


def set_active_profile_id(profile_id: str | None) -> str:
    cleaned = (profile_id or "").strip() or DEFAULT_PROFILE_ID
    os.environ[PROFILE_ENV_VAR] = cleaned
    return cleaned


def _central_fallback_profile() -> SiteProfile:
    return SiteProfile(
        profile_id=DEFAULT_PROFILE_ID,
        display_name="한국지역난방 중앙지사",
        site_type="kdhc",
        inventory_mode="shared",
        window_title="자재 관리 프로그램 (v8 - 1원단위 완벽 보정)",
        description="한국지역난방공사 중앙지사 전용 전체 기능 프로필",
        enabled_tabs=ALL_TAB_KEYS,
        companies=_CENTRAL_DEFAULT_COMPANIES,
        company_work_categories={k: list(v) for k, v in _CENTRAL_DEFAULT_WORK_CATEGORIES.items()},
        all_work_categories=_CENTRAL_DEFAULT_ALL_CATEGORIES,
        enable_hot_tapping=True,
        billing_title="가산~가평 총 계약 수량",
        billing_overrides={},
        data_dir="",
    )


def _profile_from_dict(profile_id: str, raw: dict[str, Any]) -> SiteProfile:
    site_type = str(raw.get("site_type") or ("kdhc" if profile_id == DEFAULT_PROFILE_ID else "smb_general")).strip()
    defaults = SITE_TYPE_DEFAULTS.get(site_type, SITE_TYPE_DEFAULTS["smb_general"])

    raw_tabs = raw.get("enabled_tabs")
    if isinstance(raw_tabs, list) and raw_tabs:
        enabled_tabs = tuple(k for k in raw_tabs if k in ALL_TAB_KEYS)
        if not enabled_tabs:
            enabled_tabs = tuple(defaults["enabled_tabs"])
    else:
        enabled_tabs = tuple(defaults["enabled_tabs"])

    inv_mode = str(raw.get("inventory_mode") or defaults["inventory_mode"]).strip().lower()
    if inv_mode not in ("shared", "isolated"):
        inv_mode = "shared"

    display_name = str(raw.get("display_name") or profile_id).strip()
    window_title = str(raw.get("window_title") or f"자재·NDT 관리 프로그램 [{display_name}]").strip()

    companies_raw = raw.get("companies")
    if isinstance(companies_raw, list):
        companies = tuple(str(c).strip() for c in companies_raw if str(c).strip())
    elif profile_id == DEFAULT_PROFILE_ID:
        companies = _CENTRAL_DEFAULT_COMPANIES
    else:
        companies = ()

    cats_raw = raw.get("company_work_categories")
    if isinstance(cats_raw, dict):
        company_work_categories = {str(k): [str(v) for v in vals] for k, vals in cats_raw.items() if isinstance(vals, list)}
    elif profile_id == DEFAULT_PROFILE_ID:
        company_work_categories = {k: list(v) for k, v in _CENTRAL_DEFAULT_WORK_CATEGORIES.items()}
    else:
        company_work_categories = {}

    all_cats_raw = raw.get("all_work_categories")
    if isinstance(all_cats_raw, list) and all_cats_raw:
        all_work_categories = tuple(str(v) for v in all_cats_raw)
    elif profile_id == DEFAULT_PROFILE_ID:
        all_work_categories = _CENTRAL_DEFAULT_ALL_CATEGORIES
    else:
        all_work_categories = ("", "정기", "추가", "보수", "긴급")

    enable_hot_tapping = bool(raw.get("enable_hot_tapping", defaults["enable_hot_tapping"]))
    billing_title = str(raw.get("billing_title") or f"{display_name} 계약 수량").strip()
    billing_overrides = raw.get("billing_overrides") if isinstance(raw.get("billing_overrides"), dict) else {}
    data_dir = str(raw.get("data_dir") or "").strip()

    return SiteProfile(
        profile_id=profile_id,
        display_name=display_name,
        site_type=site_type,
        inventory_mode=inv_mode,
        window_title=window_title,
        description=str(raw.get("description") or "").strip(),
        enabled_tabs=enabled_tabs,
        companies=companies,
        company_work_categories=company_work_categories,
        all_work_categories=all_work_categories,
        enable_hot_tapping=enable_hot_tapping,
        billing_title=billing_title,
        billing_overrides=billing_overrides,
        data_dir=data_dir,
    )


def load_site_profile(profile_id: str | None = None) -> SiteProfile:
    target_id = (profile_id or "").strip() or get_active_profile_id()
    json_path = os.path.join(get_profiles_root(), target_id, "profile.json")
    if os.path.isfile(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as stream:
                raw = json.load(stream)
            if isinstance(raw, dict):
                return _profile_from_dict(target_id, raw)
        except (OSError, ValueError, TypeError):
            pass
    if target_id == DEFAULT_PROFILE_ID:
        return _central_fallback_profile()
    return _profile_from_dict(target_id, {"display_name": target_id, "site_type": "smb_general"})


def get_active_profile() -> SiteProfile:
    return load_site_profile(get_active_profile_id())


def list_site_profiles() -> list[SiteProfile]:
    root = get_profiles_root()
    profiles: list[SiteProfile] = []
    seen: set[str] = set()
    if os.path.isdir(root):
        for entry in sorted(os.listdir(root)):
            json_path = os.path.join(root, entry, "profile.json")
            if os.path.isfile(json_path):
                profiles.append(load_site_profile(entry))
                seen.add(entry)
    if DEFAULT_PROFILE_ID not in seen:
        profiles.insert(0, _central_fallback_profile())
    else:
        profiles.sort(key=lambda p: (0 if p.profile_id == DEFAULT_PROFILE_ID else 1, p.display_name))
    return profiles


def create_site_profile(
    profile_id: str,
    display_name: str,
    *,
    site_type: str = "smb_general",
    inventory_mode: str = "shared",
    companies: list[str] | None = None,
    description: str = "",
) -> SiteProfile:
    cleaned_id = re.sub(r"[^a-zA-Z0-9_\-가-힣]+", "_", (profile_id or "").strip()).strip("_")
    if not cleaned_id:
        raise ValueError("프로필 ID를 입력하세요.")
    if site_type not in SITE_TYPE_DEFAULTS:
        raise ValueError(f"지원하지 않는 사업 유형입니다: {site_type}")
    if inventory_mode not in ("shared", "isolated"):
        raise ValueError(f"지원하지 않는 재고 모드입니다: {inventory_mode}")

    defaults = SITE_TYPE_DEFAULTS[site_type]
    profile_dir = os.path.join(get_profiles_root(), cleaned_id)
    data_dir = os.path.join(profile_dir, "data")
    os.makedirs(data_dir, exist_ok=True)

    payload = {
        "profile_id": cleaned_id,
        "display_name": (display_name or cleaned_id).strip(),
        "site_type": site_type,
        "inventory_mode": inventory_mode,
        "window_title": f"자재·NDT 관리 프로그램 [{(display_name or cleaned_id).strip()}]",
        "description": description.strip(),
        "enabled_tabs": list(defaults["enabled_tabs"]),
        "enable_hot_tapping": bool(defaults["enable_hot_tapping"]),
        "companies": [str(c).strip() for c in (companies or []) if str(c).strip()],
        "company_work_categories": {},
        "all_work_categories": ["", "정기", "추가", "보수", "긴급"],
        "billing_title": f"{(display_name or cleaned_id).strip()} 계약 수량",
        "billing_overrides": {},
    }
    json_path = os.path.join(profile_dir, "profile.json")
    with open(json_path, "w", encoding="utf-8") as stream:
        json.dump(payload, stream, ensure_ascii=False, indent=2)
    return load_site_profile(cleaned_id)


def get_profile_data_dir(profile: SiteProfile | None = None) -> str | None:
    """Return isolated operational data directory for non-central profiles, or None for central."""
    active = profile or get_active_profile()
    if active.is_central:
        return None
    if active.data_dir and os.path.isabs(active.data_dir):
        target = active.data_dir
    else:
        target = os.path.join(get_profiles_root(), active.profile_id, "data")
    os.makedirs(target, exist_ok=True)
    return target


def resolve_inventory_path(profile: SiteProfile, shared_inventory_path: str) -> str:
    """Resolve Material_Inventory.xlsx path according to inventory_mode ('shared' vs 'isolated')."""
    if profile.is_central or profile.inventory_mode == "shared":
        return shared_inventory_path
    data_dir = get_profile_data_dir(profile)
    assert data_dir is not None
    return os.path.join(data_dir, "Material_Inventory.xlsx")


def resolve_config_path(profile: SiteProfile, default_config_path: str) -> str:
    if profile.is_central:
        return default_config_path
    data_dir = get_profile_data_dir(profile)
    assert data_dir is not None
    return os.path.join(data_dir, "Material_Manager_Config.json")


def resolve_billing_config_path(profile: SiteProfile, default_billing_config_path: str) -> str:
    if profile.is_central:
        return default_billing_config_path
    data_dir = get_profile_data_dir(profile)
    assert data_dir is not None
    return os.path.join(data_dir, "billing_config.json")


def resolve_billing_export_path(profile: SiteProfile, default_billing_export_path: str) -> str:
    if profile.is_central:
        return default_billing_export_path
    data_dir = get_profile_data_dir(profile)
    assert data_dir is not None
    return os.path.join(data_dir, "billing_export.json")


def resolve_report_settings_path(profile: SiteProfile, default_report_settings_path: str) -> str:
    if profile.is_central:
        return default_report_settings_path
    data_dir = get_profile_data_dir(profile)
    assert data_dir is not None
    return os.path.join(data_dir, "report_settings.json")


def pick_profile_dialog() -> str | None:
    """Show a startup launcher dialog to select or create a business profile."""
    import tkinter as tk
    from tkinter import ttk, messagebox

    result: dict[str, str | None] = {"profile_id": None}
    picker = tk.Tk()
    picker.title("사업장 / 계약 프로필 선택")
    picker.geometry("620x420")
    picker.minsize(560, 360)

    container = ttk.Frame(picker, padding=14)
    container.pack(fill="both", expand=True)

    ttk.Label(
        container,
        text="실행할 사업장 / 계약 프로필을 선택하세요",
        font=("Malgun Gothic", 12, "bold"),
    ).pack(anchor="w")
    ttk.Label(
        container,
        text="회사 공용 재고(필름·약품·장비)는 공유하면서 작업·정산·성적서는 프로필별로 분리 관리됩니다.",
        foreground="#555555",
    ).pack(anchor="w", pady=(2, 10))

    cols = ("profile_id", "display_name", "site_type", "inventory_mode", "tabs")
    tree = ttk.Treeview(container, columns=cols, show="headings", height=9, selectmode="browse")
    tree.heading("profile_id", text="ID")
    tree.heading("display_name", text="사업장 / 계약명")
    tree.heading("site_type", text="사업 유형")
    tree.heading("inventory_mode", text="재고 방식")
    tree.heading("tabs", text="활성 탭 수")
    tree.column("profile_id", width=110, anchor="center")
    tree.column("display_name", width=180, anchor="w")
    tree.column("site_type", width=130, anchor="center")
    tree.column("inventory_mode", width=80, anchor="center")
    tree.column("tabs", width=75, anchor="center")
    tree.pack(fill="both", expand=True)

    desc_var = tk.StringVar(value="")
    ttk.Label(container, textvariable=desc_var, foreground="#333333", wraplength=580).pack(anchor="w", pady=(8, 6))

    profiles_by_id: dict[str, SiteProfile] = {}

    def refresh_list(select_id: str | None = None):
        tree.delete(*tree.get_children())
        profiles_by_id.clear()
        items = list_site_profiles()
        for prof in items:
            profiles_by_id[prof.profile_id] = prof
            tree.insert(
                "",
                "end",
                iid=prof.profile_id,
                values=(
                    prof.profile_id,
                    prof.display_name,
                    SITE_TYPE_LABELS.get(prof.site_type, prof.site_type),
                    "공용재고" if prof.inventory_mode == "shared" else "독립재고",
                    f"{len(prof.enabled_tabs)}개",
                ),
            )
        target = select_id if select_id in profiles_by_id else (items[0].profile_id if items else None)
        if target:
            tree.selection_set(target)
            tree.focus(target)
            on_select()

    def on_select(_event=None):
        sel = tree.selection()
        if not sel:
            desc_var.set("")
            return
        prof = profiles_by_id.get(sel[0])
        desc_var.set(prof.description if prof else "")

    def launch_selected(_event=None):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("선택 필요", "실행할 프로필을 선택하세요.", parent=picker)
            return
        result["profile_id"] = sel[0]
        picker.destroy()

    def open_create_dialog():
        dlg = tk.Toplevel(picker)
        dlg.title("새 사업장 / 계약 프로필 추가")
        dlg.geometry("420x270")
        dlg.transient(picker)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)

        id_var = tk.StringVar()
        name_var = tk.StringVar()
        type_var = tk.StringVar(value="smb_general")
        inv_var = tk.StringVar(value="shared")
        desc_in = tk.StringVar()

        ttk.Label(frm, text="프로필 ID (영문/숫자):").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=id_var, width=28).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="표시 이름:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=name_var, width=28).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="사업 유형:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm,
            textvariable=type_var,
            values=("smb_general", "annual_contract", "kdhc"),
            state="readonly",
            width=25,
        ).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="재고 관리 방식:").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm,
            textvariable=inv_var,
            values=("shared", "isolated"),
            state="readonly",
            width=25,
        ).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="설명:").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=desc_in, width=28).grid(row=4, column=1, sticky="ew", pady=4)
        frm.columnconfigure(1, weight=1)

        def do_create():
            try:
                created = create_site_profile(
                    id_var.get(),
                    name_var.get() or id_var.get(),
                    site_type=type_var.get(),
                    inventory_mode=inv_var.get(),
                    description=desc_in.get(),
                )
            except Exception as err:
                messagebox.showerror("생성 오류", str(err), parent=dlg)
                return
            dlg.destroy()
            refresh_list(select_id=created.profile_id)

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="취소", command=dlg.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="프로필 생성", command=do_create).pack(side="right")

    tree.bind("<<TreeviewSelect>>", on_select)
    tree.bind("<Double-1>", launch_selected)

    btn_row = ttk.Frame(container)
    btn_row.pack(fill="x", pady=(6, 0))
    ttk.Button(btn_row, text="+ 새 프로필 추가", command=open_create_dialog).pack(side="left")
    ttk.Button(btn_row, text="취소", command=picker.destroy).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text="선택한 프로필로 실행", command=launch_selected).pack(side="right")

    refresh_list(select_id=get_active_profile_id())
    picker.mainloop()
    return result["profile_id"]

