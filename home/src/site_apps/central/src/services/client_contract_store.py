"""Per-profile Client & Contract master store for annual contracts and SMB operations."""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime
import json
import os
import uuid
from typing import Any

from site_apps.central.src.site_profile import SiteProfile, get_active_profile, get_profile_data_dir

CONTRACT_TYPE_LABELS: dict[str, str] = {
    "annual_unit": "연간 단가계약",
    "per_job": "건별 / 일식",
    "kdhc": "지역난방 지사",
}


def get_client_contract_store_path(profile: SiteProfile | None = None) -> str:
    active = profile or get_active_profile()
    profile_dir = get_profile_data_dir(active)
    if profile_dir:
        return os.path.join(profile_dir, "clients_contracts.json")
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "data"))
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, "clients_contracts.json")


def _empty_store() -> dict[str, Any]:
    return {"version": 1, "clients": [], "contracts": []}


def load_client_contract_store(profile: SiteProfile | None = None) -> dict[str, Any]:
    path = get_client_contract_store_path(profile)
    if not os.path.isfile(path):
        return _empty_store()
    try:
        with open(path, "r", encoding="utf-8") as stream:
            data = json.load(stream)
        if not isinstance(data, dict):
            return _empty_store()
        data.setdefault("version", 1)
        data.setdefault("clients", [])
        data.setdefault("contracts", [])
        return data
    except (OSError, ValueError, TypeError):
        return _empty_store()


def save_client_contract_store(store: dict[str, Any], profile: SiteProfile | None = None) -> str:
    path = get_client_contract_store_path(profile)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp_path = path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as stream:
        json.dump(store, stream, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)
    return path


def upsert_client(
    client_data: dict[str, Any],
    *,
    profile: SiteProfile | None = None,
) -> dict[str, Any]:
    name = str(client_data.get("name") or "").strip()
    if not name:
        raise ValueError("거래처명을 입력하세요.")

    store = load_client_contract_store(profile)
    client_id = str(client_data.get("client_id") or "").strip() or f"CLI-{uuid.uuid4().hex[:8].upper()}"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    record = {
        "client_id": client_id,
        "name": name,
        "biz_no": str(client_data.get("biz_no") or "").strip(),
        "ceo": str(client_data.get("ceo") or "").strip(),
        "contact_name": str(client_data.get("contact_name") or "").strip(),
        "contact_phone": str(client_data.get("contact_phone") or "").strip(),
        "contact_email": str(client_data.get("contact_email") or "").strip(),
        "address": str(client_data.get("address") or "").strip(),
        "payment_terms": str(client_data.get("payment_terms") or "월말 마감 익월 15일").strip(),
        "note": str(client_data.get("note") or "").strip(),
        "updated_at": now,
    }

    replaced = False
    for idx, existing in enumerate(store["clients"]):
        if existing.get("client_id") == client_id:
            store["clients"][idx] = record
            replaced = True
            break
    if not replaced:
        store["clients"].append(record)

    save_client_contract_store(store, profile)
    return record


def delete_client(client_id: str, *, profile: SiteProfile | None = None) -> bool:
    store = load_client_contract_store(profile)
    before = len(store["clients"])
    store["clients"] = [c for c in store["clients"] if c.get("client_id") != client_id]
    if len(store["clients"]) == before:
        return False
    store["contracts"] = [ct for ct in store["contracts"] if ct.get("client_id") != client_id]
    save_client_contract_store(store, profile)
    return True


def _normalize_rate_items(rate_items: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for idx, item in enumerate(rate_items or [], 1):
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        code = str(item.get("code") or f"ITEM-{idx:02d}").strip()
        try:
            unit_price = int(round(float(str(item.get("unit_price") or 0).replace(",", ""))))
        except ValueError:
            unit_price = 0
        try:
            contract_qty = float(str(item.get("contract_qty") or 0).replace(",", ""))
        except ValueError:
            contract_qty = 0.0
        normalized.append(
            {
                "code": code,
                "name": name,
                "method": str(item.get("method") or "").strip().upper(),
                "spec": str(item.get("spec") or "").strip(),
                "unit": str(item.get("unit") or "EA").strip(),
                "unit_price": unit_price,
                "contract_qty": contract_qty,
            }
        )
    return normalized


def upsert_contract(
    contract_data: dict[str, Any],
    *,
    profile: SiteProfile | None = None,
) -> dict[str, Any]:
    contract_name = str(contract_data.get("contract_name") or "").strip()
    if not contract_name:
        raise ValueError("계약명을 입력하세요.")
    client_id = str(contract_data.get("client_id") or "").strip()
    if not client_id:
        raise ValueError("거래처를 선택하세요.")

    store = load_client_contract_store(profile)
    contract_id = str(contract_data.get("contract_id") or "").strip() or f"CTR-{uuid.uuid4().hex[:8].upper()}"
    contract_type = str(contract_data.get("contract_type") or "annual_unit").strip()
    if contract_type not in CONTRACT_TYPE_LABELS:
        contract_type = "annual_unit"

    def _to_int(val: Any) -> int:
        try:
            return int(round(float(str(val or 0).replace(",", ""))))
        except ValueError:
            return 0

    record = {
        "contract_id": contract_id,
        "client_id": client_id,
        "contract_name": contract_name,
        "contract_type": contract_type,
        "start_date": str(contract_data.get("start_date") or "").strip(),
        "end_date": str(contract_data.get("end_date") or "").strip(),
        "budget_amount": _to_int(contract_data.get("budget_amount")),
        "default_trip_fee": _to_int(contract_data.get("default_trip_fee")),
        "default_lump_sum": _to_int(contract_data.get("default_lump_sum")),
        "rate_items": _normalize_rate_items(contract_data.get("rate_items")),
        "note": str(contract_data.get("note") or "").strip(),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    replaced = False
    for idx, existing in enumerate(store["contracts"]):
        if existing.get("contract_id") == contract_id:
            store["contracts"][idx] = record
            replaced = True
            break
    if not replaced:
        store["contracts"].append(record)

    save_client_contract_store(store, profile)
    return record


def delete_contract(contract_id: str, *, profile: SiteProfile | None = None) -> bool:
    store = load_client_contract_store(profile)
    before = len(store["contracts"])
    store["contracts"] = [c for c in store["contracts"] if c.get("contract_id") != contract_id]
    if len(store["contracts"]) == before:
        return False
    save_client_contract_store(store, profile)
    return True


def get_client_by_id(client_id: str, *, profile: SiteProfile | None = None) -> dict[str, Any] | None:
    store = load_client_contract_store(profile)
    for client in store["clients"]:
        if client.get("client_id") == client_id:
            return deepcopy(client)
    return None


def get_contract_by_id(contract_id: str, *, profile: SiteProfile | None = None) -> dict[str, Any] | None:
    store = load_client_contract_store(profile)
    for contract in store["contracts"]:
        if contract.get("contract_id") == contract_id:
            return deepcopy(contract)
    return None

