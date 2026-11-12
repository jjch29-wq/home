"""Single-source-of-truth synchronization from PC Gyeongin masters to Google Apps Script assets."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List

from site_apps.gyeongin.src.aux_equipment import METHOD_AUX_MAP, METHOD_MAIN_EQUIP_MAP
from site_apps.gyeongin.src.company_master import (
    DEFAULT_COMPANIES,
    load_company_ndt_presets,
    merge_company_names,
)

GYEONGIN_ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = Path(__file__).resolve().parent
CONFIG_JSON_PATH = SRC_DIR / "Material_Manager_Config.json"
GAS_DIR = GYEONGIN_ROOT / "google_apps_script"
CODE_GS_PATH = GAS_DIR / "Code.gs"
INDEX_HTML_PATH = GAS_DIR / "index.html"

METHOD_DEFAULT_META: Dict[str, Dict[str, str]] = {
    "PMI": {"default_aux": ""},
    "RT": {"default_aux": "ADR, SURVEY, 콜리메타", "unit": "매"},
    "UT": {"default_aux": "사각 PROBE(70˚) MWB70-2 (56923-5617); 표준시험편 STB-A1 (09148)", "unit": "M"},
    "PT": {"default_aux": "조도계 LUX Meter 1330A (180604411)", "unit": "M"},
    "MT": {"default_aux": "자분살포기", "unit": "M"},
    "PAUT": {"default_aux": "PAUT SCANNER (MANUAL)", "unit": "M"},
}

DEFAULT_USERS = [
    "이사 곽재운", "이사 박홍규",
    "부장 주진철", "부장 박광복", "부장 이병국", "부장 이용주", "부장 주동기",
    "부장 김규년", "부장 곽상돈", "부장 정효철", "부장 노현태",
    "차장 김춘호", "차장 권인경", "차장 박원준",
    "과장 주영광", "과장 유상훈",
    "대리 김상현", "대리 윤석현", "대리 우명광", "대리 박만영",
    "계장 박상태", "계장 이준우", "계장 장승대", "계장 이은수",
    "계장 손주언", "계장 이경재", "계장 박재욱", "계장 한지원",
    "주임 김진환", "주임 황조현", "주임 김태현",
]

DEFAULT_VEHICLES = [
    "84저1259 (스타렉스)", "95가0200 (스타렉스)", "91주8839 (스타렉스)",
    "90서8616 (스타렉스)", "76마3422 (스타렉스 12인승)", "71고4405 (스타렉스 12인승)",
    "89보4028 (탑차)", "81두1580 (탑차)", "81루5100 (스타렉스)",
    "95가0175 (스타렉스)", "81도5958 (스타렉스)", "90너4889 (탑차)", "자가용/대중교통",
]

DEFAULT_MATERIALS = [
    'Carestream AA400-3⅓*12"', 'Carestream AA400-3⅓*17"', 'Carestream AA400-4½*12"',
    'Carestream AA400-10*12"', 'Carestream M100-3⅓*12"', 'Carestream M100-10*12"',
    'Carestream M100-14*17"', 'Carestream MX125-3⅓*6"', 'Carestream MX125-3⅓*12"',
    'Carestream MX125-4½*12"', 'Carestream MX125-10*12"', 'Carestream MX125-14*17"',
    'Carestream T200-3⅓*12"', 'Carestream T200-3⅓*17"', 'Carestream T200-4½*12"',
    'Carestream T200-10*12"', 'Carestream T200-14*17"',
]


def _unique_ordered(seqs: Iterable[Iterable[str]], trailing: Iterable[str] = ()) -> List[str]:
    seen = set()
    result: List[str] = []
    for seq in seqs:
        for item in seq:
            txt = str(item or "").strip()
            if txt and txt not in seen:
                seen.add(txt)
                result.append(txt)
    for item in trailing:
        txt = str(item or "").strip()
        if txt and txt not in seen:
            seen.add(txt)
            result.append(txt)
    return result


def _load_manager_config(config_path: Path = CONFIG_JSON_PATH) -> Dict[str, Any]:
    if not config_path.exists():
        return {}
    return json.loads(config_path.read_text(encoding="utf-8"))


def _indent_json(data: Any, base_indent: str) -> str:
    rendered = json.dumps(data, ensure_ascii=False, indent=2)
    lines = rendered.splitlines()
    return "\n".join([lines[0]] + [f"{base_indent}{line}" for line in lines[1:]])


def build_gas_master_blocks(config_path: Path = CONFIG_JSON_PATH) -> Dict[str, str]:
    """Build exact synchronized JS blocks for Code.gs and google_apps_script/index.html."""
    cfg = _load_manager_config(config_path)
    companies = merge_company_names(DEFAULT_COMPANIES, cfg.get("companies") or [])
    users = list(cfg.get("users") or DEFAULT_USERS)
    vehicles = list(cfg.get("vehicles") or DEFAULT_VEHICLES)
    materials = list(cfg.get("materials") or DEFAULT_MATERIALS)

    main_order = ["PAUT", "UT", "RT", "MT", "PT", "PMI"]
    equipments = _unique_ordered(
        (METHOD_MAIN_EQUIP_MAP.get(m, ()) for m in main_order),
        trailing=["해당없음"],
    )
    aux_order = ["PAUT", "UT", "RT", "MT", "PT"]
    aux_equipments = _unique_ordered(
        (METHOD_AUX_MAP.get(m, ()) for m in aux_order),
        trailing=["없음"],
    )

    method_defaults: Dict[str, Dict[str, Any]] = {}
    for method in ("PMI", "RT", "UT", "PT", "MT", "PAUT"):
        meta = METHOD_DEFAULT_META[method]
        main_list = METHOD_MAIN_EQUIP_MAP.get(method, [])
        aux_list = _unique_ordered([METHOD_AUX_MAP.get(method, ())], trailing=["없음"]) if method in METHOD_AUX_MAP else []
        entry: Dict[str, Any] = {
            "equip_main": main_list[0] if main_list else "해당없음",
            "aux_equipments": aux_list,
            "default_aux": meta["default_aux"],
        }
        if "unit" in meta:
            entry["unit"] = meta["unit"]
        method_defaults[method] = entry

    presets = load_company_ndt_presets()

    core_parts = [
        f"  companies: {_indent_json(companies, '  ')},",
        f"  users: {_indent_json(users, '  ')},",
        f"  vehicles: {_indent_json(vehicles, '  ')},",
        f"  equipments: {_indent_json(equipments, '  ')},",
        f"  aux_equipments: {_indent_json(aux_equipments, '  ')},",
        f"  method_defaults: {_indent_json(method_defaults, '  ')},",
    ]
    aux_map = {m: list(METHOD_AUX_MAP.get(m, ())) for m in ("PAUT", "UT", "RT", "MT", "PT")}

    return {
        "APP_CONFIG_CORE": "\n".join(core_parts),
        "APP_CONFIG_MATERIALS": f"  materials: {_indent_json(materials, '  ')},",
        "COMPANY_NDT_PRESETS": f"var COMPANY_NDT_PRESETS = {json.dumps(presets, ensure_ascii=False)};",
        "MAIN_EQUIP_BY_METHOD": f"    const MAIN_EQUIP_BY_METHOD = {json.dumps(METHOD_MAIN_EQUIP_MAP, ensure_ascii=False)};",
        "AUX_MASTER_DATA": f"    const AUX_MASTER_DATA = {_indent_json(aux_map, '    ')};",
    }


def replace_sync_block(text: str, block_name: str, replacement: str) -> str:
    """Replace content between // @SYNC:BEGIN <name> and // @SYNC:END <name> markers."""
    pattern = re.compile(
        rf"([ \t]*// @SYNC:BEGIN {re.escape(block_name)}\r?\n)([\s\S]*?)(\r?\n[ \t]*// @SYNC:END {re.escape(block_name)})"
    )
    match = pattern.search(text)
    if not match:
        raise ValueError(f"Sync marker not found for block '{block_name}'")
    newline = "\r\n" if "\r\n" in match.group(1) else "\n"
    normalized_body = replacement.replace("\r\n", "\n").replace("\n", newline)
    return text[: match.start(2)] + normalized_body + text[match.end(2) :]


def sync_gas_masters(
    *,
    code_gs_path: Path = CODE_GS_PATH,
    index_html_path: Path = INDEX_HTML_PATH,
    config_path: Path = CONFIG_JSON_PATH,
    write: bool = True,
) -> Dict[str, Any]:
    """Synchronize master data blocks from PC configs into Code.gs and index.html."""
    blocks = build_gas_master_blocks(config_path=config_path)
    targets = (
        (code_gs_path, ("APP_CONFIG_CORE", "APP_CONFIG_MATERIALS", "COMPANY_NDT_PRESETS")),
        (index_html_path, ("MAIN_EQUIP_BY_METHOD", "AUX_MASTER_DATA")),
    )

    changed_files: List[str] = []
    for path, keys in targets:
        original = path.read_text(encoding="utf-8")
        updated = original
        for key in keys:
            updated = replace_sync_block(updated, key, blocks[key])
        if updated != original:
            changed_files.append(str(path))
            if write:
                path.write_text(updated, encoding="utf-8", newline="")

    return {
        "status": "updated" if (write and changed_files) else ("drifted" if changed_files else "in_sync"),
        "changed_files": changed_files,
        "blocks": list(blocks.keys()),
    }


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Sync Gyeongin PC masters into Google Apps Script assets.")
    parser.add_argument("--check", action="store_true", help="Check parity without writing files; exit 1 if drifted.")
    args = parser.parse_args(argv)

    result = sync_gas_masters(write=not args.check)
    print(f"[GAS Sync] status={result['status']}, changed={len(result['changed_files'])}")
    for path_str in result["changed_files"]:
        print(f"  - {path_str}")
    if args.check and result["changed_files"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

