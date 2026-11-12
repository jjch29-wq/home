from functools import lru_cache
import json
from pathlib import Path

"""Default Gyeongin Office client company master list (extracted from Naver MYBOX archive)."""

DEFAULT_COMPANIES = (
    "경인사무소",
    "D&K켐텍",
    "K&M",
    "SG엔지니어링",
    "SK에코플랜트",
    "경성기공",
    "글로벌 크레인",
    "나라삼양감속기",
    "대성기계",
    "대성테크놀로지",
    "대화유공업",
    "대화콘베어체인",
    "도원이앤씨",
    "동남ENG",
    "동아기업DAC",
    "동해펌프",
    "디에이치건설산업(주)",
    "롯데건설",
    "리니트코리아",
    "모나스펌프",
    "부강풀리",
    "삼남엔지니어링(주)",
    "삼성파워텍",
    "삼승종합건설",
    "삼화테크",
    "서울발전본부",
    "서진인스텍",
    "성헌",
    "세진공조기계",
    "수성웹툰(주)",
    "스페이스 체인지",
    "신성플랜트",
    "아이조음",
    "아이피에스",
    "양영",
    "에버랜드",
    "에스엠테크",
    "에이스냉동공조",
    "영흥발전",
    "올스톤산업",
    "위드엔지니어링",
    "유성씨앤에프",
    "의성기어",
    "이노텍코리아",
    "자인건축",
    "전진이앤씨",
    "제이에스티",
    "제일푸리",
    "청원기연",
    "캐논코리아",
    "캐리비안베이",
    "케이엔엠",
    "코아베큠",
    "코오롱글로벌",
    "태성기전",
    "태창기계공업",
    "포스코이앤씨",
    "포스코인터내셔널",
    "피테크",
    "한국산업인력공단",
    "한양스틸",
    "함감속기제작소",
    "현대건설",
    "현대산업",
    "현성펌프",
    "현진산업",
)


def merge_company_names(*sources):
    """Return unique company names preserving '경인사무소' at index 0."""
    seen = {"경인사무소"}
    others = []
    for source in sources:
        if not source:
            continue
        for raw in source:
            name = str(raw or "").strip()
            if not name or name.lower() == "nan" or name in seen:
                continue
            seen.add(name)
            others.append(name)
    return ["경인사무소"] + sorted(others)


def _presets_file() -> Path:
    return Path(__file__).resolve().parent.parent / "resources" / "company_ndt_presets.json"


@lru_cache(maxsize=2)
def load_company_ndt_presets(presets_path: str | None = None) -> dict:
    """Load bundled per-company × inspection-method presets."""
    target = Path(presets_path) if presets_path else _presets_file()
    if not target.exists():
        return {}
    return json.loads(target.read_text(encoding="utf-8"))


def suggest_company_method_preset(company: str, method: str = "", presets: dict | None = None) -> dict:
    """Return suggested target_type, target_types, items, and default_code for (company, method)."""
    empty = {"target_type": "", "target_types": [], "items": [], "default_code": ""}
    co = str(company or "").strip()
    if not co:
        return empty
    source = presets if presets is not None else load_company_ndt_presets()
    company_map = source.get(co)
    if not isinstance(company_map, dict) or not company_map:
        return empty

    m = str(method or "").strip().upper()
    if m and m in company_map:
        entry = company_map[m]
        return {
            "target_type": str(entry.get("target_type", "") or ""),
            "target_types": list(entry.get("target_types", ())),
            "items": list(entry.get("items", ())),
            "default_code": str(entry.get("default_code", "") or ""),
        }
    if m == "PAUT" and "UT" in company_map:
        entry = company_map["UT"]
        return {
            "target_type": str(entry.get("target_type", "") or ""),
            "target_types": list(entry.get("target_types", ())),
            "items": list(entry.get("items", ())),
            "default_code": str(entry.get("default_code", "") or ""),
        }

    # Combine presets across all recorded methods for the company when method is unset/unmatched
    targets, items = [], []
    default_code = ""
    for entry in company_map.values():
        for t in entry.get("target_types", ()):
            if t and t not in targets:
                targets.append(t)
        for it in entry.get("items", ()):
            if it and it not in items:
                items.append(it)
        if not default_code and entry.get("default_code"):
            default_code = str(entry["default_code"])
    return {
        "target_type": targets[0] if targets else "",
        "target_types": targets,
        "items": items,
        "default_code": default_code,
    }


def _unique_nonempty(*sequences) -> list[str]:
    merged = []
    seen = set()
    for seq in sequences:
        if not seq:
            continue
        for raw in seq:
            txt = str(raw or "").strip()
            if txt and txt not in seen:
                seen.add(txt)
                merged.append(txt)
    return merged


def merge_company_items(
    company: str,
    method: str = "",
    history_items=(),
    fallback_items=(),
    presets: dict | None = None,
) -> list[str]:
    """Return ONLY the company's items when the company has presets/history; else fallback."""
    preset = suggest_company_method_preset(company, method, presets=presets)
    company_items = preset.get("items", ())
    if company_items or history_items:
        return _unique_nonempty(company_items, history_items)
    return _unique_nonempty(fallback_items)


METHOD_REPORT_MATERIALS = {
    "MT": (
        "MT약품 (백색페인트)",
        "MT약품 (흑색자분)",
        "MT약품 (형광자분)",
        "NAWOO 습식형광자분 (Wet Fluorescent)",
        "GS CHEM 형광자분",
    ),
    "PT": (
        "PT약품 (침투제)",
        "PT약품 (세척제)",
        "PT약품 (현상제)",
        "PT약품 (형광침투제)",
        "NAWOO NPP-2 / NPR1-3 / NPD-4 세트",
    ),
    "UT": (
        "접촉매질 (Oil)",
        "접촉매질 (Glycerin)",
        "접촉매질 (C.M.C)",
        "표준시험편 (STB-A1 / IIW)",
    ),
    "PAUT": (
        "접촉매질 (물 / 글리세린)",
        "PAUT 표준시험편 (ISO19675 / STB-A1)",
    ),
    "PMI": (
        "표준시편 (Alloy Reference / 1.25Cr 0.5Mo)",
    ),
}


def _material_matches_method(text: str, method: str) -> bool:
    u = str(text or "").upper()
    if method == "RT":
        return any(k in u for k in ("CARESTREAM", "T200", "AA400", "M100", "MX125", "FILM", "필름"))
    if method == "MT":
        return "MT약품" in u or "자분" in text or "백색페인트" in text or "형광" in text and "침투" not in text
    if method == "PT":
        return "PT약품" in u or any(k in text for k in ("침투제", "세척제", "현상제"))
    if method == "UT":
        return ("표준시험편" in text or "IIW" in u or "STB" in u) and "PAUT" not in u
    if method == "PAUT":
        return "PAUT" in u and ("BLOCK" in u or "시험편" in text or "ISO19675" in u)
    return False


def method_material_candidates(method: str = "", inventory_materials=()) -> list[str]:
    """Return method-specific consumable/media material names prioritizing inventory matches."""
    m = str(method or "").strip().upper()
    inv = [str(v or "").strip() for v in (inventory_materials or ()) if str(v or "").strip()]
    if m in ("내근/대기", "휴무", "기타"):
        return ["해당없음"]

    ordered = []
    seen = set()

    def add_all(seq):
        for item in seq:
            txt = str(item or "").strip()
            if txt and txt not in seen:
                seen.add(txt)
                ordered.append(txt)

    if m in ("RT", "MT", "PT", "UT", "PAUT", "PMI"):
        add_all(v for v in inv if _material_matches_method(v, m))
        add_all(METHOD_REPORT_MATERIALS.get(m, ()))
        add_all(["해당없음"])
    else:
        for method_key in ("RT", "MT", "PT", "UT", "PAUT", "PMI"):
            add_all(METHOD_REPORT_MATERIALS.get(method_key, ()))
        add_all(inv)
        add_all(["해당없음"])
    return ordered


def method_inspection_item_candidates(
    company: str = "",
    method: str = "",
    history_items=(),
    fallback_items=(),
    presets: dict | None = None,
) -> list[str]:
    """Return ONLY the company's items (for method or all company methods) when known; else method/defaults."""
    source = presets if presets is not None else load_company_ndt_presets()
    m = str(method or "").strip().upper()
    company_preset = suggest_company_method_preset(company, m, presets=source)
    company_items = company_preset.get("items", ())

    if company_items or history_items:
        return _unique_nonempty(company_items, history_items)

    method_items = []
    if m:
        lookup_methods = ("UT", "PAUT") if m == "PAUT" else (m,)
        for mmap in source.values():
            if not isinstance(mmap, dict):
                continue
            for mk in lookup_methods:
                for it in mmap.get(mk, {}).get("items", ()):
                    if it:
                        method_items.append(it)

    return _unique_nonempty(method_items, fallback_items)

