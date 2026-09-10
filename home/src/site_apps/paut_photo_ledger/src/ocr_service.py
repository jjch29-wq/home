from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

import numpy as np
from PIL import Image, ImageOps


_reader = None


def _get_reader():
    global _reader
    if _reader is None:
        import easyocr

        _reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    return _reader


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip()).upper()


def _extract_size(lines: list[str]) -> str:
    candidates: list[str] = []
    for line in lines:
        compact = line.replace(" ", "").replace(",", ".")
        explicit = re.search(r"[ØΦ@O](\d{0,2}\.\d{1,4})", compact)
        if explicit:
            candidates.append(explicit.group(1))
            continue
        plain = re.fullmatch(r"0?(\d{1,2}\.\d{1,4})(?:IN)?", compact)
        if plain:
            value = plain.group(1)
            if value not in {"2.5"}:
                candidates.append(value)
    if not candidates:
        return ""
    value = candidates[-1].lstrip("0")
    if value.startswith("."):
        value = "0" + value
    return f"Ø{value} IN"


def _extract_model(lines: list[str]) -> str:
    known = []
    for line in lines:
        compact = re.sub(r"[^A-Z0-9.\-/]", "", line)
        if "SA15" in compact and "N60S" in compact:
            suffix = "-AOD" if "AOD" in compact else ""
            return f"SA15-N60S-IH{suffix}"
        if "5DL16" in compact or ("A25" in compact and "12X5" in compact):
            return "5DL16-12X5-A25-P-2.5-OM"
        if compact.count("-") >= 2 and any(ch.isalpha() for ch in compact) and any(ch.isdigit() for ch in compact):
            known.append(compact)
    return max(known, key=len, default="")


def _extract_serial(lines: list[str]) -> str:
    joined = " | ".join(lines)
    match = re.search(r"(?:S/?I?N|SERIAL)\s*[:#-]?\s*([A-Z0-9-]{3,})", joined)
    return match.group(1) if match else ""


def analyze_photo(path: Path) -> dict:
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
        image.thumbnail((1600, 1600))
        array = np.asarray(image)
    raw = _get_reader().readtext(array, detail=0, paragraph=False, canvas_size=1600)
    lines = [_clean(str(text)) for text in raw if str(text).strip()]
    model = _extract_model(lines)
    size = _extract_size(lines)
    serial = _extract_serial(lines)
    joined = " ".join(lines)
    if "PROBE" in joined or model.startswith("5DL16"):
        kind = "PAUT Probe"
    elif "WEDGE" in joined or model.startswith("SA15") or size:
        kind = "Wedge"
    else:
        kind = "확인 필요"
    return {"path": path, "kind": kind, "model": model, "size": size, "serial": serial, "ocr_text": " | ".join(lines)}


def analyze_and_group(app_dir: Path, photo_refs: list[str], progress=None) -> list[dict]:
    results = []
    total = len(photo_refs)
    for index, ref in enumerate(photo_refs, 1):
        try:
            result = analyze_photo(app_dir / ref)
        except Exception as exc:
            result = {"path": app_dir / ref, "kind": "확인 필요", "model": "", "size": "", "serial": "", "ocr_text": f"인식 오류: {exc}"}
        result["ref"] = ref
        results.append(result)
        if progress:
            progress(index, total)

    wedge_models = [r["model"] for r in results if r["kind"] == "Wedge" and r["model"].startswith("SA15-")]
    common_wedge_model = Counter(wedge_models).most_common(1)[0][0] if wedge_models else ""
    for result in results:
        if result["kind"] == "Wedge" and not result["model"] and common_wedge_model:
            result["model"] = common_wedge_model

    # 글자가 거의 없는 근접 촬영 사진은 전후 2장 안의 확정 품목에 보조 사진으로 연결한다.
    for index, result in enumerate(results):
        if result["kind"] != "확인 필요" or len(result["ocr_text"].split(" | ")) > 2:
            continue
        neighbours = []
        for other_index, other in enumerate(results):
            distance = abs(index - other_index)
            if 0 < distance <= 2 and other["kind"] != "확인 필요":
                neighbours.append((distance, other_index, other))
        if neighbours:
            _, _, nearest = min(neighbours, key=lambda value: (value[0], value[1]))
            for key in ("kind", "model", "size", "serial"):
                result[key] = nearest[key]

    groups: dict[str, dict] = {}
    unknown_no = 0
    for result in results:
        if result["size"]:
            key = f"{result['model']}|{result['size']}"
        elif result["model"]:
            key = f"{result['model']}|"
        else:
            unknown_no += 1
            key = f"UNKNOWN-{unknown_no}"
        group = groups.setdefault(key, {
            "kind": result["kind"], "model": result["model"], "size": result["size"],
            "serial": result["serial"], "quantity": "", "note": "", "photos": [], "ocr_texts": []
        })
        group["photos"].append(result["ref"])
        if result["ocr_text"]:
            group["ocr_texts"].append(result["ocr_text"])
        if not group["serial"] and result["serial"]:
            group["serial"] = result["serial"]

    output = []
    for group in groups.values():
        confidence = "자동 인식 결과 - 확인 후 저장"
        if not group["model"] and not group["size"] and not group["serial"]:
            confidence = "자동 인식 불확실 - 직접 확인 필요"
        group["note"] = confidence
        group["ocr_text"] = "\n".join(group.pop("ocr_texts"))
        output.append(group)
    return output
