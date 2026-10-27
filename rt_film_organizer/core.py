from __future__ import annotations

import csv
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass, asdict
from pathlib import Path

import cv2
import numpy as np


IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class FilmRecord:
    source: str
    raw_text: str = ""
    line_no: str = ""
    joint_no: str = ""
    welder_no: str = ""
    section: str = ""
    section_source: str = ""
    sequence: int = 1
    output_name: str = ""
    confidence: float = 0.0
    status: str = "대기"


def imread(path: str | Path) -> np.ndarray:
    data = np.fromfile(str(path), dtype=np.uint8)
    image = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError(f"이미지를 읽을 수 없습니다: {path}")
    return image


def imwrite_jpg(path: str | Path, image: np.ndarray, quality: int = 95) -> None:
    ok, encoded = cv2.imencode(".jpg", image, [cv2.IMWRITE_JPEG_QUALITY, quality])
    if not ok:
        raise ValueError(f"JPG 변환에 실패했습니다: {path}")
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    encoded.tofile(str(path))


def list_images(folder: str | Path) -> list[Path]:
    root = Path(folder)
    return sorted(p for p in root.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS)


def detect_film_rect(image: np.ndarray) -> tuple[int, int, int, int]:
    """Detect the illuminated film rectangle in a dark, fixed-camera photograph."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    small_scale = min(1.0, 1400.0 / max(gray.shape))
    small = cv2.resize(gray, None, fx=small_scale, fy=small_scale, interpolation=cv2.INTER_AREA)
    blurred = cv2.GaussianBlur(small, (9, 9), 0)

    # A low percentile-derived threshold keeps dense film regions while rejecting the dark room.
    bright = blurred[blurred > 5]
    base = float(np.percentile(bright, 55)) if bright.size else 25.0
    threshold = max(18, min(100, int(base * 0.55)))
    mask = np.where(blurred >= threshold, 255, 0).astype(np.uint8)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (31, 9))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    h, w = small.shape
    candidates: list[tuple[float, tuple[int, int, int, int]]] = []
    for contour in contours:
        x, y, cw, ch = cv2.boundingRect(contour)
        area_ratio = (cw * ch) / float(w * h)
        aspect = cw / max(ch, 1)
        if area_ratio < 0.015 or cw < w * 0.18 or aspect < 1.8:
            continue
        center_penalty = abs((x + cw / 2) - w / 2) / w
        score = area_ratio * min(aspect, 5.0) - center_penalty * 0.12
        candidates.append((score, (x, y, cw, ch)))

    if not candidates:
        raise ValueError("필름 영역을 자동 검출하지 못했습니다. 수동 영역을 지정하세요.")

    _, (x, y, cw, ch) = max(candidates, key=lambda item: item[0])
    scale = 1.0 / small_scale
    x, y, cw, ch = (round(v * scale) for v in (x, y, cw, ch))

    # Trim small light leaks at the outer edge without resampling the image.
    pad_x = max(2, round(cw * 0.006))
    pad_y = max(2, round(ch * 0.012))
    x, y = x + pad_x, y + pad_y
    cw, ch = cw - 2 * pad_x, ch - 2 * pad_y
    return clamp_rect((x, y, cw, ch), image.shape[1], image.shape[0])


def clamp_rect(rect: tuple[int, int, int, int], width: int, height: int) -> tuple[int, int, int, int]:
    x, y, w, h = [int(v) for v in rect]
    x = max(0, min(x, width - 1))
    y = max(0, min(y, height - 1))
    w = max(1, min(w, width - x))
    h = max(1, min(h, height - y))
    return x, y, w, h


def crop_image(image: np.ndarray, rect: tuple[int, int, int, int]) -> np.ndarray:
    x, y, w, h = clamp_rect(rect, image.shape[1], image.shape[0])
    return image[y : y + h, x : x + w].copy()


def prepare_for_ocr(crop: np.ndarray) -> list[np.ndarray]:
    gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
    # Lead characters normally occupy the lower half. Keep a full-image variant too.
    lower = gray[int(gray.shape[0] * 0.38) :, :]
    variants = []
    for region in (lower, gray):
        scaled = cv2.resize(region, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(scaled)
        variants.append(clahe)
        variants.append(cv2.threshold(clahe, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1])
    return variants


def recognize_film_section(reader, crop: np.ndarray) -> str:
    """Read the two small lead digits: one above JY/JA and one above 100A."""
    height, width = crop.shape[:2]
    y1, y2 = int(height * 0.08), int(height * 0.62)
    zones = (
        crop[y1:y2, int(width * 0.14) : int(width * 0.34)],  # digit above JY / JA2026
        crop[y1:y2, int(width * 0.61) : int(width * 0.81)],  # digit above 100A
    )

    def read_single_digit(roi: np.ndarray) -> str:
        if roi.size == 0:
            return ""
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        enlarged = cv2.resize(gray, None, fx=6.0, fy=6.0, interpolation=cv2.INTER_CUBIC)
        variants = (
            enlarged,
            cv2.threshold(enlarged, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
            cv2.adaptiveThreshold(
                enlarged, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY, 51, 5,
            ),
        )
        candidates: list[tuple[float, str]] = []
        for variant in variants:
            results = reader.readtext(variant, detail=1, paragraph=False, allowlist="123")
            for box, text, confidence in results:
                cleaned = re.sub(r"[^123]", "", text)
                if not cleaned:
                    continue
                # The lead digit is above the main text line; prefer the uppermost detection.
                center_y = sum(float(point[1]) for point in box) / len(box)
                vertical_bonus = max(0.0, 1.0 - center_y / max(variant.shape[0], 1))
                candidates.append((float(confidence) + vertical_bonus * 0.35, cleaned[0]))
        return max(candidates, default=(0.0, ""), key=lambda item: item[0])[1]

    first, second = (read_single_digit(zone) for zone in zones)
    pair_to_section = {"12": "(1-2)", "23": "(2-3)", "31": "(3-1)"}
    return pair_to_section.get(first + second, "")


def normalize_ocr_text(text: str) -> str:
    text = text.upper().replace("—", "-").replace("_", " ")
    text = re.sub(r"[^A-Z0-9.\-/ ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_marking(text: str) -> dict[str, str]:
    raw = normalize_ocr_text(text)
    compact = re.sub(r"\s*[-/]\s*", "-", raw)
    dense = re.sub(r"[^A-Z0-9]", "", raw)

    joint = ""
    match = re.search(r"\b([SBR]\s*[0-9OI]{1,3}\s*R?)\b", raw)
    if match:
        joint = re.sub(r"\s+", "", match.group(1)).replace("O", "0").replace("I", "1")
    if not joint:
        # OCR often joins the lead characters to adjacent text (e.g. ...01504SO1SIT...).
        joint_candidates = []
        for candidate in re.findall(r"([SBR][0-9OI]{1,2}R?)", dense):
            normalized = candidate.replace("O", "0").replace("I", "1")
            repair_suffix = "R" if normalized.endswith("R") else ""
            digits = normalized[1:-1] if repair_suffix else normalized[1:]
            if digits.isdigit() and 0 < int(digits) <= 30:
                joint_candidates.append(normalized[0] + digits.zfill(2) + repair_suffix)
        if joint_candidates:
            joint = joint_candidates[0]
        elif re.search(r"\dSI(?:SPPS|SIT)", dense):
            joint = "S01"

    if joint:
        repair_suffix = "R" if joint.endswith("R") else ""
        joint_digits = joint[1:-1] if repair_suffix else joint[1:]
    else:
        repair_suffix = ""
        joint_digits = ""
    if joint and joint_digits.isdigit():
        number = int(joint_digits)
        if number > 30:
            joint = ""
        else:
            joint = joint[0] + str(number).zfill(2) + repair_suffix

    welder = ""
    match = re.search(r"\bW\s*-?\s*(?:20)?(\d{2})\s*-?\s*A\s*-?\s*(\d{1,3})\b", raw)
    if match:
        welder = f"W-20{match.group(1)}-A-{int(match.group(2)):02d}"
    else:
        # A may be read as 4 in compact lead markings: W23427 -> W23 A27.
        match = re.search(r"W(\d{2})4(\d{2})", dense)
        if match:
            welder = f"W-20{match.group(1)}-A-{int(match.group(2)):02d}"
    if not welder:
        match = re.search(r"\bW(?:ELDER)?\s*-?\s*([A-Z0-9-]{2,12})\b", raw)
        if match:
            welder = "W-" + match.group(1).strip("-")

    section = ""
    match = re.search(r"\b(?:SEC|SECTION)\.?\s*-?\s*(\d{1,3})\b", raw)
    if match:
        section = f"Sec.{int(match.group(1)):02d}"

    line_no = ""
    # Handles both spaced lead letters and already hyphenated line numbers.
    match = re.search(
        r"\b([A-Z]{1,3}\d{4})[- ]+(\d{2,3})[- ]+(\d{1,3})[- ]+(\d{1,2})(?:[- ]+(\d{1,2}))?[- ]+(\d{2,4}A)\b",
        compact,
    )
    if match:
        pieces = [p for p in match.groups() if p]
        line_no = "-".join(pieces)
    else:
        match = re.search(r"\b([A-Z]{1,3}\d{4}(?:-\d{1,3}){2,5}-\d{2,4}A)\b", compact)
        if match:
            line_no = match.group(1)

    if not line_no:
        # Field OCR frequently merges spaces and confuses J with 5/1.
        # Recover the stable year + numeric blocks, then locate a standard pipe size.
        match = re.search(r"20(\d{2})(\d{3})(\d{3})(\d{2})", dense)
        if match:
            year, block1, block2, block3 = match.groups()
            prefix_window = dense[max(0, match.start() - 2) : match.start()]
            prefix = "JA" if prefix_window in {"JA", "5A", "1A", "JY"} else "JA"
            sizes = ("1100", "1000", "900", "850", "800", "750", "700", "650", "600", "550", "500", "450", "400", "350", "300", "250", "200", "150", "125", "100", "80")
            search_tail = dense[match.end() :]
            size = next((s for s in sizes if re.search(rf"{s}A?", search_tail)), "")
            if size:
                line_no = f"{prefix}20{year}-{block1}-{block2}-{block3}-{size}A"

    return {"line_no": line_no, "joint_no": joint, "welder_no": welder, "section": section}


def safe_component(value: str, fallback: str) -> str:
    value = INVALID_FILENAME.sub("-", value.strip()).strip(" .-")
    return value or fallback


def render_filename(template: str, record: FilmRecord) -> str:
    values = {
        "line": safe_component(record.line_no, "LINE미확인"),
        "joint": safe_component(record.joint_no, "JOINT미확인"),
        "welder": safe_component(record.welder_no, "용접사미확인"),
        "section": safe_component(record.section, "구간미확인"),
        "seq": str(record.sequence),
        "original": safe_component(Path(record.source).stem, "IMAGE"),
    }
    name = template
    for key, value in values.items():
        name = name.replace("{" + key + "}", value)
    name = safe_component(name, values["original"])
    return name + ".JPG"


def assign_sequences(records: list[FilmRecord]) -> None:
    film_sections = ("(1-2)", "(2-3)", "(3-1)")
    joint_pattern = re.compile(r"[SBR]\d{2}R?")
    valid_anchors = [
        index for index, record in enumerate(records)
        if joint_pattern.fullmatch(record.joint_no or "")
    ]

    # Before OCR there are no JOINT anchors. Show only a provisional 3-film cycle.
    if not valid_anchors:
        for index, record in enumerate(records):
            record.sequence = index % 3 + 1
            record.section = film_sections[index % 3]
            record.section_source = "촬영순서(임시)"
        return

    # Fill an OCR-missed JOINT only from a nearby recognized image (maximum distance 2).
    # This recovers one 3-film set without carrying a wrong JOINT across many files.
    for index, record in enumerate(records):
        if joint_pattern.fullmatch(record.joint_no or ""):
            continue
        nearby = sorted(valid_anchors, key=lambda anchor: (abs(anchor - index), anchor))
        if nearby and abs(nearby[0] - index) <= 2:
            anchor_record = records[nearby[0]]
            record.joint_no = anchor_record.joint_no
            if not record.line_no:
                record.line_no = anchor_record.line_no
            if not record.welder_no:
                record.welder_no = anchor_record.welder_no

    # Reconcile common metadata inside each recognized JOINT group.
    grouped: defaultdict[str, list[FilmRecord]] = defaultdict(list)
    for record in records:
        if joint_pattern.fullmatch(record.joint_no or ""):
            grouped[record.joint_no].append(record)
    for group in grouped.values():
        for field in ("line_no", "welder_no"):
            values = [str(getattr(record, field, "")).strip() for record in group]
            values = [value for value in values if value]
            if not values:
                continue
            counts = defaultdict(int)
            for value in values:
                counts[value.upper()] += 1
            best = max(values, key=lambda value: (counts[value.upper()], len(value)))
            for record in group:
                if not getattr(record, field, ""):
                    setattr(record, field, best)

    counters: defaultdict[str, int] = defaultdict(int)
    for record in records:
        joint = record.joint_no if joint_pattern.fullmatch(record.joint_no or "") else ""
        if not joint:
            record.sequence = 1
            record.section = ""
            record.section_source = "확인필요(JOINT 미인식)"
            continue

        counters[joint] += 1
        record.sequence = counters[joint]
        ocr_section = record.section if record.section_source == "납마킹 OCR" else ""

        if joint.endswith("R"):
            # A repair JOINT can contain only the failed sections (one or two photos).
            # Its section cannot be inferred safely from position in the file list.
            if ocr_section:
                record.section = ocr_section
                record.section_source = "보수구간 OCR"
            elif record.section_source == "수동확인" and record.section in film_sections:
                pass
            else:
                record.section = ""
                record.section_source = "확인필요(보수구간 선택)"
            continue

        if record.sequence <= 3:
            expected_section = film_sections[record.sequence - 1]
            record.section = expected_section
            if ocr_section:
                record.section_source = (
                    "촬영순서/OCR일치"
                    if ocr_section == expected_section
                    else "확인필요(OCR불일치)"
                )
            else:
                record.section_source = "촬영순서"
        else:
            record.section = ocr_section
            record.section_source = "확인필요(원촬영 3장 초과)"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 10000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"중복되지 않는 파일명을 만들 수 없습니다: {path}")


def export_records(
    records: list[FilmRecord],
    output_dir: str | Path,
    rect: tuple[int, int, int, int],
    template: str,
    jpeg_quality: int = 95,
) -> list[Path]:
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for record in records:
        image = imread(record.source)
        cropped = crop_image(image, rect)
        record.output_name = render_filename(template, record)
        destination = unique_path(output / record.output_name)
        imwrite_jpg(destination, cropped, jpeg_quality)
        written.append(destination)

    log_path = output / "영상화_작업기록.csv"
    with log_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(asdict(records[0]).keys()) if records else [])
        if records:
            writer.writeheader()
            writer.writerows(asdict(record) for record in records)
    return written

