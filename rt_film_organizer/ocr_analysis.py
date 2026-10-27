"""Fast first pass with a detailed fallback for incomplete markings."""
import cv2

from core import parse_marking, prepare_for_ocr, recognize_film_section


def analyze_marking(reader, image):
    scale = min(1.0, 1600 / max(image.shape[:2]))
    fast = cv2.resize(image, None, fx=scale, fy=scale) if scale < 1 else image

    def read(variant):
        found = reader.readtext(variant, detail=1, paragraph=False,
                                allowlist="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-/()")
        text = " ".join(item[1] for item in found)
        confidence = sum(float(item[2]) for item in found) / max(len(found), 1)
        parsed = parse_marking(text)
        return text, confidence, parsed

    result = read(fast)
    fields = ("line_no", "joint_no", "welder_no", "section")
    if result[1] >= .65 and all(result[2][field] for field in fields):
        return (*result, result[2]["section"], False)

    candidates = [result] + [read(variant) for variant in prepare_for_ocr(image)]
    # Prefer usable fields over a long string of unrecognized characters.
    result = max(candidates, key=lambda item: (
        sum(bool(item[2][field]) for field in fields), item[1], len(item[0])))
    section = recognize_film_section(reader, image) or result[2]["section"]
    return (*result, section, True)
