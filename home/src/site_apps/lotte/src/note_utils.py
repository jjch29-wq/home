"""현장 기록 비고와 상시 메모 정리 함수."""

import re


_GENERATED_REJ_NOTE = re.compile(
    r"\s*\[당일\s*불량\(REJ\)\s+[-+]?\d+(?:\.\d+)?개\s*발생\]\s*"
)


def strip_generated_rej_note(value):
    """과거 기록에 자동 삽입된 당일 불량 안내 문구를 제거한다."""
    cleaned = _GENERATED_REJ_NOTE.sub(" ", str(value or ""))
    return " ".join(cleaned.split()).strip(" |").strip()
