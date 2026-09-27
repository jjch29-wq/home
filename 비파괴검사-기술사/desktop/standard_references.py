"""규격 관련 학습문제에 표시할 판본 확인 메타데이터."""

REFERENCE_RULES = [
    (("ks b 0896", "stb-a2", "rb-4", "전이보상", "감도보정"), "KS B 0896:2020 · 2025-12-09 확인표준 · 확인일 2026-09-27"),
    (("stb-a1", "ks b 0831"), "KS B 0831 · 적용 판본 및 세부 수치는 최신 원문 확인 필요"),
    (("iso 9712",), "ISO 9712:2021 · 인증기관의 최신 적용규정 병행 확인"),
    (("asme",), "ASME · 계약서가 지정한 Section·Article·Edition/Addenda 확인 필요"),
    (("jis",), "JIS · 적용 규격번호와 최신 판본 확인 필요"),
    (("icrp", "선량한도", "원자력안전법"), "방사선 방호·선량 기준 · 국내 최신 법령·고시 및 적용 ICRP 권고 확인 필요"),
    (("ks b ", "ks a ", "ks d "), "KS 표준 · 적용 규격번호, 확인표준 여부와 최신 판본 확인 필요"),
    (("비파괴검사기술", "비파괴 검사업", "비파괴검사업"), "비파괴검사 관련 법령 · 현행 법률·시행령·시행규칙 확인 필요"),
]


def references_for_problem(problem: dict) -> list[str]:
    source = " ".join(
        [problem.get("title", ""), problem.get("study_group", ""), *problem.get("tags", [])]
    ).casefold()
    return [description for keywords, description in REFERENCE_RULES if any(keyword in source for keyword in keywords)]
