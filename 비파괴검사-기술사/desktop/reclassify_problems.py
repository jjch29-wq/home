"""문제 색인을 학습 순서와 유사 원리 중심으로 재분류한다."""

from __future__ import annotations

import json
from pathlib import Path


INDEX = Path(__file__).with_name("problem_index.json")

WEEK_NAMES = {
    1: "초음파 탐상검사(UT) 기초",
    2: "RT 기초",
    3: "MT·PT 표면검사",
    4: "ET·LT 전자기·누설",
    5: "VT·AE·열·광학",
    6: "용접",
    7: "재료·부식·파괴",
    8: "고급 NDT·신호처리",
    9: "설비·구조물 적용",
    10: "건전성·수명평가",
    11: "법규·규격·안전",
    12: "종합·실전",
}

# 앞쪽 규칙일수록 더 구체적이며 주분류 우선순위가 높다.
GROUPS = [
    (11, "방사선 안전·법규", ["선량한도", "방사선관리구역", "원자력안전법", "피폭", "선량계", "dosimeter", "alara", "방사선 방호", "건강진단", "roentgen", "rem)"]),
    (11, "표준·자격·품질체계", ["iso 9712", "ks ", "ks-", "asme", "자격", "법률", "검사업", "tc 135", "기량검증"]),
    (10, "POD·검사신뢰성", ["pod", "probability of detection", "round robin", "신뢰도", "불확도", "오차(error)", "기량검"]),
    (10, "파괴역학·결함평가", ["파괴역학", "응력확대", "결함 특성", "손상허용", "임계", "노치계수", "응력집중계수"]),
    (10, "수명·위험도 평가", ["rbi", "잔여수명", "수명", "larson", "가속크리프", "가속수명", "건전성", "피로수명"]),
    (9, "원전·발전설비", ["원자로", "원전", "발전", "보일러", "복수기", "alloy 600", "가동중검사", "방진기"]),
    (9, "배관·탱크·압력설비", ["배관", "파이프", "저장", "탱크", "압력용기", "튜브", "관통관", "밸브", "header"]),
    (9, "콘크리트·토목구조", ["콘크리트", "교량", "schmidt", "슈미트"]),
    (8, "PAUT·TOFD·유도초음파", ["paut", "phased", "위상배열", "tofd", "todf", "guided wave", "lamb wave", "유도초음파"]),
    (8, "비접촉·고급 초음파", ["emat", "전자기 음향", "레이저 초음파", "비접촉", "초음파 현미경", "sam", "uafm", "초음파 홀로그래피"]),
    (8, "신호처리·디지털검사", ["wavelet", "fourier", "fft", "nyquist", "aliasing", "adc", "디지털", "화상처리", "신호대 잡음", "s/n"]),
    (8, "센서·SHM·스마트검사", ["smart structure", "스마트 구조", "shm", "phm", "fbg", "센서", "qn", "acoustoelastic"]),
    (5, "음향방출 AE", ["음향방출", "음향방 출", "음항방", "acoustic emission", "ernission", "카이저", "펠리시티", "ae(", "aet"]),
    (5, "육안검사 VT", ["육안", "visual", "조명", "시력", "cambridge gauge", "확대경"]),
    (5, "적외선·열화상", ["적외선", "열화상", "spate", "thermal emission", "온도계"]),
    (5, "광학·홀로그래피", ["홀로그래피", "holography", "speckle", "스펙클", "모아레", "moire", "광탄성", "photoelastic", "간섭"]),
    (4, "와전류 원리·코일", ["와전류", "eddy current", "시험코일", "리프트오프", "lift-off", "fill-factor", "충진", "rfec"]),
    (4, "누설검사", ["누설검사", "누설 검사", "누설 시험", "누설시험", "헬륨", "helium", "진공상자", "발포검사", "bubble test", "암모니아 누설", "leak"]),
    (3, "자분탐상 MT", ["자분", "자화", "요크", "리프팅파워", "누설자속", "자기이력", "탈자", "강자성체"]),
    (3, "침투탐상 PT", ["침투", "유화제", "현상제", "탐상제", "적심성", "wettability", "모세관", "과잉침투"]),
    (3, "표면검사 비교·선정", ["표면 결함", "표면결함", "표면 시험"]),
    (2, "방사선 물리·선원", ["방사선", "x선", "x-ray", "x-선", "감마", "ir-192", "co-60", "se-75", "중성자", "bremsstrahlung", "compton", "광전효과", "체렌코프", "cerenkov", "가속기", "ionization chamber"]),
    (2, "RT 촬영·상질", ["투과사진", "투과도계", "계조계", "필름", "촬영법", "radiography", "radiographic", "flash", "laminagraphy", "heel effect", "산란선", "관찰기", "imaging plate", "cr(", "형광 투시"]),
    (1, "초음파 탐상 - 파동·전파", ["초음파", "음향", "파동", "진동양식", "근거리음장", "감쇠", "굴절", "snell", "임피던스", "표면파", "위상속도", "군속도"]),
    (1, "초음파 탐상 - 탐촉자·교정", ["탐촉자", "stb-", "rb-", "감도보정", "전이보상", "대비시험편", "저면에코", "주파수의 선정"]),
    (1, "초음파 탐상 - 결함평가", ["6db", "20db", "결함 길이", "결함크기", "초음파탐상", "초음파 탐상"]),
    (6, "용접공정·시공", ["용접", "welding", "smaw", "gmaw", "mig", "tig", "saw", "esw", "wps", "pqr", "전자빔", "레이저용접", "입열", "아크"]),
    (6, "용접결함·균열", ["용접균열", "용접 균열", "언더비드", "underbead", "delayed cracking", "비드밑", "lamellar", "용입불량", "언더컷", "용접부의 결함", "용접 결함"]),
    (6, "용접응력·PWHT", ["pwht", "용접후열", "잔류응력", "용접 변형", "용접에 의한 변형", "응력제거"]),
    (7, "금속조직·열처리", ["열처리", "상태도", "ttt", "cct", "현미경 조직", "금속현미경", "복제법", "replica", "용체화", "경화능", "결정립"]),
    (7, "파괴·피로·취성", ["파괴", "피로", "취성", "shortness", "white spot", "dbtt", "인성", "notch", "노치", "응력-변형", "인장시험", "f-n("]),
    (7, "부식·환경손상", ["부식", "corrosion", "scc", "fac", "예민화", "입계", "weld decay", "침식"]),
    (7, "재료·합금·강종", ["합금", "강의 ", "탄소강", "탄소당량", "carbon equivalent", "스테인리스", "inconel", "monel", "tmcp", "sm ", "주강", "단조", "강괴", "알루미늄", "니켈", "재료", "냉간가공", "열간가공", "sulphur print"]),
    (7, "고온손상·크리프", ["크리프", "고온", "열화", "조사취화"]),
    (12, "계측·기타 NDT", ["strain gage", "스트레인게이지", "변형률", "레이더", "radar", "모스바우어", "박하우젠", "photoacoustic", "광음향", "싱크로트론", "테라헤르츠", "t-ray", "전위차", "xrd", "회절", "비파괴 검사"]),
]


def answer_type(title: str) -> str:
    t = title.lower()
    if any(k in t for k in ["비교", "차이점", "장단점", "장-단점"]):
        return "비교형"
    if any(k in t for k in ["대책", "방지", "원인", "영향을 미치는"]):
        return "원인·대책형"
    if any(k in t for k in ["절차", "방법", "적용", "선정", "분류"]):
        return "절차·적용형"
    if any(k in t for k in ["원리", "효과", "법칙", "메카니"]):
        return "원리형"
    return "용어·설명형"


def normalize_title(title: str, number: int) -> str:
    replacements = {
        "초음파탑상": "초음파 탐상",
        "초음파담상": "초음파 탐상",
        "초음파탕상": "초음파 탐상",
        "초음파람상": "초음파 탐상",
        "초음파팀상": "초음파 탐상",
        "초음파 탑상": "초음파 탐상",
        "초음파 담상": "초음파 탐상",
        "초음파 탕상": "초음파 탐상",
        "초음파 람상": "초음파 탐상",
        "자분탕상": "자분 탐상",
        "자분담상": "자분 탐상",
        "침투탕상": "침투 탐상",
        "침투담상": "침투 탐상",
        "와류담상": "와전류 탐상",
        "와전류담상": "와전류 탐상",
        "방사신투과": "방사선투과",
        "방사선두과": "방사선투과",
        "비파과검사": "비파괴검사",
        "비파과 검사": "비파괴검사",
        "비피과": "비파괴",
    }
    for wrong, correct in replacements.items():
        title = title.replace(wrong, correct)
    if number == 90:
        return "전자기음향탐촉자(EMAT)에서 초음파의 생성 원리와 실제 적용상의 문제점을 기술"
    if number == 145:
        title = title.replace("TODF", "TOFD")
    return " ".join(title.split())


def difficulty(title: str, matched_groups: list[tuple[int, str]]) -> str:
    t = title.lower()
    if any(k in t for k in ["비교", "설명하고", "평가", "수명", "파괴역학", "기량검증", "asme", "원자로"]):
        return "심화"
    if len(title) > 75 or len(matched_groups) >= 3:
        return "응용"
    return "기본"


def classify(problem: dict) -> dict:
    title = normalize_title(problem["title"], problem["number"])
    lowered = title.lower()
    matches = [(week, group) for week, group, words in GROUPS if any(word in lowered for word in words)]
    if matches:
        week, group = matches[0]
    else:
        week, group = 12, "종합·기타"
    tags = []
    for _, matched_group in matches:
        if matched_group not in tags:
            tags.append(matched_group)
    result = dict(problem)
    result.update({
        "title": title,
        "week": week,
        "week_name": WEEK_NAMES[week],
        "study_group": group,
        "tags": tags[:5] or ["종합·기타"],
        "answer_type": answer_type(title),
        "difficulty": difficulty(title, matches),
    })
    return result


def main():
    problems = json.loads(INDEX.read_text(encoding="utf-8"))
    classified = [classify(problem) for problem in problems]
    INDEX.write_text(json.dumps(classified, ensure_ascii=False, indent=2), encoding="utf-8")
    counts = {}
    for problem in classified:
        key = (problem["week"], problem["week_name"])
        counts[key] = counts.get(key, 0) + 1
    for (week, name), count in sorted(counts.items()):
        print(f"{week:02d} {name}: {count}")


if __name__ == "__main__":
    main()
