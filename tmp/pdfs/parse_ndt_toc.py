import json
import re
from pathlib import Path


text = Path("tmp/pdfs/2022-1-ocr.txt").read_text(encoding="utf-8")
start = text.index("===== PDF PAGE 11 =====")
end = text.index("===== PDF PAGE 22 =====")
lines = text[start:end].splitlines()[1:]

entries = []
current = None
for raw in lines:
    line = " ".join(raw.split())
    if not line or line.startswith("===== PDF PAGE"):
        continue
    match = re.match(r"^(\d{1,3})\s*[.,]?\s+(.+)$", line)
    number = int(match.group(1)) if match else None
    if match and 1 <= number <= 380 and (current is None or number > current["number"]):
        if current:
            entries.append(current)
        current = {"number": number, "parts": [match.group(2)]}
    elif current:
        current["parts"].append(line)
if current:
    entries.append(current)

cleaned = []
for entry in entries:
    parts = entry.pop("parts")
    page = None
    if parts and re.fullmatch(r"\d{1,3}", parts[-1]):
        candidate = int(parts.pop())
        if 1 <= candidate <= 500:
            page = candidate
    title = " ".join(parts)
    if page is None:
        trailing = re.search(r"\s+(\d{1,3})$", title)
        if trailing and 1 <= int(trailing.group(1)) <= 425:
            page = int(trailing.group(1))
            title = title[:trailing.start()]
    title = re.sub(r"\s+", " ", title).strip(" .")
    cleaned.append({**entry, "title": title, "book_page": page})

manual = {
    47: ("비파괴검사 장치로서의 가속기를 열거하고 각 가속기의 원리, 에너지 Level 및 검사대상범위를 설명", 97),
    51: ("스펙클무늬(Speckle pattern)", 99),
    62: ("와전류탐상에서 침투깊이와 주파수, 전기전도도, 투자율 사이의 상관관계", 107),
    63: ("와전류탐상의 원리와 검사방법으로서의 장단점", 107),
    64: ("와전류탐상검사", 107),
    67: ("요크 리프팅파워 시험(자분탐상시험)", 110),
    110: ("콘크리트 교량에 적용할 수 있는 비파괴시험법", 154),
    236: ("Narrow Gap Welding", 301),
    368: ("QNDE, NDT와 SHM의 차이점 및 Acoustoelastic Effect", 415),
}
existing = {item["number"] for item in cleaned}
for number, (title, book_page) in manual.items():
    if number not in existing:
        cleaned.append({"number": number, "title": title, "book_page": book_page})
cleaned.sort(key=lambda item: item["number"])

def classify(title: str) -> tuple[int, str]:
    t = title.lower()
    rules = [
        (11, "법규·규격", ["원자력안전법", "iso 9712", "ks ", "asme", "선량한도", "법률", "방사선관리구역"]),
        (9, "설비 적용", ["발전", "보일러", "원자로", "원전", "저장", "복수기", "배관", "탱크"]),
        (10, "건전성 평가", ["pod", "기량검", "rbi", "수명", "파괴역학", "건전성", "위험", "damage"]),
        (8, "고급 초음파", ["paut", "phased", "tofd", "guided", "lamb", "emat", "레이저 초음파", "비접촉"]),
        (2, "방사선투과 RT", ["방사선", "x선", "x-ray", "감마", "필름", "투과사진", "radiograph", "중성자", "선원"]),
        (3, "자분·침투 MT/PT", ["자분", "침투", "유화제", "현상제", "자화", "누설자속"]),
        (4, "와전류·누설 ET/LT", ["와전류", "eddy", "누설검", "헬륨", "진공", "leak"]),
        (5, "육안·음향방출 VT/AE", ["육안", "음향방출", "acoustic emission", "카이저", "펠리시티", "조명"]),
        (6, "용접", ["용접", "weld", "비드", "pwht", "입열", "아크"]),
        (7, "금속재료", ["부식", "취성", "피로", "파괴", "열처리", "강의 ", "합금", "응력", "크리프"]),
        (1, "초음파탐상 UT", ["초음파", "탐촉자", "음향", "에코", "snell", "표면파"]),
    ]
    for week, category, keywords in rules:
        if any(keyword in t for keyword in keywords):
            return week, category
    return 12, "종합·기타"

for item in cleaned:
    item["week"], item["category"] = classify(item["title"])

page_corrections = {
    1: 1, 2: 1, 3: 2, 4: 2, 5: 2, 6: 3, 7: 4, 8: 8, 9: 10,
    16: 51, 17: 54, 19: 67, 29: 84, 30: 84, 31: 85, 37: 86,
    38: 87, 42: 91, 43: 91, 52: 99, 238: 306, 257: 322, 380: 425,
}
for item in cleaned:
    if item["book_page"] is None and item["number"] in page_corrections:
        item["book_page"] = page_corrections[item["number"]]

out = Path("tmp/pdfs/ndt-toc-index.json")
out.write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")
numbers = {item["number"] for item in cleaned}
missing = [n for n in range(1, 381) if n not in numbers]
print(f"entries={len(cleaned)} missing={len(missing)}")
print("missing:", missing)
print("without_page:", [x["number"] for x in cleaned if x["book_page"] is None])
print(out)
