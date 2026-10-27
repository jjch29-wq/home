"""Read saved RT reports without importing or changing the report application."""
import json
from pathlib import Path

from core import FilmRecord


def load_reports(path):
    data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(data, list):
        raise ValueError("보고서 이력 또는 RT 미리보기 JSON 목록을 선택하세요.")
    if data and all(isinstance(row, dict) and "Joint" in row for row in data):
        data = [{"mode": "RT", "report_no": "", "items": data}]
    reports = []
    for report in data:
        if not isinstance(report, dict) or report.get("mode") != "RT" or report.get("deleted_at"):
            continue
        items = report.get("items")
        if not isinstance(items, list):
            continue
        # Issued reports already store their actual exported order.
        rows = [row for row in items if isinstance(row, dict) and row.get("selected", True)]
        if rows:
            reports.append(dict(report, items=rows))
    if not reports:
        raise ValueError("삭제되지 않은 RT 보고서가 없습니다.")
    return reports


def suggested_count(item):
    repair = str(item.get("Joint", "")).upper().endswith("R")
    value = item.get("RT_RE" if repair else "RT_OR")
    if value is not None and str(value).strip():
        try:
            count = int(str(value))
            if 0 <= count <= 100:
                return count
        except ValueError:
            pass
    return 0 if repair else 3


def build_slots(report, counts):
    if len(counts) != len(report["items"]):
        raise ValueError("보고서 항목 수가 일치하지 않습니다.")
    slots = []
    for index, (item, count) in enumerate(zip(report["items"], counts), 1):
        if not isinstance(count, int) or not 0 <= count <= 100:
            raise ValueError("촬영 매수는 0~100 사이의 정수로 입력하세요.")
        joint = str(item.get("Joint", "")).strip()
        line = str(item.get("Dwg", "")).strip()
        if count and (not joint or not line):
            raise ValueError(f"항목 {index}: 라인번호와 JOINT를 확인하세요.")
        repair = joint.upper().endswith("R")
        for seq in range(1, count + 1):
            section = ("(1-2)", "(2-3)", "(3-1)")[seq - 1] if count == 3 and not repair else ""
            slots.append(dict(report_no=str(item.get("Report No") or report.get("report_no") or "").strip(),
                              report_item=str(index), line_no=line, joint_no=joint,
                              welder_no=str(item.get("Welder") or ""), size=str(item.get("Size") or ""),
                              exam_date=str(item.get("Date") or ""), sequence=seq, section=section))
    return slots


def map_photos(slots, sources, assignments):
    """None marks a missing shot. Unassigned source photos are excluded explicitly."""
    if len(slots) != len(assignments):
        raise ValueError("배정 목록 수가 일치하지 않습니다.")
    used = [i for i in assignments if i is not None]
    if len(set(used)) != len(used):
        raise ValueError("같은 사진을 두 항목에 중복 배정할 수 없습니다.")
    if any(not isinstance(i, int) or not 0 <= i < len(sources) for i in used):
        raise ValueError("유효하지 않은 사진 배정입니다.")
    return [FilmRecord(source=str(sources[i]), **slot, section_source="보고서순서" if slot["section"] else "확인필요(구간 선택)",
                       status="보고서 배정·검토필요") for slot, i in zip(slots, assignments) if i is not None]
