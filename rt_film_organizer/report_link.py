"""Read saved RT reports without importing or changing the report application."""
import json
from dataclasses import dataclass, field
from pathlib import Path

from core import FilmRecord


def load_reports(path):
    data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(data, list):
        raise ValueError("보고서 이력 또는 RT 미리보기 JSON 목록을 선택하세요.")
    if data and all(isinstance(row, dict) and "Joint" in row for row in data):
        # Preview saves include hidden rows as well as the visible line/date.
        # Apply their saved filters only here; issued snapshots are authoritative.
        visible = [row for row in data if all(row.get(flag, True)
                   for flag in ('drawing_filtered', 'date_filtered', 'line_filtered'))]
        data = [{"mode": "RT", "report_no": "", "items": visible}]
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
    # Repair rows can store their own shot count in RT_OR with RT_RE=0.
    # Keep their repair identity, but use that saved count when RT_RE has none.
    fields = ('RT_RE', 'RT_OR') if repair else ('RT_OR',)
    zero = False
    for field_name in fields:
        value = item.get(field_name)
        if value is None or not str(value).strip():
            continue
        try:
            count = int(str(value))
            if 0 <= count <= 100:
                if count or not repair:
                    return count
                zero = True
        except ValueError:
            pass
    if zero:
        return 0
    return 0 if repair else 3


def saved_sections(item, count):
    """Comma positions in Film Location identify the three film sections.

    Preserve empty positions: ',N/1' is the second section, not the first.
    A count mismatch cannot safely identify which photos belong to which section.
    """
    location = item.get('Loc')
    if location is None or not str(location).strip():
        return None
    parts = str(location).split(',')
    sections = ('(1-2)', '(2-3)', '(3-1)')
    if len(parts) > len(sections):
        return [''] * count
    occupied = [sections[index] for index, value in enumerate(parts) if value.strip()]
    return occupied if len(occupied) == count else [''] * count


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
        saved = saved_sections(item, count)
        for seq in range(1, count + 1):
            section = (saved[seq - 1] if saved is not None else
                       ("(1-2)", "(2-3)", "(3-1)")[seq - 1] if count == 3 and not repair else "")
            slots.append(dict(report_no=str(item.get("Report No") or report.get("report_no") or "").strip(),
                              report_item=str(index), line_no=line, joint_no=joint,
                              welder_no=str(item.get("Welder") or ""), size=str(item.get("Size") or ""),
                              exam_date=str(item.get("Date") or ""), sequence=seq, section=section,
                              section_source=('보고서 저장 구간' if saved is not None else '보고서순서')
                              if section else '확인필요(구간 선택)'))
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
    return [FilmRecord(source=str(sources[i]), **slot,
                       status="보고서 배정·검토필요") for slot, i in zip(slots, assignments) if i is not None]


@dataclass
class ReportAssignment:
    report: dict
    counts: list
    slots: list = field(default_factory=list)
    assignments: list = field(default_factory=list)
    start: int = 0
    dirty: bool = True


class ReportBatch:
    """Each report owns its mapping so local corrections cannot shift later reports."""
    def __init__(self, sources):
        self.sources = list(sources)
        self.entries = []

    def add(self, report):
        identity = report.get('id') or json.dumps(report, sort_keys=True, ensure_ascii=False)
        for entry in self.entries:
            existing = entry.report.get('id') or json.dumps(entry.report, sort_keys=True, ensure_ascii=False)
            if existing == identity:
                raise ValueError('이미 작업 목록에 추가한 보고서입니다.')
        entry = ReportAssignment(report, [suggested_count(item) for item in report['items']])
        entry.start = sum(sum(e.counts) for e in self.entries)
        self.entries.append(entry)
        return entry

    def move(self, index, offset):
        target = index + offset
        if not 0 <= index < len(self.entries) or not 0 <= target < len(self.entries): return index
        self.entries[index], self.entries[target] = self.entries[target], self.entries[index]
        return target

    def generate(self, index, start=None):
        entry = self.entries[index]
        start = entry.start if start is None else start
        if type(start) is not int or start < 0: raise ValueError('시작 사진 번호를 확인하세요.')
        slots = build_slots(entry.report, entry.counts)
        entry.slots = slots
        entry.start = start
        entry.assignments = [i if i < len(self.sources) else None for i in range(start, start + len(slots))]
        entry.dirty = False

    def generate_all(self):
        # Validate everything before changing any existing mapping.
        for entry in self.entries: build_slots(entry.report, entry.counts)
        start = 0
        for index, entry in enumerate(self.entries):
            self.generate(index, start)
            start += len(entry.slots)

    def records(self):
        if not self.entries: raise ValueError('보고서를 작업 목록에 추가하세요.')
        if any(entry.dirty for entry in self.entries):
            raise ValueError('미배정 또는 촬영 매수가 변경된 보고서를 먼저 배정하세요.')
        return map_photos([s for e in self.entries for s in e.slots], self.sources,
                          [a for e in self.entries for a in e.assignments])
