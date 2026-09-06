"""Geometry, persistence and vector exporters for ISO Drawer."""
from __future__ import annotations

import json
import math
from dataclasses import asdict, dataclass, field
from pathlib import Path

ISO_ANGLES = (0, 30, 90, 150, 180, 210, 270, 330)


@dataclass
class Point:
    x: float
    y: float
    actual_length: float = 0.0
    component: str = "NONE"


@dataclass
class Project:
    line_no: str = "LINE-001"
    size: str = '4"'
    spec: str = ""
    company: str = "SITCO"
    project_name: str = ""
    points: list[Point] = field(default_factory=list)

    def to_dict(self):
        return {"version": 1, **asdict(self)}

    @classmethod
    def from_dict(cls, data):
        return cls(
            line_no=data.get("line_no", "LINE-001"),
            size=data.get("size", '4"'),
            spec=data.get("spec", ""),
            company=data.get("company", "SITCO"),
            project_name=data.get("project_name", ""),
            points=[Point(**p) for p in data.get("points", [])],
        )


def snap_iso(start: tuple[float, float], cursor: tuple[float, float]):
    dx, dy = cursor[0] - start[0], cursor[1] - start[1]
    distance = math.hypot(dx, dy)
    if distance == 0:
        return start
    angle = math.degrees(math.atan2(dy, dx)) % 360
    snapped = min(ISO_ANGLES, key=lambda a: abs((angle - a + 180) % 360 - 180))
    rad = math.radians(snapped)
    return start[0] + distance * math.cos(rad), start[1] + distance * math.sin(rad)


def save_project(project: Project, path):
    Path(path).write_text(json.dumps(project.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")


def load_project(path):
    return Project.from_dict(json.loads(Path(path).read_text(encoding="utf-8")))


def _bounds(points):
    if not points:
        return 0, 0, 1, 1
    xs, ys = [p.x for p in points], [p.y for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def export_dxf(project: Project, path):
    """Write an ASCII DXF R12 that virtually every CAD package can open."""
    out = ["0", "SECTION", "2", "HEADER", "0", "ENDSEC", "0", "SECTION", "2", "ENTITIES"]

    def entity(*values):
        out.extend(str(v) for v in values)

    for a, b in zip(project.points, project.points[1:]):
        entity("0", "LINE", "8", "PIPE", "10", a.x, "20", -a.y, "30", 0, "11", b.x, "21", -b.y, "31", 0)
    for i, p in enumerate(project.points):
        entity("0", "CIRCLE", "8", "NODE", "10", p.x, "20", -p.y, "30", 0, "40", 2.5)
        if i and p.actual_length:
            prev = project.points[i - 1]
            mx, my = (prev.x + p.x) / 2, -(prev.y + p.y) / 2
            entity("0", "TEXT", "8", "DIM", "10", mx, "20", my + 6, "30", 0, "40", 4, "1", f"{p.actual_length:g} mm")
        if p.component != "NONE":
            entity("0", "TEXT", "8", "COMPONENT", "10", p.x + 4, "20", -p.y + 4, "30", 0, "40", 3.5, "1", p.component)
    entity("0", "ENDSEC", "0", "EOF")
    Path(path).write_text("\n".join(out) + "\n", encoding="ascii", errors="replace")


def export_pdf(project: Project, path):
    """Dependency-free, single-page A3 landscape vector PDF."""
    width, height, margin = 1191.0, 842.0, 55.0
    x0, y0, x1, y1 = _bounds(project.points)
    spanx, spany = max(x1 - x0, 1), max(y1 - y0, 1)
    # 타이틀 블록 영역 확보 (우측 하단 310x180pt)
    tb_w, tb_h = 310, 180
    draw_w = width - 2*margin - tb_w - 15
    draw_h = height - 2*margin - 40
    scale = min(draw_w / spanx, draw_h / spany)
    # 그림 수직 중앙 정렬
    y_used = spany * scale
    y_top_offset = margin + (draw_h - y_used) / 2 + 20  # 상단 여백 포함 중앙

    def xy(p):
        pdf_x = margin + (p.x - x0) * scale
        pdf_y = height - y_top_offset - (p.y - y0) * scale
        return pdf_x, pdf_y

    def safe(s): return str(s).encode("ascii", "replace").decode("ascii").replace("(", "[").replace(")", "]")

    cmds = ["0.05 0.05 0.05 RG 0.5 w",
            f"20 20 {width-40:g} {height-40:g} re S"]
    # 배관 선 그리기
    cmds.append("0.1 0.65 0.65 RG 2 w")
    for a, b in zip(project.points, project.points[1:]):
        ax, ay = xy(a); bx, by = xy(b)
        cmds.append(f"{ax:.2f} {ay:.2f} m {bx:.2f} {by:.2f} l S")
    cmds.append("0 0 0 RG 1 w")
    cmds.append("/F1 9 Tf")
    for i, p in enumerate(project.points):
        px, py = xy(p)
        cmds.append(f"{px-2:.2f} {py-2:.2f} 4 4 re S")
        if i and p.actual_length:
            qx, qy = xy(project.points[i-1]); mx, my = (px+qx)/2, (py+qy)/2
            ldx, ldy = px - qx, py - qy
            llen = max((ldx**2 + ldy**2)**0.5, 1)
            nx, ny = -ldy/llen*10, ldx/llen*10
            if ny < 0: nx, ny = -nx, -ny
            cmds.append(f"BT {mx+nx:.2f} {my+ny+3:.2f} Td ({safe(p.actual_length)} mm) Tj ET")
        if p.component != "NONE":
            cmds.append(f"BT {px+5:.2f} {py+5:.2f} Td ({safe(p.component)}) Tj ET")

    # 우측 하단 타이틀 블록
    tbx = width - margin - tb_w
    tby = margin - 10  # 페이지 하단 기준
    # 행 구조 (아래부터): DWG NO(25) / DATE+REV(25) / SIZE+SPEC(25) / LINE NO(30) / PROJECT(30) / 구분선 / HEADER(45)
    r = [tby, tby+25, tby+50, tby+75, tby+105, tby+135, tby+tb_h]
    #    r[0]   r[1]   r[2]   r[3]    r[4]     r[5]     r[6]=top

    cmds.append("0 0 0 RG 1.5 w")
    cmds.append(f"{tbx:.1f} {r[0]:.1f} {tb_w} {tb_h} re S")  # 외부 테두리
    # 가로선
    for ry in r[1:-1]:
        cmds.append(f"{tbx:.1f} {ry:.1f} m {tbx+tb_w:.1f} {ry:.1f} l S")
    # 세로 구분선 (SIZE/SPEC, DATE/REV 행)
    mid = tbx + tb_w / 2
    cmds.append(f"{mid:.1f} {r[0]:.1f} m {mid:.1f} {r[3]:.1f} l S")

    # 헤더 영역 (r[5] ~ r[6])
    cmds.append(f"BT /F1 13 Tf {tbx+8:.1f} {r[5]+22:.1f} Td ({safe(project.company)}) Tj ET")
    cmds.append(f"BT /F1 7.5 Tf {tbx+8:.1f} {r[5]+8:.1f} Td (PIPING ISOMETRIC DRAWING) Tj ET")

    # PROJECT (r[4]~r[5], 30pt)
    cmds.append(f"BT /F1 7 Tf {tbx+5:.1f} {r[4]+20:.1f} Td (PROJECT) Tj ET")
    cmds.append(f"BT /F1 9 Tf {tbx+65:.1f} {r[4]+20:.1f} Td ({safe(project.project_name)}) Tj ET")
    cmds.append(f"BT /F1 7 Tf {tbx+5:.1f} {r[4]+7:.1f} Td () Tj ET")

    # LINE NO (r[3]~r[4], 30pt)
    cmds.append(f"BT /F1 7 Tf {tbx+5:.1f} {r[3]+20:.1f} Td (LINE NO.) Tj ET")
    cmds.append(f"BT /F1 10 Tf {tbx+65:.1f} {r[3]+18:.1f} Td ({safe(project.line_no)}) Tj ET")

    # SIZE | SPEC (r[2]~r[3], 25pt)
    cmds.append(f"BT /F1 7 Tf {tbx+5:.1f} {r[2]+16:.1f} Td (SIZE) Tj ET")
    cmds.append(f"BT /F1 9 Tf {tbx+5:.1f} {r[2]+5:.1f} Td ({safe(project.size)}) Tj ET")
    cmds.append(f"BT /F1 7 Tf {mid+5:.1f} {r[2]+16:.1f} Td (SPEC) Tj ET")
    cmds.append(f"BT /F1 9 Tf {mid+5:.1f} {r[2]+5:.1f} Td ({safe(project.spec)}) Tj ET")

    # DATE | REV (r[1]~r[2], 25pt)
    cmds.append(f"BT /F1 7 Tf {tbx+5:.1f} {r[1]+16:.1f} Td (DATE) Tj ET")
    cmds.append(f"BT /F1 9 Tf {tbx+5:.1f} {r[1]+5:.1f} Td () Tj ET")
    cmds.append(f"BT /F1 7 Tf {mid+5:.1f} {r[1]+16:.1f} Td (REV) Tj ET")
    cmds.append(f"BT /F1 9 Tf {mid+5:.1f} {r[1]+5:.1f} Td (A) Tj ET")

    # DWG NO (r[0]~r[1], 25pt)
    cmds.append(f"BT /F1 7 Tf {tbx+5:.1f} {r[0]+16:.1f} Td (DWG NO.) Tj ET")

    stream = "\n".join(cmds).encode("ascii", "replace")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width:g} {height:g}] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>".encode(),
        f"<< /Length {len(stream)} >>\nstream\n".encode() + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for i, obj in enumerate(objects, 1):
        offsets.append(len(pdf)); pdf += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
    xref = len(pdf)
    pdf += f"xref\n0 {len(objects)+1}\n0000000000 65535 f \n".encode()
    for off in offsets[1:]: pdf += f"{off:010d} 00000 n \n".encode()
    pdf += f"trailer << /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    Path(path).write_bytes(pdf)
