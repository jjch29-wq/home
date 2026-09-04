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
    points: list[Point] = field(default_factory=list)

    def to_dict(self):
        return {"version": 1, **asdict(self)}

    @classmethod
    def from_dict(cls, data):
        return cls(
            line_no=data.get("line_no", "LINE-001"),
            size=data.get("size", '4"'),
            spec=data.get("spec", ""),
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
    scale = min((width - 2 * margin) / spanx, (height - 2 * margin - 60) / spany)

    def xy(p):
        return margin + (p.x - x0) * scale, height - margin - 35 - (p.y - y0) * scale

    cmds = ["0.1 0.65 0.65 RG 2 w", f"20 20 {width-40:g} {height-40:g} re S", "0 0 0 RG"]
    for a, b in zip(project.points, project.points[1:]):
        ax, ay = xy(a); bx, by = xy(b)
        cmds.append(f"{ax:.2f} {ay:.2f} m {bx:.2f} {by:.2f} l S")
    cmds.append("/F1 12 Tf")
    title = f"ISOMETRIC  LINE: {project.line_no}  SIZE: {project.size}  SPEC: {project.spec}".replace("(", "[").replace(")", "]")
    cmds.append(f"BT 55 42 Td ({title}) Tj ET")
    for i, p in enumerate(project.points):
        px, py = xy(p)
        cmds.append(f"{px:.2f} {py:.2f} 3 0 360 arc S" if False else f"{px-2:.2f} {py-2:.2f} 4 4 re S")
        if i and p.actual_length:
            qx, qy = xy(project.points[i-1]); mx, my = (px+qx)/2, (py+qy)/2
            cmds.append(f"BT {mx:.2f} {my+7:.2f} Td ({p.actual_length:g} mm) Tj ET")
        if p.component != "NONE":
            cmds.append(f"BT {px+5:.2f} {py+5:.2f} Td ({p.component}) Tj ET")
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
