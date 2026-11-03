# -*- coding: utf-8 -*-
"""초음파탐상시험(UT) 검사설정표 및 스캔플랜 생성 모듈.

사각탐상(Angle Beam, 45°/60°/70°) 및 수직탐상(Straight Beam, 0°)을 모두 지원하며,
모재 두께, 베벨 각도, 비드 폭에 따른 0.5/1.0 Skip 거리, Sound Path, 주사 범위를 자동 계산하고
빔 다이어그램 시각화 및 openpyxl 기반 엑셀 검사설정표를 생성합니다.
"""

from __future__ import annotations

import json
import math
import os
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


# ==============================================================================
# 1. 데이터 모델 (Data Models)
# ==============================================================================

@dataclass
class UTProbeConfig:
    """초음파 탐촉자 설정."""
    probe_id: str = "P1"
    probe_type: str = "사각 (Angle)"           # "사각 (Angle)" 또는 "수직 (Straight)"
    angle_deg: float = 45.0                     # 0.0 (수직), 45.0, 60.0, 70.0 (사각)
    frequency_mhz: float = 4.0                  # 2.0, 4.0, 5.0 MHz 등
    crystal_size_mm: str = "10×10"              # "10×10", "Φ10", "Φ20" 등
    wave_type: str = "SW (횡파)"                # "SW (횡파)" 또는 "LW (종파)"
    velocity_ms: float = 3240.0                 # 탄소강: 횡파 3240 m/s, 종파 5920 m/s
    index_offset_mm: float = 10.0               # 탐촉자 전단부에서 입사점까지 거리 (사각)
    probe_model: str = "4Z10x10A45"             # 모델명
    wedge_delay_us: float = 0.0                 # 웻지 지연 시간
    is_active: bool = True                      # 활성화 여부


@dataclass
class UTScanPlanSettings:
    """UT 검사설정표 및 스캔플랜 종합 설정."""
    # 관리 정보
    project_name: str = "North Field South Project (QatarEnergy LNG)"
    report_title: str = "초음파탐상시험(UT) 검사설정표 및 스캔플랜"
    applicable_spec: str = "ASME Sec.V Art.4 / WPS-GTSM11-021"
    acceptance_criteria: str = "ASME Sec.VIII Div.1 / B31.3 (PWHT UT)"
    inspection_date: str = "2026-11-03"
    inspector_name: str = "비파괴검사원"
    inspector_cert: str = "UT Level II"

    # 시험체 정보
    component_name: str = '24" Restriction Orifice Plate (C.W.S-1~2)'
    material_spec: str = "A516 Gr.60 / A350LF2 CL.1"
    thickness_mm: float = 30.97
    pipe_od_mm: float = 609.6                   # 24인치 배관 외경 (0 또는 음수이면 평판)
    bevel_type: str = "Compound Bevel (복합 베벨)"
    bevel_angle_deg: float = 37.5               # 하부 홈 편각 (19T까지 37.5°)
    bevel_transition_mm: float = 19.0           # 개선각 전환 높이 (19 mm)
    bevel_angle_upper_deg: float = 10.0         # 상부 홈 편각 (19T 초과 10°)
    root_gap_mm: float = 2.5                    # 루트 간격 (도면: 2.5 mm)
    root_face_mm: float = 1.5                   # 루트 면
    cap_width_mm: float = 20.0                  # 비드 폭
    cap_height_mm: float = 2.0                  # 덧살 높이
    haz_width_mm: float = 10.0                  # 열영향부(HAZ) 검사 확장폭

    # 탐상 장비 및 교정 정보
    instrument_name: str = "EPOCH 650 (Olympus)"
    couplant: str = "초음파 탐상 전용 글리세린 / 페이스트"
    cal_block: str = "STB-A1 / IIW Type 1 & ASME Basic Block"
    cal_method: str = "DAC (거리진폭보정곡선)"
    reference_level: str = "DAC 100% (기준감도 H선)"
    scan_level: str = "기준감도(DAC) + 6 dB"
    transfer_correction_db: float = 2.0         # 전송 손실 보정

    # 탐촉자 목록 (기본: 사각 45°, 60°, 70° 및 수직 0°)
    probes: List[UTProbeConfig] = field(default_factory=lambda: [
        UTProbeConfig(
            probe_id="P1",
            probe_type="사각 (Angle)",
            angle_deg=45.0,
            frequency_mhz=4.0,
            crystal_size_mm="10×10",
            wave_type="SW (횡파)",
            velocity_ms=3240.0,
            index_offset_mm=10.0,
            probe_model="4Z10x10A45",
            is_active=True
        ),
        UTProbeConfig(
            probe_id="P2",
            probe_type="사각 (Angle)",
            angle_deg=60.0,
            frequency_mhz=4.0,
            crystal_size_mm="10×10",
            wave_type="SW (횡파)",
            velocity_ms=3240.0,
            index_offset_mm=11.0,
            probe_model="4Z10x10A60",
            is_active=True
        ),
        UTProbeConfig(
            probe_id="P3",
            probe_type="사각 (Angle)",
            angle_deg=70.0,
            frequency_mhz=4.0,
            crystal_size_mm="10×10",
            wave_type="SW (횡파)",
            velocity_ms=3240.0,
            index_offset_mm=12.0,
            probe_model="4Z10x10A70",
            is_active=True
        ),
        UTProbeConfig(
            probe_id="P4",
            probe_type="수직 (Straight)",
            angle_deg=0.0,
            frequency_mhz=4.0,
            crystal_size_mm="Φ10",
            wave_type="LW (종파)",
            velocity_ms=5920.0,
            index_offset_mm=0.0,
            probe_model="4Z10N",
            is_active=True
        ),
    ])

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "UTScanPlanSettings":
        data = dict(data)
        raw_probes = data.pop("probes", None)
        probes = []
        if raw_probes:
            for p in raw_probes:
                if isinstance(p, UTProbeConfig):
                    probes.append(p)
                else:
                    probes.append(UTProbeConfig(**p))
        else:
            probes = [
                UTProbeConfig("P1", "사각 (Angle)", 45.0, 4.0, "10×10", "SW (횡파)", 3240.0, 10.0, "4Z10x10A45"),
                UTProbeConfig("P2", "사각 (Angle)", 60.0, 4.0, "10×10", "SW (횡파)", 3240.0, 11.0, "4Z10x10A60"),
                UTProbeConfig("P3", "사각 (Angle)", 70.0, 4.0, "10×10", "SW (횡파)", 3240.0, 12.0, "4Z10x10A70"),
                UTProbeConfig("P4", "수직 (Straight)", 0.0, 4.0, "Φ10", "LW (종파)", 5920.0, 0.0, "4Z10N"),
            ]
        data["probes"] = probes
        return cls(**data)

    def to_dict(self) -> Dict[str, Any]:
        result = asdict(self)
        return result


# ==============================================================================
# 2. 기하학적 계산 엔진 (Geometric Calculation Engine)
# ==============================================================================

@dataclass
class UTCalculationResult:
    """단일 탐촉자에 대한 기하학 계산 결과."""
    probe_id: str
    probe_type: str
    angle_deg: float
    skip_05_mm: float           # 0.5 Skip 거리 (mm)
    skip_10_mm: float           # 1.0 Skip 거리 (mm)
    sound_path_05_mm: float     # 0.5 Skip Sound Path (mm)
    sound_path_10_mm: float     # 1.0 Skip Sound Path (mm)
    scan_min_mm: float          # 탐촉자 입사점 최소 거리 (mm, 용접중심선 기준)
    scan_max_mm: float          # 탐촉자 입사점 최대 거리 (mm, 용접중심선 기준)
    scan_range_desc: str        # 주사 범위 요약 설명
    divergence_deg: float       # 빔 확산각 (반각, 도)
    recommended_range_mm: float # 장비 시간축 권장 Range (mm)
    coverage_target: str        # 주 검출 대상 영역


def calculate_ut_geometry(settings: UTScanPlanSettings) -> List[UTCalculationResult]:
    """설정된 모든 활성 탐촉자에 대해 Skip, Sound Path, 주사 범위를 계산한다."""
    t = float(settings.thickness_mm)
    half_cap = float(settings.cap_width_mm) / 2.0
    haz = float(settings.haz_width_mm)

    results = []
    for probe in settings.probes:
        if not probe.is_active:
            continue

        theta = float(probe.angle_deg)
        if probe.probe_type.startswith("수직") or theta == 0.0:
            # 수직 탐상 (Straight Beam, 0°)
            skip_05 = 0.0
            skip_10 = 0.0
            sp_05 = t
            sp_10 = 2.0 * t
            # 모재 라미네이션 및 두께 검사: 용접 비드 양쪽 모재부
            scan_min = half_cap + float(probe.index_offset_mm)
            scan_max = half_cap + haz + 50.0
            scan_desc = f"{scan_min:.1f} ~ {scan_max:.1f} mm (양측 모재 검사)"
            rec_range = 3.0 * t
            target = "모재 라미네이션, 판두께 측정, 용접선 인접부 모재 결함"
            divergence = 2.5
        else:
            # 사각 탐상 (Angle Beam, 45°, 60°, 70° 등)
            rad = math.radians(theta)
            tan_theta = math.tan(rad)
            cos_theta = math.cos(rad)

            skip_05 = t * tan_theta
            skip_10 = 2.0 * t * tan_theta
            sp_05 = t / cos_theta
            sp_10 = (2.0 * t) / cos_theta

            # 빔 확산각 근사 계산
            # 파장 lambda = v / f
            v = probe.velocity_ms
            f = probe.frequency_mhz * 1e6
            wavelength_mm = (v / f) * 1000.0  # mm
            # 진동자 폭 파싱 (예: "10x10" -> 10.0)
            elem_w = 10.0
            size_clean = probe.crystal_size_mm.replace("×", "x").replace("Φ", "")
            if "x" in size_clean:
                try:
                    elem_w = float(size_clean.split("x")[0])
                except ValueError:
                    pass
            elif size_clean:
                try:
                    elem_w = float(size_clean)
                except ValueError:
                    pass
            try:
                divergence = math.degrees(math.asin(min(1.0, 0.7 * wavelength_mm / elem_w)))
            except Exception:
                divergence = 3.0

            # 주사 범위 산출:
            # 비드 토우 접근 한계 = half_cap + probe.index_offset_mm
            # 0.5 스킵으로 루트부 탐상 = skip_05
            # 1.0 스킵으로 상단 HAZ/베벨 탐상 = skip_10 + half_cap + haz
            scan_min = max(half_cap + float(probe.index_offset_mm), skip_05 - half_cap)
            scan_max = skip_10 + half_cap + haz
            scan_desc = f"{scan_min:.1f} ~ {scan_max:.1f} mm (0.5~1.0 Skip)"
            rec_range = sp_10 * 1.25

            if theta >= 68.0:
                target = "용접부 상층부 및 표면/언더컷, 표면 근처 HAZ"
            elif theta >= 55.0:
                target = "용접부 중간층, 융합불량(Lack of Fusion), 슬래그 혼입"
            else:
                target = "루트부(Lack of Penetration, Root Crack) 및 대향면 결함"

        results.append(UTCalculationResult(
            probe_id=probe.probe_id,
            probe_type=probe.probe_type,
            angle_deg=theta,
            skip_05_mm=round(skip_05, 2),
            skip_10_mm=round(skip_10, 2),
            sound_path_05_mm=round(sp_05, 2),
            sound_path_10_mm=round(sp_10, 2),
            scan_min_mm=round(scan_min, 2),
            scan_max_mm=round(scan_max, 2),
            scan_range_desc=scan_desc,
            divergence_deg=round(divergence, 1),
            recommended_range_mm=round(rec_range, 1),
            coverage_target=target
        ))

    return results


# ==============================================================================
# 3. 빔 다이어그램 시각화 (Pillow Beam Diagram Renderer)
# ==============================================================================

def make_ut_beam_diagram(
    settings: UTScanPlanSettings,
    output_path: str,
    width: int = 1200,
    height: int = 500,
) -> None:
    """모재 단면, 용접부, 사각(0.5/1.0 Skip) 및 수직 탐촉자 빔 경로를 정밀 도식화한다."""
    img = Image.new("RGBA", (width, height), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)

    def get_font(size: int = 12, bold: bool = False):
        font_names = [
            "malgunbd.ttf" if bold else "malgun.ttf",
            "gulim.ttc",
            "arialbd.ttf" if bold else "arial.ttf",
        ]
        for name in font_names:
            try:
                return ImageFont.truetype(name, size)
            except OSError:
                continue
        return ImageFont.load_default()

    font_title = get_font(15, bold=True)
    font_bold = get_font(11, bold=True)
    font_normal = get_font(10, bold=False)
    font_small = get_font(9, bold=False)

    t = float(settings.thickness_mm)
    bevel_deg = float(settings.bevel_angle_deg)
    root_gap = float(settings.root_gap_mm)
    root_face = float(settings.root_face_mm)
    cap_w = float(settings.cap_width_mm)
    cap_h = float(settings.cap_height_mm)
    haz_w = float(settings.haz_width_mm)

    # 1. 상단 타이틀 바
    draw.rectangle([10, 8, width - 10, 42], fill=(240, 245, 252), outline=(180, 205, 235), width=1)
    pipe_str = f"배관 OD {settings.pipe_od_mm:g}mm" if settings.pipe_od_mm > 0 else "평판"
    title_text = f"초음파탐상(UT) 스캔플랜 다이어그램  |  {settings.component_name} (두께 T={t:g}mm, {pipe_str}, 홈각도 {bevel_deg:g}°)"
    draw.text((22, 14), title_text, fill=(20, 50, 110), font=font_title)

    # 2. 모재 단면 배치 (중앙 기준)
    center_x = width // 2
    plate_top = 135
    plate_height_px = 75
    scale_y = plate_height_px / t  # px per mm
    scale_x = scale_y  # 1:1 종횡비

    plate_bottom = plate_top + plate_height_px

    plate_half_len_mm = max(110.0, t * math.tan(math.radians(70.0)) * 2.0 + haz_w + 35.0)
    plate_half_len_px = int(plate_half_len_mm * scale_x)

    left_x = max(25, center_x - plate_half_len_px)
    right_x = min(width - 25, center_x + plate_half_len_px)

    draw.rectangle([left_x, plate_top, right_x, plate_bottom], fill=(248, 250, 252), outline=(130, 140, 150), width=2)

    rg_half_px = (root_gap / 2.0) * scale_x
    rf_px = root_face * scale_y
    cap_half_px = (cap_w / 2.0) * scale_x
    cap_h_px = cap_h * scale_y
    haz_px = haz_w * scale_x

    # 복합 베벨 판정: 전환 깊이가 0보다 크고 두께보다 작으며, 상부 각도가 하부 각도와 다를 때만 적용
    trans_val = getattr(settings, "bevel_transition_mm", 0.0)
    upper_val = getattr(settings, "bevel_angle_upper_deg", bevel_deg)
    is_compound = (trans_val > 0.0) and (trans_val < t) and (abs(upper_val - bevel_deg) > 0.1)

    h_trans = trans_val if is_compound else t
    upper_angle = upper_val if is_compound else bevel_deg

    h_trans = min(t, max(root_face, h_trans))
    h1 = h_trans - root_face
    w1_px = h1 * math.tan(math.radians(bevel_deg)) * scale_x
    h1_px = h1 * scale_y
    trans_y = plate_bottom - rf_px - h1_px

    h2 = max(0.0, t - h_trans)
    w2_px = h2 * math.tan(math.radians(upper_angle)) * scale_x
    groove_top_half_px = rg_half_px + w1_px + w2_px
    trans_half_px = rg_half_px + w1_px

    # HAZ(열영향부) 채색 (복합 베벨 꺾임 반영)
    haz_poly_left = [
        (center_x - groove_top_half_px - haz_px, plate_top),
        (center_x - groove_top_half_px, plate_top),
        (center_x - trans_half_px, trans_y),
        (center_x - rg_half_px, plate_bottom - rf_px),
        (center_x - rg_half_px, plate_bottom),
        (center_x - rg_half_px - haz_px, plate_bottom),
    ]
    haz_poly_right = [
        (center_x + groove_top_half_px, plate_top),
        (center_x + groove_top_half_px + haz_px, plate_top),
        (center_x + rg_half_px + haz_px, plate_bottom),
        (center_x + rg_half_px, plate_bottom),
        (center_x + rg_half_px, plate_bottom - rf_px),
        (center_x + trans_half_px, trans_y),
    ]
    draw.polygon(haz_poly_left, fill=(255, 245, 230), outline=None)
    draw.polygon(haz_poly_right, fill=(255, 245, 230), outline=None)

    # 용접 비드 및 복합 개선부 채색
    weld_poly = [
        (center_x - cap_half_px, plate_top),
        (center_x, plate_top - cap_h_px),
        (center_x + cap_half_px, plate_top),
        (center_x + groove_top_half_px, plate_top),
        (center_x + trans_half_px, trans_y),
        (center_x + rg_half_px, plate_bottom - rf_px),
        (center_x + rg_half_px, plate_bottom),
        (center_x, plate_bottom + (cap_h_px * 0.6)),
        (center_x - rg_half_px, plate_bottom),
        (center_x - rg_half_px, plate_bottom - rf_px),
        (center_x - trans_half_px, trans_y),
        (center_x - groove_top_half_px, plate_top),
    ]
    draw.polygon(weld_poly, fill=(230, 238, 248), outline=(80, 105, 140))

    # 복합 개선면 외곽선 (꺾임 2단)
    draw.line([(center_x - groove_top_half_px, plate_top), (center_x - trans_half_px, trans_y),
               (center_x - rg_half_px, plate_bottom - rf_px), (center_x - rg_half_px, plate_bottom)],
              fill=(70, 95, 130), width=2)
    draw.line([(center_x + groove_top_half_px, plate_top), (center_x + trans_half_px, trans_y),
               (center_x + rg_half_px, plate_bottom - rf_px), (center_x + rg_half_px, plate_bottom)],
              fill=(70, 95, 130), width=2)

    # 19T 전이선 점선 및 각도 라벨
    if is_compound:
        draw.line([(center_x - trans_half_px, trans_y), (center_x + trans_half_px, trans_y)], fill=(150, 170, 190), width=1)
        draw.text((center_x + trans_half_px + 8, trans_y - 8), f"19T 전이점 (하부 37.5° / 상부 10°)", fill=(70, 90, 130), font=font_small)

    draw.line([(center_x - cap_half_px, plate_top), (center_x, plate_top - cap_h_px), (center_x + cap_half_px, plate_top)],
              fill=(50, 80, 120), width=2)

    # 용접 중심선
    for y in range(plate_top - 30, plate_bottom + 35, 7):
        draw.line([(center_x, y), (center_x, y + 4)], fill=(220, 50, 50), width=1)
    draw.text((center_x - 30, plate_bottom + 38), "용접 중심선", fill=(200, 30, 30), font=font_small)

    # 두께 치수선
    dim_x = left_x + 16
    draw.line([(dim_x, plate_top), (dim_x, plate_bottom)], fill=(80, 80, 80), width=1)
    draw.line([(dim_x - 3, plate_top), (dim_x + 3, plate_top)], fill=(80, 80, 80), width=1)
    draw.line([(dim_x - 3, plate_bottom), (dim_x + 3, plate_bottom)], fill=(80, 80, 80), width=1)
    draw.text((dim_x - 55, (plate_top + plate_bottom) // 2 - 7), f"T={t:g}mm", fill=(50, 50, 50), font=font_bold)

    # 3. 초음파 빔 경로 및 탐촉자
    calculations = calculate_ut_geometry(settings)
    calc_map = {res.angle_deg: res for res in calculations}

    color_45 = (25, 110, 220)
    color_60 = (20, 160, 65)
    color_70 = (230, 110, 15)
    color_0  = (140, 50, 180)

    # A. 0.5 Skip 빔 (우측 45도 탐촉자 -> 루트 직접 타격)
    if 45.0 in calc_map:
        c45 = calc_map[45.0]
        s05_px = int(c45.skip_05_mm * scale_x)
        idx_x = center_x + s05_px

        overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        ov_draw = ImageDraw.Draw(overlay)
        beam_poly_05 = [
            (idx_x - 6, plate_top),
            (idx_x + 6, plate_top),
            (center_x + rg_half_px + 10, plate_bottom),
            (center_x - rg_half_px - 8, plate_bottom),
        ]
        ov_draw.polygon(beam_poly_05, fill=(*color_45, 45))
        img = Image.alpha_composite(img, overlay)
        draw = ImageDraw.Draw(img)

        draw.line([(idx_x, plate_top), (center_x, plate_bottom)], fill=color_45, width=3)
        w_poly = [(idx_x - 4, plate_top), (idx_x + 22, plate_top), (idx_x + 22, plate_top - 18), (idx_x + 2, plate_top - 18)]
        draw.polygon(w_poly, fill=(240, 245, 255), outline=color_45, width=2)
        draw.text((idx_x + 4, plate_top - 32), "45° (0.5S)", fill=color_45, font=font_bold)

        y_d = plate_top - 36
        draw.line([(center_x, y_d), (idx_x, y_d)], fill=color_45, width=1)
        draw.line([(center_x, y_d - 2), (center_x, y_d + 2)], fill=color_45, width=1)
        draw.line([(idx_x, y_d - 2), (idx_x, y_d + 2)], fill=color_45, width=1)
        draw.text(((center_x + idx_x) // 2 - 25, y_d - 13), f"0.5S={c45.skip_05_mm:.1f}mm", fill=color_45, font=font_small)

    # B. 1.0 Skip 빔 (우측 60도 탐촉자 -> 바닥 반사 후 상부 개선면/비드 타격)
    if 60.0 in calc_map:
        c60 = calc_map[60.0]
        s10_px = int(c60.skip_10_mm * scale_x)
        s05_px = int(c60.skip_05_mm * scale_x)
        idx_x = center_x + s10_px
        refl_x = center_x + s05_px

        overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        ov_draw = ImageDraw.Draw(overlay)
        ov_draw.polygon([(idx_x - 7, plate_top), (idx_x + 7, plate_top),
                         (refl_x + 12, plate_bottom), (refl_x - 12, plate_bottom)], fill=(*color_60, 40))
        ov_draw.polygon([(refl_x + 12, plate_bottom), (refl_x - 12, plate_bottom),
                         (center_x - groove_top_half_px - haz_px, plate_top),
                         (center_x + groove_top_half_px, plate_top)], fill=(*color_60, 40))
        img = Image.alpha_composite(img, overlay)
        draw = ImageDraw.Draw(img)

        draw.line([(idx_x, plate_top), (refl_x, plate_bottom)], fill=color_60, width=3)
        draw.line([(refl_x, plate_bottom), (center_x, plate_top)], fill=color_60, width=3)

        w_poly = [(idx_x - 4, plate_top), (idx_x + 24, plate_top), (idx_x + 24, plate_top - 18), (idx_x + 3, plate_top - 18)]
        draw.polygon(w_poly, fill=(240, 255, 245), outline=color_60, width=2)
        draw.text((idx_x + 5, plate_top - 32), "60° (1.0S)", fill=color_60, font=font_bold)

        y_d = plate_top - 54
        draw.line([(center_x, y_d), (idx_x, y_d)], fill=color_60, width=1)
        draw.line([(center_x, y_d - 2), (center_x, y_d + 2)], fill=color_60, width=1)
        draw.line([(refl_x, y_d - 2), (refl_x, y_d + 2)], fill=color_60, width=1)
        draw.line([(idx_x, y_d - 2), (idx_x, y_d + 2)], fill=color_60, width=1)
        draw.text(((center_x + idx_x) // 2 - 25, y_d - 13), f"1.0S={c60.skip_10_mm:.1f}mm", fill=color_60, font=font_small)

    # C. 좌측 주사 (70도 탐촉자)
    if 70.0 in calc_map:
        c70 = calc_map[70.0]
        s10_px = int(c70.skip_10_mm * scale_x)
        s05_px = int(c70.skip_05_mm * scale_x)
        idx_x = center_x - s10_px
        refl_x = center_x - s05_px

        overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        ov_draw = ImageDraw.Draw(overlay)
        ov_draw.polygon([(idx_x - 8, plate_top), (idx_x + 8, plate_top),
                         (refl_x + 14, plate_bottom), (refl_x - 14, plate_bottom)], fill=(*color_70, 35))
        ov_draw.polygon([(refl_x + 14, plate_bottom), (refl_x - 14, plate_bottom),
                         (center_x - groove_top_half_px, plate_top),
                         (center_x + groove_top_half_px + haz_px, plate_top)], fill=(*color_70, 35))
        img = Image.alpha_composite(img, overlay)
        draw = ImageDraw.Draw(img)

        draw.line([(idx_x, plate_top), (refl_x, plate_bottom)], fill=color_70, width=3)
        draw.line([(refl_x, plate_bottom), (center_x, plate_top)], fill=color_70, width=3)

        w_poly = [(idx_x + 4, plate_top), (idx_x - 24, plate_top), (idx_x - 24, plate_top - 18), (idx_x - 3, plate_top - 18)]
        draw.polygon(w_poly, fill=(255, 248, 240), outline=color_70, width=2)
        draw.text((idx_x - 48, plate_top - 32), "70° (좌측)", fill=color_70, font=font_bold)

        y_d = plate_top - 46
        draw.line([(idx_x, y_d), (center_x, y_d)], fill=color_70, width=1)
        draw.line([(idx_x, y_d - 2), (idx_x, y_d + 2)], fill=color_70, width=1)
        draw.line([(center_x, y_d - 2), (center_x, y_d + 2)], fill=color_70, width=1)
        draw.text(((idx_x + center_x) // 2 - 25, y_d - 13), f"1.0S={c70.skip_10_mm:.1f}mm", fill=color_70, font=font_small)

    # D. 수직 탐촉자 (0도)
    if 0.0 in calc_map:
        c0 = calc_map[0.0]
        str_x = center_x - int((haz_w + cap_w / 2.0 + 35.0) * scale_x)
        probe_w_px = 22
        probe_h_px = 18

        draw.rectangle([str_x - probe_w_px // 2, plate_top - probe_h_px, str_x + probe_w_px // 2, plate_top],
                       fill=(245, 235, 255), outline=color_0, width=2)
        draw.text((str_x - 24, plate_top - probe_h_px - 14), "0° 수직탐상", fill=color_0, font=font_bold)

        draw.line([(str_x, plate_top), (str_x, plate_bottom)], fill=color_0, width=3)
        draw.line([(str_x - 4, plate_bottom - 7), (str_x, plate_bottom), (str_x + 4, plate_bottom - 7)], fill=color_0, width=2)
        draw.text((str_x - 22, plate_bottom + 10), "모재 라미네이션", fill=color_0, font=font_small)

    # 4. 주사 범위 화살표
    y_scan = plate_bottom + 20
    if 60.0 in calc_map:
        c60 = calc_map[60.0]
        scan_min_px = center_x + int(c60.scan_min_mm * scale_x)
        scan_max_px = center_x + int(c60.scan_max_mm * scale_x)

        draw.line([(scan_min_px, y_scan), (scan_max_px, y_scan)], fill=(40, 80, 160), width=2)
        draw.line([(scan_min_px + 5, y_scan - 3), (scan_min_px, y_scan), (scan_min_px + 5, y_scan + 3)], fill=(40, 80, 160), width=2)
        draw.line([(scan_max_px - 5, y_scan - 3), (scan_max_px, y_scan), (scan_max_px - 5, y_scan + 3)], fill=(40, 80, 160), width=2)
        draw.text(((scan_min_px + scan_max_px) // 2 - 50, y_scan + 4), f"주사범위: {c60.scan_min_mm:.1f} ~ {c60.scan_max_mm:.1f} mm",
                  fill=(20, 60, 140), font=font_bold)

    # 5. 하단 범례 테이블
    legend_top = height - 122
    draw.rectangle([10, legend_top, width - 10, height - 8], fill=(252, 253, 255), outline=(205, 215, 230), width=1)
    draw.rectangle([10, legend_top, width - 10, legend_top + 22], fill=(235, 242, 252), outline=(205, 215, 230), width=1)
    draw.text((20, legend_top + 4), "[스캔플랜 기하학 계산 및 검출 대상 영역 요약표]", fill=(20, 45, 95), font=font_bold)

    headers = [("탐촉자 구분", 110), ("굴절각", 65), ("0.5 Skip", 85), ("1.0 Skip", 85),
               ("Sound Path", 95), ("주사 범위(Xmin~Xmax)", 170), ("주요 검출 대상 영역", 380)]
    cur_x = 20
    row_y = legend_top + 28
    for h_name, w_px in headers:
        draw.text((cur_x, row_y), h_name, fill=(60, 70, 85), font=font_bold)
        cur_x += w_px

    row_y += 18
    draw.line([(15, row_y), (width - 15, row_y)], fill=(220, 225, 235), width=1)
    row_y += 4

    p_colors = {45.0: color_45, 60.0: color_60, 70.0: color_70, 0.0: color_0}
    for angle in [45.0, 60.0, 70.0, 0.0]:
        if angle not in calc_map:
            continue
        res = calc_map[angle]
        clr = p_colors.get(angle, (50, 50, 50))

        cur_x = 20
        draw.rectangle([cur_x, row_y + 2, cur_x + 10, row_y + 11], fill=clr)
        draw.text((cur_x + 16, row_y), res.probe_type, fill=(30, 30, 30), font=font_normal)
        cur_x += 110

        draw.text((cur_x, row_y), f"{res.angle_deg:g}°", fill=clr, font=font_bold)
        cur_x += 65

        s05_str = f"{res.skip_05_mm:.1f} mm" if res.angle_deg > 0 else "-"
        draw.text((cur_x, row_y), s05_str, fill=(40, 40, 40), font=font_normal)
        cur_x += 85

        s10_str = f"{res.skip_10_mm:.1f} mm" if res.angle_deg > 0 else "-"
        draw.text((cur_x, row_y), s10_str, fill=(40, 40, 40), font=font_normal)
        cur_x += 85

        sp_str = f"{res.sound_path_10_mm:.1f} mm"
        draw.text((cur_x, row_y), sp_str, fill=(40, 40, 40), font=font_normal)
        cur_x += 95

        scan_str = f"{res.scan_min_mm:.1f} ~ {res.scan_max_mm:.1f} mm"
        draw.text((cur_x, row_y), scan_str, fill=(40, 40, 40), font=font_bold)
        cur_x += 170

        draw.text((cur_x, row_y), res.coverage_target, fill=(70, 75, 85), font=font_small)
        row_y += 18

    # RGBA -> RGB 변환 저장
    img.convert("RGB").save(output_path, "PNG")


def create_ut_workbook(settings: UTScanPlanSettings, output_path: str) -> str:
    """UT 검사설정표 및 스캔플랜을 A4 1페이지에 완벽히 맞춘 엑셀 문서로 생성한다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UT 검사설정표"
    ws.views.sheetView[0].showGridLines = True

    # 스타일 정의
    font_main_title = Font(name="맑은 고딕", size=14, bold=True, color="1F4E78")
    font_sec_title = Font(name="맑은 고딕", size=10, bold=True, color="203764")
    font_header = Font(name="맑은 고딕", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="맑은 고딕", size=9, bold=True)
    font_normal = Font(name="맑은 고딕", size=8.5)
    font_small = Font(name="맑은 고딕", size=7.5)

    fill_sec_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_tbl_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_sub_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="BFBFBF")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom_side = Side(border_style="medium", color="2F5597")
    border_header = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True)

    # 열 너비 설정 (B~L 총 11열 구성)
    col_widths = {
        "A": 3, "B": 17, "C": 16, "D": 13, "E": 13,
        "F": 13, "G": 17, "H": 14, "I": 14, "J": 14,
        "K": 17, "L": 18, "M": 4
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # 1. 메인 타이틀 (Row 2)
    ws.row_dimensions[1].height = 6
    ws.merge_cells("B2:L2")
    title_cell = ws["B2"]
    title_cell.value = "초음파탐상시험(UT) 검사설정표 및 스캔플랜"
    title_cell.font = font_main_title
    title_cell.alignment = align_center
    ws.row_dimensions[2].height = 22

    current_row = 4

    def add_section_header(title: str):
        nonlocal current_row
        ws.merge_cells(f"B{current_row}:L{current_row}")
        cell = ws[f"B{current_row}"]
        cell.value = title
        cell.font = font_sec_title
        cell.fill = fill_sec_header
        cell.alignment = align_left
        ws.row_dimensions[current_row].height = 17
        current_row += 1

    def style_range(start_col: str, start_r: int, end_col: str, end_r: int, border=border_cell):
        for r in range(start_r, end_r + 1):
            for c in range(openpyxl.utils.column_index_from_string(start_col),
                           openpyxl.utils.column_index_from_string(end_col) + 1):
                ws.cell(row=r, column=c).border = border

    def add_spacer():
        nonlocal current_row
        ws.row_dimensions[current_row].height = 5
        current_row += 1

    # --------------------------------------------------------------------------
    # 2. 섹션 1: 일반 및 관리 정보
    # --------------------------------------------------------------------------
    add_section_header("1. 일반 및 관리 정보 (General Information)")

    general_info = [
        ("공 사 명", settings.project_name, "검 사 일 자", settings.inspection_date),
        ("적 용 규 격", settings.applicable_spec, "합 격 기 준", settings.acceptance_criteria),
        ("검 사 원", f"{settings.inspector_name} ({settings.inspector_cert})", "보고서 번호", "UT-SP-2026-001"),
    ]

    for label1, val1, label2, val2 in general_info:
        ws.row_dimensions[current_row].height = 16
        ws[f"B{current_row}"] = label1
        ws[f"B{current_row}"].font = font_bold
        ws[f"B{current_row}"].fill = fill_sub_header
        ws[f"B{current_row}"].alignment = align_center

        ws.merge_cells(f"C{current_row}:F{current_row}")
        ws[f"C{current_row}"] = val1
        ws[f"C{current_row}"].font = font_normal
        ws[f"C{current_row}"].alignment = align_left

        ws[f"G{current_row}"] = label2
        ws[f"G{current_row}"].font = font_bold
        ws[f"G{current_row}"].fill = fill_sub_header
        ws[f"G{current_row}"].alignment = align_center

        ws.merge_cells(f"H{current_row}:L{current_row}")
        ws[f"H{current_row}"] = val2
        ws[f"H{current_row}"].font = font_normal
        ws[f"H{current_row}"].alignment = align_left

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 3. 섹션 2: 시험체 및 용접부 제원
    # --------------------------------------------------------------------------
    add_section_header("2. 시험체 및 용접부 제원 (Component & Weld Details)")

    pipe_desc = f"OD {settings.pipe_od_mm:g} mm" if settings.pipe_od_mm > 0 else "평판 (Plate)"
    joint_info = [
        ("부품 / 구간", settings.component_name, "재 질 (Material)", settings.material_spec),
        ("모재 두께 (T)", f"{settings.thickness_mm:g} mm", "배관 외경 / 형태", pipe_desc),
        ("홈 형상 (Bevel)", settings.bevel_type, "홈 각도 (Bevel Angle)",
          f"하부 19T: {settings.bevel_angle_deg:g}° (포함 75°) / 상부: {getattr(settings, 'bevel_angle_upper_deg', 10.0):g}° (포함 20°)"
          if getattr(settings, 'bevel_transition_mm', 0) > 0 else f"{settings.bevel_angle_deg:g}° (포함각 {settings.bevel_angle_deg*2:g}°)"),
        ("루트 간격 (Gap)", f"{settings.root_gap_mm:g} mm", "루트 면 (Root Face)", f"{settings.root_face_mm:g} mm"),
        ("비드 폭 (Width)", f"{settings.cap_width_mm:g} mm", "덧살 높이 (Cap)", f"{settings.cap_height_mm:g} mm"),
        ("열영향부 (HAZ)", f"비드 양단 + {settings.haz_width_mm:g} mm", "표면 가공 상태", "쇼트 블라스트 / 그라인딩"),
    ]

    for label1, val1, label2, val2 in joint_info:
        ws.row_dimensions[current_row].height = 16
        ws[f"B{current_row}"] = label1
        ws[f"B{current_row}"].font = font_bold
        ws[f"B{current_row}"].fill = fill_sub_header
        ws[f"B{current_row}"].alignment = align_center

        ws.merge_cells(f"C{current_row}:F{current_row}")
        ws[f"C{current_row}"] = val1
        ws[f"C{current_row}"].font = font_normal
        ws[f"C{current_row}"].alignment = align_left

        ws[f"G{current_row}"] = label2
        ws[f"G{current_row}"].font = font_bold
        ws[f"G{current_row}"].fill = fill_sub_header
        ws[f"G{current_row}"].alignment = align_center

        ws.merge_cells(f"H{current_row}:L{current_row}")
        ws[f"H{current_row}"] = val2
        ws[f"H{current_row}"].font = font_normal
        ws[f"H{current_row}"].alignment = align_left

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 4. 섹션 3: 탐상 장비 및 교정 조건
    # --------------------------------------------------------------------------
    add_section_header("3. 탐상 장비 및 교정 조건 (Equipment & Calibration)")

    equip_info = [
        ("초음파 탐상기", settings.instrument_name, "접 촉 매 질", settings.couplant),
        ("대비 시험편", settings.cal_block, "교정 기법", settings.cal_method),
        ("기준 감도 (Ref.)", settings.reference_level, "검사 감도 (Scan)", settings.scan_level),
        ("전송 손실 보정", f"+ {settings.transfer_correction_db:g} dB", "주사 속도 (Max)", "150 mm/s 이하"),
    ]

    for label1, val1, label2, val2 in equip_info:
        ws.row_dimensions[current_row].height = 16
        ws[f"B{current_row}"] = label1
        ws[f"B{current_row}"].font = font_bold
        ws[f"B{current_row}"].fill = fill_sub_header
        ws[f"B{current_row}"].alignment = align_center

        ws.merge_cells(f"C{current_row}:F{current_row}")
        ws[f"C{current_row}"] = val1
        ws[f"C{current_row}"].font = font_normal
        ws[f"C{current_row}"].alignment = align_left

        ws[f"G{current_row}"] = label2
        ws[f"G{current_row}"].font = font_bold
        ws[f"G{current_row}"].fill = fill_sub_header
        ws[f"G{current_row}"].alignment = align_center

        ws.merge_cells(f"H{current_row}:L{current_row}")
        ws[f"H{current_row}"] = val2
        ws[f"H{current_row}"].font = font_normal
        ws[f"H{current_row}"].alignment = align_left

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 5. 섹션 4: 스캔플랜 및 탐촉자별 기하학 계산표
    # --------------------------------------------------------------------------
    add_section_header("4. 스캔플랜 및 기하학 계산표 (Scan Plan Calculations)")

    tbl_headers = [
        ("B", "탐촉자 구분"),
        ("C", "굴절각(°)"),
        ("D", "주파수(MHz)"),
        ("E", "진동자 크기"),
        ("F", "0.5 Skip(mm)"),
        ("G", "1.0 Skip(mm)"),
        ("H", "Sound Path(mm)"),
        ("I", "입사점(Xmin)"),
        ("J", "입사점(Xmax)"),
        ("K", "주사 범위 요약"),
        ("L", "주요 검출 대상"),
    ]

    ws.row_dimensions[current_row].height = 18
    for col, h_name in tbl_headers:
        c = ws[f"{col}{current_row}"]
        c.value = h_name
        c.font = font_header
        c.fill = fill_tbl_header
        c.alignment = align_center
        c.border = border_header
    current_row += 1

    calc_results = calculate_ut_geometry(settings)
    probe_map = {p.probe_id: p for p in settings.probes}

    for res in calc_results:
        p_cfg = probe_map.get(res.probe_id)
        ws.row_dimensions[current_row].height = 16

        ws[f"B{current_row}"] = f"{res.probe_type} ({res.probe_id})"
        ws[f"C{current_row}"] = f"{res.angle_deg:g}°"
        ws[f"D{current_row}"] = f"{p_cfg.frequency_mhz:g}" if p_cfg else "-"
        ws[f"E{current_row}"] = p_cfg.crystal_size_mm if p_cfg else "-"
        ws[f"F{current_row}"] = f"{res.skip_05_mm:.1f}" if res.angle_deg > 0 else "-"
        ws[f"G{current_row}"] = f"{res.skip_10_mm:.1f}" if res.angle_deg > 0 else "-"
        ws[f"H{current_row}"] = f"{res.sound_path_10_mm:.1f}"
        ws[f"I{current_row}"] = f"{res.scan_min_mm:.1f}"
        ws[f"J{current_row}"] = f"{res.scan_max_mm:.1f}"
        ws[f"K{current_row}"] = res.scan_range_desc
        ws[f"L{current_row}"] = res.coverage_target

        for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]:
            ws[f"{col}{current_row}"].alignment = align_center
            ws[f"{col}{current_row}"].font = font_bold if col in ["C", "F", "G", "H"] else font_normal
        ws[f"K{current_row}"].alignment = align_center
        ws[f"K{current_row}"].font = font_normal
        ws[f"L{current_row}"].alignment = align_left
        ws[f"L{current_row}"].font = font_small

        if res.angle_deg == 60.0:
            for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]:
                ws[f"{col}{current_row}"].fill = fill_highlight

        style_range("B", current_row, "L", current_row)
        current_row += 1

    add_spacer()

    # --------------------------------------------------------------------------
    # 6. 섹션 5: 스캔플랜 다이어그램 임베드
    # --------------------------------------------------------------------------
    add_section_header("5. 스캔플랜 빔 다이어그램 (Scan Plan Beam Path Diagram)")

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_img_path = tmp.name

    try:
        make_ut_beam_diagram(settings, tmp_img_path, width=1200, height=500)
        xl_img = XLImage(tmp_img_path)
        xl_img.width = 750
        xl_img.height = 312
        ws.add_image(xl_img, f"B{current_row}")
    finally:
        pass

    # 다이어그램 아래 빈 행들 높이 지정
    for offset in range(17):
        ws.row_dimensions[current_row + offset].height = 18

    # 인쇄 설정 (A4 가로 1페이지 강제 맞춤)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 여백 최소화 (인치 단위)
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1

    # 인쇄 정중앙 맞춤
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True

    # 인쇄 영역 확정
    ws.print_area = f"$B$2:$L${current_row + 16}"

    wb.save(output_path)
    return output_path
