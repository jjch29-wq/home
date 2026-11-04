# -*- coding: utf-8 -*-
"""UT 스캔플랜 모듈 단위 테스트."""

import os
import sys
import tempfile
import unittest
from pathlib import Path

# 소스 경로 추가
SRC_DIR = str(Path(__file__).resolve().parents[1] / "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from ut_scan_plan import (
    UTProbeConfig,
    UTScanPlanSettings,
    calculate_ut_geometry,
    create_ut_workbook,
    get_standard_probes_by_thickness,
    make_ut_beam_diagram,
)


class TestUTScanPlan(unittest.TestCase):
    def setUp(self):
        self.settings = UTScanPlanSettings(
            thickness_mm=15.87,
            cap_width_mm=20.0,
            bevel_angle_deg=35.0,
            haz_width_mm=10.0,
        )

    def test_probe_selection_by_thickness(self):
        """절차서 SIS H 241 SI Para 4.6.1 두께별 탐촉자 규격 자동 선정 검증."""
        # 1) 두께 T > 25.4mm (예: 30.97mm) -> 14×14 및 Φ20
        thick_probes = get_standard_probes_by_thickness(30.97)
        self.assertEqual(len(thick_probes), 4)
        p_map_thick = {p.angle_deg: p for p in thick_probes}

        self.assertEqual(p_map_thick[45.0].crystal_size_mm, "14×14")
        self.assertEqual(p_map_thick[45.0].index_offset_mm, 13.0)
        self.assertEqual(p_map_thick[60.0].crystal_size_mm, "14×14")
        self.assertEqual(p_map_thick[60.0].index_offset_mm, 14.0)
        self.assertEqual(p_map_thick[70.0].crystal_size_mm, "14×14")
        self.assertEqual(p_map_thick[70.0].index_offset_mm, 15.0)
        self.assertEqual(p_map_thick[0.0].crystal_size_mm, "Φ20")
        self.assertEqual(p_map_thick[0.0].index_offset_mm, 0.0)

        # 2) 두께 T <= 25.4mm (예: 15.87mm) -> 10×10 및 Φ10
        thin_probes = get_standard_probes_by_thickness(15.87)
        self.assertEqual(len(thin_probes), 4)
        p_map_thin = {p.angle_deg: p for p in thin_probes}

        self.assertEqual(p_map_thin[45.0].crystal_size_mm, "10×10")
        self.assertEqual(p_map_thin[45.0].index_offset_mm, 10.0)
        self.assertEqual(p_map_thin[60.0].crystal_size_mm, "10×10")
        self.assertEqual(p_map_thin[60.0].index_offset_mm, 11.0)
        self.assertEqual(p_map_thin[70.0].crystal_size_mm, "10×10")
        self.assertEqual(p_map_thin[70.0].index_offset_mm, 12.0)
        self.assertEqual(p_map_thin[0.0].crystal_size_mm, "Φ10")
        self.assertEqual(p_map_thin[0.0].index_offset_mm, 0.0)

        # 3) UTScanPlanSettings 기본값(T=30.97mm)의 탐촉자 검증
        default_settings = UTScanPlanSettings()
        def_map = {p.angle_deg: p for p in default_settings.probes}
        self.assertEqual(def_map[45.0].crystal_size_mm, "14×14")
        self.assertEqual(def_map[0.0].crystal_size_mm, "Φ20")

    def test_calculate_ut_geometry(self):
        results = calculate_ut_geometry(self.settings)
        # 기본 4개 탐촉자 (45°, 60°, 70°, 0°)
        self.assertEqual(len(results), 4)

        r_map = {r.angle_deg: r for r in results}

        # 45도 검증: tan(45) = 1.0, 0.5S = 15.87, 1.0S = 31.74
        res_45 = r_map[45.0]
        self.assertAlmostEqual(res_45.skip_05_mm, 15.87, places=1)
        self.assertAlmostEqual(res_45.skip_10_mm, 31.74, places=1)
        self.assertAlmostEqual(res_45.sound_path_05_mm, 15.87 * (2**0.5), places=1)

        # 60도 검증: tan(60) = 1.732, 0.5S = 15.87 * 1.732 = 27.49, 1.0S = 54.98
        res_60 = r_map[60.0]
        self.assertAlmostEqual(res_60.skip_05_mm, 27.5, delta=0.2)
        self.assertAlmostEqual(res_60.skip_10_mm, 55.0, delta=0.3)
        self.assertAlmostEqual(res_60.sound_path_05_mm, 15.87 / 0.5, places=1)  # cos(60) = 0.5

        # 70도 검증: tan(70) = 2.747, 0.5S = 15.87 * 2.747 = 43.6, 1.0S = 87.2
        res_70 = r_map[70.0]
        self.assertAlmostEqual(res_70.skip_05_mm, 43.6, delta=0.3)
        self.assertAlmostEqual(res_70.skip_10_mm, 87.2, delta=0.5)

        # 0도 수직 검증: skip은 0, SP는 두께와 2*두께
        res_0 = r_map[0.0]
        self.assertEqual(res_0.skip_05_mm, 0.0)
        self.assertEqual(res_0.sound_path_05_mm, 15.87)
        self.assertEqual(res_0.sound_path_10_mm, 31.74)

    def test_sus_direct_mode(self):
        """SUS(오스테나이트계 스테인리스강) 0.5 Skip 직사 전담 모드 검증."""
        sus_settings = UTScanPlanSettings(
            material_spec="304SS / 304LSS (P-No. 8, Gr. 1)",
            thickness_mm=22.23,
            pipe_od_mm=762.0,
            cap_width_mm=30.0,
            haz_width_mm=10.0,
        )
        results = calculate_ut_geometry(sus_settings)
        r_map = {r.angle_deg: r for r in results}

        # 45도 SUS 직사 주사 범위 검증: 1.0 Skip 배제, 0.5S 직사 전담
        res_45 = r_map[45.0]
        self.assertIn("0.5S 직사", res_45.scan_range_desc)
        self.assertIn("SUS", res_45.coverage_target)

        # SUS 다이어그램 생성 검증
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            make_ut_beam_diagram(sus_settings, tmp_path, width=1200, height=500)
            self.assertTrue(os.path.exists(tmp_path))
            self.assertGreater(os.path.getsize(tmp_path), 5000)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_make_ut_beam_diagram(self):
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            make_ut_beam_diagram(self.settings, tmp_path, width=1000, height=450)
            self.assertTrue(os.path.exists(tmp_path))
            self.assertGreater(os.path.getsize(tmp_path), 5000)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_create_ut_workbook(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            create_ut_workbook(self.settings, tmp_path)
            self.assertTrue(os.path.exists(tmp_path))
            self.assertGreater(os.path.getsize(tmp_path), 5000)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


if __name__ == "__main__":
    unittest.main()
