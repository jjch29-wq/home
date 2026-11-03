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
