"""Unit tests for Central branch company -> section -> line cascading relationships."""
from pathlib import Path
import sys
import unittest

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))

from site_apps.central.src.services.section_line_helper import (
    extract_company_section_line_mapping,
    get_sections_for_company,
    get_lines_for_company_and_section,
)


class CentralSectionLineCascadingTests(unittest.TestCase):
    def setUp(self):
        self.sample_history = {
            '2026-03-01': {
                'ndt_results': [
                    {'업체': 'GS네오텍', '구간': '1공구', '라인번호': '1-1L', '검사방법': 'PAUT'},
                    {'업체': 'GS네오텍', '구간': '1공구', '라인번호': '1-2L', '검사방법': 'RT'},
                    {'업체': 'GS네오텍', '구간': '2공구', '라인번호': '2-1L', '검사방법': 'PAUT'},
                    {'업체': '다림건설(주)', '구간': '3공구', '라인번호': '3-1L', '검사방법': 'MT'},
                ]
            },
            '2026-03-02': {
                'ndt_results': [
                    {'업체': 'GS네오텍', '구간': '1공구', '라인번호': '1-3L', '검사방법': 'PAUT'},
                    {'업체': '한국지역난방공사', '구간': '직영구간', '라인번호': 'MAIN-01', '검사방법': 'RT'},
                ]
            }
        }
        self.mapping = extract_company_section_line_mapping(self.sample_history)

    def test_extract_mapping_collects_all(self):
        self.assertEqual(
            self.mapping['all_companies'],
            ['GS네오텍', '다림건설(주)', '한국지역난방공사']
        )
        self.assertEqual(
            self.mapping['all_sections'],
            ['1공구', '2공구', '3공구', '직영구간']
        )
        self.assertEqual(
            self.mapping['all_lines'],
            ['1-1L', '1-2L', '1-3L', '2-1L', '3-1L', 'MAIN-01']
        )

    def test_get_sections_for_company(self):
        gs_sections = get_sections_for_company(self.mapping, 'GS네오텍')
        self.assertEqual(gs_sections, ['1공구', '2공구'])

        darim_sections = get_sections_for_company(self.mapping, '다림건설(주)')
        self.assertEqual(darim_sections, ['3공구'])

        all_sections = get_sections_for_company(self.mapping, '전체')
        self.assertEqual(all_sections, ['1공구', '2공구', '3공구', '직영구간'])

    def test_get_lines_for_company_and_section(self):
        # GS네오텍 + 1공구
        lines_1 = get_lines_for_company_and_section(self.mapping, 'GS네오텍', '1공구')
        self.assertEqual(lines_1, ['1-1L', '1-2L', '1-3L'])

        # GS네오텍 + 2공구
        lines_2 = get_lines_for_company_and_section(self.mapping, 'GS네오텍', '2공구')
        self.assertEqual(lines_2, ['2-1L'])

        # GS네오텍 + 전체 구간 -> all lines of GS네오텍
        lines_all_gs = get_lines_for_company_and_section(self.mapping, 'GS네오텍', '전체')
        self.assertEqual(lines_all_gs, ['1-1L', '1-2L', '1-3L', '2-1L'])

        # 전체 업체 + 전체 구간 -> all lines
        lines_all = get_lines_for_company_and_section(self.mapping, '전체', '전체')
        self.assertEqual(lines_all, ['1-1L', '1-2L', '1-3L', '2-1L', '3-1L', 'MAIN-01'])

    def test_get_lines_for_company_section_and_method(self):
        # GS네오텍 + 1공구 + PAUT -> only PAUT lines (1-1L, 1-3L; 1-2L is RT)
        paut_lines_1 = get_lines_for_company_and_section(self.mapping, 'GS네오텍', '1공구', 'PAUT')
        self.assertEqual(paut_lines_1, ['1-1L', '1-3L'])

        # GS네오텍 + 1공구 + RT -> only RT lines (1-2L)
        rt_lines_1 = get_lines_for_company_and_section(self.mapping, 'GS네오텍', '1공구', 'RT')
        self.assertEqual(rt_lines_1, ['1-2L'])

        # 전체 업체 + 전체 구간 + PAUT -> all PAUT lines
        all_paut = get_lines_for_company_and_section(self.mapping, '전체', '전체', 'PAUT')
        self.assertEqual(all_paut, ['1-1L', '1-3L', '2-1L'])


if __name__ == '__main__':
    unittest.main()

