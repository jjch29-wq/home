from pathlib import Path
import sys
import unittest


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))

from site_apps.gyeongin.src.daily_work_log_tab import (
    NDT_DETAIL_FIELDS,
    NDT_DIMENSION_TEMPLATES,
    NDT_ITEM_CHOICES,
    NDT_TARGET_TYPES,
    infer_ndt_target,
    ndt_item_choices,
    ndt_dimension_choices,
    ndt_visible_columns,
    parse_machine_dimension,
)


class NdtTargetViewTests(unittest.TestCase):
    def test_each_target_keeps_common_columns_and_only_relevant_geometry(self):
        for target in NDT_TARGET_TYPES:
            columns = ndt_visible_columns(target)
            self.assertIn('업체', columns)
            self.assertIn('검사방법', columns)
            self.assertIn('대상유형', columns)
            self.assertIn('검사품명', columns)
            self.assertIn('치수원문', columns)

        self.assertIn('Joint No.', ndt_visible_columns('배관'))
        self.assertNotIn('Joint No.', ndt_visible_columns('기계부품'))
        self.assertIn('시험편번호', ndt_visible_columns('판재/시험편'))
        self.assertIn('직경', ndt_visible_columns('기계부품'))
        self.assertIn('설비번호', ndt_visible_columns('구조물/설비'))
        self.assertIn('판정재질', ndt_visible_columns('PMI'))

    def test_detail_only_fields_do_not_widen_bulk_grid(self):
        for field in ('적용규격', '검사절차서', '개정번호', '도면번호', '비고'):
            self.assertIn(field, NDT_DETAIL_FIELDS)
            for target in NDT_TARGET_TYPES:
                self.assertNotIn(field, ndt_visible_columns(target))

    def test_legacy_pipeline_rows_are_inferred_as_pipe(self):
        self.assertEqual(infer_ndt_target({'Joint No.': 'J-001'}), '배관')
        self.assertEqual(infer_ndt_target({'라인번호': 'L-01'}), '배관')
        self.assertEqual(infer_ndt_target({'대상유형': 'PMI', 'Joint No.': 'old'}), 'PMI')
        self.assertEqual(infer_ndt_target({'검사방법': 'MT'}), '')

    def test_unknown_view_safely_falls_back_to_pipe(self):
        self.assertEqual(ndt_visible_columns('잘못된 값'), ndt_visible_columns('배관'))

    def test_reviewed_item_choices_follow_target_type(self):
        self.assertIn('PIPE', ndt_item_choices('배관'))
        self.assertIn('PLATE', ndt_item_choices('판재/시험편'))
        self.assertIn('SHAFT', ndt_item_choices('기계부품'))
        self.assertIn('H-PILE', ndt_item_choices('구조물/설비'))
        self.assertIn('FLANGE', ndt_item_choices('PMI'))
        self.assertNotIn('SHAFT', ndt_item_choices('배관'))

    def test_saved_custom_items_are_appended_once_without_changing_defaults(self):
        defaults = NDT_ITEM_CHOICES['기계부품']
        choices = ndt_item_choices('기계부품', ('특수 로터', 'SHAFT', '특수 로터'))
        self.assertEqual(choices[:len(defaults)], defaults)
        self.assertEqual(choices.count('특수 로터'), 1)
        self.assertEqual(choices.count('SHAFT'), 1)

    def test_machine_dimension_text_splits_diameter_length_and_quantity(self):
        self.assertEqual(
            parse_machine_dimension('Ø90×979L / 1EA'),
            {'직경': '90', '길이': '979', '수량': '1', '단위': 'EA'},
        )
        self.assertEqual(
            parse_machine_dimension('Ø90/220 * 120L / 112EA'),
            {'직경': '90/220', '길이': '120', '수량': '112', '단위': 'EA'},
        )

    def test_non_dimension_numbers_are_not_misclassified(self):
        self.assertEqual(parse_machine_dimension('1.0TON CAPA'), {})
        self.assertEqual(parse_machine_dimension('400×400×13T×21T'), {})

    def test_dimension_dropdown_combines_editable_templates_and_history(self):
        defaults = NDT_DIMENSION_TEMPLATES['기계부품']
        choices = ndt_dimension_choices(
            '기계부품', ('Ø90×979L / 1EA', defaults[0], 'Ø90×979L / 1EA')
        )
        self.assertEqual(choices[:len(defaults)], defaults)
        self.assertEqual(choices.count('Ø90×979L / 1EA'), 1)
        self.assertIn('폭 × 길이 × 두께T / 수량EA', ndt_dimension_choices('판재/시험편'))


if __name__ == '__main__':
    unittest.main()
