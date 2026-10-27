import unittest
from unittest.mock import Mock, patch

import numpy as np
import ocr_analysis as ocr


class OcrAnalysisTests(unittest.TestCase):
    def test_complete_fast_result_skips_expensive_fallback(self):
        reader = Mock()
        reader.readtext.return_value = [(None, "marking", .9)]
        fields = dict(line_no="L1", joint_no="R01", welder_no="W1", section="(1-2)")
        with patch.object(ocr, "parse_marking", return_value=fields), \
             patch.object(ocr, "prepare_for_ocr") as variants, \
             patch.object(ocr, "recognize_film_section") as section:
            result = ocr.analyze_marking(reader, np.zeros((100, 3200, 3), dtype=np.uint8))
        self.assertFalse(result[-1])
        self.assertEqual(reader.readtext.call_args.args[0].shape[1], 1600)
        self.assertEqual(reader.readtext.call_count, 1)
        variants.assert_not_called()
        section.assert_not_called()

    def test_incomplete_fast_result_uses_detailed_analysis(self):
        reader = Mock()
        reader.readtext.return_value = [(None, "partial", .9)]
        fields = dict(line_no="", joint_no="R01", welder_no="", section="")
        image = np.zeros((100, 200, 3), dtype=np.uint8)
        with patch.object(ocr, "parse_marking", return_value=fields), \
             patch.object(ocr, "prepare_for_ocr", return_value=[image, image]), \
             patch.object(ocr, "recognize_film_section", return_value="(2-3)"):
            result = ocr.analyze_marking(reader, image)
        self.assertTrue(result[-1])
        self.assertEqual(result[-2], "(2-3)")
        self.assertEqual(reader.readtext.call_count, 3)


if __name__ == "__main__":
    unittest.main()
