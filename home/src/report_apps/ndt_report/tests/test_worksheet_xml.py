from pathlib import Path
import sys
import unittest
import xml.etree.ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from worksheet_xml import normalize_print_options


class WorksheetXmlTests(unittest.TestCase):
    def test_extra_page_print_options_follow_merges_and_phonetics(self):
        xml = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>A&amp;B</t></is></c></row></sheetData>'
               '<printOptions horizontalCentered="1"/><mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>'
               '<phoneticPr fontId="2"/><pageMargins left="0.25"/><pageSetup scale="87"/>'
               '<headerFooter><oddHeader>Page 3 of 4</oddHeader></headerFooter></worksheet>')
        fixed = normalize_print_options(xml)
        names = [e.tag.split('}')[-1] for e in ET.fromstring(fixed)]
        self.assertEqual(names, ['sheetData', 'mergeCells', 'phoneticPr', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter'])
        self.assertIn('<t>A&amp;B</t>', fixed)
        self.assertIn('ref="A2:B2"', fixed)
        self.assertEqual(normalize_print_options(fixed), fixed)

    def test_no_print_options_is_unchanged(self):
        xml = '<worksheet><sheetData/></worksheet>'
        self.assertEqual(normalize_print_options(xml), xml)
