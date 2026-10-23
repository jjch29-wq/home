import ast
import os
from pathlib import Path
import tempfile
import unittest
import zipfile

import openpyxl


class MTWorkbookFormatTests(unittest.TestCase):
    def test_mt_output_extension_matches_excel_package_type(self):
        path = Path(__file__).resolve().parents[1] / 'src' / '비파괴검사보고서.py'
        source = ast.parse(path.read_text(encoding='utf-8-sig'))
        method = next(node for node in ast.walk(source)
                      if isinstance(node, ast.FunctionDef) and node.name == '_run_mt_process')
        assignments = {node.targets[0].id: node for node in ast.walk(method)
                       if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name)}
        with tempfile.TemporaryDirectory() as directory:
            for extension, expected_type in [('.xlsx', 'spreadsheetml.sheet.main+xml'),
                                              ('.xlsm', 'macroEnabled.main+xml')]:
                template = Path(directory) / ('template' + extension)
                workbook = openpyxl.Workbook()
                workbook.save(template)
                if extension == '.xlsm':
                    with zipfile.ZipFile(template) as archive:
                        contents = {name: archive.read(name) for name in archive.namelist()}
                    contents['[Content_Types].xml'] = contents['[Content_Types].xml'].replace(
                        b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                        b'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
                    with zipfile.ZipFile(template, 'w') as archive:
                        for name, data in contents.items():
                            archive.writestr(name, data)
                context = {'os': os, 'openpyxl': openpyxl, 'template_path': str(template)}
                for name in ('keep_vba', 'wb', 'output_ext'):
                    exec(compile(ast.Module(body=[assignments[name]], type_ignores=[]), str(path), 'exec'), context)
                output = Path(directory) / ('output' + context['output_ext'])
                context['wb'].save(output)
                with zipfile.ZipFile(output) as archive:
                    self.assertIn(expected_type, archive.read('[Content_Types].xml').decode())
                self.assertEqual(output.suffix, extension)
                reopened = openpyxl.load_workbook(output, keep_vba=context['keep_vba'])
                reopened.close()
                for book in (context['wb'], reopened):
                    if book.vba_archive:
                        book.vba_archive.close()
