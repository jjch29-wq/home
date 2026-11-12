from pathlib import Path
import sys
import tempfile
import unittest
from copy import deepcopy
import zipfile

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.pmi_criteria import (
    material_catalogue, normalize_ranges, save_report, restore_report,
    read_state, state_key, validate_criteria,
    saved_projects, restore_project,
)
from site_apps.gyeongin.src.cert_exporter import CertificateExporter


def criterion(name='Example steel'):
    return dict(id=name, label=name, family='기타', source='Test procedure Rev.0',
                ranges={'Mn': '0.90~1.50', 'Cr': '<=0.30', 'Ni': '-',
                        'Mo': '-', 'Cu': '<=0.40', 'Fe': 'Balance'})


class PmiCriteriaTests(unittest.TestCase):
    def test_full_report_restores_header_equipment_and_measurement_edits(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'state.json'
            batch = dict(date='2026-11-13', company='Builder', method='PMI', report_no='001',
                         project_name='Site A', customer='Customer', equipment='Analyzer A',
                         inspector='Inspector', code_procedure='ASME / SIS', aux_equipment='Light',
                         pmi_elements=['Mn'], pmi_criteria=[criterion()], pmi_source='readings.csv',
                         pmi_measurements=[{'tag': 'A-1', 'elements': {'Mn': '1.20'}}])
            save_report(batch, path)
            restored = {k: batch[k] for k in ('date', 'company', 'method', 'report_no', 'project_name')}
            restore_report(restored, path)
            self.assertEqual(restored, batch)

    def test_site_settings_reused_on_new_date_without_old_readings_or_number(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'state.json'
            base = dict(date='2026-11-13', company='Builder', method='PMI', report_no='001',
                        project_name='Site A', equipment='Analyzer A', pmi_elements=['Mn'],
                        pmi_criteria=[criterion()], pmi_source='old.csv',
                        pmi_measurements=[{'tag': 'old'}])
            save_report(base, path)
            other = deepcopy(base)
            other.update(project_name='Site B', equipment='Analyzer B')
            save_report(other, path)
            fresh = dict(date='2026-11-14', company='Builder', method='PMI', report_no='002', project_name='Site A')
            restore_report(fresh, path)
            self.assertEqual(fresh['equipment'], 'Analyzer A')
            self.assertEqual(fresh['report_no'], '002')
            self.assertEqual(fresh['date'], '2026-11-14')
            self.assertNotIn('pmi_measurements', fresh)
            self.assertNotIn('pmi_source', fresh)
            self.assertEqual(saved_projects(fresh, path), ['Site A', 'Site B'])
            self.assertTrue(restore_project(fresh, path, 'Site B'))
            self.assertEqual(fresh['equipment'], 'Analyzer B')
            fresh['company'] = 'Other builder'
            self.assertEqual(saved_projects(fresh, path), [])
            self.assertFalse(restore_project(fresh, path, 'Site A'))

    def test_legacy_criteria_remain_readable_without_guessing_a_named_site(self):
        import hashlib
        import json
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'state.json'
            identity = ['2026-11-13', 'Builder', '001']
            old_key = hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode()).hexdigest()
            path.write_text(json.dumps({'reports': {old_key: {'pmi_elements': ['Mn']}}, 'presets': {}}), encoding='utf-8')
            batch = dict(zip(('date', 'company', 'report_no'), identity))
            restore_report(batch, path)
            self.assertEqual(batch['pmi_elements'], ['Mn'])
            named = dict(zip(('date', 'company', 'report_no'), identity), project_name='Other site')
            restore_report(named, path)
            self.assertNotIn('pmi_elements', named)

    def test_changing_family_refreshes_detail_without_replacing_chosen_material(self):
        # Exercise the actual dialog callbacks without requiring installed Tcl.
        from unittest.mock import MagicMock, patch
        from contextlib import ExitStack
        from site_apps.gyeongin.src.pmi_criteria import edit_criteria_dialog

        class Variable:
            def __init__(self, value='', **kwargs):
                self.value, self.callbacks = value, []

            def get(self):
                return self.value

            def set(self, value):
                self.value = value
                for callback in self.callbacks:
                    callback()

            def trace_add(self, mode, callback):
                self.callbacks.append(callback)

        widgets = []

        def widget(kind):
            def create(*args, **kwargs):
                mock = MagicMock()
                attrs = dict(kwargs)
                mock.__getitem__.side_effect = attrs.__getitem__
                mock.__setitem__.side_effect = attrs.__setitem__
                mock.bindings = {}
                mock.bind.side_effect = lambda event, callback: mock.bindings.update({event: callback})
                mock.attrs = attrs
                widgets.append((kind, mock))
                return mock
            return create

        with tempfile.TemporaryDirectory() as tmp, ExitStack() as stack:
            stack.enter_context(patch('tkinter.StringVar', Variable))
            for name in ('Toplevel', 'Listbox'):
                stack.enter_context(patch(f'tkinter.{name}', widget(name)))
            for name in ('Frame', 'Label', 'Entry', 'Combobox', 'Button'):
                stack.enter_context(patch(f'tkinter.ttk.{name}', widget(name)))
            parent = MagicMock()
            original = criterion('Already selected')
            batch = dict(date='2026-11-13', company='QA', report_no='switch',
                         pmi_elements=['Ni', 'Cr', 'Mo'], pmi_criteria=[original])

            def interact(_):
                family, materials = [w for kind, w in widgets if kind == 'Combobox']
                listing = next(w for kind, w in widgets if kind == 'Listbox')
                query = materials.attrs['textvariable']
                query.set(next(i['label'] for i in material_catalogue()
                               if i['standard'] == 'ASTM A312' and i['grade'] == 'TP304L'))
                materials.bindings['<<ComboboxSelected>>']()
                # Even a selected full material name must not hide its siblings.
                materials.bindings['<KeyRelease>'](None)
                self.assertGreater(len(materials['values']), 1)
                listing.curselection.return_value = (0,)
                listing.bindings['<<ListboxSelect>>']()
                family.attrs['textvariable'].set('탄소강')
                candidates = {i['label']: i for i in material_catalogue()}
                self.assertEqual(candidates[query.get()]['family'], '탄소강')
                self.assertTrue(all(candidates[label]['family'] == '탄소강'
                                    for label in materials['values']))
                entries = [w for kind, w in widgets if kind == 'Entry']
                # Registered KS structural material is now first. Select a
                # pipe reference explicitly for the three chosen alloy elements.
                query.set(next(i['label'] for i in material_catalogue()
                               if i['standard'] == 'ASTM A106' and i['grade'] == 'Gr.B'))
                materials.bindings['<<ComboboxSelected>>']()
                self.assertEqual(entries[3].attrs['textvariable'].get(), '<=0.40')
                buttons = {w.attrs['text']: w for kind, w in widgets if kind == 'Button'}
                buttons['재질 추가 / 선택 항목 수정'].attrs['command']()
                buttons['목록 적용·저장'].attrs['command']()

            parent.wait_window.side_effect = interact
            edit_criteria_dialog(parent, batch, Path(tmp) / 'state.json')
            self.assertEqual(len(batch['pmi_criteria']), 2)
            self.assertEqual(batch['pmi_criteria'][0], original)
            self.assertEqual(batch['pmi_criteria'][1]['family'], '탄소강')

    def test_names_are_distinct_by_product_and_standard(self):
        items = material_catalogue()
        self.assertEqual(len(items), len({i['id'] for i in items}))
        for grade in ('STS316L', 'SUS316L', 'TP316L', 'WP316L', 'F316L', 'CF3M', 'SPP', 'SS275'):
            self.assertTrue(any(i['grade'] == grade for i in items), grade)
        self.assertTrue(any(i['standard'] == 'ASME SA-312' and i['ranges'] for i in items))
        self.assertEqual(next(i for i in items if i['grade'] == 'SS275')['ranges']['Mn'], '<=1.40')
        self.assertEqual(next(i for i in items if i['grade'] == 'SM355B')['ranges']['Mn'], '<=1.60')
        self.assertFalse(next(i for i in items if i['grade'] == 'SPP')['ranges'])
        cf3m = next(i for i in items if i['grade'] == 'CF3M')
        self.assertEqual(cf3m['ranges']['Ni'], '9.00~13.00')
        self.assertIn('사용자 제공', cf3m['source'])

    def test_public_references_have_provenance_and_product_specific_values(self):
        from site_apps.gyeongin.src.pmi_reference_data import public_material_references
        items = public_material_references()
        self.assertEqual(len(items), len({i['id'] for i in items}))
        for item in items:
            self.assertEqual(normalize_ranges(item['ranges']), item['ranges'])
            self.assertTrue(item['source_url'].startswith('https://'))
            self.assertEqual(item['reference_status'], 'public_reference')
            self.assertFalse(item['official_edition_verified'])
            self.assertTrue(item['notes'])
        def get(standard, grade):
            return next(i for i in items if i['standard'] == standard and i['grade'] == grade)
        self.assertEqual(get('ASME SA-312', 'TP316L')['ranges']['Mo'], '2.00~3.00')
        self.assertEqual(get('ASME SA-106', 'Gr.B')['ranges']['Mn'], '0.29~1.06')
        self.assertIn('1.35%', get('ASME SA-106', 'Gr.B')['notes'])
        self.assertIn('1.65%', get('ASTM A106', 'Gr.B')['notes'])
        self.assertEqual(get('ASME SA-53', 'Gr.B Type S')['ranges']['Mn'], '<=1.20')
        self.assertEqual(get('ASME SA-105', 'A105')['ranges']['Mn'], '0.600~1.050')
        self.assertEqual(get('ASME SA-403', 'WP304L')['ranges']['Ni'], '8.00~12.00')
        self.assertEqual(get('ASME SA-335', 'P11')['ranges']['Cr'], '1.00~1.50')
        self.assertNotIn('Ni', get('ASME SA-335', 'P11')['ranges'])

    def test_public_reference_notes_and_links_survive_output_and_restore(self):
        item = next(i for i in material_catalogue() if i['standard'] == 'ASME SA-106' and i['grade'] == 'Gr.B')
        with tempfile.TemporaryDirectory() as tmp:
            batch = dict(method='PMI', date='2026-11-13', company='QA', report_no='ref',
                         pmi_elements=['Mn', 'Cr', 'Mo', 'Ni', 'Cu', 'V'], pmi_criteria=[item])
            path = Path(tmp) / 'state.json'
            save_report(batch, path)
            restored = {k: batch[k] for k in ('method', 'date', 'company', 'report_no')}
            restore_report(restored, path)
            self.assertEqual(restored['pmi_criteria'][0], item)
            out = CertificateExporter().export_certificate(restored, Path(tmp) / 'ref.xlsx')
            wb = openpyxl.load_workbook(out)
            self.assertIn('1.35%', wb['PMI(갑)']['A35'].value)
            self.assertEqual(wb['PMI(갑)']['A35'].hyperlink.target, item['source_url'])
            self.assertEqual(wb['PMI(갑)']['G28'].value, '0.29~1.06')
            wb.close()

    def test_ranges_reject_ambiguous_or_reversed_limits(self):
        self.assertEqual(normalize_ranges({'Mn': '1.50% 이하'}), {'Mn': '<=1.50'})
        for invalid in ('1.50', '1.5~0.9', '-0.1~1.0', '101 이하', 'abc', '<=1.5 이상'):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                normalize_ranges({'Mn': invalid})

    def test_snapshots_restore_without_following_new_preset(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'state.json'
            batch = dict(date='2026-11-12', company='Example', report_no='001',
                         pmi_elements=['Mn'], pmi_criteria=[criterion()])
            save_report(batch, path)
            other = deepcopy(batch)
            other['report_no'] = '002'
            other['pmi_criteria'][0]['ranges']['Mn'] = '<=1.60'
            save_report(other, path)
            restored = {k: batch[k] for k in ('date', 'company', 'report_no')}
            restore_report(restored, path)
            self.assertEqual(restored['pmi_criteria'][0]['ranges']['Mn'], '0.90~1.50')
            self.assertEqual(read_state(path)['presets']['Example steel']['ranges']['Mn'], '<=1.60')
            restored['pmi_criteria'] = []
            save_report(restored, path)
            restore_report(batch, path)
            self.assertEqual(batch['pmi_criteria'], [])

    def test_changed_report_number_can_restore_under_original_identity(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'state.json'
            batch = dict(date='2026-11-12', company='Example', report_no='001')
            original = state_key(batch)
            batch.update(report_no='edited', pmi_elements=['Mn'], pmi_criteria=[criterion()])
            save_report(batch, path, original)
            self.assertIn(original, read_state(path)['reports'])

    def test_missing_new_element_blocks_output(self):
        with self.assertRaises(ValueError):
            validate_criteria([criterion()], ['Ti'])
        with tempfile.TemporaryDirectory() as tmp, self.assertRaises(ValueError):
            CertificateExporter().export_certificate(
                dict(method='PMI', pmi_elements=['Ti'], pmi_criteria=[criterion()]),
                Path(tmp) / 'invalid.xlsx')

    def test_cover_and_continuations_keep_frames_and_element_order(self):
        for count in (1, 2, 3, 4, 5, 8, 9):
            with self.subTest(count=count), tempfile.TemporaryDirectory() as tmp:
                elements = ['Mn', 'Cr', 'Ni', 'Mo', 'Cu', 'Fe']
                batch = dict(method='PMI', date='2026-11-12', pmi_elements=elements,
                             pmi_criteria=[criterion(f'Material {i}') for i in range(count)],
                             rows=[{'재질': 'Example', '수량': 1}])
                out = CertificateExporter().export_certificate(batch, Path(tmp) / 'report.xlsx')
                wb = openpyxl.load_workbook(out)
                expected_pages = 2 + (max(0, count - 4) + 3) // 4
                self.assertEqual(len(wb.worksheets), expected_pages)
                covers = [wb['PMI(갑)']] + [s for s in wb if s.title.startswith('PMI(기준')]
                heights = []
                for page_idx, sheet in enumerate(covers):
                    page_count = min(4, count - page_idx * 4)
                    dense = page_count > 2
                    first_row = 27 if dense else 28
                    self.assertEqual(sheet.cell(first_row, 1).value, 'Manganese')
                    self.assertEqual(sheet.cell(first_row, 7).value, '0.90~1.50')
                    self.assertEqual(sheet.cell(first_row + 6, 1).value, '* Tolerance ±10%')
                    self.assertEqual(sheet['E17'].value, 'N/A')
                    last_row = 64 if dense else 50
                    self.assertIn(f'$A$1:$T${last_row}', str(sheet.print_area))
                    heights.append(sum(sheet.row_dimensions[r].height or sheet.sheet_format.defaultRowHeight
                                       for r in range(1, last_row + 1)))
                    self.assertEqual(sheet.oddFooter.right.text, '&G')
                    self.assertIn(str(expected_pages), sheet['N3'].value)
                    self.assertIsNotNone(sheet['A23'].border.left.style)
                    self.assertEqual(sheet.cell(first_row, 1).alignment.wrap_text, True)
                    if dense:
                        self.assertIn('Acceptance Criteria', sheet['A38'].value)
                        self.assertEqual(sheet['A42'].value, 'Manganese')
                        self.assertEqual(sheet['K56'].value, '2026. 11. 12')
                for height in heights[1:]:
                    self.assertAlmostEqual(height, heights[0], places=3)
                if count > 1:
                    self.assertEqual(covers[0].cell(27 if count > 2 else 28, 15).value, '0.90~1.50')
                self.assertIn('H 801-1', wb['PMI(갑)'].oddFooter.left.text)
                self.assertIn('H 801-2', wb['PMI(을)'].oddFooter.left.text)
                self.assertEqual(wb['PMI(을)']['N20'].value, 'Example')
                wb.close()
                with zipfile.ZipFile(out) as package:
                    self.assertIn('xl/media/footer_sitco.png', package.namelist())


if __name__ == '__main__':
    unittest.main()
