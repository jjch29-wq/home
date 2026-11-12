from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl
import pandas as pd


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))

from site_apps.gyeongin.src.cert_exporter import (
    CertificateExporter,
    CertificateLedgerExporter,
    build_cert_batches,
    classify_standard_family,
    collect_ledger_records_from_history,
    find_daily_usage_context,
    resolve_form_code,
)


class GyeonginCertExporterTests(unittest.TestCase):
    def test_daily_usage_join_fills_site_equipment_worker_and_customer(self):
        df = pd.DataFrame([
            {
                'Date': '2026-11-12',
                '업체명': '나라삼양감속기',
                'Site': 'GY025',
                '검사방법': 'MT',
                '장비명': 'Yoke AC(MY-2 NW0453)',
                'Workers': '부장 주진철',
                '성적서번호': '',
                '적용코드': 'KS',
                '검사품명': '1ST PINION',
                '발주처': '삼양종합기계',
            }
        ])
        ctx = find_daily_usage_context(df, '2026-11-12', '나라삼양감속기', 'MT')
        self.assertEqual(ctx['site'], 'GY025')
        self.assertEqual(ctx['equipment'], 'Yoke AC(MY-2 NW0453)')
        self.assertEqual(ctx['workers'], '주진철')
        self.assertEqual(ctx['customer'], '삼양종합기계')

    def test_all_five_method_templates_export_and_reload_cleanly(self):
        df = pd.DataFrame([
            {
                'Date': '2026-11-12',
                '업체명': '나라삼양감속기',
                'Site': 'GY025',
                '검사방법': 'MT',
                '장비명': 'Yoke AC(MY-2 NW0453)',
                'Workers': '부장 주진철',
                '적용코드': 'KS',
            }
        ])
        rows = [
            {
                '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                '검사품명': '1ST PINION', '식별번호': '1ST PINION', '재질': 'SCM420',
                '치수원문': 'Ø85.2 × 555L', '수량': '1', '단위': 'EA', '결과': '합격',
                '적용규격': 'KS D 0213', '검사절차서': 'SIS H 301', '개정번호': '4',
                '발주처': '포스코 포항',
            },
            {
                '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                '검사품명': '1ST GEAR', '식별번호': '1ST GEAR', '재질': 'SCM420',
                '치수원문': 'Ø334.4 × 80L', '수량': '1', '단위': 'EA', '결과': '불합격',
            },
            {
                '업체': '제일푸리', '검사방법': 'UT', '대상유형': '기계부품',
                '검사품명': 'SHAFT', '식별번호': 'SHAFT-01', '재질': 'STS304',
                '치수원문': 'Ø90×979L', '수량': '1', '단위': 'EA', '결과': '합격',
                '적용규격': 'ASME Sec.V', '검사절차서': 'SIS H 247',
            },
            {
                '업체': '제일푸리', '검사방법': 'PT', '대상유형': '기계부품',
                '검사품명': 'SHAFT', '식별번호': 'SHAFT-02', '재질': 'STS304',
                '치수원문': 'Ø90×979L / 1EA', '수량': '1', '단위': 'EA', '결과': '합격',
            },
            {
                '업체': '한국산업인력공단', '검사방법': 'RT', '대상유형': '판재/시험편',
                '검사품명': 'PLATE', '식별번호': '301', '재질': 'SS400',
                '치수원문': '200×300×12T', '수량': '1', '단위': '매', '결과': '합격',
            },
            {
                '업체': '모나스펌프', '검사방법': 'PMI', '대상유형': 'PMI',
                '검사품명': 'CASING', '태그번호': '236-J-05A',
                '구성부품': 'CASING BODY-1', '요구재질': 'SS316',
                '판정재질': 'A351 CF3M', '수량': '1', '단위': 'EA', '결과': '합격',
            },
        ]

        batches = build_cert_batches(rows, target_date='2026-11-12', default_site='경인사무소', daily_usage_df=df)
        self.assertEqual(len(batches), 5)

        exporter = CertificateExporter()
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            paths = {}
            for batch in batches:
                out = exporter.export_certificate(batch, tmp_dir / f"{batch['company']}_{batch['method']}.xlsx")
                self.assertTrue(out.exists())
                paths[batch['method']] = out

            wb_mt = openpyxl.load_workbook(paths['MT'])
            ws_mt = wb_mt['MT(갑)']
            self.assertEqual(ws_mt['S4'].value, '포스코 포항')
            self.assertEqual(ws_mt['B9'].value, 'GY025')
            self.assertEqual(ws_mt['D15'].value, 'NW0453')
            self.assertEqual(ws_mt['B29'].value, '1ST PINION')
            self.assertEqual(ws_mt['F29'].value, 'V')
            self.assertEqual(ws_mt['H30'].value, 'V')
            self.assertIn('Ø85.2 × 555L', ws_mt['U29'].value)
            self.assertEqual(ws_mt['L31'].value, '~  이  하  여  백  ~')

            wb_ut = openpyxl.load_workbook(paths['UT'])
            self.assertEqual(wb_ut['UT(을)']['A10'].value, 'SHAFT-01')
            self.assertEqual(wb_ut['UT(을)']['E10'].value, 'NO RECORDABLE INDICATION')

            wb_pt = openpyxl.load_workbook(paths['PT'])
            self.assertEqual(wb_pt['PT(갑)']['A30'].value, 'SHAFT-02')
            self.assertEqual(wb_pt['PT(갑)']['R30'].value, 'Ø90×979L / 1EA')

            wb_rt = openpyxl.load_workbook(paths['RT'])
            self.assertEqual(wb_rt['RT(을)']['A11'].value, '301')
            self.assertEqual(wb_rt['RT(을)']['C11'].value, 'V')

            wb_pmi = openpyxl.load_workbook(paths['PMI'])
            self.assertEqual(wb_pmi['PMI(을)']['A20'].value, '236-J-05A')
            self.assertEqual(wb_pmi['PMI(을)']['E20'].value, 'CASING BODY-1')
            self.assertEqual(wb_pmi['PMI(을)']['N20'].value, 'A351 CF3M')

    def test_mt_pagination_creates_second_page_over_seven_items(self):
        rows = [
            {
                '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                '검사품명': f'GEAR-{i}', '식별번호': f'GEAR-{i}', '재질': 'SCM420',
                '치수원문': 'Ø100×50L', '수량': '1', '단위': 'EA', '결과': '합격',
            }
            for i in range(1, 10)
        ]
        batches = build_cert_batches(rows, target_date='2026-11-12')
        self.assertEqual(len(batches), 1)
        exporter = CertificateExporter()
        with tempfile.TemporaryDirectory() as tmp:
            out = exporter.export_certificate(batches[0], Path(tmp) / 'multi_mt.xlsx')
            wb = openpyxl.load_workbook(out)
            self.assertEqual(wb.sheetnames, ['MT(갑)', 'MT(갑) (2)'])
            self.assertEqual(wb['MT(갑)']['X3'].value, '2')
            self.assertEqual(wb['MT(갑) (2)']['V3'].value, '2')
            self.assertEqual(wb['MT(갑) (2)']['B29'].value, 'GEAR-8')
            self.assertEqual(wb['MT(갑) (2)']['B30'].value, 'GEAR-9')

    def test_ledger_exporter_creates_and_appends_without_duplicates(self):
        history = {
            '2026-11-12': {
                'ndt_results': [
                    {
                        '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                        '검사품명': 'PINION', '식별번호': 'P-1', '성적서번호': 'SISGI-SAMY-MT-001',
                        '결과': '합격',
                    }
                ]
            },
            '2026-11-13': {
                'ndt_results': [
                    {
                        '업체': '나라삼양감속기', '검사방법': 'UT', '대상유형': '기계부품',
                        '검사품명': 'SHAFT', '식별번호': 'S-1', '성적서번호': 'SISGI-SAMY-UT-002',
                        '결과': '합격',
                    }
                ]
            },
        }
        all_recs = collect_ledger_records_from_history(history, company_filter='나라삼양감속기')
        self.assertEqual(len(all_recs), 2)

        ledger_exporter = CertificateLedgerExporter()
        with tempfile.TemporaryDirectory() as tmp:
            ledger_path = Path(tmp) / '2026년 나라삼양감속기 성적서 발행대장.xlsx'
            ledger_exporter.export_ledger([all_recs[0]], ledger_path, company_name='나라삼양감속기', year='2026')
            wb1 = openpyxl.load_workbook(ledger_path)
            ws1 = wb1.active
            self.assertEqual(ws1['A1'].value, '2026년 나라삼양감속기 시험성적서 발행대장')
            self.assertEqual(ws1['A5'].value, '11월')
            self.assertEqual(ws1['C6'].value, 'SISGI-SAMY-MT-001')

            # Re-append both records: first must deduplicate, second must append
            ledger_exporter.export_ledger(all_recs, ledger_path, existing_ledger_path=ledger_path)
            wb2 = openpyxl.load_workbook(ledger_path)
            ws2 = wb2.active
            report_nos = [ws2.cell(r, 3).value for r in range(5, ws2.max_row + 1) if ws2.cell(r, 3).value]
            self.assertEqual(report_nos, ['SISGI-SAMY-MT-001', 'SISGI-SAMY-UT-002'])

    def test_daily_work_log_exporter_handles_non_pipe_and_mixed_rows(self):
        from site_apps.gyeongin.src.daily_work_log_exporter import DailyWorkLogExporter

        exporter = DailyWorkLogExporter()
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            non_pipe_path = tmp_dir / 'non_pipe_log.xlsx'
            exporter.generate_excel({
                'date': '2026-11-12',
                'weather': '맑음',
                'remarks': '',
                'ndt_results': [
                    {
                        '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                        '검사품명': '1ST PINION', '식별번호': 'P-01', '재질': 'SCM420',
                        '치수원문': 'Ø85.2 × 555L', '수량': '2', '단위': 'EA', '결과': '합격',
                        '적용규격': 'KS D 0213', '도면번호': 'DWG-01',
                    }
                ],
            }, str(non_pipe_path))
            wb1 = openpyxl.load_workbook(non_pipe_path)
            ws1 = wb1['작업감독일보']
            self.assertEqual(ws1['D30'].value, '대상유형')
            self.assertEqual(ws1['E32'].value, '1ST PINION')
            self.assertEqual(ws1['P32'].value, '2')
            self.assertEqual(ws1['R32'].value, 'EA')

            mixed_path = tmp_dir / 'mixed_log.xlsx'
            exporter.generate_excel({
                'date': '2026-11-12',
                'weather': '맑음',
                'remarks': '',
                'ndt_results': [
                    {
                        '업체': '현대건설', '검사방법': 'MT', '대상유형': '배관',
                        '구간': 'S-1', '라인번호': 'L-100', 'Joint No.': 'J-1', '관경': '200A',
                        '결과': '합격', 'MT': '0.6280',
                    },
                    {
                        '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                        '검사품명': 'SHAFT', '식별번호': 'S-99', '재질': 'S45C',
                        '치수원문': 'Ø90×500L', '수량': '1', '단위': 'EA', '결과': '합격',
                    },
                ],
            }, str(mixed_path))
            wb2 = openpyxl.load_workbook(mixed_path)
            self.assertIn('일반·부품 검사결과', wb2.sheetnames)
            self.assertEqual(wb2['일반·부품 검사결과']['E3'].value, 'SHAFT')

    def test_ndt_summary_exporter_outputs_non_pipe_units_and_fields(self):
        from site_apps.gyeongin.src.ndt_summary_exporter import NDTSummaryExporter

        history = {
            '2026-11-12': {
                'ndt_results': [
                    {
                        '업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품',
                        '검사품명': '1ST PINION', '식별번호': 'P-01', '재질': 'SCM420',
                        '치수원문': 'Ø85.2 × 555L', '수량': '3', '단위': 'EA', '결과': '합격',
                        '적용규격': 'KS D 0213',
                    }
                ]
            }
        }
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / 'summary.xlsx'
            NDTSummaryExporter(history).generate(str(out))
            wb = openpyxl.load_workbook(out)
            ws_sum = wb['총괄 집계표']
            self.assertEqual(ws_sum['B3'].value, '기계부품')
            self.assertEqual(ws_sum['C3'].value, '1ST PINION')
            self.assertEqual(ws_sum['G3'].value, '3 EA')
            ws_mt = wb['MT']
            self.assertEqual(ws_mt['E3'].value, '1ST PINION')
            self.assertEqual(ws_mt['L3'].value, 'EA')

    def test_monthly_report_manager_company_filter_excludes_unrelated(self):
        import json
        from unittest.mock import patch
        from site_apps.gyeongin.src.services.monthly_report_manager import MonthlyReportManager

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            hist_path = tmp_dir / 'history.json'
            tmpl_path = tmp_dir / 'tmpl.xlsx'
            openpyxl.Workbook().save(tmpl_path)
            hist_path.write_text(json.dumps({
                '2026-11-12': {
                    'ndt_results': [{'업체': '나라삼양감속기', '검사방법': 'MT', '대상유형': '기계부품', '수량': '1', '단위': 'EA'}]
                }
            }, ensure_ascii=False), encoding='utf-8')
            mgr = MonthlyReportManager(str(tmpl_path))
            res = mgr.generate_report(str(hist_path), '2026-11', str(tmp_dir / 'out.xlsx'), company_filter='현대건설')
            self.assertIsNone(res)

    def test_company_method_presets_drive_pc_and_mobile_config(self):
        import tkinter as tk
        from site_apps.gyeongin.mobile import server
        from site_apps.gyeongin.src.company_master import (
            DEFAULT_COMPANIES,
            load_company_ndt_presets,
            merge_company_items,
            suggest_company_method_preset,
        )
        from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab

        load_company_ndt_presets.cache_clear()
        all_presets = load_company_ndt_presets()
        for co in DEFAULT_COMPANIES:
            if co != '경인사무소':
                self.assertIn(co, all_presets)

        sk_rt = suggest_company_method_preset('SK에코플랜트', 'RT')
        self.assertEqual(sk_rt['target_type'], '배관')
        self.assertIn('강관 Pipe (1500A)', sk_rt['items'])

        mt_preset = suggest_company_method_preset('나라삼양감속기', 'MT')
        self.assertEqual(mt_preset['target_type'], '기계부품')
        self.assertIn('PINION', mt_preset['items'])
        ordered = merge_company_items('나라삼양감속기', 'MT', fallback_items=['OTHER'])
        self.assertEqual(ordered, mt_preset['items'])
        self.assertNotIn('OTHER', ordered)
        unknown_ordered = merge_company_items('신규미등록업체', 'MT', fallback_items=['OTHER'])
        self.assertEqual(unknown_ordered, ['OTHER'])

        cfg = server.get_config()
        self.assertIn('company_ndt_presets', cfg)
        self.assertIn('나라삼양감속기', cfg['company_ndt_presets'])

        root = tk.Tk()
        root.withdraw()
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tab = DailyWorkLogTab(root, main_app=None)
                tab.history_path = str(Path(tmp) / 'history.json')
                row = tab.ndt_grid_entries[0]
                row['업체'].set('나라삼양감속기')
                row['검사방법'].set('MT')
                tab._apply_company_method_preset(row, source='company')
                self.assertEqual(row['대상유형'].get(), '기계부품')
                self.assertEqual(tab.ndt_view_var.get(), '기계부품')
                self.assertIn('KS D 0213', row['적용규격'].get())
                choices = list(row['검사품명'].cget('values'))
                self.assertEqual(choices[0], mt_preset['items'][0])
        finally:
            root.destroy()

    def test_method_material_and_item_candidates_for_daily_tab(self):
        from site_apps.gyeongin.src.company_master import (
            method_inspection_item_candidates,
            method_material_candidates,
        )

        inv = [
            'Carestream T200-3⅓*12"',
            'MT약품 (백색페인트)',
            'PT약품 (침투제)',
            '표준시험편 /성적서 보유 여부기재 (STB-A1)',
        ]
        mt_mats = method_material_candidates('MT', inventory_materials=inv)
        self.assertIn('MT약품 (백색페인트)', mt_mats)
        self.assertIn('NAWOO 습식형광자분 (Wet Fluorescent)', mt_mats)
        self.assertNotIn('Carestream T200-3⅓*12"', mt_mats)

        rt_mats = method_material_candidates('RT', inventory_materials=inv)
        self.assertEqual(rt_mats[0], 'Carestream T200-3⅓*12"')
        ut_mats = method_material_candidates('UT', inventory_materials=inv)
        self.assertIn('접촉매질 (Oil)', ut_mats)

        mt_items = method_inspection_item_candidates('나라삼양감속기', 'MT', fallback_items=['OTHER'])
        self.assertIn('PINION', mt_items)
        self.assertNotIn('UNIVERSAL JOINT', mt_items)
        self.assertNotIn('OTHER', mt_items)

        ksmc_items = method_inspection_item_candidates('경성기공', 'MT', fallback_items=['Piping', 'Tank'])
        self.assertEqual(ksmc_items, ['UNIVERSAL JOINT', 'SHAFT', 'SPINDLE', "YOKE ASS'Y", 'FM SPINDLE'])

        unknown_items = method_inspection_item_candidates('신규미등록업체', 'MT', fallback_items=['OTHER'])
        self.assertIn('OTHER', unknown_items)

    def test_daily_usage_view_autofills_company_method_item_and_material(self):
        import tkinter as tk
        from unittest import mock
        from site_apps.gyeongin.src import data_sync, sync_google_sheet
        from site_apps.gyeongin.src.app import MaterialManager

        with (
            mock.patch.object(data_sync, '_configured_data_dir', lambda: None),
            mock.patch.object(sync_google_sheet, 'sync_all_gyeongin_sources', return_value={'status': 'skipped', 'added_count': 0}),
            mock.patch.object(MaterialManager, 'save_data', return_value=True),
        ):
            root = tk.Tk()
            root.withdraw()
            try:
                app = MaterialManager(root)
                app.cb_daily_company.set('경성기공')
                app.cb_daily_test_method.set('MT')
                app.on_method_change_auto_unit_logic()
                self.assertEqual(app.ent_daily_inspection_item.get(), 'UNIVERSAL JOINT')
                self.assertIn('SHAFT', list(app.ent_daily_inspection_item.cget('values')))
                self.assertEqual(app.inspection_target_var.get(), '배관 외')
                self.assertEqual(app.ent_daily_applied_code.get(), 'KS')
                self.assertTrue(any('MT' in m or '자분' in m for m in app.cb_daily_material.cget('values')))

                app.cb_daily_company.set('제일푸리')
                app.cb_daily_test_method.set('UT')
                app.on_method_change_auto_unit_logic()
                self.assertEqual(app.ent_daily_inspection_item.get(), 'SHAFT')
                self.assertEqual(app.ent_daily_applied_code.get(), 'ASME')
            finally:
                root.destroy()

    def test_mt_pt_ut_certificates_embed_footer_sitco_logo(self):
        import zipfile
        import openpyxl
        from site_apps.gyeongin.src.cert_exporter import CertificateExporter

        exporter = CertificateExporter()
        with tempfile.TemporaryDirectory() as tmpdir:
            for method in ('MT', 'PT', 'UT'):
                out = Path(tmpdir) / f'footer_{method}.xlsx'
                exporter.export_certificate(
                    {
                        'method': method,
                        'company': '경성기공',
                        'date': '2026-11-12',
                        'report_no': f'SIS/GI-KSMC-{method}-2611-001',
                        'customer': '경성기공',
                        'project_name': 'N/A',
                        'item_name': 'SHAFT',
                        'code_procedure': 'KS',
                        'inspector': '주진철',
                        'rows': [{'품목명': 'SHAFT', '재질': 'sm450', '수량': 1, '단위': 'EA', '구간정보': 'ALL SURFACE', '결과': '합격'}],
                    },
                    out,
                )
                with zipfile.ZipFile(out) as zf:
                    self.assertIn('xl/media/footer_sitco.png', zf.namelist())
                    self.assertIn(b'<legacyDrawingHF', zf.read('xl/worksheets/sheet1.xml'))
                wb = openpyxl.load_workbook(out)
                if method == 'MT':
                    self.assertEqual(wb['MT(갑)']['L29'].value, '관련지시 없음')
                    self.assertIn(wb['MT(갑)']['D13'].value, (None, ''))
                    self.assertIn(wb['MT(갑)']['D15'].value, (None, ''))
                elif method == 'PT':
                    self.assertEqual(wb['PT(갑)']['H30'].value, '관련지시 없음')
                elif method == 'UT':
                    self.assertEqual(wb['UT(을)']['E10'].value, '관련지시 없음')
                wb.close()

            # Multi-equipment + Black Light (fluorescent MT)
            out_fl = Path(tmpdir) / 'fluent_mt.xlsx'
            exporter.export_certificate(
                {
                    'method': 'MT',
                    'company': '경성기공',
                    'date': '2026-11-12',
                    'report_no': 'SIS/GI-KSMC-MT-2611-002',
                    'customer': '경성기공',
                    'project_name': 'N/A',
                    'item_name': 'SHAFT',
                    'code_procedure': 'KS',
                    'inspector': '주진철',
                    'equipment': 'Yoke AC(MP-A3 KI-01); Yoke AC(MY-2 NW0453)',
                    'aux_equipment': 'Black Light (베터리교체형) UV-LED 2P/SPR (NW0096); UV Meter YK-35UV (071550)',
                    'rows': [{'품목명': 'SHAFT', '재질': 'sm450', '수량': 1, '단위': 'EA', '결과': '합격'}],
                },
                out_fl,
            )
            wb_fl = openpyxl.load_workbook(out_fl)
            ws_fl = wb_fl['MT(갑)']
            self.assertEqual(ws_fl['D13'].value, 'NAWOO')
            self.assertEqual(ws_fl['D14'].value, 'MP-A3 / MY-2')
            self.assertEqual(ws_fl['D15'].value, 'KI-01, NW0453')
            self.assertEqual(ws_fl['O17'].value, 'NAWOO')
            self.assertIn('UV-LED 2P/SPR', ws_fl['O18'].value)
            wb_fl.close()

    def test_code_procedure_candidates_prioritizes_company_and_excludes_truncated(self):
        from site_apps.gyeongin.src.cert_exporter import code_procedure_candidates

        mt_codes = code_procedure_candidates('MT', '경성기공')
        self.assertGreater(len(mt_codes), 3)
        self.assertEqual(mt_codes[0], 'KS D 0213 / SIS H 301 (2급,2군)')
        self.assertIn('ASME Sec.V Art.7 / SIS H 341 (Rev.13)', mt_codes)
        self.assertTrue(all(c.count('(') == c.count(')') for c in mt_codes))

    def test_import_daily_usage_rows_to_work_log_and_build_cert(self):
        import tkinter as tk
        import pandas as pd
        from site_apps.gyeongin.src.cert_exporter import build_cert_batches
        from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab

        root = tk.Tk()
        root.withdraw()
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tab = DailyWorkLogTab(root, main_app=None)
                tab.history_path = str(Path(tmp) / 'history.json')
                daily_rows = [
                    {
                        'Date': '2026-11-12',
                        'Site': '경성기공',
                        '업체명': '경성기공',
                        '검사방법': 'MT',
                        '검사품명': 'SHAFT',
                        '검사규격': 'Ø500 × 1000L',
                        '적용코드': 'KS D 0213 / SIS H 301 (2급,2군)',
                        '성적서번호': 'SIS/GI-KSMC-MT-2611-001',
                        '장비명': 'Yoke AC(MY-2 NW0453)',
                        '보조장비': 'Black Light (베터리교체형) UV-LED 2P/SPR (NW0096)',
                        'Usage': 1.0,
                        'Unit': 'EA',
                        'User': '부장 주진철',
                    },
                    {
                        'Date': '2026-11-12',
                        'Site': '경인사무소',
                        '업체명': '경인사무소',
                        '검사방법': '내근/대기',
                        '검사품명': '사무실 내근 및 장비점검',
                    },
                ]
                added, skipped = tab.import_daily_usage_rows(daily_rows, target_date='2026-11-12')
                self.assertEqual((added, skipped), (1, 0))
                self.assertEqual(tab.ndt_grid_entries[0]['업체'].get(), '경성기공')
                self.assertEqual(tab.ndt_grid_entries[0]['검사방법'].get(), 'MT')
                self.assertEqual(tab.ndt_grid_entries[0]['검사품명'].get(), 'SHAFT')
                self.assertEqual(tab.ndt_grid_entries[0]['치수원문'].get(), 'Ø500 × 1000L')
                self.assertEqual(tab.ndt_grid_entries[0]['수량'].get(), '1')
                self.assertEqual(tab.ndt_grid_entries[0]['단위'].get(), 'EA')
                self.assertEqual(tab.ndt_grid_entries[0]['결과'].get(), '합격')
                self.assertEqual(tab.ndt_view_var.get(), '기계부품')

                # Duplicate import is skipped cleanly
                added2, skipped2 = tab.import_daily_usage_rows(daily_rows, target_date='2026-11-12')
                self.assertEqual((added2, skipped2), (0, 1))

                # User fills missing detail (재질/식별번호) and builds cert batch
                tab.ndt_grid_entries[0]['재질'].insert(0, 'SM450')
                tab.ndt_grid_entries[0]['식별번호'].insert(0, 'SH-001')
                exported_rows = tab._current_ndt_rows_for_export()
                batches = build_cert_batches(
                    exported_rows,
                    target_date='2026-11-12',
                    daily_usage_df=pd.DataFrame(daily_rows),
                )
                self.assertEqual(len(batches), 1)
                self.assertEqual(batches[0]['equipment'], 'Yoke AC(MY-2 NW0453)')
                self.assertIn('UV-LED 2P/SPR', batches[0]['aux_equipment'])
                self.assertEqual(batches[0]['material'], 'SM450')
        finally:
            root.destroy()

    def test_load_from_daily_usage_respects_date_range_and_selected_summary_row(self):
        import datetime as dt
        import tkinter as tk
        from types import SimpleNamespace
        from unittest import mock
        import pandas as pd
        from site_apps.gyeongin.src.cert_exporter import build_cert_batches
        from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab

        df = pd.DataFrame([
            {
                'Date': '2026-11-09',
                'Site': '유성정공',
                '업체명': '현대산업',
                '검사방법': 'MT',
                '검사품명': 'CONVEYOR PULLEY',
                '장비명': 'Yoke AC(MY-2 NW0453)',
                'Usage': 16.9,
                'Unit': 'M',
                'User': '계장 이경재',
            },
            {
                'Date': '2026-11-11',
                'Site': '성원에드피아',
                '업체명': '현대건설',
                '검사방법': 'PMI',
                '검사품명': 'Beam',
                '장비명': '합금분석기(X-ray) XL2 100G(200983)',
                'Usage': 15.0,
                'Unit': 'EA',
                'User': '부장 주진철',
            },
        ])

        root = tk.Tk()
        root.withdraw()
        try:
            with tempfile.TemporaryDirectory() as tmp:
                app_stub = SimpleNamespace(daily_usage_df=df)
                tab = DailyWorkLogTab(root)
                tab.main_app = app_stub
                tab.history_path = str(Path(tmp) / 'history.json')
                tab.company_start_date.set_date(dt.date(2026, 11, 9))
                tab.date_entry.set_date(dt.date(2026, 11, 12))
                tab.refresh_company_summary()
                # Select 현대산업 MT row in company_summary_tree
                target_iid = next(
                    iid for iid in tab.company_summary_tree.get_children()
                    if tab.company_summary_tree.item(iid)['values'][:2] == ['현대산업', 'MT']
                )
                tab.company_summary_tree.selection_set(target_iid)

                with mock.patch('tkinter.messagebox.showinfo'):
                    tab.load_from_daily_usage()

                self.assertEqual(tab.ndt_grid_entries[0]['업체'].get(), '현대산업')
                self.assertEqual(tab.ndt_grid_entries[0]['검사방법'].get(), 'MT')
                self.assertEqual(tab.ndt_grid_entries[1]['업체'].get(), '')

                # Exporting from 2026-11-12 still finds 2026-11-09 equipment & site via _source_date
                batches = build_cert_batches(
                    tab._current_ndt_rows_for_export(),
                    target_date='2026-11-12',
                    daily_usage_df=df,
                )
                self.assertEqual(len(batches), 1)
                self.assertEqual(batches[0]['equipment'], 'Yoke AC(MY-2 NW0453)')
                self.assertEqual(batches[0]['project_name'], '유성정공')
        finally:
            root.destroy()


    def test_mixed_ks_and_asme_rows_split_into_separate_batches(self):
        base = {
            '업체': '현대산업', '검사방법': 'MT', '대상유형': '기계부품',
            '재질': 'SCM420', '치수원문': 'Ø100', '수량': '1', '단위': 'EA', '결과': '합격',
        }
        rows = [
            dict(base, 검사품명='GEAR-KS1', 식별번호='K1', 적용규격='KS D 0213 / SIS H 301'),
            dict(base, 검사품명='GEAR-AS1', 식별번호='A1', 적용규격='ASME Sec.V Art.7 / SIS H 341'),
            dict(base, 검사품명='GEAR-KS2', 식별번호='K2'),
        ]
        batches = build_cert_batches(rows, target_date='2026-11-12')
        self.assertEqual(len(batches), 2)
        fams = sorted(b['standard_family'] for b in batches)
        self.assertEqual(fams, ['ASME', 'KS'])
        ks = next(b for b in batches if b['standard_family'] == 'KS')
        self.assertEqual(len(ks['rows']), 2)
        self.assertNotEqual(batches[0]['report_no'], batches[1]['report_no'])

    def test_standard_family_and_form_code_helpers(self):
        self.assertEqual(classify_standard_family('KS D 0213 / SIS H 301'), 'KS')
        self.assertEqual(classify_standard_family('ASME Sec.Ⅴ (SA-388) / SIS H 247'), 'ASME')
        self.assertEqual(classify_standard_family('ISO 17638'), 'ISO')
        self.assertEqual(resolve_form_code('MT', 'ASME Sec.V Art.7'), 'H 341')
        self.assertEqual(resolve_form_code('PT', 'KS B 0816'), 'H 401')
        self.assertEqual(resolve_form_code('UT', 'KS D 0248 / SIS H 203'), 'H 203')
        self.assertEqual(resolve_form_code('UT', ''), 'H 247')

    def test_footer_form_code_follows_code_procedure(self):
        def make_batch(method, code):
            rows = [{
                '업체': '현대산업', '검사방법': method, '대상유형': '기계부품',
                '검사품명': 'P-1', '식별번호': 'P-1', '재질': 'SS400', '치수원문': 'Ø50',
                '수량': '1', '단위': 'EA', '결과': '합격', '적용규격': code,
            }]
            b = build_cert_batches(rows, target_date='2026-11-12')[0]
            b['code_procedure'] = code
            return b

        exporter = CertificateExporter()
        with tempfile.TemporaryDirectory() as tmp:
            out = exporter.export_certificate(make_batch('MT', 'ASME Sec.V Art.7 / SIS H 341'), Path(tmp) / 'mt.xlsx')
            wb = openpyxl.load_workbook(out)
            self.assertIn('H 341-1', wb.worksheets[0].oddFooter.left.text)

            out = exporter.export_certificate(make_batch('PT', 'KS B 0816 / SIS H 401'), Path(tmp) / 'pt.xlsx')
            wb = openpyxl.load_workbook(out)
            for ws in wb.worksheets:
                expected = 'H 401-2' if '(을)' in ws.title else 'H 401-1'
                self.assertIn(expected, ws.oddFooter.left.text)

            out = exporter.export_certificate(make_batch('UT', 'KS D 0248 / SIS H 203'), Path(tmp) / 'ut.xlsx')
            wb = openpyxl.load_workbook(out)
            for ws in wb.worksheets:
                expected = 'H 203-2' if '(을)' in ws.title else 'H 203-1'
                self.assertIn(expected, ws.oddFooter.left.text)

    def test_family_specific_template_is_preferred(self):
        import shutil
        src_dir = CertificateExporter().template_dir
        with tempfile.TemporaryDirectory() as tmp:
            tdir = Path(tmp) / 'cert_templates'
            tdir.mkdir()
            shutil.copy(src_dir / 'Template_Cert_MT.xlsx', tdir / 'Template_Cert_MT.xlsx')
            wb = openpyxl.load_workbook(src_dir / 'Template_Cert_MT.xlsx')
            wb.worksheets[0]['AZ1'] = 'ASME_TEMPLATE_MARK'
            wb.save(tdir / 'Template_Cert_MT_ASME.xlsx')

            rows = [{
                '업체': '현대산업', '검사방법': 'MT', '대상유형': '기계부품',
                '검사품명': 'P-1', '식별번호': 'P-1', '재질': 'SS400', '치수원문': 'Ø50',
                '수량': '1', '단위': 'EA', '결과': '합격', '적용규격': 'ASME Sec.V Art.7',
            }]
            batch = build_cert_batches(rows, target_date='2026-11-12')[0]
            out = CertificateExporter(tdir).export_certificate(batch, Path(tmp) / 'out.xlsx')
            self.assertEqual(openpyxl.load_workbook(out).worksheets[0]['AZ1'].value, 'ASME_TEMPLATE_MARK')

            rows[0]['적용규격'] = 'KS D 0213'
            batch = build_cert_batches(rows, target_date='2026-11-12')[0]
            out = CertificateExporter(tdir).export_certificate(batch, Path(tmp) / 'out_ks.xlsx')
            self.assertIsNone(openpyxl.load_workbook(out).worksheets[0]['AZ1'].value)


    def test_delete_selected_ndt_row_persists_to_history(self):
        import datetime as dt
        import json
        import tkinter as tk
        from types import SimpleNamespace
        from unittest import mock
        from site_apps.gyeongin.src.daily_work_log_tab import DailyWorkLogTab

        root = tk.Tk()
        root.withdraw()
        try:
            with tempfile.TemporaryDirectory() as tmp, \
                    mock.patch('tkinter.messagebox.askyesno', return_value=True), \
                    mock.patch('tkinter.messagebox.showinfo'), \
                    mock.patch('tkinter.messagebox.showwarning'):
                tab = DailyWorkLogTab(root)
                tab.main_app = SimpleNamespace(daily_usage_df=pd.DataFrame())
                tab.history_path = str(Path(tmp) / 'history.json')
                tab.date_entry.set_date(dt.date(2026, 11, 12))
                root.update()
                rows = tab.ndt_grid_entries
                for i, (comp, method) in enumerate((('경성기공', 'MT'), ('현대산업', 'MT'))):
                    rows[i]['업체'].set(comp)
                    rows[i]['검사방법'].set(method)
                    rows[i]['검사품명'].set('SHAFT')
                self.assertTrue(tab.save_current_history())

                tab.ndt_selection_vars[0].set(True)
                tab._delete_selected_ndt_row()

                with open(tab.history_path, encoding='utf-8') as f:
                    saved = json.load(f)['2026-11-12']['ndt_results']
                companies = [r.get('업체') for r in saved if r.get('업체')]
                self.assertEqual(companies, ['현대산업'])
        finally:
            root.destroy()


    def test_pmi_measurement_file_fills_element_columns(self):
        from site_apps.gyeongin.src.cert_exporter import read_pmi_measurements

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'vanta_export.xlsx'
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(['Instrument', 'Vanta VCR'])
            ws.append([])
            ws.append(['Reading No', 'Sample Name', 'Alloy 1', 'Description',
                       'Cr', 'Cr +/-', 'Ni', 'Ni +/-', 'Mo', 'Mn', 'Cu', 'Fe', 'Ti'])
            ws.append(['1', 'LINE-01', 'SS316', 'ELBOW', '17.12', '0.1', '10.05', '0.1', '2.11', '1.40', '0.30', '68.5', '<LOD'])
            ws.append(['2', 'LINE-02', 'SS304', 'PIPE', '18.30', '0.1', '8.10', '0.1', '<LOD', '1.62', '0.25', '71.0', '0.02'])
            wb.save(src)

            measurements = read_pmi_measurements(src)
            self.assertEqual(len(measurements), 2)
            self.assertEqual(measurements[0]['tag'], 'LINE-01')
            self.assertEqual(measurements[0]['alloy'], 'SS316')
            self.assertEqual(measurements[0]['component'], 'ELBOW')
            self.assertEqual(measurements[0]['elements']['Cr'], '17.12')
            self.assertEqual(measurements[1]['elements']['Mo'], 'ND')

            batch = build_cert_batches(
                [{'업체': '현대건설', '검사방법': 'PMI', '검사품명': 'Structure', '수량': '2'}],
                target_date='2026-11-12',
            )[0]
            batch['pmi_measurements'] = measurements
            out = CertificateExporter().export_certificate(batch, Path(tmp) / 'pmi.xlsx')
            eul = openpyxl.load_workbook(out)['PMI(을)']
            self.assertEqual([eul[f'{c}17'].value for c in 'HIJKLM'], ['Ni', 'Cr', 'Mo', 'Mn', 'Cu', 'Ti'])
            self.assertEqual(eul['A20'].value, 'LINE-01')
            self.assertEqual(eul['E20'].value, 'ELBOW')
            self.assertEqual(eul['H20'].value, '10.05')
            self.assertEqual(eul['I20'].value, '17.12')
            self.assertEqual(eul['N20'].value, 'SS316')
            self.assertEqual(eul['J21'].value, 'ND')
            self.assertEqual(eul['N21'].value, 'SS304')


if __name__ == '__main__':
    unittest.main()

