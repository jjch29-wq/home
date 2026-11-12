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
    collect_ledger_records_from_history,
    find_daily_usage_context,
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
            self.assertEqual(ws_mt['B29'].value, '1ST PINION / SCM420')
            self.assertEqual(ws_mt['F29'].value, 'V')
            self.assertEqual(ws_mt['H30'].value, 'V')
            self.assertIn('Ø85.2 × 555L', ws_mt['U29'].value)
            self.assertEqual(ws_mt['L31'].value, '~  이  하  여  백  ~')

            wb_ut = openpyxl.load_workbook(paths['UT'])
            self.assertEqual(wb_ut['UT(을)']['A10'].value, 'SHAFT-01')
            self.assertEqual(wb_ut['UT(을)']['E10'].value, 'NO RECORDABLE INDICATION')

            wb_pt = openpyxl.load_workbook(paths['PT'])
            self.assertEqual(wb_pt['PT(갑)']['A30'].value, 'SHAFT-02 (STS304)')
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
            self.assertEqual(wb['MT(갑) (2)']['B29'].value, 'GEAR-8 / SCM420')
            self.assertEqual(wb['MT(갑) (2)']['B30'].value, 'GEAR-9 / SCM420')

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
        self.assertEqual(ordered[0], mt_preset['items'][0])
        self.assertEqual(ordered[-1], 'OTHER')

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
        self.assertIn('UNIVERSAL JOINT', mt_items)
        self.assertEqual(mt_items[-1], 'OTHER')


if __name__ == '__main__':
    unittest.main()

