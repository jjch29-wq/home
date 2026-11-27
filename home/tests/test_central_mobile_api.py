import unittest
import tempfile
import openpyxl
from pathlib import Path
from unittest.mock import patch
from fastapi import HTTPException

from home.src.site_apps.central.mobile.server import (
    save_record,
    get_recent_records,
    update_record,
    delete_record,
    RecordSaveRequest,
    RecordUpdateRequest,
    RecordDeleteRequest,
    WorkerItem
)

class TestCentralMobileApi(unittest.TestCase):
    def test_delete_rejects_ambiguous_time_and_targets_exact_id(self):
        entry_time = "2026-11-10 09:00:00"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "inventory.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "DailyUsage"
            sheet.append(["EntryTime", "ClientRequestID"])
            sheet.append([entry_time, "first"])
            sheet.append([entry_time, "second"])
            book.save(path)
            original = path.read_bytes()

            with patch("home.src.site_apps.central.mobile.server.get_inventory_paths", return_value=[str(path)]):
                for request, status in (
                    (RecordDeleteRequest(id=0, entry_time=entry_time), 409),
                    (RecordDeleteRequest(id=0), 404),
                    (RecordDeleteRequest(id=0, entry_time=entry_time, client_request_id="missing"), 404),
                ):
                    with self.assertRaises(HTTPException) as error:
                        delete_record(request)
                    self.assertEqual(error.exception.status_code, status)
                    self.assertEqual(path.read_bytes(), original)
                self.assertEqual(delete_record(RecordDeleteRequest(client_request_id="second"))["status"], "success")
            result = openpyxl.load_workbook(path, read_only=True)
            try:
                self.assertEqual(list(result["DailyUsage"].values)[1:], [(entry_time, "first")])
            finally:
                result.close()

    def test_mobile_crud_lifecycle(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            test_db = Path(tmp_dir) / "Material_Inventory.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DailyUsage"

            headers = [
                'Date', 'Site', '작업구분', '구분', '검사구분', '업체명', '회사코드',
                '검사방법', '규격', '관경(Inch)', '검사품명', '자재명', 'MaterialID', 'Material', '품목명',
                '적용코드', '성적서번호', '조인트수', '불량수', '수량', '검사량', 'Usage', '단위', 'Unit',
                '단가', '출장비', '일식', '검사비', '장비명', '보조장비', '작업형태', '차량번호',
                '주행거리', '차량점검', '차량비고', 'Note', 'Remark', 'DailyMemo', 'EntryTime', '입력시간',
                'User', 'WorkTime', 'OT', '제경비율', '기술료율', '핫태핑차수', '핫태핑개소', '핫태핑관경',
                'RTK_센터미스', 'RTK_농도', 'RTK_마킹미스', 'RTK_필름마크', 'RTK_취급부주의', 'RTK_고객불만',
                'RTK_기타', 'RTK_총계', 'RTK_RTK총계',
                'NDT_형광자분', 'NDT_흑색자분', 'NDT_백색페인트', 'NDT_침투제', 'NDT_세척제', 'NDT_현상제', 'NDT_형광침투제'
            ]
            ws.append(headers)
            wb.save(test_db)

            with patch("home.src.site_apps.central.mobile.server.get_inventory_paths", return_value=[str(test_db)]):
                # 1. Save record
                req_save = RecordSaveRequest(
                    date="2026-11-10",
                    site="중앙지사",
                    company="한국지역난방공사 중앙지사",
                    loc_type="핫태핑",
                    hot_tap_stage="1차",
                    hot_tap_joints="2",
                    hot_tap_pipe="300A",
                    method="PAUT",
                    pipe="300A",
                    item_name="Pipe",
                    material="PAUT 탐촉자",
                    applied_code="지역난방시방서",
                    report_no="RPT-2026-001",
                    joints=2.0,
                    qty=2.0012,
                    unit="M",
                    unit_price=325451.0,
                    travel_cost=50000.0,
                    total_fee=701292.0,
                    equip_main="PAUT MX2(OMNI2-103085)",
                    equip_aux="PAUT SCANNER (MANUAL)",
                    work_shape="주간",
                    workers=[
                        WorkerItem(user="부장 주진철", shift="주간", time="09:00~18:00", ot="0")
                    ],
                    vehicle_no="81두1580 (탑차)",
                    mileage="45",
                    vehicle_note="양호",
                    note="현장 테스트",
                    rtk_total=0.0,
                    chem_fl_mag=0.0
                )
                res_save = save_record(req_save)
                self.assertEqual(res_save["status"], "success")

                # 2. Get recent records
                res_rec = get_recent_records(date="2026-11-10")
                records = res_rec["records"]
                self.assertEqual(len(records), 1)
                entry_time = records[0]["entry_time"]
                self.assertTrue(bool(entry_time))
                self.assertEqual(records[0]["hot_tap_stage"], "1차")
                self.assertEqual(records[0]["pipe"], "300A")

                # 3. Update record
                req_upd = RecordUpdateRequest(
                    date="2026-11-10",
                    site="중앙지사",
                    company="한국지역난방공사 중앙지사",
                    loc_type="핫태핑",
                    hot_tap_stage="1차",
                    hot_tap_joints="2",
                    hot_tap_pipe="300A",
                    method="PAUT",
                    pipe="300A",
                    item_name="Pipe",
                    entry_time=entry_time,
                    id=0,
                    qty=3.5,
                    note="수정 완료된 메모"
                )
                res_upd = update_record(req_upd)
                self.assertEqual(res_upd["status"], "success")

                # Verify updated in DB
                res_rec2 = get_recent_records(date="2026-11-10")
                rec_upd = res_rec2["records"][0]
                self.assertEqual(rec_upd["qty"], 3.5)
                self.assertEqual(rec_upd["note"], "수정 완료된 메모")

                # 4. Delete record
                req_del = RecordDeleteRequest(entry_time=entry_time, id=0)
                res_del = delete_record(req_del)
                self.assertEqual(res_del["status"], "success")

                # Verify deleted in DB
                res_rec3 = get_recent_records(date="2026-11-10")
                self.assertEqual(len(res_rec3["records"]), 0)

    def test_mobile_separates_client_chemical_company_and_daily_memo_and_preserves_on_update(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            test_db = Path(tmp_dir) / "Material_Inventory.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DailyUsage"
            headers = [
                'Date', 'Site', '업체명', '발주처', '회사코드', '검사방법', '수량', '검사량', 'Usage',
                'Note', 'Remark', 'DailyMemo', 'EntryTime', '입력시간', 'ClientRequestID'
            ]
            ws.append(headers)
            wb.save(test_db)

            with patch("home.src.site_apps.central.mobile.server.get_inventory_paths", return_value=[str(test_db)]):
                res = save_record(RecordSaveRequest(
                    date="2026-11-27",
                    site="중앙지사",
                    company="다림건설(주)",
                    client="한국지역난방공사 중앙지사",
                    method="PAUT",
                    qty=4.2,
                    note="개별 작업 메모",
                ))
                self.assertEqual(res["status"], "success")

                # Simulate PC adding a shared DailyMemo and chemical manufacturer on that row
                wb_edit = openpyxl.load_workbook(test_db)
                ws_edit = wb_edit["DailyUsage"]
                row_headers = [c.value for c in ws_edit[1]]
                self.assertEqual(ws_edit.cell(row=2, column=row_headers.index('발주처') + 1).value, "한국지역난방공사 중앙지사")
                self.assertEqual(ws_edit.cell(row=2, column=row_headers.index('회사코드') + 1).value or "", "")
                self.assertEqual(ws_edit.cell(row=2, column=row_headers.index('DailyMemo') + 1).value or "", "")
                entry_time = ws_edit.cell(row=2, column=row_headers.index('EntryTime') + 1).value
                client_req_id = ws_edit.cell(row=2, column=row_headers.index('ClientRequestID') + 1).value

                ws_edit.cell(row=2, column=row_headers.index('DailyMemo') + 1, value="당일 전체 공지사항")
                ws_edit.cell(row=2, column=row_headers.index('회사코드') + 1, value="NAWOO")
                wb_edit.save(test_db)

                # Mobile edits the row's note/qty -> must NOT clobber DailyMemo or 회사코드
                upd_res = update_record(RecordUpdateRequest(
                    date="2026-11-27",
                    site="중앙지사",
                    company="다림건설(주)",
                    client="한국지역난방공사 중앙지사",
                    method="PAUT",
                    qty=5.0,
                    note="수정된 개별 메모",
                    entry_time=entry_time,
                    client_request_id=client_req_id,
                ))
                self.assertEqual(upd_res["status"], "success")

                wb_after = openpyxl.load_workbook(test_db, read_only=True)
                try:
                    ws_after = wb_after["DailyUsage"]
                    vals = dict(zip(row_headers, list(ws_after.values)[1]))
                    self.assertEqual(vals["Note"], "수정된 개별 메모")
                    self.assertEqual(vals["DailyMemo"], "당일 전체 공지사항")
                    self.assertEqual(vals["회사코드"], "NAWOO")
                    self.assertEqual(vals["발주처"], "한국지역난방공사 중앙지사")
                finally:
                    wb_after.close()

    def test_packaged_gas_excludes_other_site_warehouses_and_hardcoded_dates(self):
        root = Path(__file__).resolve().parents[2]
        code_gs = (root / "outputs" / "central-mobile-gas" / "Code.gs").read_text(encoding="utf-8")
        self.assertEqual(code_gs.count("function include("), 1)
        self.assertNotIn("recDate === '2026-11-19'", code_gs)
        self.assertNotIn("경인사무소-삼일산업", code_gs)
        self.assertNotIn("경인사무소-세종알로이", code_gs)
        self.assertIn("경인사무소-본점", code_gs)


if __name__ == "__main__":
    unittest.main()
