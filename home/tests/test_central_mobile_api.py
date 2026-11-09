import unittest
import tempfile
import openpyxl
from pathlib import Path
from unittest.mock import patch

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

if __name__ == "__main__":
    unittest.main()
