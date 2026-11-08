import sys, os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
os.environ["PMI_SITE"] = "gyeongin"

from site_apps.gyeongin.mobile.server import get_config, calculate_qty, calculate_ot, QtyCalcRequest, OtCalcRequest, RecordSaveRequest, save_record

def test_mobile_functions():
    print("[1] Test get_config:")
    cfg = get_config()
    print(f"  Methods: {cfg['methods']}")
    print(f"  Companies: {cfg['companies']}")
    print(f"  Loc types: {cfg['loc_types']}")
    print(f"  Units: {cfg['units']}")
    print(f"  Test items: {cfg['test_items']}")
    print(f"  Applied codes: {cfg['applied_codes']}")
    assert "RT" in cfg["methods"]
    assert "PT" in cfg["methods"]
    assert "MT" in cfg["methods"]
    assert "열수송관" not in cfg["loc_types"]
    assert "핫태핑" not in cfg["loc_types"]
    assert "경인사무소" in cfg["companies"]
    assert "매" in cfg["units"]
    assert "P,M,I/D" in cfg["units"]
    assert "Piping" in cfg["test_items"]
    assert "Tank" in cfg["test_items"]
    assert "KS" in cfg["applied_codes"]
    assert len(cfg["materials"]) > 0
    print("  => OK")

    print("[2] Test calculate_qty:")
    rt_res = calculate_qty(QtyCalcRequest(method="RT", pipe="250A", joints=2.0))
    assert rt_res["unit"] == "매"
    assert rt_res["qty"] == 6.0
    print(f"  RT: {rt_res}")

    pt_res = calculate_qty(QtyCalcRequest(method="PT", pipe="250A", joints=2.0))
    assert pt_res["unit"] == "M"
    assert pt_res["qty"] > 0
    print(f"  PT: {pt_res}")
    print("  => OK")

    print("[3] Test calculate_ot:")
    ot_res = calculate_ot(OtCalcRequest(worktime="09:00~22:00", date="2026-11-08"))
    assert ot_res["ot_pay"] > 0
    print(f"  OT: {ot_res}")
    print("  => OK")

    print("[4] Test index.html content:")
    html_p = Path(__file__).resolve().parents[1] / "src" / "site_apps" / "gyeongin" / "mobile" / "static" / "index.html"
    assert html_p.exists()
    content = html_p.read_text(encoding="utf-8")
    assert "경인사무소 현장 입력" in content
    assert "RTK 결함 분류" in content
    assert "NDT 자재 소모량" in content
    assert "업체명" in content
    assert "검사품명" in content
    assert "자재명" in content
    assert "단위" in content
    assert "적용코드" in content
    assert "출장비" in content
    print("  => OK")

    print("[5] Test RecordSaveRequest creation for non-pipe item:")
    req = RecordSaveRequest(
        date="2026-11-08",
        site="경인사무소",
        company="경인사무소",
        item_name="Tank",
        material="Carestream AA400-3⅓*12\"",
        applied_code="ASME",
        report_no="RPT-2026-001",
        qty=15.0,
        unit="P,M,I/D",
        unit_price=4200.0,
        travel_cost=50000.0,
        total_fee=113000.0
    )
    assert req.item_name == "Tank"
    assert req.unit == "P,M,I/D"
    assert req.travel_cost == 50000.0
    assert req.total_fee == 113000.0
    print(f"  Request validated: item={req.item_name}, unit={req.unit}, fee={req.total_fee}")
    print("  => OK")

    print("\nALL MOBILE TESTS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    test_mobile_functions()
