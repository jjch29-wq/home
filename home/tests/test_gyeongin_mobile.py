import sys, os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
os.environ["PMI_SITE"] = "gyeongin"

from site_apps.gyeongin.mobile.server import get_config, calculate_qty, calculate_ot, QtyCalcRequest, OtCalcRequest

def test_mobile_functions():
    print("[1] Test get_config:")
    cfg = get_config()
    print(f"  Methods: {cfg['methods']}")
    print(f"  Companies: {cfg['companies']}")
    print(f"  Loc types: {cfg['loc_types']}")
    assert "RT" in cfg["methods"]
    assert "PT" in cfg["methods"]
    assert "MT" in cfg["methods"]
    assert "열수송관" not in cfg["loc_types"]
    assert "핫태핑" not in cfg["loc_types"]
    assert "경인사무소" in cfg["companies"]
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
    print("  => OK")

    print("\nALL MOBILE TESTS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    test_mobile_functions()
