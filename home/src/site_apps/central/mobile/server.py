"""
중앙지사 현장 모바일 검사 입력 앱 - 백엔드 서버 (FastAPI)
"""
import os
import sys
import math
import json
import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import openpyxl

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Add central src to path
CENTRAL_SRC = Path(__file__).resolve().parents[1] / "src"
HOME_SRC = Path(__file__).resolve().parents[3]
if str(CENTRAL_SRC) not in sys.path:
    sys.path.insert(0, str(CENTRAL_SRC))
if str(HOME_SRC) not in sys.path:
    sys.path.insert(0, str(HOME_SRC))

from site_apps.central.src.services.ot_calculator import ot_breakdown
from site_apps.central.src.services.inspection_quantity import quantity_per_joint
from site_apps.central.src.data_sync import get_business_data_dir

app = FastAPI(title="중앙지사 모바일 현장 입력 시스템", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def safe_float(val, default=0.0):
    try:
        if val is None or pd.isna(val):
            return default
        f = float(val)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except (ValueError, TypeError):
        return default

def safe_str(val, default=""):
    if val is None or pd.isna(val) or str(val).strip().lower() in ('nan', 'none', 'null'):
        return default
    return str(val).strip()

# Standard definitions
SIZE_LENGTH = {
    '1100A': 3.511,  '1000A': 3.1919, '900A': 2.8727, '850A': 2.7131,
    '800A':  2.5535, '750A':  2.3939, '700A': 2.2343, '650A': 2.0747,
    '600A':  1.9151, '550A':  1.7555, '500A': 1.5959, '450A': 1.4363,
    '400A':  1.2767, '350A':  1.1172, '300A': 1.0006, '250A': 0.8401,
    '200A':  0.6795, '150A':  0.519,  '125A': 0.4392, '100A': 0.3591,
    '80A':   0.2799, '65A':   0.2397, '50A':  0.1901, '40A':  0.1527,
    '32A':   0.1341, '25A':   0.1068, '20A':  0.0855,
}

USERS = [
    "부장 주진철", "부장 이병국", "부장 박광복", "차장 권인경",
    "과장 주영광", "대리 우명광", "계장 박상태", "계장 이준우",
    "계장 장승대", "주임 김성렬", "주임 김진환"
]

VEHICLES = [
    "81두1580 (탑차)", "89보4028 (탑차)", "90너4889 (탑차)",
    "71고4405 (스타렉스)", "76마3422 (스타렉스)", "81도5958 (스타렉스)",
    "81도8458 (스타렉스)", "81루5100 (스타렉스)", "84저1259 (스타렉스)",
    "90서8616 (스타렉스)", "91주8839 (스타렉스)", "95가0175 (스타렉스)",
    "95가0200 (스타렉스)", "개인차량 (스타렉스)", "개인차량-이병국 (스타렉스)"
]

CONFIG_PATH = Path(__file__).resolve().parents[1] / "src" / "Material_Manager_Config.json"

def get_central_aux_equipments():
    items = []
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                items = data.get("aux_equipments", [])
        except Exception:
            pass
    if not items:
        items = [
            "PAUT SCANNER (MANUAL)", "COBRA Manual Scanner", "JIREH MICROBE BTG006",
            "JIREH Scanner", "ADR, SURVEY, 콜리메타", "ADR; SURVEY; 콜리메타",
            "ADR", "SURVEY", "콜리메타", "자분살포기", "조도계", "자계강도계", "온습도계", "휴대용 UV메타", "없음"
        ]
    return list(dict.fromkeys(items))

EQUIPMENTS = ["PAUT MX2(OMNI2-103085)", "Se-75", "Yoke", "PT 약품"]

METHOD_DEFAULTS = {
    "PAUT": {
        "equip_main": "PAUT MX2(OMNI2-103085)",
        "aux_equipments": [
            "PAUT SCANNER (MANUAL)",
            "COBRA Manual Scanner",
            "JIREH MICROBE BTG006",
            "JIREH Scanner",
            "없음"
        ],
        "default_aux": "PAUT SCANNER (MANUAL)",
        "unit": "M"
    },
    "RT": {
        "equip_main": "Se-75",
        "aux_equipments": [
            "ADR, SURVEY, 콜리메타",
            "ADR; SURVEY; 콜리메타",
            "ADR",
            "SURVEY",
            "콜리메타",
            "없음"
        ],
        "default_aux": "ADR, SURVEY, 콜리메타",
        "unit": "매"
    },
    "MT": {
        "equip_main": "Yoke",
        "aux_equipments": [
            "자분살포기",
            "조도계",
            "자계강도계",
            "휴대용 UV메타",
            "없음"
        ],
        "default_aux": "자분살포기",
        "unit": "P,M,I/D"
    },
    "PT": {
        "equip_main": "PT 약품",
        "aux_equipments": [
            "조도계",
            "온습도계",
            "없음"
        ],
        "default_aux": "조도계",
        "unit": "M"
    }
}

WORKTIMES = [
    "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00",
    "09:00~22:00", "09:00~23:00", "09:00~24:00", "09:00~익일01:00",
    "09:00~익일02:00", "09:00~익일03:00", "18:00~익일02:00",
    "18:00~익일03:00", "22:00~익일09:00", "12:00~익일08:00"
]

STATIC_DIR = Path(__file__).resolve().parent / "static"

def get_inventory_paths():
    """Return cloud and local inventory paths"""
    paths = []
    local_p = Path(__file__).resolve().parents[1] / "data" / "Material_Inventory.xlsx"
    if local_p.exists():
        paths.append(str(local_p))

    try:
        cloud_dir = get_business_data_dir()
        cloud_p = os.path.join(cloud_dir, "Material_Inventory.xlsx")
        if os.path.exists(cloud_p) and cloud_p not in paths:
            paths.insert(0, cloud_p)
    except Exception:
        pass

    try:
        import ctypes
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        drives = [chr(65 + i) for i in range(26) if (bitmask >> i) & 1]
        for d in ['G', 'F']:
            if d in drives:
                alt_p = os.path.join(f"{d}:\\내 드라이브\\07_Antigravity\\PMI_Data\\central\\data", "Material_Inventory.xlsx")
                if os.path.exists(alt_p) and alt_p not in paths:
                    paths.insert(0, alt_p)
    except Exception:
        pass

    return paths

class OtCalcRequest(BaseModel):
    worktime: str
    date: str

class QtyCalcRequest(BaseModel):
    method: str
    pipe: str
    joints: float

class WorkerItem(BaseModel):
    user: str
    shift: str = "주간"
    time: str = ""
    ot: str = ""

class RecordSaveRequest(BaseModel):
    date: str
    site: str = "중앙지사"
    company: str = "한국지역난방공사 중앙지사"
    company_code: str = "한국지역난방공사"
    loc_type: str = "열수송관"
    hot_tap_stage: str = ""
    hot_tap_joints: str = ""
    hot_tap_pipe: str = ""
    method: str = "PAUT"
    pipe: str = ""
    item_name: str = "Pipe"
    material: str = ""
    applied_code: str = "지역난방시방서"
    report_no: str = ""
    joints: float = 0.0
    ori_joints: float = 0.0
    rep_joints: float = 0.0
    rej_qty: float = 0.0
    qty: float = 0.0
    unit: str = "M"
    unit_price: float = 0.0
    travel_cost: float = 0.0
    lump_sum: str = ""
    total_fee: float = 0.0
    equip_main: str = ""
    equip_aux: str = ""
    work_shape: str = "주간"
    workers: List[WorkerItem] = []
    vehicle_no: str = ""
    mileage: str = ""
    vehicle_note: str = ""
    note: str = ""
    # RTK Defect Categories (8)
    rtk_center_miss: float = 0.0
    rtk_density: float = 0.0
    rtk_marking_miss: float = 0.0
    rtk_film_mark: float = 0.0
    rtk_handling: float = 0.0
    rtk_complaint: float = 0.0
    rtk_etc: float = 0.0
    rtk_total: float = 0.0
    # NDT Chemicals (7)
    chem_fl_mag: float = 0.0
    chem_black_mag: float = 0.0
    chem_white_paint: float = 0.0
    chem_penetrant: float = 0.0
    chem_cleaner: float = 0.0
    chem_developer: float = 0.0
    chem_fl_pen: float = 0.0

@app.get("/api/config")
def get_config():
    return {
        "users": USERS,
        "vehicles": VEHICLES,
        "equipments": EQUIPMENTS,
        "aux_equipments": get_central_aux_equipments(),
        "method_defaults": METHOD_DEFAULTS,
        "worktimes": WORKTIMES,
        "pipe_sizes": list(SIZE_LENGTH.keys()),
        "methods": ["PAUT", "RT", "MT", "PT"],
        "loc_types": ["열수송관", "플랜트(관리소)", "열배관", "핫태핑"],
        "units": ["M", "매", "EA", "P,M,I/D", "Point", "SET", "CAN", "KG", "Inch", "Dia", "I/D"],
        "test_items": ["Pipe", "Piping", "Tank", "Vessel", "Plate", "Beam", "Structure", "기타"],
        "applied_codes": ["지역난방시방서", "KS", "ASME", "AWS", "API", "KEPIC", "기타"],
        "materials": [
            'Carestream AA400-3⅓*12"', 'Carestream AA400-3⅓*17"', 'Carestream AA400-4½*12"',
            'Carestream AA400-10*12"', 'Carestream M100-3⅓*12"', 'Carestream M100-10*12"',
            'Carestream M100-14*17"', 'Carestream MX125-3⅓*6"', 'Carestream MX125-3⅓*12"',
            'Carestream MX125-4½*12"', 'Carestream MX125-10*12"', 'Carestream MX125-14*17"',
            'Carestream T200-3⅓*12"', 'Carestream T200-3⅓*17"', 'Carestream T200-4½*12"',
            'Carestream T200-10*12"', 'Carestream T200-14*17"'
        ]
    }

@app.post("/api/calculate_ot")
def api_calculate_ot(req: OtCalcRequest):
    try:
        res = ot_breakdown(req.worktime, req.date)
        return {
            "amount": res["amount"],
            "amount_str": f"{res['amount']:,}",
            "hours": res["hours"],
            "day": res["day"],
            "night": res["night"],
            "holiday": res["holiday"]
        }
    except Exception as e:
        return {"amount": 0, "amount_str": "0", "hours": 0, "error": str(e)}

@app.post("/api/calculate_qty")
def api_calculate_qty(req: QtyCalcRequest):
    pipe = req.pipe.strip().upper()
    if pipe.isdigit():
        pipe += "A"

    length_per_j = quantity_per_joint(req.method, pipe, SIZE_LENGTH)
    if length_per_j is None:
        unit = "매" if req.method == "RT" else ("P,M,I/D" if req.method == "MT" else "M")
        return {"qty": 0.0, "unit": unit, "unit_price": 0}

    total_qty = round(length_per_j * req.joints, 4) if req.method != "RT" else round(req.joints * 3.0, 1)
    unit_price = 325451 if req.method == "PAUT" else (205412 if req.method == "RT" else 110000)
    unit = "매" if req.method == "RT" else ("P,M,I/D" if req.method == "MT" else "M")

    return {
        "qty": total_qty,
        "unit": unit,
        "unit_price": unit_price,
        "total_cost": int(round(total_qty * unit_price))
    }

@app.get("/api/recent_records")
def get_recent_records(date: Optional[str] = None):
    paths = get_inventory_paths()
    if not paths:
        return {"records": []}

    try:
        df = pd.read_excel(paths[0], sheet_name="DailyUsage")
        if df.empty:
            return {"records": []}

        df["Date_str"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

        if date:
            filtered = df[df["Date_str"] == date]
            if filtered.empty:
                filtered = df.tail(15)
        else:
            filtered = df.tail(15)

        records = []
        for idx, row in filtered.iterrows():
            workers_str = ""
            for i in range(1, 11):
                u_col = f"User{i}" if i > 1 else "User"
                ot_col = f"User{i}_OT_Pay" if i > 1 else "User_OT_Pay"
                ot_fallback = f"OT{i}" if i > 1 else "OT"
                u = safe_str(row.get(u_col))
                if u:
                    ot_val = safe_float(row.get(ot_col, row.get(ot_fallback, 0)))
                    workers_str += f"{u} ({ot_val:,.0f}원), "

            chem_summary = []
            for col, short in [('NDT_형광자분', '형자'), ('NDT_흑색자분', '흑자'), ('NDT_백색페인트', '백페'),
                               ('NDT_침투제', '침투'), ('NDT_세척제', '세척'), ('NDT_현상제', '현상'), ('NDT_형광침투제', '형침')]:
                val = safe_float(row.get(col, 0))
                if val > 0:
                    chem_summary.append(f"{short}:{val:g}")

            rtk_summary = []
            rtk_total = safe_float(row.get("RTK_총계", row.get("RTK_RTK총계", 0)))
            if rtk_total > 0:
                rtk_summary.append(f"RTK:{rtk_total:g}")

            records.append({
                "id": int(idx),
                "date": safe_str(row.get("Date_str"), safe_str(row.get("Date"))),
                "company": safe_str(row.get("업체명"), "한국지역난방공사 중앙지사"),
                "method": safe_str(row.get("검사방법")),
                "item_name": safe_str(row.get("검사품명"), "Pipe"),
                "material": safe_str(row.get("자재명"), safe_str(row.get("MaterialID"), safe_str(row.get("품목명")))),
                "loc_type": safe_str(row.get("작업구분"), safe_str(row.get("구분"), "열수송관")),
                "pipe": safe_str(row.get("관경(Inch)"), safe_str(row.get("규격"))),
                "qty": safe_float(row.get("수량", row.get("Usage", row.get("검사량", 0)))),
                "unit": safe_str(row.get("Unit"), safe_str(row.get("단위"), "M")),
                "unit_price": safe_float(row.get("단가")),
                "travel_cost": safe_float(row.get("출장비")),
                "total_fee": safe_float(row.get("검사비")),
                "applied_code": safe_str(row.get("적용코드")),
                "report_no": safe_str(row.get("성적서번호")),
                "extra": " ".join(chem_summary + rtk_summary),
                "workers": workers_str.rstrip(", ")
            })

        records.reverse()
        return {"records": records}
    except Exception as e:
        return {"records": [], "error": str(e)}

@app.post("/api/save_record")
def save_record(req: RecordSaveRequest):
    paths = get_inventory_paths()
    if not paths:
        raise HTTPException(status_code=500, detail="Database file not found")

    calculated_fee = int(round(req.total_fee)) if req.total_fee > 0 else int(round(req.qty * req.unit_price + req.travel_cost))
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    row_dict = {
        "Date": req.date,
        "Site": req.site,
        "작업구분": req.loc_type,
        "구분": req.loc_type,
        "검사구분": req.loc_type,
        "업체명": req.company,
        "회사코드": req.company_code or "한국지역난방공사",
        "검사방법": req.method,
        "규격": req.pipe,
        "관경(Inch)": req.pipe,
        "검사품명": req.item_name,
        "자재명": req.material,
        "MaterialID": req.material,
        "Material": req.material,
        "품목명": req.material,
        "적용코드": req.applied_code,
        "성적서번호": req.report_no,
        "조인트수": req.joints or req.ori_joints,
        "불량수": req.rej_qty,
        "수량": req.qty,
        "검사량": req.qty,
        "Usage": req.qty,
        "단위": req.unit,
        "Unit": req.unit,
        "단가": req.unit_price,
        "출장비": req.travel_cost,
        "일식": req.lump_sum,
        "검사비": calculated_fee,
        "장비명": req.equip_main,
        "보조장비": req.equip_aux,
        "작업형태": req.work_shape,
        "차량번호": req.vehicle_no,
        "주행거리": req.mileage,
        "차량점검": "양호",
        "차량비고": req.vehicle_note,
        "Note": req.note,
        "Remark": req.note,
        "DailyMemo": req.note,
        "EntryTime": now_str,
        "입력시간": now_str,
        "검사자": req.workers[0].user if req.workers else "",
        "작업자": req.workers[0].user if req.workers else "",
        "User": req.workers[0].user if req.workers else "",
        "FilmCount": req.qty if req.method == 'RT' else 0,
        "제경비율": 1.10,
        "기술료율": 0.20,
        "핫태핑차수": req.hot_tap_stage,
        "핫태핑개소": req.hot_tap_joints,
        "핫태핑관경": req.hot_tap_pipe,
        # RTK (8)
        "RTK_센터미스": req.rtk_center_miss,
        "RTK_농도": req.rtk_density,
        "RTK_마킹미스": req.rtk_marking_miss,
        "RTK_필름마크": req.rtk_film_mark,
        "RTK_취급부주의": req.rtk_handling,
        "RTK_고객불만": req.rtk_complaint,
        "RTK_기타": req.rtk_etc,
        "RTK_총계": req.rtk_total,
        "RTK_RTK총계": req.rtk_total,
        # NDT Chemicals (7)
        "NDT_형광자분": req.chem_fl_mag,
        "NDT_흑색자분": req.chem_black_mag,
        "NDT_백색페인트": req.chem_white_paint,
        "NDT_침투제": req.chem_penetrant,
        "NDT_세척제": req.chem_cleaner,
        "NDT_현상제": req.chem_developer,
        "NDT_형광침투제": req.chem_fl_pen,
    }

    # Assign workers
    for idx, w in enumerate(req.workers[:10], start=1):
        row_dict[f"User{idx}" if idx > 1 else "User"] = w.user
        time_val = f"({w.shift}) {w.time}" if w.time else ""
        row_dict[f"WorkTime{idx}" if idx > 1 else "WorkTime"] = time_val
        clean_ot = safe_float(str(w.ot).replace(",", "").strip())
        row_dict[f"OT{idx}" if idx > 1 else "OT"] = clean_ot
        row_dict[f"User{idx}_OT_Pay" if idx > 1 else "User_OT_Pay"] = clean_ot

    saved_dbs = []
    for p in paths:
        try:
            wb = openpyxl.load_workbook(p)
            if "DailyUsage" not in wb.sheetnames:
                continue
            ws = wb["DailyUsage"]
            headers = [cell.value for cell in ws[1]]

            row_values = []
            for h in headers:
                val = row_dict.get(h, None)
                row_values.append(val)

            ws.append(row_values)
            wb.save(p)
            saved_dbs.append(p)
        except Exception as ex:
            print(f"Error saving to {p}: {ex}")

    if not saved_dbs:
        raise HTTPException(status_code=500, detail="Failed to save record to any database")

    return {
        "status": "success",
        "saved_databases": saved_dbs,
        "message": f"성공적으로 저장되었습니다. ({len(saved_dbs)}개 저장소 동기화 완료)"
    }

@app.get("/favicon.ico", include_in_schema=False)
@app.get("/favicon.png", include_in_schema=False)
async def favicon():
    return Response(status_code=204)

@app.get("/", response_class=HTMLResponse)
def index():
    html_file = STATIC_DIR / "index.html"
    if not html_file.exists():
        # Fallback to Google Apps Script index.html if local static is absent
        gas_index = Path(__file__).resolve().parents[1] / "google_apps_script" / "index.html"
        if gas_index.exists():
            return gas_index.read_text(encoding="utf-8")
        return "<h1>중앙지사 모바일 앱 준비 중</h1>"
    return html_file.read_text(encoding="utf-8")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
