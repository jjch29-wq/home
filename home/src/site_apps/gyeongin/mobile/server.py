"""
경인사무소 현장 모바일 검사 입력 앱 - 백엔드 서버 (FastAPI)
"""
import os
import sys
import json
from pathlib import Path
from typing import List, Dict, Any, Optional
import datetime
import pandas as pd
import openpyxl

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Add paths
GYEONGIN_SRC = Path(__file__).resolve().parents[1] / "src"
HOME_SRC = Path(__file__).resolve().parents[3]
if str(GYEONGIN_SRC) not in sys.path:
    sys.path.insert(0, str(GYEONGIN_SRC))
if str(HOME_SRC) not in sys.path:
    sys.path.insert(0, str(HOME_SRC))

# Site Isolation: Only use Gyeongin packages
from site_apps.gyeongin.src.services.ot_calculator import ot_breakdown
from site_apps.gyeongin.src.services.inspection_quantity import quantity_per_joint

app = FastAPI(title="경인사무소 모바일 현장 입력 시스템", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

def get_inventory_paths():
    """Return cloud and local inventory paths for Gyeongin Office"""
    paths = []
    local_p = Path(__file__).resolve().parents[1] / "data" / "Material_Inventory.xlsx"
    if local_p.exists():
        paths.append(str(local_p))

    try:
        # Fast check without hanging on unmounted network/virtual drives
        import ctypes
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        drives = [chr(65 + i) for i in range(26) if (bitmask >> i) & 1]
        if 'G' in drives:
            cloud_p = os.path.join(r"G:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data", "Material_Inventory.xlsx")
            if os.path.exists(cloud_p) and cloud_p not in paths:
                paths.insert(0, cloud_p)
    except Exception:
        pass

    return paths

def get_gyeongin_companies():
    """Load dynamic client company list from config and database"""
    companies = []
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                companies = data.get("companies", [])
        except Exception:
            pass
    paths = get_inventory_paths()
    if paths:
        try:
            df = pd.read_excel(paths[0], sheet_name="DailyUsage")
            if not df.empty and "업체명" in df.columns:
                existing = [str(x).strip() for x in df["업체명"].dropna() if str(x).strip() and str(x).strip() != 'nan']
                companies.extend(existing)
        except Exception:
            pass
    if not companies:
        companies = ["경인사무소"]
    return list(dict.fromkeys(companies))

def get_gyeongin_aux_equipments():
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
            "ADR, SURVEY, 콜리메타", "ADR; SURVEY; 콜리메타", "ADR", "SURVEY", "콜리메타",
            "PAUT SCANNER (MANUAL)", "COBRA Manual Scanner", "JIREH MICROBE BTG006",
            "JIREH Scanner", "자분살포기", "조도계", "자계강도계", "온습도계", "휴대용 UV메타", "없음"
        ]
    return list(dict.fromkeys(items))

EQUIPMENTS = ["Se-75", "PAUT MX2(OMNI2-103085)", "Yoke", "PT 약품"]

METHOD_DEFAULTS = {
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
    "PT": {
        "equip_main": "PT 약품",
        "aux_equipments": [
            "조도계",
            "온습도계",
            "없음"
        ],
        "default_aux": "조도계",
        "unit": "M"
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
        "unit": "M"
    },
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
    }
}

WORKTIMES = [
    "09:00~18:00", "09:00~19:00", "09:00~20:00", "09:00~21:00",
    "09:00~22:00", "09:00~23:00", "09:00~24:00", "09:00~익일01:00",
    "09:00~익일02:00", "09:00~익일03:00", "18:00~익일02:00",
    "18:00~익일03:00", "22:00~익일09:00", "12:00~익일08:00"
]

STATIC_DIR = Path(__file__).resolve().parent / "static"

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
    site: str = "경인사무소"
    company: str = "경인사무소"
    company_code: str = ""
    loc_type: str = "경인사무소"
    method: str = "RT"
    pipe: str = ""
    item_name: str = "Piping"
    joints: float = 0.0
    ori_joints: float = 0.0
    ori_qty: float = 0.0
    rep_joints: float = 0.0
    rep_qty: float = 0.0
    rej_qty: float = 0.0
    qty: float = 0.0
    unit: str = "매"
    unit_price: float = 0.0
    equip_main: str = ""
    equip_aux: str = ""
    work_shape: str = "일반"
    workers: List[WorkerItem] = []
    vehicle_no: str = ""
    mileage: str = ""
    vehicle_note: str = ""
    note: str = ""
    # RTK Defect Categories (8 items)
    rtk_center_miss: float = 0.0
    rtk_density: float = 0.0
    rtk_marking_miss: float = 0.0
    rtk_film_mark: float = 0.0
    rtk_handling: float = 0.0
    rtk_complaint: float = 0.0
    rtk_etc: float = 0.0
    rtk_total: float = 0.0
    # NDT Chemical Consumables (7 items)
    chem_fl_mag: float = 0.0      # 형광자분
    chem_black_mag: float = 0.0   # 흑색자분
    chem_white_paint: float = 0.0 # 백색페인트
    chem_penetrant: float = 0.0   # 침투제
    chem_cleaner: float = 0.0     # 세척제
    chem_developer: float = 0.0   # 현상제
    chem_fl_pen: float = 0.0      # 형광침투제

@app.get("/api/config")
def get_config():
    return {
        "users": USERS,
        "vehicles": VEHICLES,
        "equipments": EQUIPMENTS,
        "aux_equipments": get_gyeongin_aux_equipments(),
        "method_defaults": METHOD_DEFAULTS,
        "worktimes": WORKTIMES,
        "pipe_sizes": list(SIZE_LENGTH.keys()),
        "methods": ["RT", "PT", "MT", "PAUT"],
        "companies": get_gyeongin_companies(),
        "loc_types": ["경인사무소", "일반배관", "플랜트", "탱크/기기", "기타"]
    }

@app.post("/api/calculate_ot")
def calculate_ot(req: OtCalcRequest):
    try:
        dt = datetime.datetime.strptime(req.date, "%Y-%m-%d").date()
    except Exception:
        dt = datetime.date.today()
    
    res = ot_breakdown(req.worktime, dt)
    return {
        "total_hours": res.get("hours", 0.0),
        "ot_pay": res.get("amount", 0),
        "desc": f"주간 {res.get('day', 0)}h, 야간 {res.get('night', 0)}h, 심야·휴일 {res.get('holiday', 0)}h"
    }

@app.post("/api/calculate_qty")
def calculate_qty(req: QtyCalcRequest):
    method = req.method.strip().upper()
    pipe = req.pipe.strip()
    joints = float(req.joints or 0)

    if method == "RT":
        unit = "매"
        unit_price = 25500.0
        qty = round(joints * 3.0, 1) if joints > 0 else 0.0
    else:
        unit = "M"
        unit_price = 18500.0 if method == "PAUT" else 4200.0
        base_len = SIZE_LENGTH.get(pipe, 0.0)
        single_len = quantity_per_joint(method, pipe, SIZE_LENGTH)
        calc_len = single_len if single_len is not None else base_len
        qty = round(calc_len * joints, 4) if calc_len > 0 and joints > 0 else 0.0

    total_cost = int(round(qty * unit_price))
    return {
        "qty": qty,
        "unit": unit,
        "unit_price": unit_price,
        "total_cost": total_cost
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
            for i in range(1, 4):
                u = row.get(f"User{i}")
                if pd.notna(u) and str(u).strip():
                    ot_val = row.get(f"User{i}_OT_Pay", 0)
                    workers_str += f"{str(u).strip()} ({ot_val:,.0f}원), "
            
            # Form chemical summary if present
            chem_summary = []
            for col, short in [('NDT_형광자분', '형자'), ('NDT_흑색자분', '흑자'), ('NDT_백색페인트', '백페'), 
                               ('NDT_침투제', '침투'), ('NDT_세척제', '세척'), ('NDT_현상제', '현상'), ('NDT_형광침투제', '형침')]:
                val = row.get(col, 0)
                try:
                    if float(val) > 0: chem_summary.append(f"{short}:{float(val):g}")
                except Exception: pass

            rtk_summary = []
            rtk_total = row.get("RTK_총계", 0)
            try:
                if float(rtk_total) > 0: rtk_summary.append(f"RTK:{float(rtk_total):g}")
            except Exception: pass

            records.append({
                "id": int(idx),
                "date": str(row.get("Date_str", "")),
                "company": str(row.get("업체명", "경인사무소")),
                "method": str(row.get("검사방법", "")),
                "loc_type": str(row.get("작업구분", "경인사무소")),
                "pipe": str(row.get("관경(Inch)", row.get("규격", ""))),
                "qty": float(row.get("수량", 0)),
                "unit": str(row.get("단위", "M")),
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

    new_row_data = {
        "Date": req.date,
        "Site": req.site,
        "작업구분": req.loc_type,
        "구분": req.loc_type,
        "업체명": req.company,
        "회사코드": req.company_code or req.company,
        "검사방법": req.method,
        "규격": req.pipe,
        "관경(Inch)": req.pipe,
        "검사품명": req.item_name,
        "조인트수": req.joints or req.ori_joints,
        "불량수": req.rej_qty,
        "수량": req.qty,
        "단위": req.unit,
        "단가": req.unit_price,
        "검사비": int(round(req.qty * req.unit_price)),
        "장비명": req.equip_main,
        "보조장비": req.equip_aux,
        "작업형태": req.work_shape,
        "차량번호": req.vehicle_no,
        "주행거리": req.mileage,
        "차량비고": req.vehicle_note,
        "Note": req.note,
        "EntryTime": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        # RTK (8 categories)
        "RTK_센터미스": req.rtk_center_miss,
        "RTK_농도": req.rtk_density,
        "RTK_마킹미스": req.rtk_marking_miss,
        "RTK_필름마크": req.rtk_film_mark,
        "RTK_취급부주의": req.rtk_handling,
        "RTK_고객불만": req.rtk_complaint,
        "RTK_기타": req.rtk_etc,
        "RTK_총계": req.rtk_total,
        # NDT Chemicals (7 categories)
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
        new_row_data[f"User{idx}"] = w.user
        new_row_data[f"WorkTime{idx}"] = w.time
        new_row_data[f"User{idx}_WorkTime"] = w.time
        new_row_data[f"User{idx}_Work_Shape"] = w.shift
        clean_ot = float(str(w.ot).replace(",", "").strip()) if w.ot else 0.0
        new_row_data[f"User{idx}_OT_Pay"] = clean_ot
        new_row_data[f"OT{idx}"] = clean_ot

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
                val = new_row_data.get(h, None)
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

@app.get("/", response_class=HTMLResponse)
def index():
    index_file = STATIC_DIR / "index.html"
    if not index_file.exists():
        return HTMLResponse("<h1>경인사무소 모바일 웹 화면 준비 중</h1>")
    return HTMLResponse(index_file.read_text(encoding="utf-8"))

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
