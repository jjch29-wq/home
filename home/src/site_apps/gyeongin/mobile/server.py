"""
경인사무소 현장 모바일 검사 입력 앱 - 백엔드 서버 (FastAPI)
"""
import os
import sys
import json
import math
from pathlib import Path
from typing import List, Dict, Any, Optional
import datetime
import pandas as pd
import openpyxl

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response
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
    "이사 곽재운", "이사 박홍규",
    "부장 주진철", "부장 박광복", "부장 이병국", "부장 이용주", "부장 주동기", "부장 김규년", "부장 곽상돈", "부장 정효철", "부장 노현태",
    "차장 김춘호", "차장 권인경", "차장 박원준",
    "과장 주영광", "과장 유상훈",
    "대리 김상현", "대리 윤석현", "대리 우명광", "대리 박만영",
    "계장 박상태", "계장 이준우", "계장 장승대", "계장 이은수", "계장 손주언", "계장 이경재", "계장 박재욱", "계장 한지원",
    "주임 김진환", "주임 황조현", "주임 김태현"
]

VEHICLES = [
    "84저1259 (스타렉스)",
    "95가0200 (스타렉스)",
    "91주8839 (스타렉스)",
    "90서8616 (스타렉스)",
    "76마3422 (스타렉스 12인승)",
    "71고4405 (스타렉스 12인승)",
    "89보4028 (탑차)",
    "81두1580 (탑차)",
    "81루5100 (스타렉스)",
    "95가0175 (스타렉스)",
    "81도5958 (스타렉스)",
    "90너4889 (탑차)",
    "자가용/대중교통"
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
        for d in drives:
            cloud_p = os.path.join(rf"{d}:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data", "Material_Inventory.xlsx")
            if os.path.exists(cloud_p) and cloud_p not in paths:
                paths.insert(0, cloud_p)
                break
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

def get_gyeongin_materials():
    materials = []
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                materials = data.get("materials", [])
        except Exception:
            pass
    if not materials:
        materials = [
            'Carestream AA400-3⅓*12"', 'Carestream AA400-3⅓*17"', 'Carestream AA400-4½*12"',
            'Carestream AA400-10*12"', 'Carestream M100-3⅓*12"', 'Carestream M100-10*12"',
            'Carestream M100-14*17"', 'Carestream MX125-3⅓*6"', 'Carestream MX125-3⅓*12"',
            'Carestream MX125-4½*12"', 'Carestream MX125-10*12"', 'Carestream MX125-14*17"',
            'Carestream T200-3⅓*12"', 'Carestream T200-3⅓*17"', 'Carestream T200-4½*12"',
            'Carestream T200-10*12"', 'Carestream T200-14*17"'
        ]
    return materials

def get_gyeongin_units():
    units = []
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                units = data.get("daily_units", [])
        except Exception:
            pass
    if not units:
        units = ['매', 'M', 'P,M,I/D', 'EA', 'Point', 'SET', 'CAN', 'KG', 'Inch', 'Dia', 'I/D', 'Meter']
    order = ['매', 'M', 'P,M,I/D', 'EA', 'Point', 'SET', 'CAN', 'KG', 'Inch', 'Dia', 'I/D', 'Meter']
    sorted_units = [u for u in order if u in units] + [u for u in units if u not in order]
    return list(dict.fromkeys(sorted_units))

def get_gyeongin_test_items():
    items = []
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                items = data.get("test_items", [])
        except Exception:
            pass
    defaults = ['Piping', 'Tank', 'Vessel', 'Plate', 'Beam', 'Structure', 'Pipe', '사무실 내근 및 장비점검', '우천 대기', '안전교육', '휴무', '기타']
    merged = [i for i in items if i] + defaults
    return list(dict.fromkeys(merged))

def get_gyeongin_applied_codes():
    codes = []
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                codes = data.get("applied_codes", [])
        except Exception:
            pass
    defaults = ['KS', 'ASME', 'AWS', 'API', 'KEPIC', '기타']
    merged = [c for c in codes if c] + defaults
    return list(dict.fromkeys(merged))

EQUIPMENTS = [
    "PAUT MX2(OMNI2-103085)",
    "PAUT MX2(OMNI2-100055)",
    "PAUT X3(QC-0093084)",
    "Sitescan 150(I000760)",
    "Sitescan 150(I000950)",
    "Sitescan 140(1403264TC)",
    "Masterscan 380(I005587)",
    "Epoch 650(150128312)",
    "Se-75(D10692)",
    "Se-75(D10626)",
    "Ir-192(880D D10700)",
    "방사선발생장치 300EGM2",
    "Yoke AC(MY-2 NW0453)",
    "Yoke AC(MP-A3 MP31646)",
    "Yoke AC(MY-2 NW6999)",
    "Yoke DC(MP-A2D MP3403)",
    "PT 약품",
    "해당없음"
]

METHOD_DEFAULTS = {
    "RT": {
        "equip_main": "Se-75(D10692)",
        "aux_equipments": [
            "ADR, SURVEY, 콜리메타",
            "콜리메타 (120˚)",
            "콜리메타 (360˚)",
            "R/CONTROL (25ft)",
            "SURVEY METER",
            "ALARM DOSIMETER (ADR)",
            "라이트마킹",
            "없음"
        ],
        "default_aux": "ADR, SURVEY, 콜리메타",
        "unit": "매"
    },
    "UT": {
        "equip_main": "Sitescan 150(I000760)",
        "aux_equipments": [
            "UT 사각 PROBE 70˚ (2MHz 10*10)",
            "UT 사각 PROBE 60˚ (2MHz 10*10)",
            "UT 사각 PROBE 45˚ (2MHz 10*10)",
            "UT 사각 PROBE 70˚ (4MHz 10*10)",
            "UT 수직 PROBE (2MHz Φ20)",
            "UT 수직 PROBE (4MHz Φ10)",
            "대비시험편 STB-A1",
            "대비시험편 STB-A2",
            "대비시험편 RB-4",
            "접촉매질(Couplant)",
            "없음"
        ],
        "default_aux": "UT 사각 PROBE 70˚ (2MHz 10*10); 대비시험편 STB-A1",
        "unit": "M"
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
        "equip_main": "Yoke AC(MY-2 NW0453)",
        "aux_equipments": [
            "자분살포기",
            "Black Light (자외선등)",
            "조도계",
            "자계강도계",
            "Pie Gauge",
            "없음"
        ],
        "default_aux": "자분살포기",
        "unit": "M"
    },
    "PAUT": {
        "equip_main": "PAUT MX2(OMNI2-103085)",
        "aux_equipments": [
            "PAUT SCANNER (MANUAL)",
            "MICROBE (스캐너 Body-137331, Encoder-136861)",
            "COBRA (스캐너 139310)",
            "PAUT PROBE 5L64-A2",
            "PAUT PROBE 2.25L64-A2",
            "PAUT PROBE 7.5CCEV35-16-A15",
            "PAUT PROBE 5L32-A31",
            "PAUT WEDGE SA2-N55S-IHC",
            "PAUT WEDGE SA2-OL-IHC",
            "PAUT WEDGE SA15-N60S-IHC",
            "PAUT WEDGE SA10-N55S",
            "대비시험편 STB-A1",
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
    item_name: str = "Piping"        # 검사품명
    material: str = ""               # 자재명 / 필름명
    applied_code: str = "KS"         # 적용코드
    report_no: str = ""              # 성적서번호
    joints: float = 0.0
    ori_joints: float = 0.0
    ori_qty: float = 0.0
    rep_joints: float = 0.0
    rep_qty: float = 0.0
    rej_qty: float = 0.0
    qty: float = 0.0
    unit: str = "매"
    unit_price: float = 0.0
    travel_cost: float = 0.0         # 출장비
    lump_sum: str = ""               # 일식
    total_fee: float = 0.0           # 검사비
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

class RecordUpdateRequest(RecordSaveRequest):
    id: Optional[int] = None
    entry_time: Optional[str] = None

class RecordDeleteRequest(BaseModel):
    id: Optional[int] = None
    entry_time: Optional[str] = None

def build_row_data(req: RecordSaveRequest, preserve_entry_time: Optional[str] = None) -> dict:
    calculated_fee = int(round(req.total_fee)) if req.total_fee > 0 else int(round(req.qty * req.unit_price + req.travel_cost))
    stamp = preserve_entry_time if preserve_entry_time else datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_row_data = {
        "Date": req.date,
        "Site": req.site,
        "작업구분": req.loc_type,
        "구분": req.loc_type,
        "검사구분": req.loc_type,
        "업체명": req.company,
        "회사코드": req.company_code or req.company,
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
        "EntryTime": stamp,
        "입력시간": stamp,
        "검사자": req.workers[0].user if req.workers else "",
        "User": req.workers[0].user if req.workers else "",
        "FilmCount": req.qty if req.method == 'RT' else 0,
        # RTK (8 categories)
        "RTK_센터미스": req.rtk_center_miss,
        "RTK_농도": req.rtk_density,
        "RTK_마킹미스": req.rtk_marking_miss,
        "RTK_필름마크": req.rtk_film_mark,
        "RTK_취급부주의": req.rtk_handling,
        "RTK_고객불만": req.rtk_complaint,
        "RTK_기타": req.rtk_etc,
        "RTK_총계": req.rtk_total,
        "RTK_RTK총계": req.rtk_total,
        # NDT Chemicals (7 categories)
        "NDT_형광자분": req.chem_fl_mag,
        "NDT_흑색자분": req.chem_black_mag,
        "NDT_백색페인트": req.chem_white_paint,
        "NDT_침투제": req.chem_penetrant,
        "NDT_세척제": req.chem_cleaner,
        "NDT_현상제": req.chem_developer,
        "NDT_형광침투제": req.chem_fl_pen,
    }

    # Reset worker columns up to 10
    for idx in range(1, 11):
        new_row_data[f"User{idx}"] = ""
        new_row_data[f"WorkTime{idx}"] = ""
        new_row_data[f"User{idx}_WorkTime"] = ""
        new_row_data[f"User{idx}_Work_Shape"] = ""
        new_row_data[f"User{idx}_OT_Pay"] = 0.0
        new_row_data[f"OT{idx}"] = 0.0

    # Assign workers
    for idx, w in enumerate(req.workers[:10], start=1):
        new_row_data[f"User{idx}"] = w.user
        new_row_data[f"WorkTime{idx}"] = w.time
        new_row_data[f"User{idx}_WorkTime"] = w.time
        new_row_data[f"User{idx}_Work_Shape"] = w.shift
        clean_ot = float(str(w.ot).replace(",", "").strip()) if w.ot else 0.0
        new_row_data[f"User{idx}_OT_Pay"] = clean_ot
        new_row_data[f"OT{idx}"] = clean_ot

    return new_row_data

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
        "loc_types": ["경인사무소", "일반배관", "플랜트", "탱크/기기", "기타"],
        "units": get_gyeongin_units(),
        "materials": get_gyeongin_materials(),
        "test_items": get_gyeongin_test_items(),
        "applied_codes": get_gyeongin_applied_codes()
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

def safe_str(val, default=""):
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    return default if s.lower() in ("nan", "none", "<na>", "") else s

def safe_float(val, default=0.0):
    try:
        if val is None or pd.isna(val):
            return default
        f = float(str(val).replace(",", "").strip())
        return f if math.isfinite(f) else default
    except Exception:
        return default

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
        unit = "P,M,I/D" if method == "MT" else "M"
        unit_price = 18500.0 if method == "PAUT" else 4200.0
        base_len = SIZE_LENGTH.get(pipe, 0.0)
        single_len = quantity_per_joint(method, pipe, SIZE_LENGTH)
        calc_len = single_len if single_len is not None else base_len
        qty = round(calc_len * joints, 4) if calc_len > 0 and joints > 0 else (joints if joints > 0 else 0.0)

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
            workers_summary = ""
            workers_list = []
            for i in range(1, 11):
                u_col = f"User{i}" if i > 1 else "User"
                wt_col = f"WorkTime{i}" if i > 1 else "WorkTime"
                ws_col = f"User{i}_Work_Shape" if i > 1 else "User_Work_Shape"
                ot_col = f"User{i}_OT_Pay" if i > 1 else "User_OT_Pay"
                ot_fallback = f"OT{i}" if i > 1 else "OT"

                u = safe_str(row.get(u_col))
                if u:
                    time_val = safe_str(row.get(wt_col), "09:00~18:00")
                    shift_val = safe_str(row.get(ws_col), "일반")
                    ot_val = safe_float(row.get(ot_col, row.get(ot_fallback, 0)))
                    workers_summary += f"{u} ({ot_val:,.0f}원), "
                    workers_list.append({
                        "user": u,
                        "shift": shift_val,
                        "time": time_val,
                        "ot": str(int(ot_val)) if ot_val == int(ot_val) else f"{ot_val:.1f}"
                    })
            
            # Form chemical summary if present
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

            entry_time_val = safe_str(row.get("EntryTime"), safe_str(row.get("입력시간")))

            records.append({
                "id": int(idx),
                "entry_time": entry_time_val,
                "date": safe_str(row.get("Date_str"), safe_str(row.get("Date"))),
                "site": safe_str(row.get("Site"), "경인사무소"),
                "company": safe_str(row.get("업체명"), "경인사무소"),
                "method": safe_str(row.get("검사방법")),
                "item_name": safe_str(row.get("검사품명"), "Piping"),
                "material": safe_str(row.get("자재명"), safe_str(row.get("MaterialID"), safe_str(row.get("품목명")))),
                "loc_type": safe_str(row.get("작업구분"), safe_str(row.get("구분"), "경인사무소")),
                "pipe": safe_str(row.get("관경(Inch)"), safe_str(row.get("규격"))),
                "qty": safe_float(row.get("수량", row.get("Usage", row.get("검사량", 0)))),
                "unit": safe_str(row.get("Unit"), safe_str(row.get("단위"), "M")),
                "unit_price": safe_float(row.get("단가")),
                "travel_cost": safe_float(row.get("출장비")),
                "lump_sum": safe_str(row.get("일식")),
                "total_fee": safe_float(row.get("검사비")),
                "applied_code": safe_str(row.get("적용코드")),
                "report_no": safe_str(row.get("성적서번호")),
                "equip_main": safe_str(row.get("장비명")),
                "equip_aux": safe_str(row.get("보조장비")),
                "vehicle_no": safe_str(row.get("차량번호")),
                "mileage": safe_str(row.get("주행거리")),
                "note": safe_str(row.get("Note"), safe_str(row.get("DailyMemo"))),
                "joints": safe_float(row.get("조인트수")),
                "ori_joints": safe_float(row.get("조인트수")),
                "rej_qty": safe_float(row.get("불량수")),
                # RTK
                "rtk_center_miss": safe_float(row.get("RTK_센터미스")),
                "rtk_density": safe_float(row.get("RTK_농도")),
                "rtk_marking_miss": safe_float(row.get("RTK_마킹미스")),
                "rtk_film_mark": safe_float(row.get("RTK_필름마크")),
                "rtk_handling": safe_float(row.get("RTK_취급부주의")),
                "rtk_complaint": safe_float(row.get("RTK_고객불만")),
                "rtk_etc": safe_float(row.get("RTK_기타")),
                "rtk_total": safe_float(row.get("RTK_총계", row.get("RTK_RTK총계"))),
                # Chemicals
                "chem_fl_mag": safe_float(row.get("NDT_형광자분")),
                "chem_black_mag": safe_float(row.get("NDT_흑색자분")),
                "chem_white_paint": safe_float(row.get("NDT_백색페인트")),
                "chem_penetrant": safe_float(row.get("NDT_침투제")),
                "chem_cleaner": safe_float(row.get("NDT_세척제")),
                "chem_developer": safe_float(row.get("NDT_현상제")),
                "chem_fl_pen": safe_float(row.get("NDT_형광침투제")),
                "extra": " ".join(chem_summary + rtk_summary),
                "workers_summary": workers_summary.rstrip(", "),
                "workers": workers_list
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

    new_row_data = build_row_data(req)

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

@app.post("/api/update_record")
def update_record(req: RecordUpdateRequest):
    paths = get_inventory_paths()
    if not paths:
        raise HTTPException(status_code=500, detail="Database file not found")

    new_row_data = build_row_data(req, preserve_entry_time=req.entry_time)

    updated_dbs = []
    for p in paths:
        try:
            wb = openpyxl.load_workbook(p)
            if "DailyUsage" not in wb.sheetnames:
                continue
            ws = wb["DailyUsage"]
            headers = [cell.value for cell in ws[1]]

            target_row_idx = None

            # 1. Match by EntryTime if available
            if req.entry_time and ("EntryTime" in headers or "입력시간" in headers):
                time_col = headers.index("EntryTime") + 1 if "EntryTime" in headers else headers.index("입력시간") + 1
                for r in range(2, ws.max_row + 1):
                    val = str(ws.cell(row=r, column=time_col).value or "").strip()
                    if val == req.entry_time.strip():
                        target_row_idx = r
                        break

            # 2. Fallback to row index id (id is 0-based index from DataFrame -> row is id + 2)
            if target_row_idx is None and req.id is not None:
                candidate_row = req.id + 2
                if 2 <= candidate_row <= ws.max_row:
                    target_row_idx = candidate_row

            if target_row_idx is not None:
                for col_idx, h in enumerate(headers, start=1):
                    if h in new_row_data:
                        ws.cell(row=target_row_idx, column=col_idx, value=new_row_data[h])
                wb.save(p)
                updated_dbs.append(p)
            else:
                print(f"Record not found for update in {p}")
        except Exception as ex:
            print(f"Error updating {p}: {ex}")

    if not updated_dbs:
        raise HTTPException(status_code=404, detail="수정할 일보 내역을 찾을 수 없습니다.")

    return {
        "status": "success",
        "updated_databases": updated_dbs,
        "message": f"성공적으로 수정되었습니다. ({len(updated_dbs)}개 저장소 반영 완료)"
    }

@app.post("/api/delete_record")
def delete_record(req: RecordDeleteRequest):
    paths = get_inventory_paths()
    if not paths:
        raise HTTPException(status_code=500, detail="Database file not found")

    deleted_dbs = []
    for p in paths:
        try:
            wb = openpyxl.load_workbook(p)
            if "DailyUsage" not in wb.sheetnames:
                continue
            ws = wb["DailyUsage"]
            headers = [cell.value for cell in ws[1]]

            target_row_idx = None
            if req.entry_time and ("EntryTime" in headers or "입력시간" in headers):
                time_col = headers.index("EntryTime") + 1 if "EntryTime" in headers else headers.index("입력시간") + 1
                for r in range(2, ws.max_row + 1):
                    val = str(ws.cell(row=r, column=time_col).value or "").strip()
                    if val == req.entry_time.strip():
                        target_row_idx = r
                        break

            if target_row_idx is None and req.id is not None:
                candidate_row = req.id + 2
                if 2 <= candidate_row <= ws.max_row:
                    target_row_idx = candidate_row

            if target_row_idx is not None:
                ws.delete_rows(target_row_idx, 1)
                wb.save(p)
                deleted_dbs.append(p)
            else:
                print(f"Record not found for deletion in {p}")
        except Exception as ex:
            print(f"Error deleting from {p}: {ex}")

    if not deleted_dbs:
        raise HTTPException(status_code=404, detail="삭제할 일보 내역을 찾을 수 없습니다.")

    return {
        "status": "success",
        "deleted_databases": deleted_dbs,
        "message": f"성공적으로 삭제되었습니다. ({len(deleted_dbs)}개 저장소 반영 완료)"
    }

@app.get("/favicon.ico", include_in_schema=False)
@app.get("/favicon.png", include_in_schema=False)
async def favicon():
    return Response(status_code=204)

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
