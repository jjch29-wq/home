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

from site_apps.central.src.services.ot_calculator import ot_breakdown
from site_apps.central.src.services.inspection_quantity import quantity_per_joint

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

EQUIPMENTS = ["PAUT MX2(OMNI2-103085)", "Se-75", "Yoke"]

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
        "default_aux": "ADR, SURVEY, 콜리메타"
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
        "default_aux": "PAUT SCANNER (MANUAL)"
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
        "default_aux": "자분살포기"
    },
    "PT": {
        "equip_main": "PT 약품",
        "aux_equipments": [
            "조도계",
            "온습도계",
            "없음"
        ],
        "default_aux": "조도계"
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
    """Return cloud and local inventory paths for Gyeongin Office"""
    cloud_p = os.path.join(r"G:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data", "Material_Inventory.xlsx")
    local_p = Path(__file__).resolve().parents[1] / "data" / "Material_Inventory.xlsx"
    return [p for p in [cloud_p, str(local_p)] if os.path.exists(p)]

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
    loc_type: str = "열수송관"
    method: str
    pipe: str = ""
    item_name: str = "Pipe"
    joints: float = 0.0
    qty: float = 0.0
    unit: str = "M"
    unit_price: float = 0.0
    equip_main: str = ""
    equip_aux: str = ""
    work_shape: str = "주간"
    workers: List[WorkerItem]
    vehicle_no: str = ""
    mileage: str = ""
    vehicle_note: str = ""
    note: str = ""

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
        "methods": ["PAUT", "RT", "MT", "PT"],
        "loc_types": ["열수송관", "플랜트(관리소)", "열배관", "핫태핑"]
    }

@app.post("/api/calculate_ot")
def calculate_ot(req: OtCalcRequest):
    try:
        dt = datetime.datetime.strptime(req.date, "%Y-%m-%d").date()
    except Exception:
        dt = datetime.date.today()
    
    total_ot_hours, ot_pay, desc = ot_breakdown(req.worktime, dt)
    return {
        "total_hours": total_ot_hours,
        "ot_pay": ot_pay,
        "desc": desc
    }

@app.post("/api/calculate_qty")
def calculate_qty(req: QtyCalcRequest):
    method = req.method.strip().upper()
    pipe = req.pipe.strip()
    joints = float(req.joints or 0)

    if method == "RT":
        qty = round(joints * 3.0, 4) if joints > 0 else 0.0
        unit = "매"
        unit_price = 205412.0
    else:
        unit = "M"
        unit_price = 325451.0 if method == "PAUT" else 110000.0
        base_len = SIZE_LENGTH.get(pipe, 0.0)
        qty = round(quantity_per_joint(base_len) * joints, 4) if base_len > 0 and joints > 0 else 0.0

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
            
            records.append({
                "id": int(idx),
                "date": str(row.get("Date_str", "")),
                "method": str(row.get("검사방법", "")),
                "loc_type": str(row.get("작업구분", "열수송관")),
                "pipe": str(row.get("규격", "")),
                "qty": float(row.get("수량", 0)),
                "unit": str(row.get("단위", "M")),
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
        "업체명": req.company,
        "검사방법": req.method,
        "규격": req.pipe,
        "검사품명": req.item_name,
        "조인트수": req.joints,
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
        "EntryTime": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # Assign workers
    for idx, w in enumerate(req.workers[:5], start=1):
        new_row_data[f"User{idx}"] = w.user
        new_row_data[f"User{idx}_WorkTime"] = w.time
        new_row_data[f"User{idx}_Work_Shape"] = w.shift
        clean_ot = float(str(w.ot).replace(",", "").strip()) if w.ot else 0.0
        new_row_data[f"User{idx}_OT_Pay"] = clean_ot

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

