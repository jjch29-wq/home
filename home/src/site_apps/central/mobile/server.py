"""
중앙지사 현장 모바일 검사 입력 앱 - 백엔드 서버 (FastAPI)
"""
import os
import sys
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
            "ADR, SURVEY, 콜리메타", "ADR; SURVEY; 콜리메타", "ADR", "SURVEY", "콜리메타",
            "PAUT SCANNER (MANUAL)", "COBRA Manual Scanner", "JIREH MICROBE BTG006",
            "JIREH Scanner", "자분살포기", "조도계", "자계강도계", "온습도계", "BLACK LIGHT", "없음"
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
            "BLACK LIGHT",
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
    """Return cloud and local inventory paths"""
    cloud_dir = get_business_data_dir()
    cloud_p = os.path.join(cloud_dir, "Material_Inventory.xlsx")
    local_p = Path(__file__).resolve().parents[1] / "data" / "Material_Inventory.xlsx"
    backup_p = Path(__file__).resolve().parents[3] / "data" / "Material_Inventory.xlsx"
    return [p for p in [cloud_p, str(local_p), str(backup_p)] if os.path.exists(p)]

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
        "aux_equipments": get_central_aux_equipments(),
        "method_defaults": METHOD_DEFAULTS,
        "worktimes": WORKTIMES,
        "pipe_sizes": list(SIZE_LENGTH.keys()),
        "methods": ["PAUT", "RT", "MT", "PT"],
        "loc_types": ["열수송관", "플랜트(관리소)", "열배관", "핫태핑"]
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
        return {"qty": 0.0, "unit": "매" if req.method == "RT" else "M", "unit_price": 0}
    
    total_qty = round(length_per_j * req.joints, 4)
    unit_price = 325451 if req.method == "PAUT" else (205412 if req.method == "RT" else 110000)
    
    return {
        "qty": total_qty,
        "unit": "매" if req.method == "RT" else "M",
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
        if date:
            target = pd.to_datetime(date, errors="coerce").date()
            dates = pd.to_datetime(df["Date"], errors="coerce").dt.date
            df = df[dates == target]
        else:
            df = df.tail(15)
            
        records = []
        for idx, r in df.iterrows():
            workers = []
            for i in range(1, 11):
                u = r.get("User" if i == 1 else f"User{i}")
                ot = r.get("OT" if i == 1 else f"OT{i}")
                wt = r.get("WorkTime" if i == 1 else f"WorkTime{i}")
                if pd.notna(u) and str(u).strip() and str(u) != "nan":
                    workers.append(f"{u} ({ot or '0'}원)")
            
            records.append({
                "id": int(idx),
                "date": str(r.get("Date", ""))[:10],
                "method": str(r.get("검사방법", "")),
                "loc_type": str(r.get("구분", "")),
                "pipe": str(r.get("관경(Inch)", "")),
                "qty": r.get("Usage") or r.get("검사량") or 0,
                "unit": str(r.get("Unit", "M")),
                "workers": ", ".join(workers[:3]) + (f" 외 {len(workers)-3}명" if len(workers) > 3 else "")
            })
        return {"records": list(reversed(records))}
    except Exception as e:
        return {"records": [], "error": str(e)}

@app.post("/api/save_record")
def save_record(req: RecordSaveRequest):
    paths = get_inventory_paths()
    if not paths:
        raise HTTPException(status_code=500, detail="Database file not found")
        
    try:
        # Prepare row data dictionary
        row_dict = {
            "Date": pd.to_datetime(req.date),
            "Site": req.site,
            "업체명": req.company,
            "구분": req.loc_type,
            "검사방법": req.method,
            "검사품명": req.item_name,
            "관경(Inch)": req.pipe,
            "조인트수": req.joints,
            "Usage": req.qty,
            "검사량": req.qty,
            "Unit": req.unit,
            "단가": req.unit_price,
            "검사비": int(round(req.qty * req.unit_price)),
            "장비명": req.equip_main,
            "보조장비": req.equip_aux,
            "작업형태": req.work_shape,
            "차량번호": req.vehicle_no,
            "주행거리": req.mileage,
            "차량비고": req.vehicle_note,
            "Note": req.note,
            "EntryTime": datetime.datetime.now()
        }
        
        # Populate workers
        for i, w in enumerate(req.workers[:10], 1):
            u_key = "User" if i == 1 else f"User{i}"
            wt_key = "WorkTime" if i == 1 else f"WorkTime{i}"
            ot_key = "OT" if i == 1 else f"OT{i}"
            time_val = f"({w.shift}) {w.time}" if w.time else ""
            row_dict[u_key] = w.user
            row_dict[wt_key] = time_val
            row_dict[ot_key] = w.ot
            
        # Append to all available inventory Excel paths
        saved_count = 0
        for p in paths:
            try:
                wb = openpyxl.load_workbook(p)
                ws = wb["DailyUsage"]
                header = [cell.value for cell in ws[1]]
                
                new_row = []
                for col in header:
                    val = row_dict.get(col, None)
                    new_row.append(val)
                ws.append(new_row)
                wb.save(p)
                saved_count += 1
            except Exception as ex:
                print(f"Error saving to {p}: {ex}")
                
        return {"status": "success", "saved_databases": saved_count, "message": "성공적으로 저장되었습니다."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"저장 실패: {str(e)}")

@app.get("/", response_class=HTMLResponse)
def index():
    html_path = STATIC_DIR / "index.html"
    if html_path.exists():
        return html_path.read_text(encoding="utf-8")
    return "<h1>중앙지사 모바일 앱 준비 중</h1>"

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

