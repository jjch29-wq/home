"""
한국지역난방공사 중앙지사 24시간 모바일 클라우드(Google Apps Script / Sheets) 동기화 유틸리티
- 퇴근 후 또는 외부에서 스마트폰으로 구글 스프레드시트에 저장된 작업일보 내역을
  사무실 PC의 Material_Inventory.xlsx DailyUsage 시트로 가져와 동기화합니다.
- 중복 저장을 방지하기 위해 EntryTime 및 (Date, User, 검사방법, 수량, 관경) 복합키를 검증합니다.
"""

import json
import logging
import os
import shutil
import urllib.request
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd

logger = logging.getLogger("central_sync")

CONFIG_PATH = Path(__file__).resolve().parent.parent / "data" / "google_sync_config.json"
DEFAULT_INVENTORY_PATH = Path(__file__).resolve().parent.parent / "data" / "Material_Inventory.xlsx"

def get_all_inventory_paths() -> List[Path]:
    """중앙지사의 모든 재고 DB 경로 탐색 (로컬 및 Google Drive)"""
    paths = []
    if DEFAULT_INVENTORY_PATH.exists():
        paths.append(DEFAULT_INVENTORY_PATH)

    try:
        from site_apps.central.src.data_sync import get_business_data_dir
        cloud_dir = get_business_data_dir()
        cloud_p = Path(cloud_dir) / "Material_Inventory.xlsx"
        if cloud_p.exists() and cloud_p not in paths:
            paths.append(cloud_p)
    except Exception:
        pass

    try:
        import ctypes
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        drives = [chr(65 + i) for i in range(26) if (bitmask >> i) & 1]
        for d in ['G', 'F']:
            if d in drives:
                alt_p = Path(f"{d}:/내 드라이브/07_Antigravity/PMI_Data/central/data/Material_Inventory.xlsx")
                if alt_p.exists() and alt_p not in paths:
                    paths.append(alt_p)
    except Exception:
        pass

    return paths

def load_sync_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"설정 파일 읽기 오류: {e}")
    return {
        "web_app_url": "",
        "sheet_id": "",
        "auto_sync_on_startup": True,
        "last_sync_time": ""
    }

def save_sync_config(cfg: dict) -> bool:
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        logger.error(f"설정 파일 저장 실패: {e}")
        return False

def fetch_records_from_web_app(web_app_url: str) -> Tuple[List[str], List[Dict[str, Any]]]:
    """Google Apps Script 배포 URL (Web App)에서 전체 레코드 획득"""
    if not web_app_url:
        raise ValueError("Google Apps Script 배포 웹 앱 URL이 지정되지 않았습니다.")

    url = web_app_url.strip()
    delimiter = "&" if "?" in url else "?"
    sync_url = f"{url}{delimiter}api=sync"

    req = urllib.request.Request(
        sync_url,
        headers={"User-Agent": "Central-Office-PC/1.0"}
    )

    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    if data.get("status") != "success":
        raise RuntimeError(f"클라우드 API 반환 오류: {data.get('detail', 'Unknown error')}")

    headers = data.get("headers", [])
    rows = data.get("rows", [])
    return headers, rows

def sync_records_to_excel(
    rows: List[Dict[str, Any]],
    inventory_path: Optional[Path] = None
) -> Dict[str, Any]:
    """구글 시트에서 가져온 레코드들을 로컬/클라우드 Material_Inventory.xlsx의 DailyUsage에 추가"""
    target_path = Path(inventory_path) if inventory_path else DEFAULT_INVENTORY_PATH
    if not target_path.exists():
        raise FileNotFoundError(f"재고 파일이 존재하지 않습니다: {target_path}")

    # 백업 생성
    backup_dir = target_path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"Material_Inventory_pre_sync_{timestamp}.xlsx"
    shutil.copy2(target_path, backup_path)

    # 엑셀 파일 읽기
    with pd.ExcelFile(target_path) as excel_file:
        all_sheets = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}

    if "DailyUsage" not in all_sheets:
        raise ValueError("Material_Inventory.xlsx에 DailyUsage 시트가 없습니다.")

    daily_df = all_sheets["DailyUsage"].copy()

    # 기존 행들의 고유 키 집합 생성 (중복 방지)
    existing_keys = set()
    for _, row in daily_df.iterrows():
        entry_time = str(row.get("EntryTime", "")).strip()
        date_str = str(row.get("Date", "")).split()[0].strip()
        user_str = str(row.get("User", "")).strip()
        method_str = str(row.get("검사방법", "")).strip()
        qty_str = str(row.get("수량", row.get("Usage", 0))).strip()
        pipe_str = str(row.get("관경(Inch)", row.get("규격", ""))).strip()

        if entry_time and entry_time != "nan":
            existing_keys.add(f"TIME:{entry_time}")
        composite_key = f"COMP:{date_str}|{user_str}|{method_str}|{qty_str}|{pipe_str}"
        existing_keys.add(composite_key)

    new_rows = []
    skipped_count = 0

    for r in rows:
        entry_time = str(r.get("EntryTime", "")).strip()
        date_str = str(r.get("Date", "")).split()[0].strip()
        user_str = str(r.get("User", "")).strip()
        method_str = str(r.get("검사방법", "")).strip()
        qty_str = str(r.get("수량", r.get("Usage", 0))).strip()
        pipe_str = str(r.get("관경(Inch)", r.get("규격", ""))).strip()

        time_key = f"TIME:{entry_time}" if (entry_time and entry_time != "nan") else None
        comp_key = f"COMP:{date_str}|{user_str}|{method_str}|{qty_str}|{pipe_str}"

        if (time_key and time_key in existing_keys) or (comp_key in existing_keys):
            skipped_count += 1
            continue

        new_rows.append(r)
        if time_key:
            existing_keys.add(time_key)
        existing_keys.add(comp_key)

    if not new_rows:
        return {
            "status": "success",
            "added_count": 0,
            "skipped_count": skipped_count,
            "total_cloud_rows": len(rows),
            "message": f"동기화할 새로운 작업 내역이 없습니다. (이미 최신 상태 / 중복 {skipped_count}건 제외)"
        }

    new_df = pd.DataFrame(new_rows)
    for col in daily_df.columns:
        if col not in new_df.columns:
            new_df[col] = None

    aligned_new_df = new_df[[col for col in daily_df.columns if col in new_df.columns]].copy()
    updated_daily_df = pd.concat([daily_df, aligned_new_df], ignore_index=True)

    all_sheets["DailyUsage"] = updated_daily_df

    with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return {
        "status": "success",
        "added_count": len(new_rows),
        "skipped_count": skipped_count,
        "total_cloud_rows": len(rows),
        "backup_path": str(backup_path),
        "message": f"총 {len(new_rows)}건의 클라우드 작업 내역이 PC 엑셀(DailyUsage)에 성공적으로 동기화되었습니다."
    }

def run_sync(web_app_url: Optional[str] = None, inventory_path: Optional[Path] = None) -> Dict[str, Any]:
    """통합 동기화 실행 함수"""
    cfg = load_sync_config()
    target_url = web_app_url or cfg.get("web_app_url")
    if not target_url:
        return {
            "status": "error",
            "message": "Google Apps Script 배포 웹 앱 URL이 설정되지 않았습니다. google_sync_config.json에 등록하세요."
        }

    try:
        headers, rows = fetch_records_from_web_app(target_url)
        res = sync_records_to_excel(rows, inventory_path)

        cfg["last_sync_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if web_app_url:
            cfg["web_app_url"] = web_app_url
        save_sync_config(cfg)

        return res
    except Exception as e:
        logger.error(f"동기화 실패: {e}")
        return {
            "status": "error",
            "message": f"동기화 중 오류 발생: {str(e)}"
        }

def sync_local_and_drive_files() -> Dict[str, Any]:
    """로컬 Material_Inventory.xlsx와 Google Drive의 파일 간 DailyUsage 양방향 무손실 동기화"""
    paths = get_all_inventory_paths()
    if len(paths) < 2:
        return {
            "status": "success",
            "added_count": 0,
            "message": "단일 데이터베이스 모드 (구글 드라이브 미연결)"
        }

    p1, p2 = paths[0], paths[1]
    try:
        with pd.ExcelFile(p1) as f1:
            df1 = f1.parse("DailyUsage") if "DailyUsage" in f1.sheet_names else pd.DataFrame()
        with pd.ExcelFile(p2) as f2:
            df2 = f2.parse("DailyUsage") if "DailyUsage" in f2.sheet_names else pd.DataFrame()

        rows1 = df1.to_dict(orient="records") if not df1.empty else []
        rows2 = df2.to_dict(orient="records") if not df2.empty else []

        res1 = sync_records_to_excel(rows2, inventory_path=p1)
        res2 = sync_records_to_excel(rows1, inventory_path=p2)

        total_added = res1.get("added_count", 0) + res2.get("added_count", 0)
        return {
            "status": "success",
            "added_count": total_added,
            "message": f"중앙지사 로컬 및 구글 드라이브 동기화 완료 (추가: {total_added}건)"
        }
    except Exception as e:
        logger.error(f"로컬/드라이브 동기화 실패: {e}")
        return {"status": "error", "message": str(e), "added_count": 0}

def sync_all_central_sources() -> Dict[str, Any]:
    """
    중앙지사 전체 데이터 소스 (Google Apps Script Web App + Google Drive + Local) 자동 통합 동기화
    휴대폰에서 저장된 내역을 앱 실행 시 또는 백그라운드에서 자동으로 감지하여 엑셀에 반영합니다.
    """
    total_added = 0
    messages = []

    # 1. Google Drive 및 Local 간 상호 동기화
    res_drive = sync_local_and_drive_files()
    if res_drive.get("added_count", 0) > 0:
        total_added += res_drive["added_count"]
        messages.append(res_drive["message"])

    # 2. Google Apps Script 클라우드 Web App 동기화 (설정된 경우)
    cfg = load_sync_config()
    target_url = cfg.get("web_app_url")
    if target_url:
        try:
            for p in get_all_inventory_paths():
                res_web = run_sync(target_url, inventory_path=p)
                if res_web.get("added_count", 0) > 0:
                    total_added += res_web["added_count"]
                    messages.append(res_web["message"])
        except Exception as e:
            messages.append(f"웹 앱 동기화 건너뜀: {e}")

    summary_msg = " / ".join(messages) if messages else "이미 최신 상태입니다."
    return {
        "status": "success",
        "added_count": total_added,
        "message": summary_msg
    }

if __name__ == "__main__":
    import sys
    url = sys.argv[1] if len(sys.argv) > 1 else None
    result = run_sync(url) if url else sync_all_central_sources()
    print(json.dumps(result, ensure_ascii=False, indent=2))
