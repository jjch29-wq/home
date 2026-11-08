"""
경인사무소 24시간 모바일 클라우드(Google Apps Script / Sheets) 동기화 유틸리티
- 퇴근 후 또는 외부에서 스마트폰으로 구글 스프레드시트에 저장된 작업일보 내역을
  사무실 PC의 Material_Inventory.xlsx DailyUsage 시트로 가져와 동기화합니다.
- 중복 저장을 방지하기 위해 EntryTime 및 (Date, User, 검사방법, 수량, 검사품명) 복합키를 검증합니다.
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

logger = logging.getLogger("gyeongin_sync")

CONFIG_PATH = Path(__file__).resolve().parent.parent / "data" / "google_sync_config.json"
DEFAULT_INVENTORY_PATH = Path(__file__).resolve().parent.parent / "data" / "Material_Inventory.xlsx"

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
        "auto_sync_on_startup": False,
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
    """
    Google Apps Script 배포 URL (Web App)에서 전체 레코드 획득
    """
    if not web_app_url:
        raise ValueError("Google Apps Script 배포 웹 앱 URL이 지정되지 않았습니다.")
    
    url = web_app_url.strip()
    delimiter = "&" if "?" in url else "?"
    sync_url = f"{url}{delimiter}api=sync"

    req = urllib.request.Request(
        sync_url,
        headers={"User-Agent": "Gyeongin-Office-PC/1.0"}
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
    """
    구글 시트에서 가져온 레코드들을 로컬 Material_Inventory.xlsx의 DailyUsage에 추가
    """
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
        item_str = str(row.get("검사품명", "")).strip()

        if entry_time and entry_time != "nan":
            existing_keys.add(f"TIME:{entry_time}")
        composite_key = f"COMP:{date_str}|{user_str}|{method_str}|{qty_str}|{item_str}"
        existing_keys.add(composite_key)

    new_rows = []
    skipped_count = 0

    for r in rows:
        entry_time = str(r.get("EntryTime", "")).strip()
        date_str = str(r.get("Date", "")).split()[0].strip()
        user_str = str(r.get("User", "")).strip()
        method_str = str(r.get("검사방법", "")).strip()
        qty_str = str(r.get("수량", r.get("Usage", 0))).strip()
        item_str = str(r.get("검사품명", "")).strip()

        time_key = f"TIME:{entry_time}" if (entry_time and entry_time != "nan") else None
        comp_key = f"COMP:{date_str}|{user_str}|{method_str}|{qty_str}|{item_str}"

        if (time_key and time_key in existing_keys) or (comp_key in existing_keys):
            skipped_count += 1
            continue

        # 신규 레코드
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

    # DataFrame 생성 및 컬럼 정합성 유지
    new_df = pd.DataFrame(new_rows)
    for col in daily_df.columns:
        if col not in new_df.columns:
            new_df[col] = None

    # 기존 열 순서대로 정렬
    aligned_new_df = new_df[[col for col in daily_df.columns if col in new_df.columns]].copy()
    updated_daily_df = pd.concat([daily_df, aligned_new_df], ignore_index=True)

    all_sheets["DailyUsage"] = updated_daily_df

    # 원본 엑셀에 저장
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
    """
    통합 동기화 실행 함수
    """
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

if __name__ == "__main__":
    import sys
    url = sys.argv[1] if len(sys.argv) > 1 else None
    result = run_sync(url)
    print(json.dumps(result, ensure_ascii=False, indent=2))
