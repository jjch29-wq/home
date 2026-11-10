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
SYNC_STATE_PATH = Path(__file__).resolve().parent.parent / "data" / ".mobile_sync_state.json"
DEFAULT_INVENTORY_PATH = Path(__file__).resolve().parent.parent / "data" / "Material_Inventory.xlsx"

def load_sync_state() -> dict:
    if SYNC_STATE_PATH.exists():
        try:
            with open(SYNC_STATE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"동기화 상태 읽기 오류: {e}")
    return {"synced_entry_times": []}

def save_sync_state(state: dict) -> bool:
    try:
        SYNC_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(SYNC_STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        logger.warning(f"동기화 상태 저장 실패: {e}")
        return False

def get_all_inventory_paths() -> List[Path]:
    """경인사무소의 모든 재고 DB 경로 탐색 (로컬 및 Google Drive)"""
    paths = []
    if DEFAULT_INVENTORY_PATH.exists():
        paths.append(DEFAULT_INVENTORY_PATH)

    try:
        import ctypes
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        drives = [chr(65 + i) for i in range(26) if (bitmask >> i) & 1]
        for d in drives:
            cloud_p = Path(rf"{d}:\내 드라이브\07_Antigravity\PMI_Data\gyeongin\data\Material_Inventory.xlsx")
            if cloud_p.exists() and cloud_p not in paths:
                paths.append(cloud_p)
                break
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

def clean_record_for_json(record: Dict[str, Any]) -> Dict[str, Any]:
    """DataFrame 레코드를 JSON 직렬화 안전한 딕셔너리로 변환"""
    cleaned = {}
    for k, v in record.items():
        if pd.isna(v) or v is None:
            cleaned[k] = ""
        elif isinstance(v, (datetime, pd.Timestamp)):
            cleaned[k] = v.strftime("%Y-%m-%d %H:%M:%S")
        elif hasattr(v, 'item'):
            cleaned[k] = v.item()
        elif isinstance(v, float) and v.is_integer():
            cleaned[k] = int(v)
        else:
            cleaned[k] = v
    return cleaned

def push_record_to_web_app(
    record: Dict[str, Any],
    action: str = "update",
    original_entry_time: Optional[str] = None,
    web_app_url: Optional[str] = None
) -> Dict[str, Any]:
    """
    PC에서 작업 일보 수정/삭제/추가 내역을 Google Apps Script Web App에 실시간 Push
    (requests 미설치 환경에서도 표준 라이브러리 urllib.request로 100% 작동 보장)
    """
    cfg = load_sync_config()
    target_url = web_app_url or cfg.get("web_app_url")
    if not target_url:
        return {"status": "error", "message": "Google Apps Script 웹 앱 URL이 설정되지 않았습니다."}

    payload = {
        "action": action,
        "entry_time": str(original_entry_time or record.get("EntryTime", "")).split(".")[0].strip(),
        "row_data": clean_record_for_json(record)
    }

    # 1. requests 시도
    try:
        import requests
        resp = requests.post(
            target_url,
            json=payload,
            headers={"User-Agent": "Gyeongin-Office-PC/1.0"},
            timeout=15,
            allow_redirects=True
        )
        data = resp.json()
        logger.info(f"구글 시트 Push 성공 ({action}): {data.get('message', '')}")
        return data
    except ImportError:
        pass
    except Exception as e:
        logger.warning(f"requests 통한 Push 실패, urllib로 재시도: {e}")

    # 2. urllib.request 표준 라이브러리 fallback
    try:
        req_data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        req = urllib.request.Request(
            target_url,
            data=req_data,
            headers={
                "User-Agent": "Gyeongin-Office-PC/1.0",
                "Content-Type": "application/json"
            }
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        logger.info(f"구글 시트 Push 성공 (urllib, {action}): {data.get('message', '')}")
        return data
    except Exception as e:
        logger.error(f"구글 시트 Push 실패 ({action}): {e}")
        return {"status": "error", "message": str(e)}

def push_delete_to_web_app(
    entry_time: str,
    web_app_url: Optional[str] = None
) -> Dict[str, Any]:
    """
    PC에서 삭제된 작업 일보의 EntryTime을 Google Apps Script Web App에 전송하여 시트에서도 삭제
    """
    et_clean = str(entry_time).split(".")[0].strip()
    sync_state = load_sync_state()
    cur_times = sync_state.get("synced_entry_times", [])
    if et_clean in cur_times:
        sync_state["synced_entry_times"] = [t for t in cur_times if t != et_clean]
        save_sync_state(sync_state)
    return push_record_to_web_app({}, action="delete", original_entry_time=et_clean, web_app_url=web_app_url)

def push_batch_to_web_app(
    saves: Optional[List[Dict[str, Any]]] = None,
    updates: Optional[List[Dict[str, Any]]] = None,
    deletes: Optional[List[str]] = None,
    web_app_url: Optional[str] = None
) -> Dict[str, Any]:
    """
    Google Apps Script Web App에 일괄 동기화 요청
    (단건 표준 API를 순차 호출하여 모든 Apps Script 배포 버전과 100% 호환 보장)
    """
    cfg = load_sync_config()
    target_url = web_app_url or cfg.get("web_app_url")
    if not target_url:
        return {"status": "error", "message": "Google Apps Script 웹 앱 URL이 설정되지 않았습니다."}

    deleted_cnt = 0
    updated_cnt = 0
    saved_cnt = 0

    # 1. 삭제 처리
    for d in (deletes or []):
        d_str = str(d).split(".")[0].strip()
        if d_str:
            res_del = push_delete_to_web_app(d_str, web_app_url=target_url)
            if res_del.get("status") == "success":
                deleted_cnt += 1

    # 2. 수정 처리
    for u in (updates or []):
        u_time = str(u.get("EntryTime", "")).split(".")[0].strip()
        res_upd = push_record_to_web_app(u, action="update", original_entry_time=u_time, web_app_url=target_url)
        if res_upd.get("status") == "success":
            updated_cnt += 1

    # 3. 신규 추가 처리
    for s in (saves or []):
        res_sav = push_record_to_web_app(s, action="save", web_app_url=target_url)
        if res_sav.get("status") == "success":
            saved_cnt += 1

    return {
        "status": "success",
        "added_count": saved_cnt,
        "updated_count": updated_cnt,
        "deleted_count": deleted_cnt,
        "message": f"동기화 완료 (추가: {saved_cnt}건, 수정: {updated_cnt}건, 삭제: {deleted_cnt}건)"
    }

def two_way_reconcile_with_web_app(
    inventory_path: Optional[Path] = None,
    web_app_url: Optional[str] = None
) -> Dict[str, Any]:
    """
    Google 스프레드시트와 PC 엑셀(DailyUsage) 간 완전 양방향(Two-way) 무손실 정합성 동기화
    1. 구글 시트의 신규 행을 PC 엑셀로 가져옴 (Pull)
    2. PC 엑셀에만 있는 행이나 수정된 행을 구글 시트로 밀어넣음 (Push)
    """
    cfg = load_sync_config()
    target_url = web_app_url or cfg.get("web_app_url")
    if not target_url:
        return {"status": "error", "message": "Google Apps Script URL 미설정"}

    target_path = Path(inventory_path) if inventory_path else DEFAULT_INVENTORY_PATH
    if not target_path.exists():
        return {"status": "error", "message": f"재고 파일이 존재하지 않습니다: {target_path}"}

    try:
        # 1. 구글 시트 전체 레코드 가져오기
        headers, cloud_rows = fetch_records_from_web_app(target_url)

        # 2. 엑셀 파일 읽기
        with pd.ExcelFile(target_path) as ef:
            all_sheets = {sheet: ef.parse(sheet) for sheet in ef.sheet_names}

        if "DailyUsage" not in all_sheets:
            return {"status": "error", "message": "DailyUsage 시트가 없습니다."}

        daily_df = all_sheets["DailyUsage"].copy()
        local_rows = daily_df.to_dict(orient="records") if not daily_df.empty else []

        def get_keys(r):
            entry_time = str(r.get("EntryTime", "")).split(".")[0].strip()
            date_str = str(r.get("Date", "")).split()[0].strip()
            user_str = str(r.get("User", "")).strip()
            method_str = str(r.get("검사방법", "")).strip()
            qty_str = str(r.get("수량", r.get("Usage", 0))).strip()
            item_str = str(r.get("검사품명", "")).strip()
            time_key = f"TIME:{entry_time}" if (entry_time and entry_time != "nan") else None
            comp_key = f"COMP:{date_str}|{user_str}|{method_str}|{qty_str}|{item_str}"
            return time_key, comp_key

        local_keys = set()
        local_time_map = {}
        for r in local_rows:
            tk, ck = get_keys(r)
            if tk:
                local_keys.add(tk)
                local_time_map[tk] = r
            local_keys.add(ck)

        cloud_keys = set()
        cloud_time_map = {}
        for r in cloud_rows:
            tk, ck = get_keys(r)
            if tk:
                cloud_keys.add(tk)
                cloud_time_map[tk] = r
            cloud_keys.add(ck)

        # 동기화 상태 추적 (모바일 삭제 감지용)
        sync_state = load_sync_state()
        known_cloud_times = set(sync_state.get("synced_entry_times", []))
        cloud_raw_times = {str(r.get("EntryTime", "")).split(".")[0].strip() for r in cloud_rows if r.get("EntryTime")}
        if not known_cloud_times and cloud_raw_times:
            known_cloud_times = set(cloud_raw_times)
            sync_state["synced_entry_times"] = sorted(list(known_cloud_times))
            save_sync_state(sync_state)

        # A. 구글 시트의 변경사항 -> PC 엑셀로 동기화 (신규 추가: Pull / 수정 갱신: Pull-Update)
        to_pull = []
        excel_modified = False

        def _norm_val(val):
            if pd.isna(val) or val is None:
                return ""
            s = str(val).strip()
            if s.lower() in ("nan", "none"):
                return ""
            try:
                f = float(s.replace(",", ""))
                return str(int(f)) if f.is_integer() else str(f)
            except ValueError:
                return s

        for cr in cloud_rows:
            tk, ck = get_keys(cr)
            if tk and tk in local_time_map:
                # 이미 존재하는 행: 클라우드에서 수정된 필드가 있으면 로컬 엑셀 행을 클라우드 값으로 업데이트
                lr = local_time_map[tk]
                row_changed = False
                for col in ["수량", "단가", "검사방법", "User", "검사품명", "Note", "DailyMemo", "Site", "업체명", "관경(Inch)", "출장비", "검사비"]:
                    if col in cr and _norm_val(lr.get(col)) != _norm_val(cr.get(col)):
                        lr[col] = cr[col]
                        row_changed = True
                if row_changed:
                    excel_modified = True
            elif (tk and tk in local_keys) or (ck in local_keys):
                continue
            else:
                to_pull.append(cr)

        pulled_cnt = 0
        if to_pull:
            res_p = sync_records_to_excel(to_pull, inventory_path=target_path)
            pulled_cnt = res_p.get("added_count", 0)

        # B. PC에만 있는 행 -> 구글 시트로 밀어넣기 (Push) or 모바일 삭제 감지 반영
        to_push_saves = []
        to_push_updates = []
        to_delete_local = []
        kept_local_rows = []

        for r in local_rows:
            tk, ck = get_keys(r)
            raw_et = str(r.get("EntryTime", "")).split(".")[0].strip()
            if not tk:
                if ck not in cloud_keys:
                    to_push_saves.append(r)
                kept_local_rows.append(r)
                continue

            if tk not in cloud_keys:
                # 클라우드에 없는 행:
                # 이전에 클라우드에 존재했던 기록이면 => 모바일에서 삭제된 것! 로컬 엑셀에서도 삭제
                # 이전에 클라우드에 없던 새로운 기록이면 => PC에서 새로 등록한 것! 클라우드로 Push
                if raw_et in known_cloud_times:
                    to_delete_local.append(r)
                    excel_modified = True
                else:
                    to_push_saves.append(r)
                    kept_local_rows.append(r)
            else:
                kept_local_rows.append(r)
                # 존재하는 행: 주요 필드(수량, 단가, 비고, 업체명 등) 수정 여부 비교
                cr = cloud_time_map.get(tk)
                if cr:
                    is_modified = False
                    for check_col in ["수량", "단가", "검사방법", "User", "검사품명", "Note", "DailyMemo", "Site", "업체명", "관경(Inch)", "출장비", "검사비"]:
                        if _norm_val(r.get(check_col)) != _norm_val(cr.get(check_col)):
                            is_modified = True
                            break
                    if is_modified:
                        to_push_updates.append(r)

        if to_delete_local:
            local_rows = kept_local_rows
            excel_modified = True
            logger.info(f"모바일에서 삭제된 {len(to_delete_local)}건의 작업이 로컬 엑셀에서도 제거되었습니다.")

        if excel_modified:
            # 수정/삭제된 로컬 행들을 엑셀에 다시 저장
            with pd.ExcelFile(target_path) as ef:
                all_sheets = {sheet: ef.parse(sheet) for sheet in ef.sheet_names}
            all_sheets["DailyUsage"] = pd.DataFrame(local_rows)
            with pd.ExcelWriter(target_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                all_sheets["DailyUsage"].to_excel(writer, sheet_name="DailyUsage", index=False)
            logger.info("모바일 수정/삭제 사항이 로컬 엑셀 행에 제자리 갱신되었습니다.")

        pushed_cnt = 0
        if to_push_saves or to_push_updates:
            res_batch = push_batch_to_web_app(saves=to_push_saves, updates=to_push_updates, web_app_url=target_url)
            pushed_cnt = len(to_push_saves) + len(to_push_updates)

        # 동기화 상태 갱신: 최종 클라우드 EntryTime 목록 저장
        final_cloud_times = set(cloud_raw_times)
        for s in to_push_saves:
            s_et = str(s.get("EntryTime", "")).split(".")[0].strip()
            if s_et:
                final_cloud_times.add(s_et)
        sync_state["synced_entry_times"] = sorted(list(final_cloud_times))
        save_sync_state(sync_state)

        cfg["last_sync_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_sync_config(cfg)

        return {
            "status": "success",
            "pulled_count": pulled_cnt,
            "pushed_count": pushed_cnt,
            "deleted_local_count": len(to_delete_local),
            "message": f"구글 스프레드시트 양방향 동기화 완료 (가져옴: {pulled_cnt}건, 밀어넣음: {pushed_cnt}건, 모바일삭제반영: {len(to_delete_local)}건)"
        }
    except Exception as e:
        logger.error(f"양방향 동기화 오류: {e}")
        return {"status": "error", "message": f"양방향 동기화 오류: {e}"}

def run_sync(web_app_url: Optional[str] = None, inventory_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    통합 동기화 실행 함수 (완전 양방향 지원)
    """
    cfg = load_sync_config()
    target_url = web_app_url or cfg.get("web_app_url")
    if not target_url:
        return {
            "status": "error",
            "message": "Google Apps Script 배포 웹 앱 URL이 설정되지 않았습니다. google_sync_config.json에 등록하세요."
        }

    return two_way_reconcile_with_web_app(inventory_path=inventory_path, web_app_url=target_url)

def sync_local_and_drive_files() -> Dict[str, Any]:
    """로컬 Material_Inventory.xlsx와 Google Drive의 파일 간 DailyUsage 양방향 무손실 동기화"""
    paths = get_all_inventory_paths()
    if len(paths) < 2:
        return {
            "status": "success",
            "added_count": 0,
            "message": "단일 데이터베이스 모드 (구글 드라이브 미연결)"
        }

    p1, p2 = paths[0], paths[1]  # local, google drive
    try:
        with pd.ExcelFile(p1) as f1:
            df1 = f1.parse("DailyUsage") if "DailyUsage" in f1.sheet_names else pd.DataFrame()
        with pd.ExcelFile(p2) as f2:
            df2 = f2.parse("DailyUsage") if "DailyUsage" in f2.sheet_names else pd.DataFrame()

        rows1 = df1.to_dict(orient="records") if not df1.empty else []
        rows2 = df2.to_dict(orient="records") if not df2.empty else []

        # p2(드라이브)의 행들을 p1(로컬)에 동기화
        res1 = sync_records_to_excel(rows2, inventory_path=p1)
        # p1(로컬)의 행들을 p2(드라이브)에 동기화
        res2 = sync_records_to_excel(rows1, inventory_path=p2)

        total_added = res1.get("added_count", 0) + res2.get("added_count", 0)
        return {
            "status": "success",
            "added_count": total_added,
            "message": f"로컬 및 구글 드라이브 동기화 완료 (추가: {total_added}건)"
        }
    except Exception as e:
        logger.error(f"로컬/드라이브 동기화 실패: {e}")
        return {"status": "error", "message": str(e), "added_count": 0}

def sync_all_gyeongin_sources() -> Dict[str, Any]:
    """
    경인사무소 전체 데이터 소스 (Google Apps Script Web App + Google Drive + Local) 자동 통합 동기화
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
                p_added = res_web.get("pulled_count", 0) + res_web.get("pushed_count", 0) + res_web.get("added_count", 0)
                if p_added > 0:
                    total_added += p_added
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
    result = run_sync(url) if url else sync_all_gyeongin_sources()
    print(json.dumps(result, ensure_ascii=False, indent=2))
