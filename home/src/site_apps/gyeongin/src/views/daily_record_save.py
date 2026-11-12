"""Explicit source tracking and transactional single-entry replacement."""
from tkinter import messagebox
import pandas as pd


def remember_source(app, index):
    app._daily_edit_source = (index, app.daily_usage_df.loc[index].copy(deep=True))


def save_prepared(app, plan):
    source = getattr(app, '_daily_edit_source', None)
    replace = None
    snapshot = None
    if source is not None:
        index, snapshot = source
        if index not in app.daily_usage_df.index or not app.daily_usage_df.loc[index].equals(snapshot):
            raise ValueError('원본 기록이 변경되었습니다. 최근 기록에서 다시 불러와주세요.')
        answer = messagebox.askyesnocancel('기록 재저장', '불러온 기존 기록을 덮어쓸까요?\n예: 덮어쓰기 / 아니오: 새 기록 추가 / 취소: 저장 취소', parent=app.root)
        if answer is None:
            return 0
        if answer:
            replace = index
    else:
        # Without a source, matching dates alone cannot identify an edit target.
        for record in plan['records_to_save']:
            candidates = app.daily_usage_df
            if candidates.empty:
                break
            match = pd.to_datetime(candidates['Date'], errors='coerce').dt.date == pd.to_datetime(record['Date']).date()
            for key in ('Site', '검사방법', '검사품명', 'MaterialID', '검사대상', '검사규격', '관경(Inch)', 'Usage', '단가'):
                if key not in candidates:
                    continue
                def canonical(value):
                    if pd.isna(value): return ''
                    if key in ('Usage', '단가'):
                        try: return float(str(value).replace(',', '') or 0)
                        except ValueError: pass
                    return str(value).strip().removesuffix('.0') if key == 'MaterialID' else str(value).strip()
                match &= candidates[key].map(canonical) == canonical(record.get(key, ''))
            if match.any():
                if not messagebox.askyesno('중복 기록 확인', '동일한 작업 기록이 있습니다. 별도 기록으로 추가할까요?\n기존 기록 수정은 최근 기록에서 불러온 뒤 저장해주세요.', parent=app.root):
                    return 0
                break
    names = ('daily_usage_df', 'transactions_df', 'materials_df', 'monthly_usage_df')
    frames = {name:getattr(app, name).copy(deep=True) for name in names}
    lists = {name:list(getattr(app, name)) for name in ('sites', 'users', 'vehicles', 'daily_units', 'equipments')}
    try:
        if replace is not None:
            # EntryTime is the cloud reconciliation identity, not the edit time.
            original_time = snapshot.get('EntryTime')
            if pd.notna(original_time) and str(original_time).strip():
                plan = dict(plan)
                plan['common_data'] = dict(plan.get('common_data', {}), EntryTime=original_time)
                plan['records_to_save'] = [dict(record, EntryTime=original_time)
                                           for record in plan['records_to_save']]
            app.daily_usage_df = app.daily_usage_df.drop(replace)
        count = app._commit_usage_plan(plan, auto_save=False)
        if not count or not app.save_data(require_primary=True):
            raise RuntimeError('원본 파일에 저장하지 못했습니다. 입력 내용은 유지됩니다.')
        app._daily_edit_source = None

        # 일보 저장 시 신규 업체명이 입력된 경우 PC 마스터 목록(companies)에 자동 등록
        try:
            companies_list = getattr(app, 'companies', None)
            if isinstance(companies_list, list):
                added_company = False
                for rec in plan.get('records_to_save', []):
                    co_name = str(rec.get('업체명', '')).strip()
                    if co_name and co_name.lower() != 'nan' and co_name not in companies_list:
                        companies_list.append(co_name)
                        added_company = True
                if added_company:
                    if hasattr(app, 'refresh_ui_for_list_change'):
                        app.refresh_ui_for_list_change('companies')
                    if hasattr(app, 'save_tab_config'):
                        app.save_tab_config()
        except Exception:
            pass

        # [경인 클라우드 실시간 Push] 구글 스프레드시트 웹앱으로 수정/신규 내역 비동기 전송
        try:
            import threading
            from site_apps.gyeongin.src.sync_google_sheet import push_record_to_web_app

            records = list(plan.get('records_to_save', []))
            is_update = (replace is not None)
            orig_time = str(snapshot.get('EntryTime', '')).strip() if (is_update and snapshot is not None) else None

            def _async_push():
                for rec in records:
                    push_record_to_web_app(rec, action='update' if is_update else 'save', original_entry_time=orig_time)

            threading.Thread(target=_async_push, daemon=True).start()
        except Exception:
            pass

        return count
    except Exception:
        for name, frame in frames.items(): setattr(app, name, frame)
        for name, values in lists.items(): getattr(app, name)[:] = values
        raise
