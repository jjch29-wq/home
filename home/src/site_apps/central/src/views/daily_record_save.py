"""Explicit source tracking and transactional single-entry replacement."""
from tkinter import messagebox
import pandas as pd


def remember_source(app, index):
    app._daily_edit_source = (index, app.daily_usage_df.loc[index].copy(deep=True))


def save_prepared(app, plan):
    source = getattr(app, '_daily_edit_source', None)
    replace = None
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
        # Without a source, check if an existing record matches the core job (Date, Site, Method, Pipe, Usage)
        for record in plan.get('records_to_save', []):
            candidates = app.daily_usage_df
            if candidates.empty:
                break
            match = pd.to_datetime(candidates['Date'], errors='coerce').dt.date == pd.to_datetime(record['Date']).date()
            
            def canon(val, is_num=False):
                if pd.isna(val): return ''
                if is_num:
                    try: return float(str(val).replace(',', '') or 0)
                    except ValueError: pass
                return str(val).strip()

            core_match = match.copy()
            for k, is_num in [('Site', False), ('검사방법', False), ('관경(Inch)', False), ('Usage', True)]:
                if k in candidates:
                    core_match &= candidates[k].map(lambda v: canon(v, is_num)) == canon(record.get(k, ''), is_num)

            if core_match.any():
                matched_idx = candidates[core_match].index[0]
                matched_row = candidates.loc[matched_idx]
                date_s = str(record.get('Date', '')).split()[0]
                method_s = str(record.get('검사방법', ''))
                pipe_s = str(record.get('관경(Inch)', ''))
                usage_s = str(record.get('Usage', ''))
                
                ans = messagebox.askyesnocancel(
                    '기존 작업 덮어쓰기',
                    f"{date_s}에 동일한 규격·수량의 작업이 이미 등록되어 있습니다.\n({method_s} {pipe_s} {usage_s})\n\n"
                    "기존 기록을 수정(덮어쓰기)할까요?\n\n"
                    "• [예]: 기존 기록 덮어쓰기 (수정 반영)\n"
                    "• [아니오]: 별도 새 작업으로 추가\n"
                    "• [취소]: 저장 취소",
                    parent=app.root
                )
                if ans is None:
                    return 0
                if ans:
                    replace = matched_idx
                    snapshot = matched_row.copy(deep=True)
                    source = (replace, snapshot)
                break
    names = ('daily_usage_df', 'transactions_df', 'materials_df', 'monthly_usage_df')
    frames = {name:getattr(app, name).copy(deep=True) for name in names}
    lists = {name:list(getattr(app, name)) for name in ('sites', 'users', 'vehicles', 'daily_units', 'equipments')}
    try:
        if replace is not None:
            # Preserve original cloud EntryTime and ClientRequestID so cloud background sync does not duplicate it
            orig_entry_time = snapshot.get('EntryTime') if source is not None else None
            orig_client_id = snapshot.get('ClientRequestID') if source is not None else None
            if orig_entry_time is not None and not pd.isna(orig_entry_time) and str(orig_entry_time).strip():
                for rec in plan.get('records_to_save', []):
                    rec['EntryTime'] = orig_entry_time
                    if orig_client_id:
                        rec['ClientRequestID'] = orig_client_id

            # Drop the replaced row and any sibling rows sharing the exact same EntryTime group
            replace_targets = [replace]
            if orig_entry_time is not None and not pd.isna(orig_entry_time) and str(orig_entry_time).strip() and 'EntryTime' in app.daily_usage_df:
                matched_same_entry = app.daily_usage_df[
                    (app.daily_usage_df.get('Date') == snapshot.get('Date')) &
                    (app.daily_usage_df['EntryTime'] == orig_entry_time)
                ].index.tolist()
                if replace in matched_same_entry:
                    replace_targets = matched_same_entry

            app.daily_usage_df = app.daily_usage_df.drop(replace_targets)
        count = app._commit_usage_plan(plan, auto_save=False)
        if not count or not app.save_data(require_primary=True):
            raise RuntimeError('원본 파일에 저장하지 못했습니다. 입력 내용은 유지됩니다.')
        app._daily_edit_source = None
        return count
    except Exception:
        for name, frame in frames.items(): setattr(app, name, frame)
        for name, values in lists.items(): getattr(app, name)[:] = values
        raise
