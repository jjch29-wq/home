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
        # Without a source, matching dates alone cannot identify an edit target.
        for record in plan['records_to_save']:
            candidates = app.daily_usage_df
            if candidates.empty:
                break
            match = pd.to_datetime(candidates['Date'], errors='coerce').dt.date == pd.to_datetime(record['Date']).date()
            for key in ('Site', '검사방법', '검사품명', 'MaterialID', '관경(Inch)', 'Usage', '단가'):
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
            app.daily_usage_df = app.daily_usage_df.drop(replace)
        count = app._commit_usage_plan(plan, auto_save=False)
        if not count or not app.save_data(require_primary=True):
            raise RuntimeError('원본 파일에 저장하지 못했습니다. 입력 내용은 유지됩니다.')
        app._daily_edit_source = None
        return count
    except Exception:
        for name, frame in frames.items(): setattr(app, name, frame)
        for name, values in lists.items(): getattr(app, name)[:] = values
        raise
