"""Switch the daily entry date without carrying the previous job forward."""
from tkinter import messagebox


def show_recent_scope(app, selected_only=False):
    """Change only the record list, never the date or unsaved editor contents."""
    app._daily_recent_date_filter = app.ent_daily_date.get_date() if selected_only else None
    app._daily_recent_show_all = not selected_only
    app.update_recent_entries_view()


def has_daily_input(app):
    data = app._get_usage_session_data()
    # 1. 작업 목록에 대기 중인 작업이 있는 경우
    batch = data.get('inspection_batch')
    if batch and len(batch) > 0:
        return True
    
    # 2. 사용자가 직접 기입한 유의미한 필드 (자동 세팅되는 기본값 applied_code, method, unit_price 등 제외)
    user_filled_fields = ('report_no', 'inspection_item', 'note', 'main_memo', 'material')
    if any(str(data.get(key, '') or '').strip() for key in user_filled_fields):
        return True
        
    # 검사량이 0 초과로 기입된 경우
    amount_str = str(data.get('amount', '') or '').strip()
    try:
        if amount_str and float(amount_str) > 0:
            return True
    except ValueError:
        if amount_str:
            return True

    # 3. 화학약품 / RTK에 기입된 값이 있는 경우
    if any(str(value or '').strip() for value in data.get('rtk', {}).values() if str(value or '').strip() not in ('', '0')):
        return True
    for co in data.get('ndt_companies', []):
        for k, v in co.items():
            if k != '_company' and str(v or '').strip() not in ('', '0'):
                return True
                
    # 4. 차량 점검에 주행거리가 입력된 경우
    for row in data.get('vehicles', []):
        km = str(row.get('운행거리', '') or row.get('mileage', '') or '').strip()
        if km and km not in ('0', '0.0'):
            return True

    return any(getattr(app, f'worker_group{i}', None) is not None
               and getattr(app, f'worker_group{i}').ent_meal.get().strip()
               for i in range(1, 11))


def change_daily_date(app, force=False):
    selected = app.ent_daily_date.get_date()
    previous = getattr(app, '_last_daily_date', None)
    if selected == previous and not force:
        return True
    if has_daily_input(app) and not messagebox.askyesno(
            '작업 날짜 변경',
            '입력 내용과 등록할 작업 목록을 비우고 날짜를 변경할까요?\n'
            '저장된 기록은 유지됩니다. 작성 중인 내용은 사라집니다.', parent=app.root):
        if previous is not None:
            app.ent_daily_date.set_date(previous)
        return False
    company, site = app.cb_daily_company.get(), app.cb_daily_site.get()
    batch = getattr(app, 'inspection_batch', None)
    if batch is not None:
        batch.clear()
    app.clear_daily_usage_form_all(keep_date=True, keep_workers=True)
    fixed = getattr(app, 'fixed_vehicle_widget', None)
    if fixed is not None:
        fixed.clear_photo()
    app.cb_daily_company.set(company)
    app.cb_daily_site.set(site)
    # Keep the crew roster available, but no previous day's participation/pay.
    for i in range(1, 11):
        group = getattr(app, f'worker_group{i}', None)
        if group is None:
            continue
        group.participation_var.set(False)
        group.cb_shift.set('주간')
        group.ent_worktime.set('')
        group.ent_ot.delete(0, 'end')
        group.ent_meal.delete(0, 'end')
    if hasattr(app, 'main_memo_text'):
        app.main_memo_text.delete('1.0', 'end')
    if batch is not None:
        batch.reset_job_fields()
    app._last_daily_date = selected
    app._daily_recent_date_filter = selected
    status = getattr(app, 'daily_save_status', None)
    if status is not None:
        status.config(text=f'{selected:%Y-%m-%d} · 새 작업 입력')
    app.update_recent_entries_view()
    return True


def open_memo_date(app, selected, memo):
    """Navigate through the normal discard guard, then restore the full saved memo."""
    app.ent_daily_date.set_date(selected)
    if not change_daily_date(app, force=True):
        return False
    app.main_memo_text.delete('1.0', 'end')
    app.main_memo_text.insert('1.0', memo)
    app._daily_recent_show_all = False
    return True


def load_recent_record(app, record):
    """Replace editor contents only after explicit consent to discard current input."""
    if has_daily_input(app) and not messagebox.askyesno(
            '최근 기록 불러오기',
            '작성 중인 입력과 등록할 작업 목록을 비우고 선택한 기록을 불러올까요?\n'
            '아직 저장하지 않은 내용은 사라집니다.', parent=app.root):
        return False
    batch = getattr(app, 'inspection_batch', None)
    if batch is not None:
        batch.clear()
    app.clear_daily_usage_form_all(keep_date=True, keep_workers=True)
    app.load_daily_usage_to_form(record)
    return True
