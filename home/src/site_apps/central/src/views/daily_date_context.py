"""Switch the daily entry date without carrying the previous job forward."""
from tkinter import messagebox


def show_recent_scope(app, selected_only=False):
    """Change only the record list, never the date or unsaved editor contents."""
    app._daily_recent_date_filter = app.ent_daily_date.get_date() if selected_only else None
    app._daily_recent_show_all = not selected_only
    app.update_recent_entries_view()


def has_daily_input(app):
    data = app._get_usage_session_data()
    fields = ('applied_code', 'report_no', 'equip', 'material', 'inspection_item',
              'method', 'amount', 'unit_price', 'travel_cost', 'meal_cost',
              'test_fee', 'note', 'main_memo')
    if any(str(data.get(key, '') or '').strip() for key in fields):
        return True
    if data.get('inspection_batch'):
        return True
    if any(worker.get('time') or worker.get('ot') for worker in data.get('workers', [])):
        return True
    if any(str(value or '').strip() for row in data.get('vehicles', [])
           for key, value in row.items() if key != '_raw_mileage' or str(value) not in ('', '0')):
        return True
    fixed = getattr(app, 'fixed_vehicle_widget', None)
    if fixed is not None and any(str(value or '').strip() for value in fixed.get_data().values()):
        return True
    if any(str(value or '').strip() for value in data.get('rtk', {}).values()):
        return True
    if any(str(value or '').strip() for row in data.get('ndt_companies', [])
           for key, value in row.items() if key != '_company'):
        return True
    return any(getattr(app, f'worker_group{i}', None) is not None
               and getattr(app, f'worker_group{i}').ent_meal.get().strip()
               for i in range(1, 11))


def change_daily_date(app):
    selected = app.ent_daily_date.get_date()
    previous = getattr(app, '_last_daily_date', None)
    if selected == previous:
        return
    if has_daily_input(app) and not messagebox.askyesno(
            '작업 날짜 변경',
            '입력 내용과 등록할 작업 목록을 비우고 날짜를 변경할까요?\n'
            '저장된 기록은 유지됩니다. 작성 중인 내용은 사라집니다.', parent=app.root):
        if previous is not None:
            app.ent_daily_date.set_date(previous)
        return
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
