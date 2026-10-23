"""Stage independent inspections and commit their daily records together."""
from copy import deepcopy
import json
import math
import re
from tkinter import ttk, messagebox

import pandas as pd


def number(value, label, required=False):
    text = str(value).strip().replace(',', '')
    try:
        result = float(text or ('nan' if required else '0'))
        if not math.isfinite(result) or result < 0:
            raise ValueError
        return result
    except ValueError:
        raise ValueError(f'{label}: 0 이상의 숫자를 입력하세요.')


def json_copy(value):
    def normalize(obj):
        if isinstance(obj, dict):
            return {key: normalize(item) for key, item in obj.items()}
        if isinstance(obj, (list, tuple)):
            return [normalize(item) for item in obj]
        if obj is None or isinstance(obj, (str, bool, int)):
            return obj
        # json's default hook is never called for float NaN. Normalize nested
        # Excel/pandas missing values before encoding (including vehicle data).
        if pd.isna(obj):
            return None
        if isinstance(obj, float):
            if not math.isfinite(obj):
                raise ValueError('작업 데이터에 무한대 값이 있습니다. 입력값을 확인하세요.')
            return obj
        if hasattr(obj, 'isoformat'):
            return obj.isoformat()
        if hasattr(obj, 'item'):
            return normalize(obj.item())
        raise TypeError(type(obj).__name__)
    return json.loads(json.dumps(normalize(value), ensure_ascii=False, allow_nan=False))


def worker_fields(index):
    suffix = '' if index == 1 else str(index)
    return [prefix + suffix for prefix in ('User', 'WorkTime', 'OT', 'Meal')]


def participants(plan):
    data = plan['common_data']
    return [{key: data.get(key, '') for key in worker_fields(i)} for i in range(1, 11)]


def participation_record(record):
    """Restore original job inputs without changing the stored accounting fields."""
    result = record.copy()
    raw = record.get('작업별인원')
    if isinstance(raw, str) and raw:
        for person in json.loads(raw):
            result.update(person)
    return result


def work_interval(value):
    times = re.findall(r'(\d{1,2}):(\d{2})', str(value))
    if len(times) != 2:
        return None
    start, end = [int(h) * 60 + int(m) for h, m in times]
    if any(int(m) >= 60 for _, m in times):
        return None
    if end < start or ('익일' in str(value) and end < 1440):
        end += 1440
    return (start, end) if end > start else None


def ot_amount(value):
    text = str(value or '').strip()
    if not text:
        return 0.0
    match = re.search(r'\(([\d,]+)\s*원\)', text)
    return number(match.group(1) if match else text, '작업자 OT 금액')


def allocate_labor(records, date, ot_calculator=None):
    """Assign each worker's time union once; keep source participation separately."""
    people = {}
    meals = {}
    for record in records:
        for i in range(1, 11):
            user, time, ot, meal = worker_fields(i)
            name = str(record.get(user, '') or '').strip()
            if not name:
                record[time] = record[ot] = record[meal] = ''
                continue
            key = re.sub(r'\s+', '', name)
            interval = work_interval(record.get(time, ''))
            amount = ot_amount(record.get(ot, ''))
            if record.get(time) and interval is None:
                raise ValueError(f'{name}: 작업시간을 시작~종료 형식으로 입력하세요.')
            meals[key] = max(meals.get(key, 0), number(record.get(meal, ''), '작업자 일비'))
            record[meal] = ''
            people.setdefault(key, []).append((interval, amount, record, time, ot, meal))
    for key, items in people.items():
        # Highest OT first for identical ranges, so repeated daily amounts count once.
        items.sort(key=lambda x: (x[0][0] if x[0] else float('inf'),
                                  -(x[0][1] if x[0] else 0), -x[1]))
        end_seen = -1
        for interval, amount, record, time, ot, meal in items:
            if interval is None:
                if amount:
                    raise ValueError('OT 금액이 있는 작업자는 작업시간도 입력하세요.')
                continue
            start, end = interval
            remaining_start = max(start, end_seen)
            end_seen = max(end_seen, end)
            if remaining_start >= end:
                record[time] = record[ot] = ''
            elif remaining_start > start:
                def clock(minutes):
                    return f'{minutes // 60:02d}:{minutes % 60:02d}'
                prefix = re.match(r'\s*(\([^)]*\)\s*)?', str(record[time])).group(0)
                original = record[time]
                record[time] = f'{prefix}{clock(remaining_start)}~{clock(end)}'
                ratio = (end - remaining_start) / (end - start)
                if ot_calculator:
                    _, full_cost = ot_calculator(original, pd.to_datetime(date))
                    _, remaining_cost = ot_calculator(record[time], pd.to_datetime(date))
                    if full_cost > 0:
                        ratio = max(0, min(1, remaining_cost / full_cost))
                record[ot] = str(round(amount * ratio)) if amount else ''
        items[0][2][items[0][5]] = str(meals[key]) if meals[key] else ''
    return sum(meals.values())


def combine_plans(plans, shared, ot_calculator=None):
    """Preserve each inspection's crew, sharing only daily site/vehicle expenses."""
    result = deepcopy(plans)
    first = True
    labor_records = []
    for plan in result:
        common = plan['common_data']
        for key in ('Date', 'Site', '업체명'):
            common[key] = shared[key]
        crew = participants(plan)
        names = [str(person.get(worker_fields(i)[0], '') or '').strip() for i, person in enumerate(crew, 1)]
        plan['workers_list'] = [name for name in names if name]
        plan['all_workers'] = ', '.join(plan['workers_list'])
        for row_index, record in enumerate(plan['records_to_save']):
            for key in ('Date', 'Site', '업체명'):
                record[key] = shared[key]
            for key in list(record):
                if re.fullmatch(r'(User|WorkTime|OT|Meal)\d*', key) or key in (
                    '차량번호', '주행거리', '차량점검', '차량비고', 'DailyMemo', 'OT시간', 'OT금액'):
                    record[key] = ''
            record['출장비'] = record['일식'] = 0.0
            if row_index == 0:
                for person in crew:
                    record.update(person)
                record['작업별인원'] = json.dumps(crew, ensure_ascii=False)
                record['참여인원'] = plan['all_workers']
                record['참여시간'] = ' / '.join(f"{names[i-1]}: {person.get(worker_fields(i)[1], '')}" for i, person in enumerate(crew, 1) if names[i-1])
                record['참여OT'] = ' / '.join(f"{names[i-1]}: {person.get(worker_fields(i)[2], '') or '0'}" for i, person in enumerate(crew, 1) if names[i-1])
                labor_records.append(record)
            if first:
                record.update({key: value for key, value in shared.items()
                               if key not in ('Date', 'Site', '업체명') and not re.fullmatch(r'(User|WorkTime|OT|Meal)\d*', key)})
                record['검사비'] += shared['출장비']
                first = False
    meals = allocate_labor(labor_records, shared['Date'], ot_calculator)
    if labor_records:
        labor_records[0]['일식'] = meals if meals else shared.get('일식', 0)
    return result


def commit_plans(app, plans):
    """One database save; restore in-memory records and stock on failure."""
    frame_names = ('daily_usage_df', 'transactions_df', 'materials_df', 'monthly_usage_df')
    frames = {name: getattr(app, name).copy(deep=True) for name in frame_names}
    list_names = ('sites', 'users', 'vehicles', 'daily_units', 'equipments')
    lists = {name: list(getattr(app, name)) for name in list_names}
    try:
        count = sum(app._commit_usage_plan(plan, auto_save=False) for plan in plans)
        if not app.save_data():
            raise RuntimeError('일괄 저장하지 못했습니다. 입력 목록은 유지됩니다. 파일 상태를 확인한 뒤 다시 저장하세요.')
        return count
    except Exception:
        for name, value in frames.items():
            setattr(app, name, value)
        for name, value in lists.items():
            getattr(app, name)[:] = value
        raise


class InspectionBatch(ttk.LabelFrame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=5)
        self.app = app
        self.plans = []
        self.editing = None
        bar = ttk.Frame(self)
        bar.pack(fill='x')
        ttk.Button(bar, text='선택 작업 수정', command=self.edit).pack(side='left', padx=3)
        ttk.Button(bar, text='선택 삭제', command=self.remove).pack(side='left')
        ttk.Button(bar, text='목록 비우기', command=self.clear).pack(side='left', padx=3)
        self.summary = ttk.Label(bar)
        self.summary.pack(side='left', padx=8)
        ttk.Button(bar, text='사용 안내', command=self.show_help).pack(side='right')
        cols = ('method', 'pipe', 'item', 'material', 'quantity', 'unit', 'price', 'fee', 'condition', 'workers', 'times', 'ots')
        area = ttk.Frame(self)
        area.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(area, columns=cols, show='headings', height=5, selectmode='extended')
        for key, title, width in zip(cols, ('방법', '관경', '검사품명', '품목/장비', '검사량', '단위', '단가', '검사비', '계산조건', '참여 인원', '참여 시간', '참여 OT'),
                                     (65, 75, 100, 140, 90, 75, 90, 110, 170, 160, 230, 150)):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=45, stretch=True)
        self.tree.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(area, orient='vertical', command=self.tree.yview)
        scroll.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scroll.set)
        horizontal = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        horizontal.pack(fill='x')
        self.tree.configure(xscrollcommand=horizontal.set)
        save_bar = ttk.Frame(getattr(app, 'daily_save_bar', self))
        save_bar.pack(fill='x', pady=(6, 0))
        self.footer_summary = ttk.Label(save_bar)
        self.footer_summary.pack(side='left')
        app.btn_daily_save = ttk.Button(save_bar, command=app.add_daily_usage_entry, style='Action.TButton')
        app.btn_daily_save.pack(side='right')
        self.tree.bind('<Double-1>', lambda event: self.edit())
        self.render()

    def serialize(self):
        return json_copy(self.plans)

    def show_help(self):
        messagebox.showinfo('작업 목록 사용 안내',
            '방법·관경별 입력 → 자동 산출 또는 단가 입력 → 작업 목록에 추가 → 목록 확인 → 저장\n\n'
            '참여 체크된 인원·시간·OT만 작업별 보관합니다. 체크를 해제해도 명단은 유지됩니다.\n'
            '겹치는 시간·OT는 중복 집계하지 않습니다. 원래 입력값은 참여 인원·시간·OT 열에서 확인하세요.',
            parent=self)

    def restore(self, plans):
        # Sessions contain staged work, never executable commands.
        restored = json_copy(plans)
        if not isinstance(restored, list) or any(not isinstance(p, dict) or
                not isinstance(p.get('common_data'), dict) or not p.get('records_to_save')
                for p in restored):
            raise ValueError('일괄 작업 세션 형식이 올바르지 않습니다.')
        self.plans = restored
        self.editing = None
        self.render()

    def snapshot(self):
        data = self.app._get_usage_session_data()
        data.pop('inspection_batch', None)
        data['ndt_vars'] = {key: getattr(self.app, key).get() for key in (
            'ndt_work_time_var', 'ndt_source_var', 'ndt_thickness_var', 'ndt_pipe_var',
            'ndt_overhead_var', 'ndt_tech_var', 'ndt_ori_joint_var', 'ndt_ori_qty_var',
            'ndt_rep_joint_var', 'ndt_rep_qty_var', 'ndt_rej_joint_var', 'ndt_report_pipe_var')}
        data['ndt_costs'] = {key: getattr(self.app, key, 0) for key in (
            '_last_ndt_factor', '_last_ndt_adj_qty', '_last_ndt_mat_cost',
            '_last_ndt_lab_cost', '_last_ndt_overhead', '_last_ndt_tech')}
        data['worker_meals'] = {str(i): getattr(self.app, f'worker_group{i}').get_meal().strip()
                                for i in range(1, 11) if hasattr(self.app, f'worker_group{i}')}
        return data

    def add(self):
        app = self.app
        try:
            site = app.cb_daily_site.get().strip()
            method = app.cb_daily_test_method.get().strip().upper()
            if not site or not method:
                raise ValueError('현장명과 검사방법을 입력하세요.')
            if not app.ent_daily_inspection_item.get().strip():
                raise ValueError('검사품명을 입력하세요.')
            qty = number(app.ent_daily_test_amount.get(), '검사량', True)
            rate = number(app.ent_daily_unit_price.get(), '단가', True)
            if qty <= 0 or not app.cb_daily_unit.get().strip():
                raise ValueError('검사량은 0보다 커야 하며 단위를 입력해야 합니다.')
            number(qty * rate, '검사비')
            for key in ('ndt_ori_joint_var', 'ndt_ori_qty_var', 'ndt_rep_joint_var', 'ndt_rep_qty_var', 'ndt_rej_joint_var'):
                number(getattr(app, key).get(), 'ORI/REP/REJ 수량')
            if app.ndt_loc_type_var.get() == '핫태핑':
                raise ValueError('핫태핑은 기존 단건 저장을 이용하세요. 이 목록은 일반 검사 작업용입니다.')
            if method != 'RT' and not app.ndt_rep_joint_var.get().strip():
                app.ndt_ori_qty_var.set(str(qty))
            # Resolve only known material IDs here; staging must never register or save stock.
            display = app.cb_daily_material.get().strip()
            mat_id = display
            for _, row in app.materials_df.iterrows():
                if app.get_material_display_name(row['MaterialID']) == display:
                    mat_id = row['MaterialID']
                    break
            if display and mat_id == display and app._is_consumable_material(display, method):
                raise ValueError('소모품은 품목 검색에서 등록된 자재를 선택하세요. 그래야 일괄 저장 시 재고가 차감됩니다.')
            plan = app._add_single_usage_record_logic(mat_id, app.ent_daily_date.get_date(), site,
                                                     auto_save=False, prepare_only=True)
            seen_workers = set()
            for i in range(1, 11):
                user, time, ot, meal = worker_fields(i)
                name = str(plan['common_data'].get(user, '') or '').strip()
                if not name:
                    continue
                normalized = re.sub(r'\s+', '', name)
                if normalized in seen_workers:
                    raise ValueError(f'{name}: 같은 작업의 참여 인원에 중복 입력되어 있습니다.')
                seen_workers.add(normalized)
                time_value = plan['common_data'].get(time, '')
                amount = ot_amount(plan['common_data'].get(ot, ''))
                number(plan['common_data'].get(meal, ''), '작업자 일비')
                if (time_value or amount) and work_interval(time_value) is None:
                    raise ValueError(f'{name}: 작업시간을 시작~종료 형식으로 입력하세요.')
            primary = next((r for r in plan['records_to_save'] if r.get('검사구분') != 'REP'), plan['records_to_save'][0])
            primary['검사비'] = int(qty * rate)
            # PAUT/UT never carry chemicals or RT retake counts from a previously selected method.
            if app.cb_daily_test_method.get().strip().upper() in ('PAUT', 'UT'):
                for record in plan['records_to_save']:
                    for key in record:
                        if key.startswith(('NDT_', 'RTK_')):
                            record[key] = 0.0
                for company in plan['company_data_list']:
                    company['ndt_data'] = {}
            plan['editor'] = self.snapshot()
            plan['material_display'] = display
            plan = json_copy(plan)
            if self.editing is None:
                self.plans.append(plan)
            else:
                self.plans[self.editing] = plan
                self.editing = None
            self.render()
            self.reset_job_fields()
        except Exception as exc:
            messagebox.showwarning('작업 담기', str(exc))

    def edit(self):
        selected = self.tree.selection()
        if len(selected) != 1:
            messagebox.showinfo('작업 수정', '수정할 작업 하나를 선택하세요.')
            return
        index = int(selected[0])
        data = deepcopy(self.plans[index]['editor'])
        current = self.app._get_usage_session_data()
        for key in ('date', 'site', 'company', 'travel_cost', 'meal_cost', 'vehicles', 'main_memo'):
            data[key] = current[key]
        self.app.cb_daily_test_method.set(data['method'])
        self.app.cb_daily_test_method.event_generate('<<ComboboxSelected>>')
        self.app._set_usage_session_data(data)
        for i in range(1, 11):
            group = getattr(self.app, f'worker_group{i}', None)
            if group is not None and hasattr(group, 'ent_meal'):
                group.ent_meal.delete(0, 'end')
                group.ent_meal.insert(0, data.get('worker_meals', {}).get(str(i), ''))
        for key, value in data['ndt_vars'].items():
            getattr(self.app, key).set(value)
        for key, value in data['ndt_costs'].items():
            setattr(self.app, key, value)
        # Diameter traces may recalculate: preserve explicitly entered quantities/prices.
        for name, key in (('ent_daily_test_amount', 'amount'), ('ent_daily_unit_price', 'unit_price'),
                          ('ent_daily_test_fee', 'test_fee')):
            widget = getattr(self.app, name)
            widget.delete(0, 'end')
            widget.insert(0, data[key])
        # Method selection changes unit; restore the saved unit/equipment afterwards.
        self.app.cb_daily_unit.set(data['unit'])
        self.app.cb_daily_equip.delete(0, 'end')
        self.app.cb_daily_equip.insert(0, data['equip'])
        self.editing = index
        self.render()
        self.app.entry_canvas.yview_moveto(0)

    def reset_job_fields(self):
        """Leave date/site/workers/expenses/vehicles available for the next job."""
        app = self.app
        for name in ('ent_daily_test_amount', 'ent_daily_unit_price', 'ent_daily_test_fee',
                     'ent_daily_report_no', 'ent_daily_note'):
            getattr(app, name).delete(0, 'end')
        for name in ('ndt_ori_joint_var', 'ndt_ori_qty_var', 'ndt_rep_joint_var',
                     'ndt_rep_qty_var', 'ndt_rej_joint_var', 'ndt_report_pipe_var'):
            getattr(app, name).set('')
        # Quantity traces can write zero while the joint fields are being cleared.
        for name in ('ent_daily_test_amount', 'ent_daily_unit_price', 'ent_daily_test_fee'):
            getattr(app, name).delete(0, 'end')
        for name in ('_last_ndt_factor', '_last_ndt_adj_qty', '_last_ndt_mat_cost',
                     '_last_ndt_lab_cost', '_last_ndt_overhead', '_last_ndt_tech'):
            setattr(app, name, 1.0 if name == '_last_ndt_factor' else 0)
        for company in getattr(app, 'ndt_company_entries', []):
            for key, widget in company.items():
                if key != '_company':
                    widget.delete(0, 'end')
        for key, widget in app.rtk_entries.items():
            if key != '총계':
                widget.delete(0, 'end')
        app.calculate_rtk_total()

    def remove(self):
        for index in sorted((int(i) for i in self.tree.selection()), reverse=True):
            self.plans.pop(index)
        self.editing = None
        self.render()

    def clear(self):
        self.plans.clear()
        self.editing = None
        self.render()

    def render(self):
        self.tree.delete(*self.tree.get_children())
        total = 0
        for i, plan in enumerate(self.plans):
            data = plan['common_data']
            fee = sum(r['검사비'] for r in plan['records_to_save'])
            total += fee
            crew = participants(plan)
            names = [str(p.get(worker_fields(i)[0], '') or '').strip() for i, p in enumerate(crew, 1)]
            times = ' / '.join(f"{names[i-1]}: {p.get(worker_fields(i)[1], '')}" for i, p in enumerate(crew, 1) if names[i-1])
            ots = ' / '.join(f"{names[i-1]}: {p.get(worker_fields(i)[2], '') or '0'}" for i, p in enumerate(crew, 1) if names[i-1])
            self.tree.insert('', 'end', iid=str(i), values=(data['검사방법'], data.get('관경(Inch)', ''),
                data['검사품명'], plan.get('material_display', ''), data['검사량'], data['Unit'],
                f"{data['단가']:,.0f}", f'{fee:,.0f}', f"{data.get('조건1', '')} / {data.get('조건2', '')}",
                ', '.join(n for n in names if n), times, ots))
        suffix = ' · 수정 중: 작업 담기로 반영' if self.editing is not None else ''
        self.summary.config(text=f'{len(self.plans)}개 작업 / 검사비 {total:,.0f}원 (출장비 별도){suffix}')
        footer = getattr(self, 'footer_summary', None)
        if footer is not None:
            footer.config(text=f'{len(self.plans)}개 작업 · 검사비 합계 {total:,.0f}원 (출장비 별도)'
                          if self.plans else '목록이 비어 있으면 현재 입력 1건을 저장합니다.')
        self.app.btn_daily_save.config(text='③ 목록 전체 저장' if self.plans else '③ 현재 입력 저장', width=18)
        add_button = getattr(self.app, 'btn_daily_add', None)
        if add_button is not None:
            add_button.config(text='수정 내용 반영' if self.editing is not None else '＋ 작업 목록에 추가')

    def shared(self):
        app = self.app
        shared = {'Date': app.ent_daily_date.get_date(), 'Site': app.cb_daily_site.get().strip(),
                  '업체명': app.cb_daily_company.get().strip(),
                  '출장비': number(app.ent_daily_travel_cost.get(), '출장비'),
                  '일식': number(app.ent_daily_meal_cost.get(), '일식'),
                  'DailyMemo': app.main_memo_text.get('1.0', 'end-1c') if hasattr(app, 'main_memo_text') else ''}
        if not shared['Site']:
            raise ValueError('공통 현장명을 입력하세요.')
        boxes = list(getattr(app, 'vehicle_boxes', []))
        fixed = getattr(app, 'fixed_vehicle_widget', None)
        if fixed is not None and fixed not in boxes:
            boxes.insert(0, fixed)
        vehicle = {key: [] for key in ('차량번호', '주행거리', '차량점검', '차량비고')}
        for box in boxes:
            if not box.winfo_exists():
                continue
            data = box.get_data()
            values = (data.get('vehicle_info', ''), data.get('_raw_mileage', data.get('mileage', '')),
                      '|'.join(f'{k}:{v}' for k, v in data.items() if v and k not in
                               ('vehicle_info', 'mileage', 'remarks', '_raw_mileage')), data.get('remarks', ''))
            if any(values):
                for key, value in zip(vehicle, values):
                    vehicle[key].append(str(value))
        shared.update({key: ' || '.join(values) for key, values in vehicle.items()})
        return shared

    def save(self):
        if self.editing is not None:
            messagebox.showwarning('작업 수정 중', '작업 담기를 눌러 수정 내용을 반영한 뒤 저장하세요.')
            return
        try:
            shared = self.shared()
            if self.app.ent_daily_test_amount.get().strip():
                raise ValueError('현재 입력 중인 작업이 있습니다. 작업 담기를 눌러 목록에 추가한 뒤 저장하세요.')
            plans = combine_plans(self.plans, shared, getattr(self.app, '_calculate_ot_from_worktime', None))
            # Warn on exact repeats without deleting any existing daily/stock records.
            existing = self.app.daily_usage_df
            import pandas as pd
            keys = ('Site', '검사방법', '검사품명', 'MaterialID', '관경(Inch)', '검사구분', '회사코드', 'Usage', '단가', '조건1', '조건2')
            def signature(record):
                date = pd.to_datetime(record.get('Date'), errors='coerce')
                if pd.isna(date):
                    return None
                values = []
                for key in keys:
                    value = record.get(key, '')
                    if pd.isna(value):
                        value = ''
                    if key in ('Usage', '단가'):
                        value = float(value or 0)
                    elif key == 'MaterialID':
                        try:
                            value = float(value)
                        except (ValueError, TypeError):
                            pass
                    values.append(str(value).strip())
                return (str(date.date()),) + tuple(values)
            seen = {signature(r) for r in existing.to_dict('records')} if not existing.empty else set()
            duplicate = False
            for plan in plans:
                for record in plan['records_to_save']:
                    sig = signature(record)
                    duplicate |= sig in seen
                    seen.add(sig)
            if duplicate and not messagebox.askyesno('중복 작업 확인', '같은 날짜·방법·관경·수량의 작업이 있습니다. 별도 기록으로 추가 저장하시겠습니까?'):
                return
            count = commit_plans(self.app, plans)
        except Exception as exc:
            messagebox.showerror('일괄 저장', str(exc))
            return
        # Persistence succeeded: never retain a retryable batch after this point.
        self.clear()
        try:
            self.app.save_tab_config()
            self.app.clear_daily_usage_form_all(keep_date=True, keep_workers=True)
            for field in ('ent_daily_start_date', 'ent_daily_end_date'):
                widget = getattr(self.app, field)
                if (field.endswith('start_date') and shared['Date'] < widget.get_date()) or (
                        field.endswith('end_date') and shared['Date'] > widget.get_date()):
                    widget.set_date(shared['Date'])
            for method in ('update_daily_usage_view', 'update_recent_entries_view', 'refresh_inquiry_filters', 'update_stock_view'):
                getattr(self.app, method)()
        except Exception as exc:
            messagebox.showwarning('저장 완료 · 화면 갱신 확인', f'{count}건은 저장되었습니다. 화면을 다시 조회하세요.\n{exc}')
            return
        messagebox.showinfo('일괄 저장 완료', f'{len(plans)}개 작업, {count}건의 기록을 저장했습니다.')
