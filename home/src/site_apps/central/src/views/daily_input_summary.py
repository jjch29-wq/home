"""Read-only preview of the current editor; never saves or changes its values."""
import math
import re
import tkinter as tk
from tkinter import ttk

from site_apps.central.src.aux_equipment import names
from site_apps.central.src.inspection_batch import ot_amount, work_interval


def _get(app, name):
    field = getattr(app, name, None)
    return str(field.get()).strip() if field is not None else ''


def _number(text):
    try:
        value = float(str(text).replace(',', '').strip())
        return value if math.isfinite(value) and value >= 0 else None
    except (ValueError, TypeError):
        return None


def build_input_summary(app):
    warnings = []
    for attribute, title in (('cb_daily_site', '현장명'), ('cb_daily_test_method', '검사방법'),
                             ('ent_daily_inspection_item', '검사품명'), ('cb_daily_unit', '단위')):
        if not _get(app, attribute):
            warnings.append(f'{title} 미입력')
    quantity = _number(_get(app, 'ent_daily_test_amount'))
    rate = _number(_get(app, 'ent_daily_unit_price'))
    if quantity is None or quantity <= 0:
        warnings.append('검사량 확인')
    if rate is None:
        warnings.append('단가 확인')
    workers = []
    seen = set()
    ot_total = meal_total = 0.0
    ot_valid = meal_valid = True
    for index in range(1, 11):
        worker = getattr(app, f'worker_group{index}', None)
        if worker is None or not worker.participation_var.get():
            continue
        name = worker.get_worker().strip()
        if not name:
            continue
        key = re.sub(r'\s+', '', name)
        if key in seen:
            warnings.append(f'{name} 중복 선택')
        seen.add(key)
        time = worker.get_time().strip()
        if not time or work_interval(time) is None:
            warnings.append(f'{name} 작업시간 확인')
        try:
            ot = ot_amount(worker.get_ot())
        except ValueError:
            ot = None
            ot_valid = False
            warnings.append(f'{name} OT 확인')
        meal_text = worker.get_meal().strip()
        meal = _number(meal_text or '0')
        if meal is None:
            meal_valid = False
            warnings.append(f'{name} 일비 확인')
        ot_total += ot or 0
        meal_total += meal or 0
        workers.append(f"{name} · {time or '시간 미입력'} · OT {ot:,.0f}원" if ot is not None
                       else f"{name} · {time or '시간 미입력'} · OT 확인 필요")

    def amount(attribute):
        text = _get(app, attribute)
        value = _number(text)
        if text and value is None:
            warnings.append({'ent_daily_test_fee': '검사비', 'ent_daily_travel_cost': '출장비',
                             'ent_daily_meal_cost': '일식'}[attribute] + ' 확인')
        return f'{value:,.0f}원' if value is not None else ('미입력' if not text else '확인 필요')

    ot_text = f'{ot_total:,.0f}원' if ot_valid else '확인 필요'
    meal_text = f'{meal_total:,.0f}원' if meal_valid else '확인 필요'
    money = (f"검사비 {amount('ent_daily_test_fee')}  ·  출장비 {amount('ent_daily_travel_cost')}\n"
             f"OT 합계 {ot_text}  ·  작업자 일비 {meal_text}\n"
             f"일식(공통) {amount('ent_daily_meal_cost')}  ·  현재 입력값 기준")
    blocks = []
    for title, attribute in (('주 장비', 'cb_daily_equip'), ('보조장비', 'daily_aux_equipment_var')):
        selected = names(_get(app, attribute))
        blocks.append((f'{title} · {len(selected)}개', '\n'.join(selected) or '선택 없음'))
    blocks.append((f'참여 작업자 · {len(workers)}명', '\n'.join(workers) or '참여 작업자 없음'))
    status = ' / '.join(warnings) if warnings else '기본 입력 확인 완료 · 상세 조건은 작업 추가 시 확인합니다.'
    return dict(blocks=blocks, money=money, status=status, warnings=warnings)


class DailyInputSummary(ttk.LabelFrame):
    def __init__(self, parent, app, form):
        super().__init__(parent, text='현재 입력 요약', padding=8, width=400, height=250)
        self.app = app
        self.form = form
        self.grid_propagate(False)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        body = ttk.Frame(self)
        body.grid(row=0, column=0, sticky='nsew')
        self.text = tk.Text(body, width=34, height=8, wrap='word', state='disabled',
                            font=('Malgun Gothic', 10), relief='flat', padx=6, pady=4,
                            background='#F6F8FA', foreground='#273444')
        scrollbar = ttk.Scrollbar(body, command=self.text.yview)
        scrollbar.pack(side='right', fill='y')
        self.text.configure(yscrollcommand=scrollbar.set)
        self.text.pack(fill='both', expand=True)
        self.text.tag_configure('heading', font=('Malgun Gothic', 10, 'bold'), foreground='#234D70')
        self.money = ttk.Label(self, justify='left')
        self.money.grid(row=1, column=0, sticky='ew', pady=(5, 2))
        self.status = ttk.Label(self, justify='left')
        self.status.grid(row=2, column=0, sticky='ew')
        self._snapshot = None
        self._timer = None
        parent.bind('<Configure>', self.layout, add='+')
        self.bind('<Configure>', self._wrap_labels, add='+')
        self.bind('<Map>', lambda event: self.refresh(), add='+')
        self.bind('<Destroy>', self._destroy, add='+')
        self.after_idle(self.layout)
        self.refresh()

    def layout(self, event=None):
        available = self.master.winfo_width()
        if available < 2:
            return
        if available < self.form.winfo_reqwidth() + 380:
            self.grid_remove()
        else:
            self.grid(row=0, column=1, sticky='nsew', padx=(10, 8), pady=8)
        self._wrap_labels()

    def _wrap_labels(self, event=None):
        wrap = max(220, self.winfo_width() - 24)
        self.money.configure(wraplength=wrap)
        self.status.configure(wraplength=wrap)

    def refresh(self):
        if self._timer is not None:
            self.after_cancel(self._timer)
            self._timer = None
        if self.winfo_ismapped():
            data = build_input_summary(self.app)
            if data != self._snapshot:
                self.text.configure(state='normal')
                self.text.delete('1.0', 'end')
                for title, content in data['blocks']:
                    self.text.insert('end', title + '\n', 'heading')
                    self.text.insert('end', content + '\n\n')
                self.text.configure(state='disabled')
                self.text.yview_moveto(0)
                self.money.configure(text=data['money'])
                self.status.configure(text=data['status'], foreground='#A15300' if data['warnings'] else '#28733B')
                self._snapshot = data
        self._timer = self.after(300, self.refresh)

    def _destroy(self, event):
        if event.widget is self and self._timer is not None:
            self.after_cancel(self._timer)
            self._timer = None
