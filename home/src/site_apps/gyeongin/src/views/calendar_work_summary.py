"""Read-only daily work previews for the entry calendar."""
import datetime
import json
import tkinter as tk
from tkinter import ttk

import pandas as pd


def clean(value):
    if value is None or pd.isna(value):
        return ''
    return str(value).strip()


def short(value, limit=90):
    text = clean(value).replace('\n', ' ')
    return text if len(text) <= limit else text[:limit] + '…'


def daily_summary(day, site, records, include_times=False):
    lines = [f'{day:%Y-%m-%d} · {site or "전체 현장"} · 저장된 기록']
    if not records:
        return '\n'.join(lines + ['저장된 작업 없음'])
    jobs, workers, vehicles, memos, times = [], [], [], [], []
    for row in records:
        if clean(row.get('MaterialID')) != '차량점검' and clean(row.get('User')) != '차량점검':
            method = clean(row.get('검사방법'))
            item = clean(row.get('검사품명'))
            if method or item:
                qty = row.get('검사량', row.get('Usage', ''))
                try:
                    qty = f'{float(qty):g}' if pd.notna(qty) else ''
                except (TypeError, ValueError):
                    qty = clean(qty)
                detail = ' · '.join(filter(None, [method, clean(row.get('관경(Inch)')),
                    item, clean(row.get('검사구분')), f'{qty}{clean(row.get("Unit"))}']))
                if not site:
                    detail = f'{clean(row.get("Site"))}: {detail}'
                jobs.append(short(detail))
        names = []
        raw = row.get('작업별인원')
        if isinstance(raw, str) and raw:
            try:
                people = json.loads(raw)
                if isinstance(people, list):
                    for person in people:
                        if isinstance(person, dict):
                            names.extend(clean(v) for k, v in person.items()
                                         if k == 'User' or (k.startswith('User') and k[4:].isdigit()))
            except (TypeError, ValueError):
                pass
        names.extend(clean(row.get('User' + (str(i) if i > 1 else ''))) for i in range(1, 11))
        if include_times:
            people = [row]
            try:
                original = json.loads(raw) if isinstance(raw, str) and raw else []
                if isinstance(original, list):
                    people += [person for person in original if isinstance(person, dict)]
            except (TypeError, ValueError):
                pass
            for person in people:
                for i in range(1, 11):
                    suffix = str(i) if i > 1 else ''
                    name = clean(person.get('User' + suffix))
                    time = clean(person.get('WorkTime' + suffix))
                    detail = f'{name}: {time}'
                    if name and name != '차량점검' and time and detail not in times:
                        times.append(detail)
        for name in names:
            if name and name != '차량점검' and name not in workers:
                workers.append(name)
        number = clean(row.get('차량번호'))
        if number:
            mileage = clean(row.get('주행거리'))
            check = '점검 기록 있음' if clean(row.get('차량점검')) else '점검 내용 없음'
            vehicle = f'{number} · {mileage + "km · " if mileage else ""}{check}'
            if vehicle not in vehicles:
                vehicles.append(vehicle)
        memo = short(row.get('DailyMemo'))
        if memo and memo not in memos:
            memos.append(memo)
    lines += [f'검사 작업 {len(jobs)}건'] + jobs[:6]
    if len(jobs) > 6:
        lines.append(f'외 {len(jobs) - 6}건')
    if workers:
        lines.append('참여자: ' + short(', '.join(workers), 120))
    lines += ['근무시간: ' + short(time) for time in times[:4]]
    if len(times) > 4:
        lines.append(f'근무시간 외 {len(times) - 4}건')
    lines += ['차량: ' + short(vehicle) for vehicle in vehicles[:3]]
    if len(vehicles) > 3:
        lines.append(f'차량 외 {len(vehicles) - 3}건')
    lines += ['메모: ' + memo for memo in memos[:2]]
    return '\n'.join(lines)


def work_log_summary(data):
    """Describe saved report inputs without implying that Excel was exported."""
    if data is None:
        return '작업일보: 저장 이력 없음'
    lines = ['작업일보: 저장됨 (출력 여부는 별도)']
    weather = short(data.get('weather'))
    if weather:
        lines.append('날씨: ' + weather)
    jobs = []
    for row in data.get('ndt_results', []):
        if not isinstance(row, dict) or not clean(row.get('검사방법')):
            continue
        parts = [clean(row.get(key)) for key in
                 ('검사방법', '관경', '구간', '라인번호', '근무구분', '결과')]
        quantities = [f'{key} {clean(row.get(key))}' for key in ('RT_OR', 'RT_RE', 'PAUT', 'MT', 'PT')
                      if clean(row.get(key)) not in ('', '0', '0.0')]
        jobs.append(short(' · '.join(filter(None, parts + quantities))))
    lines += [f'일보 검사 내역 {len(jobs)}건'] + jobs[:5]
    if len(jobs) > 5:
        lines.append(f'외 {len(jobs) - 5}건')
    personnel = data.get('personnel_data', {})
    counts = [f'{key.replace("_", " ")} {clean(personnel.get(key))}명'
              for key in ('검사원_인원', '검사원_현장대리인', '안전_인원', '안전_현장대리인')
              if clean(personnel.get(key)) not in ('', '0', '0.0')]
    if counts:
        lines.append('인원: ' + ', '.join(counts))
    remarks = short(data.get('remarks'))
    if remarks:
        lines.append('일보 비고: ' + remarks)
    return '\n'.join(lines)


class CalendarWorkSummary:
    def __init__(self, entry, app, site_getter=None, frame_getter=None, history_getter=None):
        self.app = app
        self.entry = entry
        self.site_getter = site_getter
        self.frame_getter = frame_getter
        self.history_getter = history_getter
        # DateEntry exposes no public accessor for its popup Calendar.
        self.calendar = entry._calendar
        self.event_ids = []
        self.hover_timer = None
        self.hover_popup = None
        self.hover_widget = None
        self.hover_texts = {}
        self.bound_hover_cells = set()
        self.popup_parent = entry._top_cal
        self.calendar.configure(tooltipdelay=450, tooltipalpha=1.0,
                                tooltipbackground='#fffbe6', tooltipforeground='#202020')
        self.calendar.tooltip_wrapper.configure(wraplength=550, justify='left')
        self.calendar.bind('<Map>', self.refresh, add='+')
        self.calendar.bind('<<CalendarMonthChanged>>', self.refresh, add='+')
        entry._top_cal.bind('<Unmap>', self.hide, add='+')
        entry._top_cal.bind('<Escape>', self.close_calendar, add='+')
        entry.bind('<Escape>', self.close_calendar, add='+')

    def close_calendar(self, event=None):
        if not self.popup_parent.winfo_ismapped():
            return
        self.hide()
        self.popup_parent.withdraw()
        self.entry.focus_set()
        return 'break'

    def hide(self, event=None):
        # Toplevel bindings also receive child events. A tooltip's own Unmap
        # must not cancel the next day's preview.
        if event is not None and event.widget is not self.popup_parent:
            return
        if getattr(self, 'hover_timer', None) is not None:
            self.calendar.after_cancel(self.hover_timer)
            self.hover_timer = None
        self.hover_widget = None
        if getattr(self, 'hover_popup', None) is not None:
            self.hover_popup.withdraw()
        wrapper = self.calendar.tooltip_wrapper
        wrapper.current_widget = None
        wrapper.tooltip.withdraw()

    def queue_hover(self, event):
        if event.widget is getattr(self, 'hover_widget', None):
            return
        self.hide()
        self.hover_widget = event.widget
        self.hover_timer = self.calendar.after(350, self.show_hover)

    def show_hover(self):
        self.hover_timer = None
        widget = self.hover_widget
        text = self.hover_texts.get(widget)
        if not text or not self.calendar.winfo_ismapped():
            return
        if self.hover_popup is None:
            self.hover_popup = tk.Toplevel(self.popup_parent)
            self.hover_popup.withdraw()
            self.hover_popup.overrideredirect(True)
            self.hover_popup.attributes('-topmost', True)
            self.hover_popup.bind('<Escape>', self.close_calendar, add='+')
            self.hover_label = ttk.Label(self.hover_popup, padding=10, justify='left',
                wraplength=520, background='#fffbe6', foreground='#202020',
                relief='solid', borderwidth=1)
            self.hover_label.pack()
        self.hover_label.config(text=text)
        self.hover_popup.update_idletasks()
        x = min(widget.winfo_rootx() + widget.winfo_width() + 12,
                widget.winfo_screenwidth() - self.hover_popup.winfo_reqwidth() - 10)
        y = min(widget.winfo_rooty() + 12,
                widget.winfo_screenheight() - self.hover_popup.winfo_reqheight() - 10)
        self.hover_popup.geometry(f'+{max(0, x)}+{max(0, y)}')
        self.hover_popup.deiconify()
        self.hover_popup.lift()

    def bind_hover_cells(self, previews):
        # Use additive cell bindings, independent of tkcalendar's tooltip
        # bindings, which are replaced whenever its day styles are refreshed.
        self.calendar.tooltip_wrapper.remove_all()
        self.hover_texts = {}
        for day, text in previews.items():
            week, weekday = self.calendar._get_day_coords(day)
            if week is None:
                continue
            cell = self.calendar._calendar[week][weekday]
            if not cell.cget('text'):
                continue
            self.hover_texts[cell] = text
            if cell not in self.bound_hover_cells:
                cell.bind('<Enter>', self.queue_hover, add='+')
                cell.bind('<Motion>', self.queue_hover, add='+')
                cell.bind('<Leave>', lambda event: self.hide(), add='+')
                cell.bind('<ButtonPress-1>', lambda event: self.hide(), add='+')
                self.bound_hover_cells.add(cell)

    def refresh(self, event=None, history=None):
        self.hide()
        if self.event_ids:
            self.calendar.calevent_remove(*self.event_ids)
        self.event_ids = []
        month, year = self.calendar.get_displayed_month()
        # Include adjacent-month cells; tkcalendar may show the preceding week.
        start = datetime.date(year, month, 1) - datetime.timedelta(days=7)
        end = start + datetime.timedelta(days=49)
        site_getter = getattr(self, 'site_getter', None)
        frame_getter = getattr(self, 'frame_getter', None)
        history_getter = getattr(self, 'history_getter', None)
        site = site_getter() if site_getter else self.app.cb_daily_site.get().strip()
        frame = frame_getter() if frame_getter else getattr(self.app, 'daily_usage_df', None)
        if not history_getter:
            self.calendar.tag_config('saved_daily_work', background='#2E7D32', foreground='white')
        if history_getter:
            if history is None:
                history = history_getter()
            self.calendar.tag_config('saved_work_log', background='#2E7D32', foreground='white')
        grouped = {}
        if frame is not None and not frame.empty and 'Date' in frame:
            dates = pd.to_datetime(frame['Date'], errors='coerce').dt.date
            selected = dates.notna() & (dates >= start) & (dates < end)
            if site:
                selected &= frame['Site'].fillna('').astype(str).str.strip().eq(site) if 'Site' in frame else False
            for day, row in zip(dates[selected], frame.loc[selected].to_dict('records')):
                grouped.setdefault(day, []).append(row)
        previews = {}
        for offset in range(49):
            day = start + datetime.timedelta(days=offset)
            text = daily_summary(day, site, grouped.get(day, []), include_times=bool(history_getter))
            tags = []
            if not history_getter and any(
                    clean(row.get('MaterialID')) != '차량점검'
                    and clean(row.get('User')) != '차량점검'
                    and (clean(row.get('검사방법')) or clean(row.get('검사품명')))
                    for row in grouped.get(day, [])):
                tags = ['saved_daily_work']
            if history_getter:
                saved = history.get(day.isoformat())
                text = work_log_summary(saved) + '\n\n' + text
                if saved is not None:
                    tags = ['saved_work_log']
            self.event_ids.append(self.calendar.calevent_create(day, text, tags))
            previews[day] = text
        # Event tags overwrite day styles; restore the existing selection style
        # without selecting a date or navigating away from the displayed month.
        self.calendar._display_selection()
        if hasattr(self, 'bound_hover_cells'):
            self.bind_hover_cells(previews)
