"""One embedded editor per vehicle, with lossless tab switching."""
import uuid
from copy import deepcopy
import tkinter as tk
from tkinter import ttk, messagebox

from site_apps.central.src.views.components import VehicleInspectionWidget


def signature(data):
    return tuple(sorted((key, str(value or '')) for key, value in data.items()
                        if key not in ('_raw_mileage',)))


class VehiclePanel(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.editors = []
        self.saved = {}
        self.sources = {}
        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=3, pady=3)
        self.summary = ttk.Label(toolbar, text='차량점검 · 1대')
        self.summary.pack(side='left')
        ttk.Button(toolbar, text='＋ 차량 추가', command=self.add).pack(side='right')
        ttk.Button(toolbar, text='기록 편집', command=self.open_history).pack(side='right', padx=3)
        ttk.Button(toolbar, text='차량 목록 설정', command=lambda: app.open_list_management_dialog(
            '차량 목록 관리', getattr(app, 'vehicles', []), 'vehicles')).pack(side='right', padx=3)
        self.notebook = ttk.Notebook(self)
        self.install_tab_close_buttons()
        self.notebook.pack(fill='both', expand=True, padx=3)
        self.notebook.bind('<<NotebookTabChanged>>', lambda event: self.refresh())
        app.cb_daily_site.bind('<<ComboboxSelected>>', lambda event: self.refresh(), add='+')
        footer = ttk.Frame(self)
        footer.pack(fill='x', padx=3, pady=3)
        self.edit_status = ttk.Label(footer, text='새 차량 입력')
        self.edit_status.pack(side='left')
        ttk.Button(footer, text='차량 삭제', command=self.delete_selected).pack(side='right')
        self.save_button = ttk.Button(footer, text='이 차량 저장', command=self.save_selected)
        self.save_button.pack(side='right', padx=3)
        self.add(primary=True, reveal=False)

    def install_tab_close_buttons(self):
        """Add a real clickable close element without changing other notebooks."""
        style = ttk.Style(self)
        name = f'VehicleTabs{uuid.uuid4().hex}.TNotebook'
        element = f'{name}.close'
        self._close_images = []
        for color in ('#606060', '#c42b1c'):
            image = tk.PhotoImage(master=self, width=20, height=18)
            for offset in range(7):
                for width in (0, 1):
                    image.put(color, (6 + offset, 5 + offset + width))
                    image.put(color, (12 - offset, 5 + offset + width))
            self._close_images.append(image)
        style.element_create(element, 'image', self._close_images[0],
                             ('active', self._close_images[1]), sticky='')
        layout = deepcopy(style.layout('TNotebook.Tab'))

        def insert_close(nodes):
            for index, (part, options) in enumerate(nodes):
                if part.endswith('.label'):
                    options['side'] = 'left'
                    nodes.insert(index + 1, (element, {'side': 'left', 'sticky': ''}))
                    return True
                if insert_close(options.get('children', [])):
                    return True
            return False

        if not insert_close(layout):
            raise RuntimeError('차량 탭 닫기 버튼을 구성하지 못했습니다.')
        style.layout(f'{name}.Tab', layout)
        self.notebook.configure(style=name)
        self._close_pressed = None
        self.notebook.bind('<ButtonPress-1>', self.on_tab_close_press, add='+')
        self.notebook.bind('<ButtonRelease-1>', self.on_tab_close_release, add='+')
        self.notebook.bind('<Motion>', lambda event: self.notebook.configure(
            cursor='hand2' if self.notebook.identify(event.x, event.y).endswith('.close') else ''), add='+')
        self.notebook.bind('<Leave>', lambda event: self.notebook.configure(cursor=''), add='+')

    def close_target(self, event):
        if not self.notebook.identify(event.x, event.y).endswith('.close'):
            return None
        try:
            tab = self.notebook.tabs()[self.notebook.index(f'@{event.x},{event.y}')]
        except (tk.TclError, IndexError):
            return None
        return next((editor for editor in self.editors if str(editor) == tab), None)

    def on_tab_close_press(self, event):
        self._close_pressed = self.close_target(event)
        if self._close_pressed is not None:
            return 'break'

    def on_tab_close_release(self, event):
        pressed = self._close_pressed
        self._close_pressed = None
        if pressed is not None:
            if self.close_target(event) is pressed:
                self.request_close(pressed)
            return 'break'

    def reveal(self):
        notebook = getattr(self.app, 'daily_work_notebook', None)
        page = getattr(self.app, 'daily_history_panel', None)
        if notebook is not None and page is not None:
            notebook.select(page)

    def add(self, initial_data=None, key=None, primary=False, reveal=True):
        # Restore legacy layout keys into the panel instead of floating windows.
        if key and key in getattr(self.app, 'vehicle_inspections', {}):
            widget = self.app.vehicle_inspections[key]
            if widget in self.editors:
                if initial_data:
                    widget.set_data(initial_data)
                return widget
        widget = VehicleInspectionWidget(
            self.notebook, theme_bg=getattr(self.app, 'theme_bg', '#f0f0f0'),
            vehicle_list=getattr(self.app, 'vehicles', []), on_save=self.save_one)
        widget.save_button.pack_forget()
        widget._panel_key = None if primary else (key or f'vehicle_inspection_{uuid.uuid4().hex}')
        self.editors.append(widget)
        self.notebook.add(widget, text='새 차량 · 작성 중')
        if primary:
            self.app.fixed_vehicle_widget = widget
            self.app.vehicle_widget = widget
        else:
            self.app.vehicle_inspections[widget._panel_key] = widget
            self.app.vehicle_boxes.append(widget)
        # Text variables catch typing, pasted values, and programmatic restoration.
        widget._panel_variables = []
        for field in (widget.cb_vehicle_info, widget.ent_mileage, widget.ent_remarks):
            variable = tk.StringVar(widget, value=field.get())
            field.configure(textvariable=variable)
            variable.trace_add('write', lambda *args, editor=widget: self.refresh(editor))
            widget._panel_variables.append(variable)
        for variable in widget.vars.values():
            variable.trace_add('write', lambda *args, editor=widget: self.refresh(editor))
        widget.bind('<<VehiclePhotoChanged>>', lambda event: self.refresh(widget))
        if initial_data:
            widget.set_data(initial_data)
        self.notebook.select(widget)
        self.refresh(widget)
        if reveal:
            self.reveal()
            widget.cb_vehicle_info.focus_set()
        return widget

    def context(self):
        return str(self.app.ent_daily_date.get_date()), self.app.cb_daily_site.get().strip()

    def mark_saved(self, widget):
        self.saved[widget] = (self.context(), signature(widget.get_data()))
        self.refresh(widget)

    def refresh(self, widget=None):
        for editor in ([widget] if widget is not None else self.editors):
            if editor not in self.editors:
                continue
            data = editor.get_data()
            number = data.get('vehicle_info', '').strip().split(' (')[0] or '새 차량'
            distance = data.get('mileage', '').strip()
            saved = self.saved.get(editor) == (self.context(), signature(data))
            state = '수정 중' if editor in getattr(self, 'sources', {}) else '작성 중'
            label = f"{number}{' · ' + distance + 'km' if distance else ''} · {'저장 완료' if saved else state}"
            self.notebook.tab(editor, text=label)
        self.summary.configure(text=f'차량점검 · {len(self.editors)}대')
        active = self.selected()
        source = getattr(self, 'sources', {}).get(active)
        if hasattr(self, 'save_button'):
            self.save_button.configure(text='수정 저장' if source else '이 차량 저장')
            self.edit_status.configure(text=f"{source['date']} · {source['site']} · 기록 편집" if source else '새 차량 입력')

    def selected(self):
        selected = self.notebook.select()
        return next((w for w in self.editors if str(w) == selected), None)

    def save_selected(self):
        widget = self.selected()
        if widget is not None:
            self.save_one(widget)

    def check_pending_edits(self):
        for editor in self.sources:
            if self.saved.get(editor) != (self.context(), signature(editor.get_data())):
                raise ValueError('차량 기록 수정 내용이 있습니다. 차량 패널의 수정 저장을 먼저 눌러주세요.')

    def save_one(self, widget):
        if self.saved.get(widget) == (self.context(), signature(widget.get_data())):
            messagebox.showinfo('차량점검', '현재 내용은 이미 저장되어 있습니다.', parent=self)
            return
        source = getattr(self, 'sources', {}).get(widget)
        if source:
            from site_apps.central.src.vehicle_records import persist_vehicle_edit
            try:
                self.sources[widget] = persist_vehicle_edit(self.app, source, widget.get_data())
            except Exception as exc:
                messagebox.showerror('수정 저장', str(exc), parent=self)
                return
            self.mark_saved(widget)
            try:
                for method in ('update_daily_usage_view', 'update_recent_entries_view', 'refresh_inquiry_filters'):
                    getattr(self.app, method)()
            except Exception as exc:
                messagebox.showwarning('수정 저장 완료', f'기록은 저장되었습니다. 화면을 다시 조회해주세요.\n{exc}', parent=self)
            return
        if self.app.save_single_vehicle_data(widget):
            self.mark_saved(widget)

    def open_history(self):
        from site_apps.central.src.vehicle_records import record_choices
        choices = record_choices(self.app.daily_usage_df)
        dialog = tk.Toplevel(self)
        dialog.title('저장된 차량 기록 편집')
        dialog.geometry('850x440')
        dialog.transient(self.winfo_toplevel())
        filter_row = ttk.Frame(dialog)
        filter_row.pack(fill='x', padx=10, pady=8)
        selected_date = tk.BooleanVar(value=True)
        query = tk.StringVar()
        ttk.Checkbutton(filter_row, text='선택 날짜만', variable=selected_date,
                        command=lambda: render()).pack(side='left')
        ttk.Label(filter_row, text='차량·현장 검색:').pack(side='left', padx=(10, 3))
        entry = ttk.Entry(filter_row, textvariable=query, width=30)
        entry.pack(side='left', fill='x', expand=True)
        ttk.Label(dialog, text='선택한 차량의 주행거리·점검·비고·사진을 수정합니다. 검사량과 비용은 유지됩니다.').pack(anchor='w', padx=10)
        body = ttk.Frame(dialog)
        body.pack(fill='both', expand=True, padx=10, pady=5)
        columns = ('date', 'site', 'vehicle', 'mileage', 'method', 'copies')
        tree = ttk.Treeview(body, columns=columns, show='headings', selectmode='browse')
        for key, title, width in zip(columns, ('날짜', '현장', '차량정보', '주행거리(km)', '검사방법', '연결 기록'),
                                     (105, 120, 240, 95, 80, 75)):
            tree.heading(key, text=title)
            tree.column(key, width=width)
        scrollbar = ttk.Scrollbar(body, command=tree.yview)
        scrollbar.pack(side='right', fill='y')
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(fill='both', expand=True)
        count_label = ttk.Label(dialog)
        count_label.pack(anchor='w', padx=10)

        def render():
            tree.delete(*tree.get_children())
            count = 0
            for index, choice in enumerate(choices):
                if selected_date.get() and choice['date'] != self.context()[0]:
                    continue
                data = choice['data']
                if query.get().strip().lower() not in f"{choice['site']} {data.get('vehicle_info', '')}".lower():
                    continue
                tree.insert('', 'end', iid=str(index), values=(choice['date'], choice['site'],
                    data.get('vehicle_info', ''), data.get('mileage', ''), choice['method'],
                    len(choice['source']['fingerprints'])))
                count += 1
            count_label.configure(text=f'{count}건 · 더블클릭하거나 편집을 누르세요.')

        def load(event=None):
            selection = tree.selection()
            if not selection:
                return
            self.edit_record(choices[int(selection[0])])
            dialog.destroy()

        query.trace_add('write', lambda *args: render())
        tree.bind('<Double-1>', load)
        buttons = ttk.Frame(dialog)
        buttons.pack(pady=8)
        ttk.Button(buttons, text='선택 기록 편집', command=load).pack(side='left', padx=4)
        ttk.Button(buttons, text='닫기', command=dialog.destroy).pack(side='left', padx=4)
        render()
        dialog.grab_set()

    def edit_record(self, choice):
        for editor, source in self.sources.items():
            if source == choice['source']:
                self.notebook.select(editor)
                self.reveal()
                return editor
        editor = self.add(choice['data'])
        self.sources[editor] = choice['source']
        self.mark_saved(editor)
        self.refresh(editor)
        return editor

    def delete_selected(self):
        self.request_close(self.selected())

    def request_close(self, editor):
        if editor not in self.editors:
            return
        data = editor.get_data()
        if any(str(v or '').strip() for k, v in data.items() if k != '_raw_mileage'):
            if not messagebox.askyesno('차량 탭 닫기', '이 차량 탭을 닫을까요?\n저장하지 않은 입력 내용은 사라지며, 저장된 기록은 유지됩니다.', parent=self):
                return
        self.remove(editor)

    def remove(self, editor):
        if editor not in self.editors:
            return
        if editor is self.app.fixed_vehicle_widget:
            # Promote another tab to keep the legacy fixed-widget contract intact.
            if len(self.editors) == 1:
                getattr(self, 'sources', {}).pop(editor, None)
                self.saved.pop(editor, None)
                editor.reset_fields()
                editor.clear_photo()
                self.refresh(editor)
                return
            replacement = next(w for w in self.editors if w is not editor)
            self.app.vehicle_inspections.pop(replacement._panel_key, None)
            self.app.vehicle_boxes.remove(replacement)
            replacement._panel_key = None
            self.app.fixed_vehicle_widget = replacement
            self.app.vehicle_widget = replacement
        else:
            self.app.vehicle_inspections.pop(editor._panel_key, None)
            if editor in self.app.vehicle_boxes:
                self.app.vehicle_boxes.remove(editor)
        self.saved.pop(editor, None)
        getattr(self, 'sources', {}).pop(editor, None)
        self.editors.remove(editor)
        self.notebook.forget(editor)
        editor.destroy()
        self.refresh()

    def serialize(self):
        result = []
        for editor in self.editors:
            data = editor.get_data()
            if editor in getattr(self, 'sources', {}):
                data['_edit_source'] = self.sources[editor]
            result.append(data)
        return result

    def restore(self, records):
        self.reset()
        for i, record in enumerate(records):
            editor = self.editors[0] if i == 0 else self.add(reveal=False)
            editor.set_data(record)
            if record.get('_edit_source'):
                self.sources[editor] = record['_edit_source']
        self.notebook.select(self.editors[0])
        self.refresh()

    def reset(self):
        for editor in list(self.editors[1:]):
            self.remove(editor)
        self.saved.clear()
        getattr(self, 'sources', {}).clear()
        self.editors[0].reset_fields()
        self.editors[0].clear_photo()
        self.refresh()
