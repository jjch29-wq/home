import sys
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk
from tkinter import ttk
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.daily_input_summary import build_input_summary, DailyInputSummary


class Field:
    def __init__(self, value):
        self.value = value
    def get(self):
        return self.value


def worker(name, time, ot, meal, selected=True):
    return SimpleNamespace(participation_var=Field(selected), get_worker=lambda: name,
                           get_time=lambda: time, get_ot=lambda: ot, get_meal=lambda: meal)


def fixture():
    values = dict(cb_daily_site='중앙지사', cb_daily_test_method='PT', ent_daily_inspection_item='Pipe',
                  cb_daily_unit='M', ent_daily_test_amount='2', ent_daily_unit_price='4,200',
                  ent_daily_test_fee='8,400', ent_daily_travel_cost='50,000', ent_daily_meal_cost='0',
                  cb_daily_equip='MX2 (SN01); MX2 (SN02)', daily_aux_equipment_var='Probe (P01); Scanner (S01)')
    app = SimpleNamespace(**{key: Field(value) for key, value in values.items()})
    app.worker_group1 = worker('A', '(야간) 20:00~24:00', '18000', '12000')
    app.worker_group2 = worker('B', '20:00~24:00', '99999', '99999', False)
    return app


class InputSummaryTests(unittest.TestCase):
    def test_shows_full_equipment_and_counts_selected_workers_only(self):
        app = fixture()
        summary = build_input_summary(app)
        self.assertIn('MX2 (SN01)\nMX2 (SN02)', summary['blocks'][0][1])
        self.assertIn('Probe (P01)\nScanner (S01)', summary['blocks'][1][1])
        self.assertEqual(summary['blocks'][2][0], '참여 작업자 · 1명')
        self.assertIn('18,000원', summary['money'])
        self.assertIn('12,000원', summary['money'])
        self.assertNotIn('99,999', summary['money'])
        self.assertEqual(summary['warnings'], [])
        self.assertEqual(app.ent_daily_test_fee.get(), '8,400')

    def test_invalid_values_and_duplicate_workers_are_explicit(self):
        app = fixture()
        app.cb_daily_site.value = ''
        app.ent_daily_test_amount.value = 'nan'
        app.worker_group2 = worker('A', '', 'invalid', '-1')
        summary = build_input_summary(app)
        for text in ('현장명 미입력', '검사량 확인', 'A 중복 선택', 'A 작업시간 확인', 'A OT 확인', 'A 일비 확인'):
            self.assertIn(text, summary['status'])
        self.assertIn('OT 합계 확인 필요', summary['money'])
        self.assertIn('작업자 일비 확인 필요', summary['money'])

    def test_widget_refresh_and_resizing_preserve_editor(self):
        root = tk.Tk()
        root.withdraw()
        self.addCleanup(root.destroy)
        parent = ttk.Frame(root)
        form = ttk.Frame(parent, width=900, height=250)
        form.grid(row=0, column=0)
        app = fixture()
        view = DailyInputSummary(parent, app, form)
        root.update_idletasks()
        parent.winfo_width = lambda: 1900
        view.layout()
        self.assertEqual(view.winfo_manager(), 'grid')
        view.winfo_ismapped = lambda: True
        view.refresh()
        self.assertEqual(str(view.text.cget('state')), 'disabled')
        self.assertIn('SN02', view.text.get('1.0', 'end'))
        app.cb_daily_equip.value = 'Custom (SN03)'
        view.refresh()
        self.assertIn('SN03', view.text.get('1.0', 'end'))
        self.assertNotIn('SN02', view.text.get('1.0', 'end'))
        parent.winfo_width = lambda: 1000
        view.layout()
        self.assertEqual(view.winfo_manager(), '')
        self.assertEqual(app.cb_daily_equip.get(), 'Custom (SN03)')
        view.destroy()
        root.update_idletasks()


if __name__ == '__main__':
    unittest.main()
