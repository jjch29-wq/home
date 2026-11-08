"""One schedule rule for every OT consumer and consistent budget conversion."""
import ast
from pathlib import Path
import sys
from types import SimpleNamespace
from unittest.mock import Mock
import unittest

SRC=Path(__file__).resolve().parents[1]/'src'
sys.path.insert(0,str(SRC))
from site_apps.central.src.services.ot_calculator import ot_breakdown
from site_apps.central.src.models.worker_model import _calculate_ot_from_worktime_impl


class Field:
    def __init__(self,value=''): self.value=str(value)
    def get(self): return self.value
    def delete(self,*args): self.value=''
    def insert(self,index,value): self.value=str(value)


class CommonOtTests(unittest.TestCase):
    def test_overnight_ordinary_weekday_is_eleven_hours_without_rest_deduction(self):
        result=ot_breakdown('(야간) 22:00~익일09:00','2026-10-22')
        self.assertEqual((result['hours'],result['night'],result['amount']),(11,11,55000))
        self.assertEqual(_calculate_ot_from_worktime_impl(None,'22:00~익일09:00',__import__('datetime').date(2026,10,22)),(11,55000))

    def test_friday_dawn_and_weekend_use_same_holiday_rate(self):
        result=ot_breakdown('(야간) 22:00~익일09:00','2026-10-23')
        self.assertEqual((result['night'],result['holiday'],result['amount']),(2,9,77500))
        self.assertEqual(ot_breakdown('09:00~18:00','2026-10-24')['amount'],67500)

    def test_explicit_holiday_and_fractional_start(self):
        self.assertEqual(ot_breakdown('(휴일) 09:00~18:00','2026-10-22')['amount'],67500)
        result=ot_breakdown('18:30~23:30','2026-10-22')
        self.assertEqual((result['day'],result['night'],result['amount']),(3.5,1.5,21500))
        self.assertEqual(result['amount'],result['day_amount']+result['night_amount']+result['holiday_amount'])

    def test_normal_day_has_no_ot_and_recorded_amount_does_not_override_schedule(self):
        self.assertEqual(ot_breakdown('09:00~18:00','2026-10-22','99999')['amount'],0)
        self.assertEqual(ot_breakdown('22:00~익일09:00','2026-10-22','50000')['amount'],55000)

    def test_monthly_apply_updates_actual_only_without_one_won_loss(self):
        source=SRC/'site_apps/central/src/app.py'
        fn=next(n for n in ast.walk(ast.parse(source.read_text(encoding='utf-8'))) if isinstance(n,ast.FunctionDef) and n.name=='apply_worker_shift_hours_to_budget')
        ns={'messagebox':Mock()}
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(source),'exec'),ns)
        entries={key:{field:Field() for field in ('personnel','period','unit_price')} for key in ('연장근무','야간근무','휴일근무')}
        tree=Mock();tree.get_children.return_value=list(range(6))
        tree.item.side_effect=lambda item,what: () if what=='tags' else [f'worker{item}',1,0,110/6,0,0,0,0,0,0]
        planned=Mock()
        app=SimpleNamespace(worker_summary_tree=tree,actual_labor_detail_widget=SimpleNamespace(entries=entries,_on_input_change=Mock(),calculate_all=Mock()),labor_detail_widget=planned,notebook=Mock())
        ns['apply_worker_shift_hours_to_budget'](app)
        night=entries['야간근무']
        self.assertEqual(round(float(night['personnel'].get())*float(night['period'].get())*float(night['unit_price'].get())),550000)
        self.assertEqual(entries['연장근무']['personnel'].get(),'0')
        self.assertEqual(planned.mock_calls,[])

if __name__=='__main__': unittest.main()
