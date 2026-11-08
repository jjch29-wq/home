"""Budget history mutations must validate inputs and restore failed persistence."""
import ast
import json
import math
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import unittest
import pandas as pd

SOURCE=Path(__file__).resolve().parents[1]/'src/site_apps/central/src/app.py'
TREE=ast.parse(SOURCE.read_text(encoding='utf-8'))


class BudgetHistoryTests(unittest.TestCase):
    def setUp(self):
        self.messages=Mock()
        self.ns={'pd':pd,'json':json,'math':math,'messagebox':self.messages}
        for name in ('save_budget_entry','delete_budget_entry','_load_budget_to_form','clear_budget_form'):
            fn=next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef) and n.name==name)
            exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),self.ns)
        self.site=Mock(); self.site.get.return_value='A'
        self.revenue=Mock(); self.revenue.get.return_value='150'
        self.app=SimpleNamespace(root=Mock(),cb_budget_site=self.site,ent_budget_revenue=self.revenue,
            budget_df=pd.DataFrame([{'Site':'A','Revenue':100.0}]),save_data=Mock(return_value=False),
            _capture_budget_input=lambda:{'fields':{'revenue':'draft'},'details':{}},clear_budget_form=Mock())

    def test_reset_clears_planned_and_actual_detail_values(self):
        class Detail:
            def __init__(self): self.values = {'amount':100}
            def reset(self): self.values = {}
        names=('labor_detail_widget','material_detail_widget','expense_detail_widget',
               'actual_labor_detail_widget','actual_material_detail_widget','actual_expense_detail_widget')
        app=SimpleNamespace(budget_widgets={},_update_budget_kpis=Mock(),
                            **{name:Detail() for name in names})
        self.ns['clear_budget_form'](app)
        self.assertTrue(all(getattr(app,name).values == {} for name in names))
        app._update_budget_kpis.assert_called_once()

    def test_invalid_amount_never_reaches_save(self):
        for value in ('wrong','NaN','inf'):
            self.revenue.get.return_value=value
            self.ns['save_budget_entry'](self.app)
        self.app.save_data.assert_not_called()
        self.assertEqual(self.app.budget_df.iloc[0]['Revenue'],100)

    def test_save_failure_restores_original(self):
        original=self.app.budget_df.copy(deep=True)
        self.ns['save_budget_entry'](self.app)
        pd.testing.assert_frame_equal(self.app.budget_df,original)
        self.app.save_data.assert_called_once_with(require_primary=True)

    def test_delete_failure_restores_original(self):
        original=self.app.budget_df.copy(deep=True)
        self.ns['delete_budget_entry'](self.app)
        pd.testing.assert_frame_equal(self.app.budget_df,original)
        self.app.clear_budget_form.assert_not_called()

    def test_missing_budget_does_not_clear_draft(self):
        self.app._budget_loaded_site='A'
        self.ns['_load_budget_to_form'](self.app,'B',silent=True)
        self.app.clear_budget_form.assert_not_called()
        self.site.set.assert_called_with('A')

    def test_cancel_loading_existing_budget_keeps_draft(self):
        self.app._budget_loaded_site='previous'
        self.app._budget_baseline={'fields':{'revenue':'saved'},'details':{}}
        self.messages.askyesno.return_value=False
        self.ns['_load_budget_to_form'](self.app,'A',silent=True)
        self.app.clear_budget_form.assert_not_called()
        self.site.set.assert_called_with('previous')

if __name__=='__main__': unittest.main()
