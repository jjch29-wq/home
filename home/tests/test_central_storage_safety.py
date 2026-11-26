"""Use disposable files to verify primary-workbook protection and safe imports."""
import ast
import datetime
import json
import os
from pathlib import Path
import sys
import tempfile
import time
from types import SimpleNamespace
from unittest.mock import Mock, mock_open, patch
import unittest
import pandas as pd

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'home/src'))
from site_apps.central.src.inspection_batch import commit_plans
from site_apps.central.src.runtime_config import ensure_initial_config


def method(path,name,ns):
    fn=next(n for n in ast.walk(ast.parse(path.read_text(encoding='utf-8'))) if isinstance(n,ast.FunctionDef) and n.name==name)
    exec(compile(ast.Module(body=[fn],type_ignores=[]),str(path),'exec'),ns)
    return ns[name]


class StorageSafetyTests(unittest.TestCase):
    def test_first_run_creates_config_without_replacing_existing_settings(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'tmp') as directory:
            path = Path(directory) / 'nested' / 'Material_Manager_Config.json'
            self.assertTrue(ensure_initial_config(path))
            self.assertEqual(json.loads(path.read_text(encoding='utf-8')), {})

            original = b'{"layout_locked": true}\n'
            path.write_bytes(original)
            self.assertFalse(ensure_initial_config(path))
            self.assertEqual(path.read_bytes(), original)

    def test_tab_change_retries_failed_save(self):
        from site_apps.central.src.controllers.event_controller import on_tab_changed_impl

        app = SimpleNamespace(root=Mock(), notebook=Mock(), is_ready=True,
                              _data_dirty=True, save_data=Mock(side_effect=[False, True]))
        app.notebook.select.return_value = ''
        on_tab_changed_impl(app)
        self.assertTrue(app._data_dirty)
        on_tab_changed_impl(app)
        self.assertFalse(app._data_dirty)
        on_tab_changed_impl(app)
        self.assertEqual(app.save_data.call_count, 2)

    def saver(self):
        return method(ROOT/'home/src/site_apps/central/src/services/data_loader.py','save_data_impl',
                      {'pd':pd,'os':os,'time':time,'datetime':datetime})

    def app(self,path):
        return SimpleNamespace(db_path=str(path), materials_df=pd.DataFrame({'value':[1]}),
            daily_usage_df=pd.DataFrame(),transactions_df=pd.DataFrame(),monthly_usage_df=pd.DataFrame(),
            budget_df=pd.DataFrame(),show_error_dialog=Mock())

    def test_serialization_failure_preserves_original_bytes(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'tmp') as directory:
            path=Path(directory)/'data.xlsx'
            pd.DataFrame({'original':[123]}).to_excel(path,index=False)
            original=path.read_bytes()
            app=self.app(path)
            app.materials_df=pd.DataFrame({'bad':['\x00']})
            self.assertFalse(self.saver()(app,require_primary=True))
            self.assertEqual(path.read_bytes(),original)
            self.assertEqual(list(Path(directory).glob('.material-save-*')),[])

    def test_success_replaces_with_complete_workbook(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'tmp') as directory:
            path=Path(directory)/'data.xlsx'
            app=self.app(path)
            self.assertTrue(self.saver()(app,require_primary=True))
            with pd.ExcelFile(path) as book:
                self.assertEqual(set(book.sheet_names),{'Materials','Transactions','MonthlyUsage','DailyUsage','Budget'})

    def test_external_change_blocks_stale_save_and_preserves_memory(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'tmp') as directory:
            path=Path(directory)/'data.xlsx'
            app=self.app(path)
            save=self.saver()
            self.assertTrue(save(app))
            pd.DataFrame({'new_mobile_record':[99]}).to_excel(path,index=False)
            newer=path.read_bytes()
            self.assertFalse(save(app))
            self.assertEqual(path.read_bytes(),newer)
            self.assertEqual(app.materials_df.iloc[0,0],1)
            app.show_error_dialog.assert_called_once()

    def test_change_during_serialization_blocks_replace(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'tmp') as directory:
            path=Path(directory)/'data.xlsx'
            app=self.app(path)
            save=self.saver()
            self.assertTrue(save(app))
            original=pd.DataFrame.to_excel
            newer=[]
            def external_change(frame,*args,**kwargs):
                result=original(frame,*args,**kwargs)
                if not newer:
                    path.write_bytes(b'external newer version')
                    newer.append(path.read_bytes())
                return result
            with patch.object(pd.DataFrame,'to_excel',external_change):
                self.assertFalse(save(app))
            self.assertEqual(path.read_bytes(),newer[0])

    def test_locked_primary_keeps_original_and_returns_failure_with_backup(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'tmp') as directory:
            path=Path(directory)/'data.xlsx'
            pd.DataFrame({'original':[123]}).to_excel(path,index=False)
            original=path.read_bytes()
            app=self.app(path)
            replace=os.replace
            def locked(source, destination):
                if os.path.abspath(destination)==os.path.abspath(path):
                    raise PermissionError('locked')
                return replace(source,destination)
            with patch('os.replace',side_effect=locked), patch('time.sleep'):
                self.assertFalse(self.saver()(app))
            self.assertEqual(path.read_bytes(),original)
            backups=list(Path(directory).glob('*_Conflict_*.xlsx'))
            self.assertEqual(len(backups),1)
            with pd.ExcelFile(backups[0]) as book:
                self.assertIn('Budget',book.sheet_names)

    def test_batch_requires_primary_save_and_restores_failed_attempt(self):
        app=self.app('unused')
        for name in ('sites','users','vehicles','daily_units','equipments'): setattr(app,name,[])
        app._commit_usage_plan=Mock(return_value=1)
        app.save_data=Mock(return_value=False)
        with self.assertRaises(RuntimeError): commit_plans(app,[{}])
        app.save_data.assert_called_once_with(require_primary=True)

    def test_corrupt_project_import_never_clears_existing_records(self):
        path=ROOT/'home/src/site_apps/central/src/ndt_billing_tab.py'
        fn=method(path,'load_project',{'json':json,'deepcopy':__import__('copy').deepcopy,
            'filedialog':SimpleNamespace(askopenfilenames=lambda **kw:['bad.ndt']),'messagebox':Mock()})
        app=SimpleNamespace(records=[{'qty':10}],_capture_billing_state=lambda:{'round':1},clear_records=Mock())
        with patch('builtins.open',mock_open(read_data='{broken')): fn(app)
        self.assertEqual(app.records,[{'qty':10}])
        app.clear_records.assert_not_called()

    def test_material_and_transaction_writes_rollback_when_save_fails(self):
        from site_apps.central.src.views.inout_view import InoutViewMixin
        from site_apps.central.src.views.stock_management_view import (
            delete_selected_material_impl,
            save_material_edits_impl,
        )
        class Field:
            def __init__(self, v=''): self.v = v
            def get(self): return self.v
            def set(self, v): self.v = v
            def delete(self, *a): self.v = ''
            def insert(self, i, v): self.v = v

        base_materials = pd.DataFrame([
            {'MaterialID': 1, '품목명': '세척제', '모델명': 'M1', 'SN': '', '창고': '본사', '수량': 10.0, 'Active': 1}
        ])
        base_trans = pd.DataFrame(columns=['Date', 'MaterialID', 'Type', 'Quantity', 'Note', 'User', 'Site'])
        app = SimpleNamespace(
            root=Mock(),
            materials_df=base_materials.copy(deep=True),
            transactions_df=base_trans.copy(deep=True),
            equipments=[], sites=[], users=[],
            cb_co_code=Field('NAWOO'), cb_item_name=Field('신규약품'), cb_eq_name=Field(''),
            cb_eq_code=Field(''), ent_sn=Field(''), cb_class=Field(''), cb_spec=Field(''),
            cb_unit=Field('CAN'), cb_supplier=Field(''), cb_mfr=Field('NAWOO'), cb_origin=Field(''),
            ent_init=Field('5'), ent_reorder=Field('1'), ent_price=Field('1000'), ent_cost=Field('800'),
            cb_model=Field('M2'),
            cb_material=Field('세척제'), cb_type=Field('OUT'), ent_user=Field('작업자'),
            ent_qty=Field('3'), ent_note=Field('출고메모'), cb_warehouse=Field('현장창고'), cb_trans_site=Field('중앙지사'),
            extract_sn_from_model=lambda m, s: (m, s),
            normalize_id=lambda v: str(v).strip(),
            save_data=Mock(return_value=False),
            update_material_combo=Mock(), update_registration_combos=Mock(),
            update_stock_view=Mock(), update_transaction_view=Mock(), show_error_dialog=Mock(),
            stock_tree=SimpleNamespace(selection=lambda: ['i1'], item=lambda i, k: [1] if k == 'values' else {}),
        )
        with (
            patch('site_apps.central.src.views.inout_view.messagebox') as io_msg,
            patch('site_apps.central.src.views.stock_management_view.messagebox') as st_msg,
        ):
            st_msg.askyesno.return_value = True
            InoutViewMixin.register_material(app)
            pd.testing.assert_frame_equal(app.materials_df, base_materials)
            self.assertEqual(app.cb_item_name.get(), '신규약품')
            io_msg.showerror.assert_called_once()

            io_msg.reset_mock()
            InoutViewMixin.add_transaction(app)
            pd.testing.assert_frame_equal(app.transactions_df, base_trans)
            pd.testing.assert_frame_equal(app.materials_df, base_materials)
            self.assertEqual(app.ent_qty.get(), '3')
            io_msg.showerror.assert_called_once()

            self.assertFalse(save_material_edits_impl(app, 1, {'수량': '99', '모델명': 'CHANGED'}))
            pd.testing.assert_frame_equal(app.materials_df, base_materials)

            delete_selected_material_impl(app)
            pd.testing.assert_frame_equal(app.materials_df, base_materials)

if __name__=='__main__': unittest.main()
