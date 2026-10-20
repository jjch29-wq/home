"""Exercise the actual daily record builder/committer without opening Tk or the live database."""
import ast
from copy import deepcopy
import datetime
import json
import math
import re
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import pandas as pd

SRC = Path(__file__).resolve().parents[1] / 'src'
sys.path.insert(0, str(SRC))
from site_apps.central.src.inspection_batch import (
    InspectionBatch, combine_plans, commit_plans, json_copy, number,
    participation_record, work_interval,
)
from site_apps.central.src.worker_selection import worker_is_selected, select_workers


class Field:
    def __init__(self, value=''):
        self.value = str(value)
        self.callbacks = []

    def get(self, *args):
        return self.value

    def set(self, value):
        self.value = str(value)
        for callback in self.callbacks:
            callback()

    def trace_add(self, mode, callback):
        self.callbacks.append(callback)

    def delete(self, *args):
        self.value = ''

    def insert(self, index, value):
        self.value = str(value)


class Flag(Field):
    def __init__(self, value=True):
        self.value = bool(value)

    def set(self, value):
        self.value = bool(value)


def actual_methods():
    tree = ast.parse((SRC / 'site_apps/central/src/app.py').read_text(encoding='utf-8'))
    cls = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == 'MaterialManager')
    methods = [node for node in cls.body if isinstance(node, ast.FunctionDef) and node.name in (
        '_add_single_usage_record_logic', '_commit_usage_plan', '_init_ndt_vars_and_traces',
        '_get_usage_session_data', '_set_usage_session_data')]
    namespace = {'pd': pd, 'datetime': datetime, 'ttk': SimpleNamespace(Combobox=type('Combo', (), {})),
                 'tk': SimpleNamespace(StringVar=lambda value='': Field(value), END='end')}
    module = ast.Module(body=[ast.ClassDef(name='App', bases=[], keywords=[], body=methods, decorator_list=[])], type_ignores=[])
    exec(compile(ast.fix_missing_locations(module), '<actual daily record methods>', 'exec'), namespace)
    return namespace['App']


def fixture():
    app = actual_methods()()
    fields = {
        'cb_daily_equip': 'OMNI', 'cb_daily_test_method': 'PAUT',
        'ent_daily_inspection_item': 'PIPE', 'ent_daily_applied_code': 'CODE',
        'ent_daily_report_no': 'REPORT', 'ent_daily_test_amount': '1.6802',
        'ent_daily_unit_price': '100', 'ent_daily_travel_cost': '500',
        'ent_daily_meal_cost': '1000', 'ent_daily_test_fee': '168.02',
        'cb_daily_material': 'OMNI', 'ent_daily_note': '',
        'cb_daily_company': '발주처', 'cb_daily_unit': 'M,I/D',
        'ndt_work_time_var': '일반', 'ndt_source_var': 'Ir-192 (1.0)',
        'ndt_pipe_var': '250A (1.0)', 'ndt_thickness_var': '조건없음 (1.0)',
        'ndt_loc_type_var': '열수송관', 'hot_tap_stage_var': '1차',
        'hot_tap_location_var': '', 'hot_tap_size_var': '600A',
        'ndt_report_pipe_var': '250A', 'ndt_ori_joint_var': '2',
        'ndt_ori_qty_var': '1.6802', 'ndt_rep_joint_var': '', 'ndt_rep_qty_var': '',
        'ndt_rej_joint_var': '', 'ndt_overhead_var': '110', 'ndt_tech_var': '20',
    }
    for key, value in fields.items():
        setattr(app, key, Field(value))
    app.worker_group1 = SimpleNamespace(get_worker=lambda: '작업자', get_time=lambda: '09:00~18:00',
                                       get_ot=lambda: '12000', get_meal=lambda: '1000')
    app.rtk_entries = {'기타': Field('1')}
    app.ndt_materials_all = ['세척제']
    app.ndt_company_entries = [{'_company': Field('A'), '세척제': Field('2')}]
    app._get_merged_memo_and_note = lambda: '작업 비고'
    app.daily_usage_df = pd.DataFrame()
    app.transactions_df = pd.DataFrame(columns=['MaterialID', 'Quantity'])
    app.materials_df = pd.DataFrame([{'MaterialID': 10, '품목명': 'Carestream T200'}])
    app.monthly_usage_df = pd.DataFrame()
    for name in ('sites', 'users', 'vehicles', 'daily_units', 'equipments'):
        setattr(app, name, [])
    app.refresh_ui_for_list_change = Mock()
    app.calculate_rtk_total = Mock()
    app.save_tab_config = Mock()
    app.save_data = Mock(return_value=True)
    app.get_material_display_name = lambda mid: 'Carestream T200' if mid == 10 else 'PAUT장비'
    app._is_consumable_material = lambda name, method: name == 'Carestream T200'
    app._auto_reconcile_and_register_ndt = Mock()
    def stock(date, mid, kind, qty, site, workers, note):
        app.transactions_df = pd.concat([app.transactions_df, pd.DataFrame([{'MaterialID': mid, 'Quantity': -qty}])], ignore_index=True)
    app._create_manual_stock_transaction = Mock(side_effect=stock)
    return app


def three_plans(app):
    day = datetime.date(2026, 10, 20)
    first = app._add_single_usage_record_logic('OMNI', day, '중앙지사', prepare_only=True)
    app.ndt_report_pipe_var.set('400A')
    app.ndt_pipe_var.set('300A 이상 (1.0)')
    app.ent_daily_test_amount.set('2.5534')
    app.ndt_ori_qty_var.set('2.5534')
    app.ent_daily_unit_price.set('80')
    app.ent_daily_test_fee.set('204.272')
    second = app._add_single_usage_record_logic('OMNI', day, '중앙지사', prepare_only=True)
    app.cb_daily_test_method.set('RT')
    app.cb_daily_equip.set('IR-192')
    app.cb_daily_unit.set('매')
    app.ndt_report_pipe_var.set('250A')
    app.ent_daily_test_amount.set('54')
    app.ndt_ori_qty_var.set('54')
    app.ent_daily_unit_price.set('200')
    app.ent_daily_test_fee.set('10800')
    third = app._add_single_usage_record_logic(10, day, '중앙지사', prepare_only=True)
    return [first, second, third]


def shared():
    return {'Date': datetime.date(2026, 10, 20), 'Site': '중앙지사', '업체명': '발주처',
            'User': '작업자', 'WorkTime': '09:00~18:00', 'OT': '12000', 'Meal': '1000',
            '출장비': 500, '일식': 1000, '차량번호': '차량', '주행거리': '30',
            '차량점검': '양호', '차량비고': '', 'DailyMemo': '공통 메모'}


class BatchTests(unittest.TestCase):
    def test_worker_checkboxes_filter_saved_crew_without_erasing_roster(self):
        app = fixture()
        app.worker_group1.participation_var = Flag(True)
        app.worker_group2 = SimpleNamespace(participation_var=Flag(True), get_worker=lambda: 'B',
                get_time=lambda: '09:00~18:00', get_ot=lambda: '12000', get_meal=lambda: '1000')
        rt = app._add_single_usage_record_logic(10, datetime.date(2026, 10, 20), '중앙지사', prepare_only=True)
        app.worker_group2.participation_var.set(False)
        paut = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 20), '중앙지사', prepare_only=True)
        self.assertEqual(rt['workers_list'], ['작업자', 'B'])
        self.assertEqual(paut['workers_list'], ['작업자'])
        self.assertEqual(paut['common_data']['User2'], '')
        self.assertEqual(paut['common_data']['WorkTime2'], '')
        self.assertEqual(paut['common_data']['OT2'], '')
        self.assertEqual(paut['common_data']['Meal2'], '')
        self.assertEqual(app.worker_group2.get_worker(), 'B')
        self.assertEqual(app.worker_group2.get_time(), '09:00~18:00')
        self.assertEqual(app.worker_group2.get_ot(), '12000')
        app.worker_group2.participation_var.set(True)
        again = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 20), '중앙지사', prepare_only=True)
        self.assertEqual(again['workers_list'], ['작업자', 'B'])
        app.save_data.assert_not_called()

    def test_select_all_and_none_keep_input_values(self):
        app = fixture()
        app.worker_group1.participation_var = Flag(True)
        select_workers(app, False)
        self.assertFalse(worker_is_selected(app.worker_group1))
        self.assertEqual(app.worker_group1.get_worker(), '작업자')
        select_workers(app, True)
        self.assertTrue(worker_is_selected(app.worker_group1))
        self.assertTrue(worker_is_selected(SimpleNamespace()))

    def test_worker_selection_session_roundtrip_and_old_session_default(self):
        app = fixture()
        app.cb_daily_site = Field('중앙지사')
        app.ent_daily_date = SimpleNamespace(get_date=lambda: datetime.date(2026, 10, 20), set_date=Mock())
        app.ndt_company_entries[0].update({name: Field() for name in
                ('형광자분', '흑색자분', '백색페인트', '침투제', '현상제', '형광침투제')})
        app.worker_group1 = SimpleNamespace(cb_name=Field('A'), cb_shift=Field('야간'),
                ent_worktime=Field('22:00~익일09:00'), ent_ot=Field('55000'), participation_var=Flag(False))
        data = json_copy(app._get_usage_session_data())
        self.assertFalse(data['workers'][0]['participating'])
        self.assertEqual(data['workers'][0]['name'], 'A')
        app.worker_group1.participation_var.set(True)
        app.worker_group1.cb_name.set('B')
        app._set_usage_session_data(data)
        self.assertFalse(worker_is_selected(app.worker_group1))
        self.assertEqual(app.worker_group1.cb_name.get(), 'A')
        del data['workers'][0]['participating']
        app._set_usage_session_data(data)
        self.assertTrue(worker_is_selected(app.worker_group1))
    def test_each_job_keeps_its_crew_and_shared_worker_is_counted_once(self):
        app = fixture()
        app.worker_group1 = SimpleNamespace(get_worker=lambda: 'A', get_time=lambda: '09:00~18:00',
                                           get_ot=lambda: '', get_meal=lambda: '1000')
        app.worker_group2 = SimpleNamespace(get_worker=lambda: 'B', get_time=lambda: '09:00~21:00',
                                           get_ot=lambda: '12000', get_meal=lambda: '1000')
        first = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 20), '중앙지사', prepare_only=True)
        app.cb_daily_test_method.set('RT')
        app.worker_group1 = app.worker_group2
        app.ent_daily_test_amount.set('3')
        app.ndt_ori_qty_var.set('3')
        app.worker_group2 = SimpleNamespace(get_worker=lambda: 'C', get_time=lambda: '13:00~21:00',
                                           get_ot=lambda: '12000', get_meal=lambda: '1000')
        second = app._add_single_usage_record_logic(10, datetime.date(2026, 10, 20), '중앙지사', prepare_only=True)
        plans = combine_plans([first, second], dict(shared(), User='저장 시 다른 사람', WorkTime='00:00~24:00'))
        a, b = [p['records_to_save'][0] for p in plans]
        self.assertEqual((a['User'], a['User2']), ('A', 'B'))
        self.assertEqual((b['User'], b['User2']), ('B', 'C'))
        self.assertEqual(plans[1]['workers_list'], ['B', 'C'])
        self.assertEqual(b['WorkTime'], '')
        self.assertEqual(b['WorkTime2'], '13:00~21:00')
        self.assertEqual(b['OT'], '')
        self.assertEqual(b['OT2'], '12000')
        self.assertEqual(participation_record(b)['WorkTime'], '09:00~21:00')
        self.assertEqual(participation_record(b)['OT'], '12000')
        self.assertEqual(a['일식'] + b['일식'], 3000)
        hours = sum((span[1] - span[0]) / 60 for r in (a, b)
                    for key in ('WorkTime', 'WorkTime2') if (span := work_interval(r[key])))
        self.assertEqual(hours, 29)
        commit_plans(app, plans)
        self.assertEqual(app.daily_usage_df.iloc[1]['참여인원'], 'B, C')
        self.assertIn('B: 09:00~21:00', app.daily_usage_df.iloc[1]['참여시간'])

    def test_partial_overlap_uses_existing_ot_rates_for_remaining_time(self):
        app = fixture()
        first, second = three_plans(app)[:2]
        for plan, time, amount in ((first, '(야간) 18:00~22:00', '16000'),
                                   (second, '(야간) 20:00~24:00', '18000')):
            plan['common_data'].update(User='A', WorkTime=time, OT=amount)
        tree = ast.parse((SRC / 'site_apps/central/src/models/worker_model.py').read_text(encoding='utf-8'))
        method = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_calculate_ot_from_worktime_impl')
        namespace = {'MARKER_PATTERN': re.compile(r'\(.*?\)\s*|익일')}
        exec(compile(ast.fix_missing_locations(ast.Module(body=[method], type_ignores=[])), '<actual OT rules>', 'exec'), namespace)
        calculate = lambda value, date: namespace['_calculate_ot_from_worktime_impl'](None, value, date)
        plans = combine_plans([first, second], shared(), calculate)
        a, b = [p['records_to_save'][0] for p in plans]
        self.assertEqual(b['WorkTime'], '(야간) 22:00~24:00')
        self.assertEqual(float(a['OT']) + float(b['OT']), 26000)
        self.assertEqual(participation_record(b)['WorkTime'], '(야간) 20:00~24:00')
        self.assertEqual(participation_record(b)['OT'], '18000')

    def test_night_shift_overlap_and_disjoint_intervals(self):
        app = fixture()
        first, second, third = three_plans(app)
        for plan, time, amount in ((first, '(야간) 22:00~익일09:00', '55000'),
                                   (second, '(야간) 23:00~익일06:00', '35000'),
                                   (third, '(주간) 09:00~18:00', '')):
            plan['common_data'].update(User='A', WorkTime=time, OT=amount)
        plans = combine_plans([first, second, third], shared())
        a, b, c = [p['records_to_save'][0] for p in plans]
        self.assertEqual(a['OT'], '55000')
        self.assertEqual(b['WorkTime'], '')
        self.assertEqual(b['OT'], '')
        self.assertEqual(c['WorkTime'], '(주간) 09:00~18:00')
        self.assertEqual(work_interval(a['WorkTime']), (1320, 1980))

    def test_same_time_uses_highest_entered_ot_without_losing_originals(self):
        app = fixture()
        plans = three_plans(app)[:2]
        plans[0]['common_data']['OT'] = '12000'
        plans[1]['common_data']['OT'] = '15000'
        combined = combine_plans(plans, shared())
        records = [p['records_to_save'][0] for p in combined]
        self.assertEqual(sum(float(r['OT'] or 0) for r in records), 15000)
        self.assertEqual(participation_record(records[0])['OT'], '12000')
        self.assertEqual(participation_record(records[1])['OT'], '15000')
        self.assertEqual(json_copy(combined)[1]['records_to_save'][0]['참여인원'], '작업자')

    def test_rt_80_total_and_rep_6_are_saved_without_double_counting(self):
        app = fixture()
        app.cb_daily_test_method.set('RT')
        app.ndt_report_pipe_var.set('100A')
        app.ndt_ori_joint_var.set('26')
        app.ndt_ori_qty_var.set('78')
        app.ndt_rep_joint_var.set('2')
        app.ndt_rep_qty_var.set('6')
        app.ndt_rej_joint_var.set('2')
        app.ent_daily_test_amount.set('80')
        app.ent_daily_unit_price.set('200')
        app.ent_daily_test_fee.set('16000')
        plan = app._add_single_usage_record_logic(10, datetime.date(2026, 10, 20), '중앙지사', prepare_only=True)
        ori, rep = plan['records_to_save']
        self.assertEqual(ori['Usage'], 80)
        self.assertEqual(ori['원본매수'], 78)
        self.assertEqual(ori['불량수'], '2')
        self.assertEqual(rep['재검사매수'], 6)
        self.assertEqual(rep['조인트수'], '2')
        self.assertEqual(rep['Usage'], 0)
        self.assertEqual(rep['검사량'], 0)
        self.assertEqual(rep['검사비'], 0)
        self.assertEqual(sum(r['Usage'] for r in plan['records_to_save']), 80)
        self.assertEqual(commit_plans(app, [plan]), 2)
        self.assertEqual(list(app.transactions_df['Quantity']), [-80])
        self.assertEqual(app.daily_usage_df['검사비'].sum(), 16000)

    def test_json_copy_handles_nested_excel_missing_values(self):
        import numpy as np
        data = {'vehicles': [{'mileage': float('nan'), 'remarks': pd.NA}],
                'optional': (np.float64('nan'), pd.NaT),
                'quantity': np.float64(8.9369), 'count': np.int64(7),
                'date': datetime.date(2026, 10, 19), 'note': 'nan'}
        result = json_copy(data)
        self.assertEqual(result['vehicles'], [{'mileage': None, 'remarks': None}])
        self.assertEqual(result['optional'], [None, None])
        self.assertEqual(result['quantity'], 8.9369)
        self.assertEqual(result['count'], 7)
        self.assertEqual(result['date'], '2026-10-19')
        self.assertEqual(result['note'], 'nan')
        self.assertTrue(math.isnan(data['vehicles'][0]['mileage']))
        json.dumps(result, allow_nan=False)
        with self.assertRaises(ValueError):
            json_copy({'invalid': float('inf')})

    def test_rt_100a_joint_traces_update_ori_rep_and_total(self):
        app = fixture()
        app.cb_daily_test_method.set('RT')
        app.SIZE_LENGTH = {'100A': .3591, '150A': .519, '250A': .8401}
        app.ndt_report_pipe_var.set('100A')
        app.ndt_ori_qty_var.set('80')
        app._calculate_daily_ndt_fee = Mock()
        app._init_ndt_vars_and_traces()
        app.ndt_ori_joint_var.set('26')
        self.assertEqual(app.ndt_ori_qty_var.get(), '78')
        self.assertEqual(app.ent_daily_test_amount.get(), '78')
        app._calculate_daily_ndt_fee.assert_called_once()
        app.ndt_rep_joint_var.set('2')
        self.assertEqual(app.ndt_rep_qty_var.get(), '6')
        self.assertEqual(app.ent_daily_test_amount.get(), '78')
        app.ndt_rej_joint_var.set('2')
        self.assertEqual(app.ent_daily_test_amount.get(), '78')
        self.assertEqual(app.ndt_ori_qty_var.get(), '78')
        self.assertEqual(app.ndt_rep_qty_var.get(), '6')
        app.ndt_rep_joint_var.set('')
        self.assertEqual(app.ndt_rep_qty_var.get(), '')
        self.assertEqual(app.ent_daily_test_amount.get(), '78')
        app.ndt_rej_joint_var.set('')
        self.assertEqual(app.ent_daily_test_amount.get(), '78')
        app.ndt_ori_joint_var.set('')
        self.assertEqual(app.ndt_ori_qty_var.get(), '')
        self.assertEqual(app.ent_daily_test_amount.get(), '0')

    def test_rt_large_pipe_does_not_use_paut_circumference(self):
        from site_apps.central.src.services.inspection_quantity import quantity_per_joint
        sizes = {'100A': .3591, '150A': .519, '250A': .8401}
        self.assertEqual(quantity_per_joint('RT', '100', sizes), 3)
        self.assertEqual(quantity_per_joint('RT', '150A', sizes), 3)
        self.assertIsNone(quantity_per_joint('RT', '250A', sizes))
        self.assertAlmostEqual(quantity_per_joint('PAUT', '250A', sizes), .8401)

    def test_save_flow_clears_only_after_success(self):
        for success in (True, False):
            with self.subTest(success=success):
                app = fixture()
                plans = three_plans(app)
                app.save_data.return_value = success
                app.ent_daily_test_amount.set('')
                app.clear_daily_usage_form_all = Mock()
                for method in ('update_daily_usage_view', 'update_recent_entries_view', 'refresh_inquiry_filters', 'update_stock_view'):
                    setattr(app, method, Mock())
                app.ent_daily_start_date = app.ent_daily_end_date = SimpleNamespace(get_date=lambda: shared()['Date'])
                batch = InspectionBatch.__new__(InspectionBatch)
                batch.app, batch.plans, batch.editing = app, plans, None
                batch.render = Mock()
                batch.shared = lambda: shared()
                with patch('site_apps.central.src.inspection_batch.messagebox') as dialogs:
                    batch.save()
                    if success:
                        dialogs.showerror.assert_not_called()
                        self.assertEqual(batch.plans, [])
                        app.clear_daily_usage_form_all.assert_called_once()
                    else:
                        dialogs.showerror.assert_called_once()
                        self.assertEqual(len(batch.plans), 3)
                        app.clear_daily_usage_form_all.assert_not_called()
                app.save_data.assert_called_once()

    def test_edit_restores_manual_price_after_diameter_recalculation(self):
        app = fixture()
        batch = InspectionBatch.__new__(InspectionBatch)
        batch.app, batch.editing = app, None
        current = {'date': '2026-10-20', 'site': '중앙지사', 'company': '발주처',
                   'workers': [], 'travel_cost': '500', 'meal_cost': '1000', 'vehicles': [], 'main_memo': ''}
        app._get_usage_session_data = Mock(return_value=current)
        app._set_usage_session_data = Mock()
        app.cb_daily_test_method.event_generate = Mock()
        app.entry_canvas = Mock()
        editor = dict(current, method='PAUT', amount='1.5', unit_price='999', test_fee='1498',
                      unit='M,I/D', equip='OMNI', ndt_vars={'ndt_report_pipe_var': '250A'}, ndt_costs={},
                      workers=[{'name': 'A', 'shift': '야간', 'time': '18:00~22:00', 'ot': '16000'}],
                      worker_meals={'1': '1000'})
        app.worker_group1.ent_meal = Field('2000')
        def diameter_changes_price(value):
            app.ent_daily_unit_price.set('10000')
            app.ent_daily_test_amount.set('10')
        app.ndt_report_pipe_var.set = diameter_changes_price
        batch.plans = [{'editor': editor}]
        batch.tree = SimpleNamespace(selection=lambda: ('0',))
        batch.render = Mock()
        batch.edit()
        self.assertEqual(app.ent_daily_unit_price.get(), '999')
        self.assertEqual(app.ent_daily_test_amount.get(), '1.5')
        self.assertEqual(app.ent_daily_test_fee.get(), '1498')
        self.assertEqual(batch.editing, 0)
        self.assertEqual(app._set_usage_session_data.call_args.args[0]['workers'][0]['name'], 'A')
        self.assertEqual(app.worker_group1.ent_meal.get(), '1000')

    def test_paut_diameter_selects_correct_contract_and_circumference(self):
        app = fixture()
        app.SIZE_LENGTH = {'250A': 0.8401, '400A': 1.2767, '100A': .3591}
        view = ast.parse((SRC / 'site_apps/central/src/views/daily_usage_view.py').read_text(encoding='utf-8'))
        setup = next(n for n in view.body if isinstance(n, ast.FunctionDef) and n.name == 'setup_daily_usage_tab_impl')
        functions = [n for n in setup.body if isinstance(n, ast.FunctionDef) and n.name in ('_calculate_ndt_fee', 'on_batch_pipe_change')]
        rates = ast.parse((SRC / 'site_apps/central/src/ndt_billing_tab.py').read_text(encoding='utf-8'))
        config = next(ast.literal_eval(n.value) for n in rates.body if isinstance(n, ast.Assign)
                      and any(isinstance(t, ast.Name) and t.id == 'DEFAULT_CONFIG' for t in n.targets))
        namespace = {'self': app, 'tk': SimpleNamespace(END='end')}
        exec(compile(ast.fix_missing_locations(ast.Module(body=functions, type_ignores=[])), '<actual PAUT calculation>', 'exec'), namespace)
        module = SimpleNamespace(MATERIAL_COST=config['MATERIAL_COST'], LABOR_COST=config['LABOR_COST'])
        with patch.dict(sys.modules, {'site_apps.central.src.ndt_billing_tab': module}):
            namespace['on_batch_pipe_change']()
            price250 = float(app.ent_daily_unit_price.get().replace(',', ''))
            self.assertEqual(app.ndt_pipe_var.get(), '250A (1.0)')
            self.assertEqual(app.ent_daily_test_amount.get(), '1.6802')
            app.ndt_report_pipe_var.set('400A')
            namespace['on_batch_pipe_change']()
            price400 = float(app.ent_daily_unit_price.get().replace(',', ''))
            self.assertEqual(app.ndt_pipe_var.get(), '300A 이상 (1.0)')
            self.assertEqual(app.ent_daily_test_amount.get(), '2.5534')
            self.assertGreater(price250, price400)
            app.cb_daily_test_method.set('RT')
            app.ndt_report_pipe_var.set('100A')
            app.ndt_ori_joint_var.set('26')
            app.ndt_rep_joint_var.set('2')
            app.ndt_ori_qty_var.set('80')
            app.ndt_rej_joint_var.set('2')
            namespace['_calculate_ndt_fee']()
            self.assertEqual(app.ndt_ori_qty_var.get(), '78')
            self.assertEqual(app.ndt_rep_qty_var.get(), '6')
            self.assertEqual(app.ent_daily_test_amount.get(), '78')

    def test_staging_three_methods_keeps_each_job_and_clears_next_quantity(self):
        app = fixture()
        app.cb_daily_site = Field('중앙지사')
        app.ent_daily_date = SimpleNamespace(get_date=lambda: datetime.date(2026, 10, 20))
        batch = InspectionBatch.__new__(InspectionBatch)
        batch.app, batch.plans, batch.editing = app, [], None
        batch.render = Mock()
        batch.snapshot = Mock(return_value={'method': 'PAUT', 'vehicles': [{'mileage': float('nan')}]})
        with patch('site_apps.central.src.inspection_batch.messagebox') as dialogs:
            batch.add()
            self.assertEqual(len(batch.plans), 1)
            self.assertEqual(app.ent_daily_test_amount.get(), '')
            self.assertEqual(batch.plans[0]['records_to_save'][0]['NDT_세척제'], 0)
            app.ndt_report_pipe_var.set('400A')
            app.ndt_pipe_var.set('300A 이상 (1.0)')
            app.ent_daily_test_amount.set('2.5534')
            app.ent_daily_unit_price.set('80')
            batch.add()
            app.cb_daily_test_method.set('RT')
            app.cb_daily_material.set('Carestream T200')
            app.ent_daily_test_amount.set('54')
            app.ent_daily_unit_price.set('200')
            app.cb_daily_unit.set('매')
            batch.add()
            dialogs.showwarning.assert_not_called()
        self.assertEqual(len(batch.plans), 3)
        self.assertEqual([p['common_data']['검사방법'] for p in batch.plans], ['PAUT', 'PAUT', 'RT'])
        self.assertEqual(batch.plans[2]['common_data']['MaterialID'], 10)
        self.assertIsNone(batch.plans[0]['editor']['vehicles'][0]['mileage'])
        app.save_data.assert_not_called()

    def test_prepare_is_read_only(self):
        app = fixture()
        plans = three_plans(app)
        self.assertEqual([p['common_data']['검사방법'] for p in plans], ['PAUT', 'PAUT', 'RT'])
        self.assertEqual([p['records_to_save'][0]['Usage'] for p in plans], [1.6802, 2.5534, 54])
        self.assertEqual(plans[1]['common_data']['조건1'], '300A 이상 (1.0)')
        self.assertTrue(app.daily_usage_df.empty)
        self.assertTrue(app.transactions_df.empty)
        self.assertEqual(app.sites, [])
        app.save_data.assert_not_called()
        app.save_tab_config.assert_not_called()

    def test_three_jobs_commit_once_and_stock_only_for_rt(self):
        app = fixture()
        original = three_plans(app)
        plans = combine_plans(original, shared())
        count = commit_plans(app, plans)
        self.assertEqual(count, 3)
        app.save_data.assert_called_once()
        df = app.daily_usage_df
        self.assertEqual(list(df['검사방법']), ['PAUT', 'PAUT', 'RT'])
        self.assertEqual(list(df['단가']), [100, 80, 200])
        self.assertEqual(list(df['조건1'])[:2], ['250A (1.0)', '300A 이상 (1.0)'])
        self.assertEqual(list(df['관경(Inch)']), ['250A', '400A', '250A'])
        self.assertEqual(list(df['출장비']), [500, 0, 0])
        self.assertEqual(list(df['일식']), [1000, 0, 0])
        self.assertEqual(list(df['WorkTime']), ['09:00~18:00', '', ''])
        self.assertEqual(list(df['OT']), ['12000', '', ''])
        self.assertEqual(list(df['주행거리']), ['30', '', ''])
        self.assertAlmostEqual(df['검사비'].sum(), 168.02 + 204.272 + 10800 + 500)
        self.assertEqual(list(app.transactions_df['Quantity']), [-54])
        app._auto_reconcile_and_register_ndt.assert_called_once()
        self.assertEqual(original[1]['records_to_save'][0]['WorkTime'], '09:00~18:00')

    def test_failure_restores_stock_and_daily_records_and_retry_is_clean(self):
        app = fixture()
        plans = combine_plans(three_plans(app), shared())
        before = app.materials_df.copy(deep=True)
        app.save_data.return_value = False
        with self.assertRaises(RuntimeError):
            commit_plans(app, plans)
        self.assertTrue(app.daily_usage_df.empty)
        self.assertTrue(app.transactions_df.empty)
        pd.testing.assert_frame_equal(app.materials_df, before)
        self.assertEqual(app.sites, [])
        self.assertEqual(app.daily_units, [])
        app.save_data.return_value = True
        self.assertEqual(commit_plans(app, plans), 3)
        self.assertEqual(len(app.daily_usage_df), 3)
        self.assertEqual(len(app.transactions_df), 1)

    def test_sessions_roundtrip_independent_conditions(self):
        app = fixture()
        result = json.loads(json.dumps(json_copy(three_plans(app))))
        self.assertEqual(result[0]['common_data']['Date'], '2026-10-20')
        self.assertEqual(result[1]['common_data']['관경(Inch)'], '400A')
        self.assertEqual(result[2]['common_data']['MaterialID'], 10)
        self.assertEqual(commit_plans(app, combine_plans(result, shared())), 3)

    def test_legacy_single_save_still_works(self):
        app = fixture()
        count = app._add_single_usage_record_logic('OMNI', datetime.date(2026, 10, 20), '중앙지사')
        self.assertEqual(count, 1)
        app.save_data.assert_called_once()
        self.assertEqual(app.daily_usage_df.iloc[0]['WorkTime'], '09:00~18:00')

    def test_invalid_numbers(self):
        for value in ('nan', 'inf', '-1', 'bad'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                number(value, '수량')
        with self.assertRaises(ValueError):
            number('', '단가', required=True)
        self.assertEqual(number('12,000', '단가'), 12000)


if __name__ == '__main__':
    unittest.main()
