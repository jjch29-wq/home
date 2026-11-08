"""Billing history persistence must preserve records across failure and migration."""
import ast
from copy import deepcopy
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, mock_open, patch
import unittest

SOURCE=Path(__file__).resolve().parents[1]/'src/site_apps/central/src/ndt_billing_tab.py'
TREE=ast.parse(SOURCE.read_text(encoding='utf-8'))


def functions(names, ns):
    for name in names:
        fn=next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef) and n.name==name)
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),ns)
    return ns


class BillingPersistenceTests(unittest.TestCase):
    def test_legacy_config_keeps_records_and_existing_prices(self):
        default={'BILLING_RECORDS':[], 'MATERIAL_COST':{'new':3,'old':2},'LABOR_COST':{}}
        ns=functions(['load_config'],{'json':json,'deepcopy':deepcopy,
            'os':SimpleNamespace(path=SimpleNamespace(exists=lambda _:True)),
            'CONFIG_FILE':'unused','DEFAULT_CONFIG':default})
        opened=mock_open(read_data=json.dumps({'BILLING_RECORDS':[{'qty':10}],'MATERIAL_COST':{'old':99}}))
        with patch('builtins.open',opened): result=ns['load_config']()
        self.assertEqual(result['BILLING_RECORDS'],[{'qty':10}])
        self.assertEqual(result['MATERIAL_COST'],{'new':3,'old':99})
        self.assertFalse(any(len(c.args)>1 and c.args[1]=='w' for c in opened.call_args_list))

    def test_failed_save_keeps_global_records_and_redraws_saved_history(self):
        config={'BILLING_RECORDS':[{'qty':10}]}
        messages=Mock()
        ns=functions(['save_billing_records'],{'deepcopy':deepcopy,'CONFIG':config,
            'save_config':Mock(side_effect=OSError('locked')),'messagebox':messages})
        app=SimpleNamespace(records=[],_capture_billing_state=lambda:{'round':2},load_billing_records=Mock())
        self.assertFalse(ns['save_billing_records'](app))
        self.assertEqual(config['BILLING_RECORDS'],[{'qty':10}])
        app.load_billing_records.assert_called_once()
        messages.showerror.assert_called_once()

    def test_success_saves_round_and_records_in_one_snapshot(self):
        config={'BILLING_RECORDS':[{'qty':10}]}
        save=Mock()
        ns=functions(['save_billing_records'],{'deepcopy':deepcopy,'CONFIG':config,
            'save_config':save,'messagebox':Mock()})
        state={'round':2,'contract':{'RT':{'p_qty':'10','prev':'100'}}}
        app=SimpleNamespace(records=[],_capture_billing_state=lambda:state)
        self.assertTrue(ns['save_billing_records'](app))
        self.assertEqual(save.call_args.args[0]['BILLING_STATE'],state)
        self.assertEqual(config['BILLING_RECORDS'],[])
        app.records.append({'qty':99})
        self.assertEqual(config['BILLING_RECORDS'],[])

    def test_atomic_write_replaces_only_after_serialization_success(self):
        os=Mock()
        ns=functions(['save_config'],{'os':os,'json':json,'CONFIG_FILE':'unused'})
        with patch('builtins.open',mock_open()):
            with self.assertRaises(TypeError): ns['save_config']({'bad':object()})
        os.replace.assert_not_called()

if __name__=='__main__': unittest.main()
