"""Work-log history failures and cancellation must preserve saved data and drafts."""
import ast
import json
import os
import shutil
import tempfile
import uuid
from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, mock_open, patch
import unittest

SOURCE=Path(__file__).resolve().parents[1]/'src/site_apps/central/src/daily_work_log_tab.py'
TREE=ast.parse(SOURCE.read_text(encoding='utf-8'))


class WorkLogSafetyTests(unittest.TestCase):
    def setUp(self):
        self.messages=Mock()
        self.dialog=Mock()
        self.ns={'json':json,'uuid':uuid,'deepcopy':deepcopy,'messagebox':self.messages,'filedialog':self.dialog}
        for name in ('load_history','on_date_change','export_excel','save_history','delete_current_date_data'):
            fn=next(n for n in ast.walk(TREE) if isinstance(n,ast.FunctionDef) and n.name==name)
            exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),self.ns)

    def test_corrupt_json_is_not_treated_as_empty(self):
        with patch('builtins.open',mock_open(read_data='{broken')):
            with self.assertRaises(json.JSONDecodeError):
                self.ns['load_history'](SimpleNamespace(history_path='unused'))
        self.messages.showerror.assert_called_once()

    def test_delete_from_stale_editor_does_not_remove_newer_date(self):
        date = '2026-10-23'
        latest = {date: {'remarks': 'newer'}}
        app = SimpleNamespace(date_entry=Mock(get=Mock(return_value=date)), load_history=lambda:latest,
                              _loaded_log_date=date, _saved_log_fields={'remarks':'old'}, save_history=Mock())
        self.ns['delete_current_date_data'](app)
        self.messages.showerror.assert_called_once()
        self.messages.askyesno.assert_not_called()
        app.save_history.assert_not_called()
        self.assertEqual(latest[date]['remarks'], 'newer')

    def test_history_changed_during_serialization_is_not_overwritten(self):
        self.ns.update(os=os, shutil=shutil, get_data_connection_error=lambda:None)
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'history.json'
            path.write_text('{"2026-10-23":{"remarks":"old"}}', encoding='utf-8')
            version = path.stat().st_mtime
            app = SimpleNamespace(history_path=str(path), _history_cache_mtime=version)
            newer = '{"2026-10-23":{"remarks":"newer"}}'
            serialize = json.dump

            def serialize_then_change(data, stream, **kwargs):
                serialize(data, stream, **kwargs)
                path.write_text(newer, encoding='utf-8')
                os.utime(path, (version + 2, version + 2))

            with patch.object(json, 'dump', side_effect=serialize_then_change):
                with self.assertRaisesRegex(OSError, '다른 곳에서 작업일보를 변경'):
                    self.ns['save_history'](app, {'2026-10-23': {'remarks': 'draft'}})
            self.assertEqual(path.read_text(encoding='utf-8'), newer)
            self.assertEqual(list(path.parent.glob('*.tmp')), [])
            self.messages.showerror.assert_called_once()

    def test_missing_first_run_history_is_empty(self):
        with patch('builtins.open',side_effect=FileNotFoundError):
            self.assertEqual(self.ns['load_history'](SimpleNamespace(history_path='unused')), {})

    def test_cancel_date_change_preserves_editor(self):
        app=SimpleNamespace(_selected_date_key=lambda:'2026-10-24', _loaded_log_date='2026-10-23',
            _log_baseline={'remarks':'saved'},_capture_log_input=lambda:{'remarks':'draft'},
            date_entry=Mock(),_clear_ndt_selection=Mock(),load_history=Mock())
        self.messages.askyesno.return_value=False
        self.ns['on_date_change'](app)
        app.date_entry.set_date.assert_called_once_with('2026-10-23')
        app.load_history.assert_not_called()
        app._clear_ndt_selection.assert_not_called()

    def test_photo_save_keeps_unsaved_editor_baseline_and_date(self):
        os=Mock()
        os.path.isfile.return_value=False
        os.path.exists.return_value=False
        os.path.getmtime.side_effect=FileNotFoundError
        self.ns.update(os=os,shutil=Mock(),get_data_connection_error=lambda:None)
        baseline={'remarks':'saved'}
        app=SimpleNamespace(history_path='unused',_log_baseline=baseline.copy(),
            _loaded_log_date='2026-10-23',_capture_log_input=lambda:{'remarks':'draft'},
            _selected_date_key=lambda:'2026-10-23',
            _refresh_ndt_history_choices=Mock(),_refresh_saved_date_markers=Mock(),_update_date_status=Mock())
        with patch('builtins.open',mock_open()):
            self.ns['save_history'](app,{'2026-10-23':{'remarks':'saved','process_photos':[]}})
        self.assertEqual(app._log_baseline,baseline)
        app._selected_date_key=lambda:'2026-10-24'
        app.date_entry=Mock()
        app.load_history=Mock()
        self.messages.askyesno.return_value=False
        self.ns['on_date_change'](app)
        self.messages.askyesno.assert_called_once()
        app.load_history.assert_not_called()

    def test_full_editor_save_updates_baseline_only_after_success(self):
        cls=next(n for n in TREE.body if isinstance(n,ast.ClassDef) and n.name=='DailyWorkLogTab')
        fn=[n for n in cls.body if isinstance(n,ast.FunctionDef) and n.name=='save_current_history'][-1]
        self.ns['tk']=SimpleNamespace(END='end')
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),self.ns)
        app=SimpleNamespace(_validate_rt_rows=lambda:True,_validate_hot_rows=lambda:True,
            load_history=lambda:{},_selected_date_key=lambda:'2026-10-23',
            weather_entry=SimpleNamespace(get=lambda:'sunny'),remarks_text=SimpleNamespace(get=lambda *args:'draft'),
            qty_entries={},equip_entries={},personnel_entries={},ndt_grid_entries=[],save_history=Mock(),
            _capture_log_input=lambda:{'remarks':'draft'},_log_baseline={'remarks':'saved'})
        app.save_history.side_effect=OSError('locked')
        with patch('builtins.open',mock_open()):
            with self.assertRaises(OSError): self.ns['save_current_history'](app)
        self.assertEqual(app._log_baseline,{'remarks':'saved'})
        app.save_history.side_effect=None
        with patch('builtins.open',mock_open()): self.ns['save_current_history'](app)
        self.assertEqual(app._log_baseline,{'remarks':'draft'})

    def test_stale_editor_does_not_overwrite_newer_saved_date(self):
        cls=next(n for n in TREE.body if isinstance(n,ast.ClassDef) and n.name=='DailyWorkLogTab')
        fn=[n for n in cls.body if isinstance(n,ast.FunctionDef) and n.name=='save_current_history'][-1]
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(SOURCE),'exec'),self.ns)
        date='2026-10-23'
        latest={date:{'remarks':'newer'}}
        app=SimpleNamespace(_validate_rt_rows=lambda:True, load_history=lambda:latest,
                            _selected_date_key=lambda:date, _loaded_log_date=date,
                            _saved_log_fields={'remarks':'original'}, save_history=Mock())
        with self.assertRaisesRegex(OSError, '다른 곳에서 변경'):
            self.ns['save_current_history'](app)
        app.save_history.assert_not_called()
        self.assertEqual(latest[date]['remarks'],'newer')
        self.messages.showerror.assert_called_once()

    def test_cancel_export_does_not_save_history(self):
        app=SimpleNamespace(_validate_rt_rows=lambda:True,_validate_hot_rows=lambda:True,
            date_entry=SimpleNamespace(get=lambda:'2026-10-23'),
            auto_calculate_and_save=Mock(),save_current_history=Mock())
        self.dialog.asksaveasfilename.return_value=''
        self.ns['export_excel'](app)
        app.auto_calculate_and_save.assert_not_called()
        app.save_current_history.assert_not_called()

if __name__=='__main__': unittest.main()
