"""Resaving a loaded job must target its source, even after quantity changes."""
from pathlib import Path
import sys
from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest
import pandas as pd
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.daily_record_save import remember_source, save_prepared


class DailyRecordSaveTests(unittest.TestCase):
    def setUp(self):
        self.app = SimpleNamespace(root=Mock(), daily_usage_df=pd.DataFrame([
            {'Date':'2026-10-23','Site':'A','Usage':10},
            {'Date':'2026-10-23','Site':'A','Usage':20}]),
            transactions_df=pd.DataFrame(),materials_df=pd.DataFrame(),monthly_usage_df=pd.DataFrame(),
            sites=[],users=[],vehicles=[],daily_units=[],equipments=[],save_data=Mock(return_value=True))
        self.plan = {'records_to_save':[{'Date':'2026-10-23','Site':'A','Usage':15}]}
        def commit(plan, auto_save=False):
            self.app.daily_usage_df = pd.concat([self.app.daily_usage_df, pd.DataFrame(plan['records_to_save'])],ignore_index=True)
            return len(plan['records_to_save'])
        self.app._commit_usage_plan = commit
        remember_source(self.app, 0)

    def test_quantity_change_overwrites_only_loaded_record(self):
        with patch('site_apps.central.src.views.daily_record_save.messagebox.askyesnocancel',return_value=True):
            self.assertEqual(save_prepared(self.app,self.plan),1)
        self.assertEqual(list(self.app.daily_usage_df['Usage']),[20,15])
        self.app.save_data.assert_called_once_with(require_primary=True)

    def test_choose_append_preserves_both_originals(self):
        with patch('site_apps.central.src.views.daily_record_save.messagebox.askyesnocancel',return_value=False):
            save_prepared(self.app,self.plan)
        self.assertEqual(list(self.app.daily_usage_df['Usage']),[10,20,15])

    def test_cancel_does_not_mutate(self):
        with patch('site_apps.central.src.views.daily_record_save.messagebox.askyesnocancel',return_value=None):
            self.assertEqual(save_prepared(self.app,self.plan),0)
        self.assertEqual(list(self.app.daily_usage_df['Usage']),[10,20])
        self.app.save_data.assert_not_called()

    def test_failed_overwrite_restores_original_and_source_for_retry(self):
        self.app.save_data.return_value = False
        with patch('site_apps.central.src.views.daily_record_save.messagebox.askyesnocancel',return_value=True):
            with self.assertRaises(RuntimeError): save_prepared(self.app,self.plan)
        self.assertEqual(list(self.app.daily_usage_df['Usage']),[10,20])
        self.assertIsNotNone(self.app._daily_edit_source)

    def test_stale_source_cannot_overwrite_another_record(self):
        self.app.daily_usage_df.loc[0,'Usage'] = 99
        with self.assertRaises(ValueError): save_prepared(self.app,self.plan)
        self.app.save_data.assert_not_called()

    def test_fresh_exact_duplicate_requires_confirmation(self):
        self.app._daily_edit_source = None
        plan = {'records_to_save':[{'Date':'2026-10-23','Site':'A','Usage':10}]}
        with patch('site_apps.central.src.views.daily_record_save.messagebox.askyesno',return_value=False) as ask:
            self.assertEqual(save_prepared(self.app,plan),0)
        ask.assert_called_once()
        self.app.save_data.assert_not_called()

if __name__ == '__main__': unittest.main()
