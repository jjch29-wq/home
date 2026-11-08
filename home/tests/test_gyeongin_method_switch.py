import sys, os
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
os.environ["PMI_SITE"] = "gyeongin"

import tkinter as tk
from site_apps.gyeongin.src.app import MaterialManager
import site_apps.gyeongin.src.data_sync as data_sync

def test_view_logic():
    # Mock data dir so it doesn't query unmounted G: drive in sandbox
    with mock.patch.object(data_sync, '_configured_data_dir', lambda: None):
        root = tk.Tk()
        root.withdraw()
        try:
            print("[INIT] Creating MaterialManager...")
            app = MaterialManager(root)
            print("[1] Initial state check:")
            curr_m = app.cb_daily_test_method.get().strip().upper()
            print(f"  Method: {curr_m}")
            assert curr_m == 'RT', f"Expected RT, got {curr_m}"
            
            rtk_info = app.rtk_grid.grid_info()
            ndt_info = app.ndt_frame.grid_info()
            guide_info = app.empty_guide_frame.grid_info()
            calc_info = app.ndt_calc_frame.grid_info()
            
            assert bool(rtk_info), "RTK grid should be gridded"
            assert not bool(ndt_info), "NDT chemical frame should NOT be gridded"
            assert not bool(guide_info), "Guide frame should NOT be gridded"
            assert bool(calc_info), "NDT calc frame should be gridded"
            print("  => OK (RT has RTK & NDT calc)")
            
            print("[2] Switch to PT:")
            app.cb_daily_test_method.set('PT')
            app.on_method_change_auto_unit_logic()
            
            rtk_info = app.rtk_grid.grid_info()
            ndt_info = app.ndt_frame.grid_info()
            guide_info = app.empty_guide_frame.grid_info()
            calc_info = app.ndt_calc_frame.grid_info()
            
            assert not bool(rtk_info), "RTK grid should NOT be gridded"
            assert bool(ndt_info), "NDT chemical frame SHOULD be gridded"
            assert not bool(guide_info), "Guide frame should NOT be gridded"
            assert bool(calc_info), "NDT calc frame SHOULD be gridded"
            print(f"  NDT company entries count: {len(app.ndt_company_entries)}")
            assert len(app.ndt_company_entries) >= 1
            print("  => OK (PT shows chemical frame)")
            
            print("[3] Switch to MT:")
            app.cb_daily_test_method.set('MT')
            app.on_method_change_auto_unit_logic()
            
            rtk_info = app.rtk_grid.grid_info()
            ndt_info = app.ndt_frame.grid_info()
            guide_info = app.empty_guide_frame.grid_info()
            calc_info = app.ndt_calc_frame.grid_info()
            
            assert not bool(rtk_info), "RTK grid should NOT be gridded"
            assert bool(ndt_info), "NDT chemical frame SHOULD be gridded"
            assert not bool(guide_info), "Guide frame should NOT be gridded"
            print("  => OK (MT shows chemical frame)")

            print("[4] Switch to PAUT:")
            app.cb_daily_test_method.set('PAUT')
            app.on_method_change_auto_unit_logic()
            
            rtk_info = app.rtk_grid.grid_info()
            ndt_info = app.ndt_frame.grid_info()
            guide_info = app.empty_guide_frame.grid_info()
            calc_info = app.ndt_calc_frame.grid_info()
            
            assert not bool(rtk_info), "RTK grid should NOT be gridded"
            assert not bool(ndt_info), "NDT chemical frame should NOT be gridded"
            assert bool(guide_info), "Guide frame SHOULD be gridded"
            print("  => OK (PAUT shows guide frame)")

            print("[5] Switch back to RT:")
            app.cb_daily_test_method.set('RT')
            app.on_method_change_auto_unit_logic()
            
            assert bool(app.rtk_grid.grid_info()), "RTK grid should be gridded"
            assert not bool(app.ndt_frame.grid_info()), "NDT chemical frame should NOT be gridded"
            assert not bool(app.empty_guide_frame.grid_info()), "Guide frame should NOT be gridded"
            print("  => OK (Back to RT)")
            
            print("[6] Test Add / Remove Company in Chemical Frame:")
            initial_comp_count = len(app.ndt_company_entries)
            app.add_ndt_company_section()
            assert len(app.ndt_company_entries) == initial_comp_count + 1
            app.remove_last_ndt_company()
            assert len(app.ndt_company_entries) == initial_comp_count
            print("  => OK (Company add/remove works)")
            
            print("[7] Test clear_daily_usage_form_all:")
            app.clear_daily_usage_form_all()
            assert app.cb_daily_test_method.get().strip().upper() == 'RT'
            assert bool(app.rtk_grid.grid_info()), "RTK grid should remain gridded after clear"
            assert bool(app.ndt_calc_frame.grid_info()), "NDT calc frame should remain gridded after clear"
            print("  => OK (Clear restores RT and RTK/calc panels)")
            
            print("\nALL METHOD SWITCHING & PANEL TESTS PASSED SUCCESSFULLY!")
        finally:
            root.destroy()

if __name__ == '__main__':
    test_view_logic()
