import sys
from pathlib import Path
import tkinter as tk
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.tk_lifecycle import install_root_timer_cleanup


class RootTimerCleanupTests(unittest.TestCase):
    def test_root_close_cancels_pending_child_and_root_callbacks(self):
        root = tk.Tk()
        root.withdraw()
        install_root_timer_cleanup(root)
        install_root_timer_cleanup(root)
        child = tk.Frame(root)
        root.after(60000, lambda: None)
        child.after(60000, lambda: None)
        root.after_idle(lambda: None)
        root.destroy()
        self.assertEqual(root.tk.splitlist(root.tk.call('after', 'info')), ())

    def test_closing_child_keeps_other_windows_callbacks(self):
        root = tk.Tk()
        root.withdraw()
        try:
            install_root_timer_cleanup(root)
            child = tk.Toplevel(root)
            timer = root.after(60000, lambda: None)
            child.destroy()
            self.assertIn(timer, root.tk.splitlist(root.tk.call('after', 'info')))
        finally:
            root.destroy()


if __name__ == '__main__':
    unittest.main()
