from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.views.tree_column_resize import enable_column_resize


class Tree:
    def __init__(self):
        self.bindings = {}
        self.region = 'separator'
        self.columns = {'#2': {'width': 75, 'minwidth': 60}}

    def bind(self, event, callback, **kwargs): self.bindings[event] = callback
    def identify_region(self, x, y): return self.region
    def identify_column(self, x): return '#2'
    def column(self, column, option=None, **kwargs):
        self.columns[column].update(kwargs)
        return self.columns[column].get(option)


class ColumnResizeTests(unittest.TestCase):
    def event(self, root_x): return SimpleNamespace(x=80, y=5, x_root=root_x)

    def test_drag_resizes_selected_column_and_respects_minimum(self):
        tree = Tree()
        enable_column_resize(tree)
        self.assertEqual(tree.bindings['<Button-1>'](self.event(100)), 'break')
        tree.bindings['<B1-Motion>'](self.event(160))
        self.assertEqual(tree.columns['#2']['width'], 135)
        tree.bindings['<B1-Motion>'](self.event(0))
        self.assertEqual(tree.columns['#2']['width'], 60)
        self.assertEqual(tree.bindings['<ButtonRelease-1>'](self.event(0)), 'break')
        self.assertIsNone(tree.bindings['<B1-Motion>'](self.event(300)))
        self.assertEqual(tree.columns['#2']['width'], 60)

    def test_normal_header_and_row_clicks_keep_default_behavior(self):
        tree = Tree()
        enable_column_resize(tree)
        for region in ('heading', 'cell', 'nothing'):
            tree.region = region
            self.assertIsNone(tree.bindings['<Button-1>'](self.event(100)))
            self.assertIsNone(tree.bindings['<B1-Motion>'](self.event(200)))
        self.assertEqual(tree.columns['#2']['width'], 75)


if __name__ == '__main__': unittest.main()
