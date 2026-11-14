"""Detached inventory uses the main view's current column schema."""
import ast
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import tkinter as tk
from tkinter import ttk
import unittest


class StockPopupTests(unittest.TestCase):
    def test_popup_preserves_columns_order_and_quantity_alignment(self):
        root = tk.Tk()
        root.withdraw()
        self.addCleanup(root.destroy)
        source = Path(__file__).resolve().parents[1] / 'src/site_apps/central/src/app.py'
        fn = next(n for n in ast.walk(ast.parse(source.read_text(encoding='utf-8')))
                  if isinstance(n, ast.FunctionDef) and n.name == 'open_detached_stock_view')
        ns = {'tk':tk, 'ttk':ttk}
        exec(compile(ast.Module(body=[fn],type_ignores=[]),str(source),'exec'),ns)
        columns = ('No.', '회사코드', '공급업체', '입고수량', '사용수량', '재고수량')
        tree = ttk.Treeview(root, columns=columns, show='headings')
        for c in columns:
            tree.heading(c,text=c)
            tree.column(c,width=123)
        tree['displaycolumns'] = tuple(reversed(columns))
        app = SimpleNamespace(root=root,stock_tree=tree,treeview_sort_column=Mock(),
                              enable_tree_column_drag=Mock())
        row = (1,'central','supplier',100,20,80,'hidden-id')
        def refresh():
            app.detached_windows['stock']['tree'].insert('', 'end',values=row)
        app.update_stock_view=refresh
        ns['open_detached_stock_view'](app)
        popup=app.detached_windows['stock']['tree']
        self.assertEqual(popup['columns'],tree['columns'])
        self.assertEqual(popup['displaycolumns'],tree['displaycolumns'])
        item=popup.get_children()[0]
        self.assertEqual(popup.set(item,'재고수량'),'80')
        self.assertEqual(popup.set(item,'공급업체'),'supplier')
        self.assertEqual(popup.column('재고수량','width'),123)


if __name__ == '__main__': unittest.main()
