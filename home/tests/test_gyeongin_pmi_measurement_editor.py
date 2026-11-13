import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from tkinter import ttk

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.gyeongin.src.pmi_measurement_editor import open_measurement_editor, import_measurement_files
from site_apps.gyeongin.src.cert_exporter import read_pmi_measurements


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


class MeasurementEditorTests(unittest.TestCase):
    def test_selected_rows_bind_to_the_selected_material_criterion(self):
        root = tk.Tk()
        root.withdraw()
        criteria = [dict(id='304', label='TP304', ranges={'Ni': '8~11'}),
                    dict(id='316', label='TP316', ranges={'Ni': '10~14'})]
        batch = dict(method='PMI', pmi_elements=['Ni'], pmi_criteria=criteria,
                     pmi_measurements=[dict(elements={'Ni': '8'})])
        try:
            with tempfile.TemporaryDirectory() as tmp:
                open_measurement_editor(root, batch, Path(tmp) / 'state.json', lambda: True, lambda: None)
                win = next(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))
                tree = next(w for w in descendants(win) if isinstance(w, ttk.Treeview))
                self.assertEqual(tree.set('0', 'result'), 'HOLD')
                tree.selection_set('0')
                next(w for w in descendants(win) if isinstance(w, ttk.Button)
                     and str(w.cget('text')) == '선택 행에 적용').invoke()
                self.assertEqual(tree.set('0', 'criterion'), 'TP304')
                self.assertEqual(tree.set('0', 'result'), 'PASS')
                next(w for w in descendants(win) if isinstance(w, ttk.Button)
                     and str(w.cget('text')) == '저장').invoke()
                self.assertEqual(batch['pmi_measurements'][0]['criterion_id'], '304')
        finally:
            root.destroy()

    def test_editor_displays_only_selected_elements_in_selected_order(self):
        root = tk.Tk()
        root.withdraw()
        batch = dict(method='PMI', pmi_elements=['Cr', 'Ni'], pmi_measurements=[
            dict(tag='T1', elements={'Ni': '8.2', 'Cr': '18.5', 'Mo': '0.1', 'Fe': '70'})])
        try:
            with tempfile.TemporaryDirectory() as tmp:
                open_measurement_editor(root, batch, Path(tmp) / 'state.json', lambda: True, lambda: None)
                win = next(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))
                tree = next(w for w in descendants(win) if isinstance(w, ttk.Treeview))
                self.assertEqual(tuple(tree['columns']),
                                 ('include', 'reading_no', 'date', 'source', 'tag', 'component', 'alloy', 'criterion', 'result', 'Cr', 'Ni'))
                tree.selection_set('0')
                next(w for w in descendants(win) if isinstance(w, ttk.Button)
                     and str(w.cget('text')) == '선택 행 편집').invoke()
                editor = next(w for w in win.winfo_children() if isinstance(w, tk.Toplevel))
                labels = [str(w.cget('text')) for w in descendants(editor) if isinstance(w, ttk.Label)]
                self.assertIn('Cr', labels)
                self.assertIn('Ni', labels)
                self.assertNotIn('Mo', labels)
                self.assertNotIn('Fe', labels)
                editor.destroy()
                next(w for w in descendants(win) if isinstance(w, ttk.Button)
                     and str(w.cget('text')) == '저장').invoke()
                self.assertEqual(batch['pmi_measurements'][0]['elements']['Fe'], '70')
        finally:
            root.destroy()

    def test_import_registers_without_editor_and_rolls_back_failed_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'readings.csv'
            path.write_text('Time,Reading No,Ni,Cr,Mo\n2026-02-04,12,8.2,18.5,0.1\n', encoding='utf-8')
            batch = dict(method='PMI')
            with patch('site_apps.gyeongin.src.pmi_measurement_editor.filedialog.askopenfilenames',
                       return_value=(str(path),)), patch(
                       'site_apps.gyeongin.src.pmi_measurement_editor.tk.Toplevel') as window:
                imported = []
                import_measurement_files(None, batch, lambda: False, lambda: imported.append(True))
                self.assertEqual(batch, dict(method='PMI'))
                self.assertEqual(imported, [])
                import_measurement_files(None, batch, lambda: True, lambda: imported.append(True))
                self.assertEqual(imported, [True])
                self.assertEqual(batch['pmi_measurements'][0]['reading_no'], '12')
                self.assertEqual(batch['pmi_measurements'][0]['measurement_date'], '2026-02-04')
                import_measurement_files(None, batch, lambda: True, lambda: None)
                self.assertEqual(len(batch['pmi_measurements']), 1)
                window.assert_not_called()

    def test_reimport_fills_missing_source_metadata_without_replacing_edits(self):
        root = tk.Tk()
        root.withdraw()
        try:
            with tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / 'old.csv'
                path.write_text('Timestamp,Reading_No,Tag,Ni,Cr,Mo\n2026-02-04 09:30,42,T1,8.2,18.5,0.1\n', encoding='utf-8')
                old_row = read_pmi_measurements(path)[0]
                old_row.pop('reading_no')
                old_row['measurement_date'] = ''
                old_row['source_file'] = path.name
                old_row['source_row'] = 1
                old_row['elements']['Ni'] = '9.5'
                batch = dict(method='PMI', pmi_measurements=[old_row])
                with patch('site_apps.gyeongin.src.pmi_measurement_editor.filedialog.askopenfilenames',
                           return_value=(str(path),)), patch(
                           'site_apps.gyeongin.src.pmi_measurement_editor.messagebox.showinfo'):
                    open_measurement_editor(root, batch, Path(tmp) / 'state.json',
                                            lambda: True, lambda: None, import_files=True)
                win = next(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))
                tree = next(w for w in descendants(win) if isinstance(w, ttk.Treeview))
                self.assertEqual(tree.get_children(), ('0',))
                self.assertEqual(tree.set('0', 'reading_no'), '42')
                self.assertEqual(tree.set('0', 'date'), '2026-02-04 09:30')
                self.assertEqual(tree.set('0', 'Ni'), '9.5')
                self.assertEqual(batch['pmi_measurements'][0]['measurement_date'], '')
                next(w for w in descendants(win) if isinstance(w, ttk.Button)
                     and str(w.cget('text')) == '저장').invoke()
                self.assertEqual(batch['pmi_measurements'][0]['reading_no'], '42')
        finally:
            root.destroy()

    def test_file_loads_directly_and_inline_cell_is_saved(self):
        root = tk.Tk()
        root.withdraw()
        callbacks = {}
        original_bind = ttk.Treeview.bind

        def capture_bind(tree, sequence=None, func=None, add=None):
            callbacks[sequence] = func
            return original_bind(tree, sequence, func, add)

        batch = dict(method='PMI')
        saved = []
        try:
            with tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / 'readings.csv'
                path.write_text('Time,Reading No,Tag,Alloy,Ni,Cr,Mo\n2026-11-01 10:23:45,17,T1,304,8.2,18.5,0.1\n', encoding='utf-8')
                with patch.object(ttk.Treeview, 'bind', capture_bind), patch(
                    'site_apps.gyeongin.src.pmi_measurement_editor.filedialog.askopenfilenames',
                    return_value=(str(path),)):
                    open_measurement_editor(root, batch, Path(tmp) / 'state.json',
                                            lambda: saved.append(True) or True, lambda: None,
                                            import_files=True)
                win = next(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))
                root.deiconify()
                win.deiconify()
                root.update()
                self.assertFalse(any(isinstance(w, tk.Toplevel) for w in win.winfo_children()))
                tree = next(w for w in descendants(win) if isinstance(w, ttk.Treeview))
                self.assertEqual(tree.get_children(), ('0',))
                self.assertEqual(tree.set('0', 'Ni'), '8.20')
                self.assertEqual(tree.set('0', 'reading_no'), '17')
                self.assertEqual(tree.set('0', 'date'), '2026-11-01 10:23:45')
                add_file = next(w for w in descendants(win) if isinstance(w, ttk.Button)
                                and str(w.cget('text')) == '파일 추가')
                with patch('site_apps.gyeongin.src.pmi_measurement_editor.filedialog.askopenfilenames',
                           return_value=(str(path),)), patch(
                           'site_apps.gyeongin.src.pmi_measurement_editor.messagebox.showinfo') as info:
                    add_file.invoke()
                    info.assert_called_once()
                self.assertEqual(tree.get_children(), ('0',))
                x, y, width, height = tree.bbox('0', 'Ni')
                callbacks['<Double-1>'](SimpleNamespace(x=x + width // 2, y=y + height // 2))
                entry = next(w for w in tree.winfo_children() if isinstance(w, ttk.Entry))
                entry.delete(0, 'end')
                entry.insert(0, '101')
                save = next(w for w in descendants(win) if isinstance(w, ttk.Button)
                            and str(w.cget('text')) == '저장')
                with patch('site_apps.gyeongin.src.pmi_measurement_editor.messagebox.showerror') as error:
                    save.invoke()
                    error.assert_called_once()
                self.assertEqual(saved, [])
                self.assertNotIn('pmi_measurements', batch)
                entry.delete(0, 'end')
                entry.insert(0, '9.5')
                save.invoke()
                self.assertEqual(saved, [True])
                row = batch['pmi_measurements'][0]
                self.assertEqual(row['elements']['Ni'], '9.5')
                self.assertEqual(row['_original']['elements']['Ni'], '8.20')
        finally:
            root.destroy()


if __name__ == '__main__':
    unittest.main()
