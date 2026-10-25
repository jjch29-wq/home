"""Popup dismissal must work before focus arrives and remove temporary bindings."""
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from site_apps.central.src.views.detail_popup import DetailPopupDismissal


class Widget:
    def __init__(self, top=None):
        self.tags = ('widget', 'class', 'toplevel', 'all')
        self.children = []
        self.top = top or self
        self.exists = True
        self.bindings = {}
        self.focus_force = Mock()
        self.focus_set = Mock()
        self.lift = Mock()

    def bindtags(self, tags=None):
        if tags is not None: self.tags = tags
        return self.tags

    def winfo_children(self): return self.children
    def winfo_toplevel(self): return self.top
    def winfo_exists(self): return self.exists
    def bind(self, event, callback, **kwargs): self.bindings[event] = callback

    def destroy(self):
        self.exists = False
        self.bindings['<Destroy>'](SimpleNamespace(widget=self))


class Root(Widget):
    def __init__(self):
        super().__init__()
        self.pending = {}
        self.current_focus = None
        self.bind_class = Mock(return_value='escape-command')
        self.unbind_class = Mock()
        self.deletecommand = Mock()
        self.counter = 0

    def focus_get(self): return self.current_focus

    def after_idle(self, callback):
        self.counter += 1
        identifier = str(self.counter)
        self.pending[identifier] = callback
        return identifier

    def after_cancel(self, identifier): self.pending.pop(identifier, None)

    def idle(self):
        callbacks = list(self.pending.values())
        self.pending.clear()
        for callback in callbacks: callback()


class DetailPopupTests(unittest.TestCase):
    def fixture(self):
        root, popup = Root(), Widget()
        origin, content, scrollbar = Widget(root), Widget(popup), Widget(popup)
        popup.children = [content, scrollbar]
        root.current_focus = origin
        cleanup = Mock()
        handler = DetailPopupDismissal(root, popup, content, origin, cleanup)
        return root, popup, origin, content, scrollbar, handler, cleanup

    def test_escape_from_original_list_before_focus_arrives_closes_only_popup(self):
        root, popup, origin, content, _, handler, cleanup = self.fixture()
        self.assertEqual(origin.bindtags()[0], handler.tag)
        self.assertEqual(root.bind_class.call_args.args[1], '<Escape>')
        self.assertEqual(handler.escape(), 'break')
        self.assertFalse(popup.exists)
        self.assertTrue(root.exists)
        origin.focus_set.assert_called_once()
        self.assertEqual(origin.bindtags(), ('widget', 'class', 'toplevel', 'all'))
        root.idle()
        content.focus_force.assert_not_called()
        cleanup.assert_called_once()

    def test_mapped_content_gets_focus_and_escape_outranks_widget_handlers(self):
        root, popup, origin, content, scrollbar, handler, _ = self.fixture()
        root.idle()
        content.focus_force.assert_called_once()
        popup.lift.assert_called_once()
        self.assertEqual(content.bindtags()[0], handler.tag)
        self.assertEqual(scrollbar.bindtags()[0], handler.tag)

    def test_focus_between_popup_children_does_not_close_it(self):
        root, popup, origin, content, scrollbar, handler, _ = self.fixture()
        root.idle()
        root.current_focus = scrollbar
        handler.focus_out()
        root.idle()
        self.assertTrue(popup.exists)

    def test_click_outside_closes_without_stealing_focus(self):
        root, popup, origin, content, scrollbar, handler, _ = self.fixture()
        root.idle()
        root.current_focus = origin
        handler.focus_out()
        root.idle()
        self.assertFalse(popup.exists)
        origin.focus_set.assert_not_called()

    def test_external_destroy_removes_only_own_bindings_and_pending_callbacks(self):
        root, popup, origin, content, scrollbar, handler, cleanup = self.fixture()
        content.bindtags(('other-feature',) + content.bindtags())
        popup.destroy()
        handler.cleanup()
        self.assertEqual(content.bindtags(), ('other-feature', 'widget', 'class', 'toplevel', 'all'))
        root.unbind_class.assert_called_once_with(handler.tag, '<Escape>')
        root.deletecommand.assert_called_once_with('escape-command')
        self.assertEqual(root.pending, {})
        cleanup.assert_called_once()


if __name__ == '__main__':
    unittest.main()
