"""Keyboard and focus lifecycle for the borderless record detail popup."""
import tkinter as tk
import uuid


class DetailPopupDismissal:
    def __init__(self, root, popup, content, origin, on_close):
        self.root, self.popup, self.content, self.origin = root, popup, content, origin
        self.on_close = on_close
        self.closed = False
        self.pending = set()
        self.tag = f'RecordDetailEscape_{uuid.uuid4().hex}'
        self.widgets = []
        focused = root.focus_get()
        candidates = [root, origin, focused]

        def collect(widget):
            candidates.append(widget)
            for child in widget.winfo_children():
                collect(child)

        collect(popup)
        for widget in candidates:
            if widget is not None and widget not in self.widgets:
                widget.bindtags((self.tag,) + tuple(widget.bindtags()))
                self.widgets.append(widget)
        self.escape_id = root.bind_class(self.tag, '<Escape>', self.escape)
        popup.bind('<FocusOut>', self.focus_out)
        popup.bind('<Double-1>', self.close)
        popup.bind('<Destroy>', self.destroyed, add='+')
        # Windows may ignore focus_set() on an unmapped, borderless Toplevel.
        self.schedule(self.focus_content)

    def schedule(self, callback):
        pending_id = None

        def run():
            self.pending.discard(pending_id)
            if not self.closed:
                callback()

        pending_id = self.root.after_idle(run)
        self.pending.add(pending_id)

    def focus_content(self):
        if self.popup.winfo_exists():
            self.popup.lift()
            self.content.focus_force()

    def escape(self, event=None):
        self.close()
        if self.origin.winfo_exists():
            self.origin.focus_set()
        return 'break'

    def focus_out(self, event=None):
        self.schedule(self.close_if_outside)

    def close_if_outside(self):
        focused = self.root.focus_get()
        if focused is not None:
            try:
                if focused.winfo_toplevel() == self.popup:
                    return
            except tk.TclError:
                pass
        self.close()

    def cleanup(self):
        if self.closed:
            return
        self.closed = True
        for pending_id in self.pending:
            try:
                self.root.after_cancel(pending_id)
            except tk.TclError:
                pass
        self.pending.clear()
        for widget in self.widgets:
            try:
                widget.bindtags(tuple(tag for tag in widget.bindtags() if tag != self.tag))
            except tk.TclError:
                pass
        self.root.unbind_class(self.tag, '<Escape>')
        self.root.deletecommand(self.escape_id)
        self.on_close()

    def close(self, event=None):
        self.cleanup()
        if self.popup.winfo_exists():
            self.popup.destroy()
        return 'break'

    def destroyed(self, event):
        if event.widget == self.popup:
            self.cleanup()
