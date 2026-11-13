"""Cancel interpreter timers when the Gyeongin application's root closes."""
def install_root_timer_cleanup(widget):
    root = widget._root()
    if getattr(root, '_gyeongin_timer_cleanup', False):
        return

    def cancel_pending(event):
        if event.widget is not root:
            return
        # Children have already removed their Tcl callback commands. Cancel
        # timers directly so cancellation does not try to delete them twice.
        for timer in root.tk.splitlist(root.tk.call('after', 'info')):
            root.tk.call('after', 'cancel', timer)

    root.bind('<Destroy>', cancel_pending, add='+')
    root._gyeongin_timer_cleanup = True
