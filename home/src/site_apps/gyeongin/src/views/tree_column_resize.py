"""Explicit separator dragging for fixed-width, horizontally scrollable tables."""


def enable_column_resize(tree):
    drag = {}

    def press(event):
        if tree.identify_region(event.x, event.y) != 'separator':
            return
        column = tree.identify_column(event.x)
        drag.update(column=column, x=event.x_root, width=int(tree.column(column, 'width')))
        return 'break'

    def motion(event):
        if not drag:
            return
        column = drag['column']
        minimum = int(tree.column(column, 'minwidth'))
        tree.column(column, width=max(minimum, drag['width'] + event.x_root - drag['x']), stretch=False)
        return 'break'

    def release(event):
        if drag:
            drag.clear()
            return 'break'

    tree.bind('<Button-1>', press, add='+')
    tree.bind('<B1-Motion>', motion, add='+')
    tree.bind('<ButtonRelease-1>', release, add='+')
