"""Shared Treeview utilities."""

from __future__ import annotations


def _sort_value(value):
    text = str(value).replace(",", "").replace("Hrs", "").replace(" ", "")
    for token in ("시간", "원", "%", "(", ")"):
        text = text.replace(token, "")
    try:
        return 0, float(text or 0)
    except ValueError:
        return 1, str(value).casefold()


def sort_treeview_column(tree, column, reverse=False):
    rows = [(tree.set(item, column), item) for item in tree.get_children("")]
    totals = [row for row in rows if "total" in tree.item(row[1], "tags")]
    data = [row for row in rows if row not in totals]
    data.sort(key=lambda row: _sort_value(row[0]), reverse=reverse)
    for index, (_, item) in enumerate(data):
        tree.move(item, "", index)
    for _, item in totals:
        tree.move(item, "", "end")
    tree.heading(column, command=lambda: sort_treeview_column(tree, column, not reverse))
