"""Compatibility fixes for tkcalendar on Windows and recent Python versions.

All patches are deliberately opt-in and idempotent.  Keeping them here prevents
the application entry point from mutating third-party classes while it imports.
"""

from __future__ import annotations

import logging
import tkinter as tk

from tkcalendar import Calendar, DateEntry

LOGGER = logging.getLogger(__name__)
_PATCHED = False


def _guard(method, *, restore_date: bool = False):
    """Return a callback wrapper that handles known Tk lifecycle failures."""

    def guarded(instance, *args, **kwargs):
        saved_date = getattr(instance, "_date", None) if restore_date else None
        try:
            return method(instance, *args, **kwargs)
        except (tk.TclError, RuntimeError, AttributeError) as exc:
            if restore_date and saved_date is not None:
                instance._date = saved_date
            LOGGER.debug("Ignored tkcalendar lifecycle error in %s: %s", method.__name__, exc)
            return None

    return guarded


def apply_tkcalendar_compatibility_patches() -> None:
    """Apply the small set of compatibility patches required by the UI."""

    global _PATCHED
    if _PATCHED:
        return

    original_calendar_init = Calendar.__init__

    def calendar_init(instance, *args, **kwargs):
        kwargs.setdefault("select_on_nav", False)
        original_calendar_init(instance, *args, **kwargs)

        def widen_headers():
            for name, width in (("_header_month", 25), ("_header_year", 10)):
                widget = getattr(instance, name, None)
                if widget is not None:
                    try:
                        widget.configure(width=width)
                    except tk.TclError as exc:
                        LOGGER.debug("Could not resize tkcalendar header: %s", exc)

        instance._widen_headers = widen_headers
        instance.after_idle(widen_headers)

    Calendar.__init__ = calendar_init

    for owner, method_name, restore_date in (
        (Calendar, "_display_calendar", False),
        (Calendar, "_setup_style", False),
        (Calendar, "_prev_year", True),
        (Calendar, "_next_year", True),
        (Calendar, "_prev_month", True),
        (Calendar, "_next_month", True),
        (DateEntry, "_show_calendar", False),
        (DateEntry, "drop_down", False),
        (DateEntry, "_on_b1_press", False),
        (DateEntry, "_on_calendar_selection", False),
        (DateEntry, "_setup_style", False),
        (DateEntry, "_determine_downarrow_name", False),
        (DateEntry, "_on_focus_out_cal", False),
    ):
        method = getattr(owner, method_name, None)
        if method is not None:
            setattr(owner, method_name, _guard(method, restore_date=restore_date))

    _PATCHED = True

