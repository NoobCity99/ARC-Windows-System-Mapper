import tkinter as tk

import pytest

from main_view import MainView


def _tk_available() -> bool:
    try:
        root = tk.Tk()
        root.withdraw()
        root.destroy()
        return True
    except tk.TclError:
        return False


@pytest.mark.gui
def test_main_view_smoke() -> None:
    if not _tk_available():
        pytest.skip("Tkinter unavailable")
    root = tk.Tk()
    root.withdraw()
    callbacks = {name: (lambda *args, **kwargs: None) for name in (
        "on_scan",
        "on_export",
        "on_import",
        "on_open_json",
        "on_save_json",
        "on_open_settings",
        "on_clear_scan",
        "on_close_reference",
        "on_view_change",
        "on_sort_change",
        "on_sort_toggle",
        "on_sort_column",
        "on_filter_change",
        "on_clear_filter",
        "on_group_double_click",
        "on_related_double_click",
        "on_related_add_files",
        "on_related_add_folder",
        "on_related_deep_scan",
        "on_related_remove_manual",
        "on_related_unassign",
        "on_related_reassign",
        "on_website_click",
        "on_open_install_location",
        "on_view_related_for_app",
        "on_row_select",
        "on_close",
    )}
    MainView(
        root,
        columns=[("name", "Name", tk.W, True)],
        widths=[200],
        related_columns=[("path", "Path", tk.W, True)],
        related_widths=[400],
        sort_labels=["Name"],
        callbacks=callbacks,
    )
    root.destroy()
