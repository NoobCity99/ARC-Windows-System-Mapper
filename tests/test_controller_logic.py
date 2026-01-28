import tkinter as tk

import pytest

import main_controller
from models import AppEntry, RelatedFile


class _DummyVar:
    def __init__(self, value=None) -> None:
        self._value = value

    def get(self):
        return self._value

    def set(self, value) -> None:
        self._value = value


class _FakeMainView:
    def __init__(self, *args, **kwargs) -> None:
        sort_labels = kwargs.get("sort_labels")
        callbacks = kwargs.get("callbacks")
        if sort_labels is None and len(args) >= 6:
            sort_labels = args[5]
        if callbacks is None and len(args) >= 7:
            callbacks = args[6]
        self.callbacks = callbacks or {}
        labels = sort_labels or []
        self.sort_desc_var = _DummyVar(False)
        self.sort_var = _DummyVar(labels[0] if labels else "")

    def set_sort_desc(self, _value: bool) -> None:
        pass

    def set_sort_label(self, _label: str) -> None:
        pass

    def set_view_mode(self, _mode: str) -> None:
        pass

    def set_sort_enabled(self, _enabled: bool) -> None:
        pass

    def set_export_enabled(self, _enabled: bool) -> None:
        pass

    def set_save_json_enabled(self, _enabled: bool) -> None:
        pass

    def set_clear_scan_enabled(self, _enabled: bool) -> None:
        pass

    def set_close_reference_enabled(self, _enabled: bool) -> None:
        pass

    def set_related_view_enabled(self, _enabled: bool) -> None:
        pass

    def set_map_view_enabled(self, _enabled: bool) -> None:
        pass

    def set_map_group_colors(self, _colors) -> None:
        pass

    def set_map_style(self, _style) -> None:
        pass

    def set_tree_tag_colors(self, _installed: str, _missing: str) -> None:
        pass

    def update_group_editor_values(self, _groups, _no_group_label) -> None:
        pass

    def clear_system_tree(self) -> None:
        pass

    def populate_tree(self, _rows) -> None:
        pass

    def populate_related_tree(self, _groups, preserve_expansion: bool = False) -> None:
        pass

    def populate_system_map(self, _payload) -> None:
        pass

    def set_status(self, _text: str) -> None:
        pass

    def set_filter(self, _value: str) -> None:
        pass

    def set_scan_enabled(self, _enabled: bool) -> None:
        pass

    def consume_view_request(self, _mode: str) -> bool:
        return False


class _FakeSettingsView:
    def __init__(self, _root, callbacks) -> None:
        self.callbacks = callbacks
        self.window = None

    def show(self, *_args, **_kwargs) -> None:
        pass

    def destroy(self) -> None:
        pass

    def refresh_groups(self, _groups, _colors) -> None:
        pass


def _tk_available() -> bool:
    try:
        root = tk.Tk()
        root.withdraw()
        root.destroy()
        return True
    except tk.TclError:
        return False


@pytest.mark.gui
def test_controller_map_helpers(monkeypatch) -> None:
    if not _tk_available():
        pytest.skip("Tkinter unavailable")
    monkeypatch.setattr(main_controller, "MainView", _FakeMainView)
    monkeypatch.setattr(main_controller, "SettingsView", _FakeSettingsView)
    root = tk.Tk()
    root.withdraw()
    controller = main_controller.MainController(root)
    assert controller._map_label_for_path(r"C:\Program Files\App\app.exe") == "app.exe"
    assert controller._drive_for_path(r"C:\Program Files") == "C:"
    controller.gui_settings["map_max_related"] = 999
    assert controller._map_max_related() == 50
    controller.gui_settings["map_max_related"] = -5
    assert controller._map_max_related() == 0
    root.destroy()


@pytest.mark.gui
def test_related_filter_index(monkeypatch) -> None:
    if not _tk_available():
        pytest.skip("Tkinter unavailable")
    monkeypatch.setattr(main_controller, "MainView", _FakeMainView)
    monkeypatch.setattr(main_controller, "SettingsView", _FakeSettingsView)
    root = tk.Tk()
    root.withdraw()
    controller = main_controller.MainController(root)
    app = AppEntry(name="Alpha", publisher="Vendor")
    app.related_files = [
        RelatedFile(path="C:\\Temp\\alpha.cfg", kind="file", source="config_file"),
        RelatedFile(path="C:\\Temp\\beta.cfg", kind="file", source="config_file"),
    ]
    controller.all_apps = [app]
    controller._refresh_app_index()
    groups, count = controller._build_related_groups(controller.all_apps, "alpha")
    assert count == 1
    assert len(groups) == 1
    root.destroy()
