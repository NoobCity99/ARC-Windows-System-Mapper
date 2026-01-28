import datetime as _dt

from compare import build_scan_name_set, is_installed
from models import AppEntry, RelatedFile


def test_compare_helpers() -> None:
    scan = [AppEntry(name="App A"), AppEntry(name="App B")]
    names = build_scan_name_set(scan)
    assert is_installed(AppEntry(name="App A"), names, True) is True
    assert is_installed(AppEntry(name="Missing"), names, True) is False
    assert is_installed(AppEntry(name="Anything"), names, False) is True


def test_appentry_keys_and_search() -> None:
    app = AppEntry(name="Example", version="1.0", publisher="Vendor")
    assert app.key().startswith("7|Example|3|1.0")
    assert app.legacy_key() == '["Example", "1.0"]'
    assert app.name_key() == "example"
    assert "example" in app.search_blob()
    assert "vendor" in app.search_blob()


def test_install_date_value() -> None:
    app = AppEntry(name="Example", install_date="2024-01-31")
    assert app.install_date_value() == _dt.date(2024, 1, 31)
    app.install_date = "bad"
    assert app.install_date_value() is None


def test_relatedfile_search_blob() -> None:
    rel = RelatedFile(path="C:\\Temp\\file.cfg", source="appdata", confidence="High", marked="keep")
    blob = rel.search_blob()
    assert "c:\\temp\\file.cfg" in blob
    assert "appdata" in blob
