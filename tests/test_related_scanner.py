from models import AppEntry
from related_scanner import RelatedFileScanner


def test_related_scanner_install_location_and_configs(tmp_path) -> None:
    root = tmp_path / "MyApp"
    root.mkdir()
    config = root / "settings.ini"
    config.write_text("x")
    app = AppEntry(name="MyApp", install_location=str(root))
    scanner = RelatedFileScanner(max_dirs=5, max_files=10, max_depth=2)
    related = scanner.scan_for_app(app, deep_scan=False, include_files=True)
    paths = {item.path for item in related}
    assert str(root) in paths
    assert str(config) in paths
