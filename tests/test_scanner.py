import os

from scanner import AppScanner, SizeScanLimits


def test_normalize_install_location() -> None:
    assert AppScanner._normalize_install_location(' "C:\\Program Files\\App\\app.exe" ') == "C:\\Program Files\\App\\app.exe"
    assert AppScanner._normalize_install_location("unknown") == ""
    assert AppScanner._normalize_install_location("") == ""


def test_extract_path_from_command() -> None:
    assert AppScanner._extract_path_from_command('"C:\\App\\app.exe" /quiet') == "C:\\App\\app.exe"
    assert AppScanner._extract_path_from_command("C:\\App\\app.exe /quiet") == "C:\\App\\app.exe"
    assert AppScanner._extract_path_from_command("not a path") == ""


def test_compute_install_size_mb(tmp_path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    file_a = root / "a.bin"
    file_b = root / "b.bin"
    file_a.write_bytes(b"a" * 1024 * 1024)
    file_b.write_bytes(b"b" * 512 * 1024)
    scanner = AppScanner()
    size = scanner.compute_install_size_mb(str(root), SizeScanLimits(max_files=10, max_depth=2, max_seconds=2.0))
    assert size == 1


def test_resolve_install_location_with_fallback(tmp_path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    exe = root / "app.exe"
    exe.write_text("x")
    scanner = AppScanner()
    resolved = scanner._resolve_install_location("", {"DisplayIcon": f'"{exe}"'})
    assert os.path.normcase(resolved) == os.path.normcase(str(exe))
