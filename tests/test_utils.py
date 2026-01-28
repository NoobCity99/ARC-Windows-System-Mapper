from utils import normalize_date, normalize_registry_size, normalize_url, parse_size_mb, unique_casefold


def test_normalize_date_basic() -> None:
    assert normalize_date("20240131") == "2024-01-31"
    assert normalize_date("2024-1-5") == "2024-01-05"
    assert normalize_date("2024/01/05") == "2024-01-05"
    assert normalize_date("bad") == ""
    assert normalize_date("") == ""


def test_normalize_registry_size() -> None:
    assert normalize_registry_size(2048) == 2
    assert normalize_registry_size(0) == 0
    assert normalize_registry_size(None) is None
    assert normalize_registry_size("nope") is None


def test_parse_size_mb() -> None:
    assert parse_size_mb("12") == 12
    assert parse_size_mb("12.5") == 12
    assert parse_size_mb("-5") == 0
    assert parse_size_mb("") is None
    assert parse_size_mb(None) is None


def test_unique_casefold() -> None:
    assert unique_casefold(["Alpha", "alpha", "Beta", "", "  "]) == ["Alpha", "Beta"]


def test_normalize_url() -> None:
    assert normalize_url("https://example.com") == "https://example.com"
    assert normalize_url("www.example.com") == "https://www.example.com"
    assert normalize_url("example.com") == "https://example.com"
    assert normalize_url("not a url") == ""
