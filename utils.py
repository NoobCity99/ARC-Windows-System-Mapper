import datetime as _dt
import re
from typing import Iterable, List, Optional


def normalize_date(raw: str) -> str:
    """Return YYYY-MM-DD or blank for unsupported formats (Win32 stores YYYYMMDD)."""
    if not raw:
        return ""
    raw = raw.strip()
    patterns = [
        r"^(?P<y>\d{4})(?P<m>\d{2})(?P<d>\d{2})$",
        r"^(?P<y>\d{4})[-/](?P<m>\d{1,2})[-/](?P<d>\d{1,2})$",
    ]
    for pat in patterns:
        match = re.match(pat, raw)
        if match:
            try:
                dt = _dt.date(int(match.group("y")), int(match.group("m")), int(match.group("d")))
            except ValueError:
                return ""
            return dt.isoformat()
    return ""


def normalize_registry_size(value: Optional[int]) -> Optional[int]:
    if value is None:
        return None
    try:
        return max(int(value) // 1024, 0)
    except (ValueError, TypeError):
        return None


def parse_size_mb(value) -> Optional[int]:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return max(int(float(text)), 0)
    except ValueError:
        return None


def unique_casefold(values: Iterable[str]) -> List[str]:
    seen = set()
    result: List[str] = []
    for value in values:
        name = str(value).strip()
        if not name:
            continue
        folded = name.casefold()
        if folded in seen:
            continue
        seen.add(folded)
        result.append(name)
    return result


def normalize_url(raw: str) -> str:
    url = raw.strip()
    if not url:
        return ""
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url):
        return url
    if url.startswith("www."):
        return f"https://{url}"
    if "." in url and " " not in url:
        return f"https://{url}"
    return ""

