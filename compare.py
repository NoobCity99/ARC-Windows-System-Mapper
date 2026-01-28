from typing import Iterable, Set

from models import AppEntry


def build_scan_name_set(scan: Iterable[AppEntry]) -> Set[str]:
    return {app.name_key() for app in scan}


def is_installed(entry: AppEntry, scan_names: Set[str], has_reference: bool) -> bool:
    if not has_reference:
        return True
    if not scan_names:
        return False
    return entry.name_key() in scan_names
