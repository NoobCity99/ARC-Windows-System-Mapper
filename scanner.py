from dataclasses import dataclass
import os
import time
from typing import Dict, List, Optional, Tuple

from models import AppEntry
from utils import normalize_date, normalize_registry_size

try:
    import winreg
except ImportError as exc:  # pragma: no cover - Windows only
    raise RuntimeError("This script must run on Windows") from exc


# Enumerate uninstall keys for both registry views plus HKCU without elevation.
REG_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_64KEY),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_32KEY),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_64KEY),
    (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.KEY_WOW64_32KEY),
]


@dataclass(frozen=True)
class SizeScanLimits:
    max_files: int = 20000
    max_depth: int = 6
    max_seconds: float = 2.0


class AppScanner:
    """Enumerates Win32 uninstall entries across 32/64-bit registry views."""

    VALUE_MAP = {
        "DisplayName": "name",
        "DisplayVersion": "version",
        "InstallDate": "install_date",
        "EstimatedSize": "size_mb",
        "Publisher": "publisher",
        "InstallLocation": "install_location",
    }
    FALLBACK_LOCATION_KEYS = ("DisplayIcon", "UninstallString", "QuietUninstallString", "ModifyPath")
    INVALID_LOCATION_TOKENS = {"unknown", "n/a", "na", "none", "null"}
    BLOCKED_EXE_NAMES = {
        "msiexec.exe",
        "rundll32.exe",
        "regsvr32.exe",
        "cmd.exe",
        "powershell.exe",
        "pwsh.exe",
    }
    WEBSITE_KEYS = ("URLInfoAbout", "URLUpdateInfo", "HelpLink")
    APP_PATHS = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths", winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths", winreg.KEY_WOW64_32KEY),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths", winreg.KEY_WOW64_64KEY),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths", winreg.KEY_WOW64_32KEY),
    ]

    def __init__(self) -> None:
        self.apps: List[AppEntry] = []
        self._app_path_cache: Dict[str, str] = {}

    def scan(self, include_sizes: bool = False, size_limits: Optional[SizeScanLimits] = None) -> List[AppEntry]:
        raw_entries: Dict[Tuple[str, str], AppEntry] = {}
        for hive, path, view in REG_PATHS:
            # KEY_WOW64_* flags target the desired registry view without needing elevation.
            access = winreg.KEY_READ | view
            try:
                base = winreg.OpenKey(hive, path, 0, access)
            except OSError:
                continue
            with base:
                for idx in range(self._subkey_count(base)):
                    try:
                        sub_name = winreg.EnumKey(base, idx)
                        sub_key = winreg.OpenKey(base, sub_name)
                    except OSError:
                        continue
                    with sub_key:
                        app = self._read_entry(sub_key)
                    if not app.name:
                        continue
                    key = (app.name, app.version)
                    # Prefer 64-bit entries over 32-bit duplicates.
                    is_64 = view == winreg.KEY_WOW64_64KEY
                    if key in raw_entries:
                        existing = raw_entries[key]
                        existing_is_64 = getattr(existing, "_64bit", False)
                        if existing_is_64 and not is_64:
                            continue
                    setattr(app, "_64bit", is_64)
                    raw_entries[key] = app
        if include_sizes:
            for entry in raw_entries.values():
                if entry.size_mb is None:
                    entry.size_mb = self._compute_install_size_mb(entry.install_location, size_limits)
        self.apps = sorted(raw_entries.values(), key=lambda x: (x.name or "").lower())
        return self.apps

    @staticmethod
    def _subkey_count(key) -> int:
        try:
            info = winreg.QueryInfoKey(key)
            return info[0]
        except OSError:
            return 0

    def _read_entry(self, handle) -> AppEntry:
        entry = AppEntry(name="")
        raw_values: Dict[str, str] = {}
        for value_name, target in self.VALUE_MAP.items():
            try:
                value, _ = winreg.QueryValueEx(handle, value_name)
            except OSError:
                continue
            if value_name == "InstallDate":
                setattr(entry, target, normalize_date(str(value)))
            elif value_name == "EstimatedSize":
                entry.size_mb = normalize_registry_size(value)
            else:
                setattr(entry, target, str(value))
            if value_name == "InstallLocation":
                raw_values[value_name] = str(value)
        for value_name in self.FALLBACK_LOCATION_KEYS:
            try:
                value, _ = winreg.QueryValueEx(handle, value_name)
            except OSError:
                continue
            raw_values[value_name] = str(value)
        for key in self.WEBSITE_KEYS:
            if entry.website:
                break
            try:
                value, _ = winreg.QueryValueEx(handle, key)
            except OSError:
                continue
            entry.website = str(value).strip()
        entry.install_location = self._resolve_install_location(entry.install_location, raw_values)
        return entry

    def compute_install_size_mb(self, raw_path: str, limits: Optional[SizeScanLimits] = None) -> Optional[int]:
        return self._compute_install_size_mb(raw_path, limits)

    @staticmethod
    def _compute_install_size_mb(raw_path: str, limits: Optional[SizeScanLimits] = None) -> Optional[int]:
        path = AppScanner._normalize_install_location(raw_path)
        if not path or not os.path.exists(path):
            return None
        if os.path.isfile(path):
            if path.lower().endswith(".exe"):
                path = os.path.dirname(path)
                if not path or not os.path.exists(path):
                    return None
            else:
                try:
                    return max(os.path.getsize(path) // (1024 * 1024), 0)
                except OSError:
                    return None
        size_bytes = AppScanner._dir_size_bytes(path, limits)
        return max(size_bytes // (1024 * 1024), 0) if size_bytes is not None else None

    @staticmethod
    def _normalize_install_location(raw_path: str) -> str:
        if not raw_path:
            return ""
        path = raw_path.strip().strip('"')
        if not path:
            return ""
        lower = path.casefold()
        if lower in AppScanner.INVALID_LOCATION_TOKENS:
            return ""
        if lower.startswith("unknown"):
            return ""
        if ".exe" in lower:
            idx = lower.find(".exe")
            path = path[: idx + 4].strip()
        path = os.path.expandvars(os.path.expanduser(path))
        return path

    def _resolve_install_location(self, current: str, raw_values: Dict[str, str]) -> str:
        normalized = self._normalize_install_location(current)
        if normalized:
            return normalized

        for key in self.FALLBACK_LOCATION_KEYS:
            candidate = self._extract_path_from_command(raw_values.get(key, ""))
            if not candidate:
                continue
            normalized = self._normalize_install_location(candidate)
            if normalized:
                return normalized

        for key in self.FALLBACK_LOCATION_KEYS:
            exe_name = self._extract_exe_name(raw_values.get(key, ""))
            if not exe_name:
                continue
            app_path = self._lookup_app_path(exe_name)
            if app_path:
                normalized = self._normalize_install_location(app_path)
                if normalized:
                    return normalized

        return ""

    @classmethod
    def _extract_path_from_command(cls, raw: str) -> str:
        text = (raw or "").strip()
        if not text:
            return ""
        if text.startswith('"'):
            end = text.find('"', 1)
            if end > 1:
                candidate = text[1:end]
                return candidate
        lowered = text.casefold()
        idx = lowered.find(".exe")
        if idx != -1:
            candidate = text[: idx + 4]
            candidate = candidate.split(",")[0].strip()
            if cls._is_blocked_exe(candidate):
                return ""
            return candidate
        if len(text) >= 2 and text[1] == ":":
            token = text.split(" ")[0].split(",")[0].strip()
            if cls._is_blocked_exe(token):
                return ""
            return token
        return ""

    @classmethod
    def _extract_exe_name(cls, raw: str) -> str:
        candidate = cls._extract_path_from_command(raw)
        if not candidate:
            return ""
        name = os.path.basename(candidate)
        if name.casefold() in cls.BLOCKED_EXE_NAMES:
            return ""
        if not name.casefold().endswith(".exe"):
            return ""
        return name

    @classmethod
    def _is_blocked_exe(cls, path: str) -> bool:
        name = os.path.basename((path or "").strip().strip('"')).casefold()
        if name in cls.BLOCKED_EXE_NAMES:
            return True
        return False

    def _lookup_app_path(self, exe_name: str) -> str:
        if not exe_name:
            return ""
        cached = self._app_path_cache.get(exe_name)
        if cached is not None:
            return cached
        for hive, base_path, view in self.APP_PATHS:
            try:
                access = winreg.KEY_READ | view
                key = winreg.OpenKey(hive, f"{base_path}\\{exe_name}", 0, access)
            except OSError:
                continue
            with key:
                try:
                    value, _ = winreg.QueryValueEx(key, "")
                except OSError:
                    continue
                resolved = str(value)
                self._app_path_cache[exe_name] = resolved
                return resolved
        self._app_path_cache[exe_name] = ""
        return ""

    @staticmethod
    def _dir_size_bytes(root: str, limits: Optional[SizeScanLimits]) -> Optional[int]:
        total = 0
        stack: List[Tuple[str, int]] = [(root, 0)]
        file_count = 0
        max_files = limits.max_files if limits else None
        max_depth = limits.max_depth if limits else None
        deadline = time.monotonic() + limits.max_seconds if limits and limits.max_seconds else None
        while stack:
            current, depth = stack.pop()
            if deadline and time.monotonic() >= deadline:
                return None
            try:
                with os.scandir(current) as entries:
                    for entry in entries:
                        try:
                            if entry.is_symlink():
                                continue
                            if entry.is_dir(follow_symlinks=False):
                                child_depth = depth + 1
                                if max_depth is not None and child_depth > max_depth:
                                    continue
                                stack.append((entry.path, child_depth))
                            else:
                                if max_files is not None and file_count >= max_files:
                                    return None
                                file_count += 1
                                total += entry.stat(follow_symlinks=False).st_size
                        except OSError:
                            continue
            except OSError:
                continue
        return total
