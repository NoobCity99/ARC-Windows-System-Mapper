import json
import os
import shutil
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from utils import unique_casefold

APP_NAME = "ARC"
ENV_DATA_DIR = "ARC_DATA_DIR"
CONFIG_FILENAME = "arc_poc_config.json"
CONFIG_KEY = "data_dir"


@dataclass
class StoredState:
    geometry: str = ""
    sort_column: str = "name"
    sort_reverse: bool = False
    gui_settings: Dict[str, object] = field(default_factory=dict)
    groups: List[str] = field(default_factory=list)
    app_groups: Dict[str, str] = field(default_factory=dict)
    scan_drives: List[str] = field(default_factory=list)
    group_colors: Dict[str, str] = field(default_factory=dict)
    size_cache: Dict[str, Dict[str, object]] = field(default_factory=dict)
    related_overrides: Dict[str, str] = field(default_factory=dict)
    related_manual: Dict[str, List[Dict[str, str]]] = field(default_factory=dict)
    related_ignore: Dict[str, List[str]] = field(default_factory=dict)
    related_unassigned: Dict[str, List[str]] = field(default_factory=dict)
    install_location_overrides: Dict[str, str] = field(default_factory=dict)
    version_overrides: Dict[str, str] = field(default_factory=dict)
    install_date_overrides: Dict[str, str] = field(default_factory=dict)


def app_data_dir(app_name: str = APP_NAME) -> str:
    if os.name == "nt":
        base = os.getenv("LOCALAPPDATA") or os.getenv("APPDATA")
    else:
        base = os.getenv("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    if not base:
        base = os.path.expanduser("~")
    return os.path.join(base, app_name)


def _config_path() -> str:
    return os.path.join(app_data_dir(), CONFIG_FILENAME)


def _load_config() -> Dict[str, str]:
    path = _config_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    config: Dict[str, str] = {}
    for key, value in payload.items():
        if not isinstance(key, str):
            continue
        if isinstance(value, str):
            config[key] = value
    return config


def _save_config(config: Dict[str, str]) -> None:
    path = _config_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(config, fh, indent=2)
    except OSError:
        pass


def set_configured_data_dir(data_dir: str) -> None:
    if not isinstance(data_dir, str):
        return
    data_dir = data_dir.strip()
    if not data_dir:
        return
    config = _load_config()
    config[CONFIG_KEY] = data_dir
    _save_config(config)


def _legacy_state_path(filename: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


def _maybe_migrate_legacy_state(target_path: str, filename: str) -> None:
    if os.path.exists(target_path):
        return
    legacy = _legacy_state_path(filename)
    if legacy == target_path or not os.path.exists(legacy):
        return
    try:
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        shutil.copyfile(legacy, target_path)
    except OSError:
        pass


def resolve_data_dir(prompt_for_dir: Optional[Callable[[], str]] = None) -> str:
    override = os.getenv(ENV_DATA_DIR)
    if override:
        return override
    config = _load_config()
    data_dir = str(config.get(CONFIG_KEY) or "").strip()
    if not data_dir and prompt_for_dir:
        selected = str(prompt_for_dir() or "").strip()
        if selected:
            data_dir = selected
    if not data_dir:
        data_dir = app_data_dir()
    if prompt_for_dir:
        config[CONFIG_KEY] = data_dir
        _save_config(config)
    return data_dir


def default_state_path(filename: str, prompt_for_dir: Optional[Callable[[], str]] = None) -> str:
    data_dir = resolve_data_dir(prompt_for_dir)
    path = os.path.join(data_dir, filename)
    _maybe_migrate_legacy_state(path, filename)
    return path


def load_state(path: str, default_gui_settings: Dict[str, object]) -> StoredState:
    state = StoredState(gui_settings=dict(default_gui_settings))
    if not os.path.exists(path):
        return state
    try:
        with open(path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
    except (OSError, json.JSONDecodeError):
        return state
    if isinstance(payload, dict):
        state.geometry = str(payload.get("geometry") or "")
        state.sort_column = str(payload.get("sort_column") or state.sort_column)
        state.sort_reverse = bool(payload.get("sort_reverse", state.sort_reverse))
        saved_gui = payload.get("gui_settings", {})
        if isinstance(saved_gui, dict):
            for key, value in saved_gui.items():
                if key not in default_gui_settings:
                    continue
                if key == "font_size":
                    try:
                        size = int(value)
                    except (TypeError, ValueError):
                        continue
                    state.gui_settings[key] = max(6, min(72, size))
                elif isinstance(value, str) or isinstance(value, (int, float)):
                    state.gui_settings[key] = value
        groups = payload.get("groups", [])
        if isinstance(groups, list):
            state.groups = unique_casefold(groups)
        app_groups = payload.get("app_groups", {})
        if isinstance(app_groups, dict):
            valid = set(state.groups)
            pruned: Dict[str, str] = {}
            for key, value in app_groups.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, str):
                    continue
                if value not in valid:
                    continue
                pruned[key] = value
            state.app_groups = pruned
        scan_drives = payload.get("scan_drives", [])
        if isinstance(scan_drives, list):
            state.scan_drives = [str(value) for value in scan_drives if isinstance(value, str)]
        group_colors = payload.get("group_colors", {})
        if isinstance(group_colors, dict):
            valid = set(state.groups)
            pruned_colors: Dict[str, str] = {}
            for key, value in group_colors.items():
                if not isinstance(key, str):
                    continue
                if key not in valid:
                    continue
                if not isinstance(value, str):
                    continue
                pruned_colors[key] = value
            state.group_colors = pruned_colors
        size_cache = payload.get("size_cache", {})
        if isinstance(size_cache, dict):
            pruned_cache: Dict[str, Dict[str, object]] = {}
            for key, value in size_cache.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, dict):
                    continue
                size = value.get("size_mb")
                location = value.get("install_location")
                if not isinstance(location, str):
                    continue
                if isinstance(size, bool):
                    continue
                if isinstance(size, (int, float)):
                    size_val = int(size)
                    if size_val < 0:
                        continue
                else:
                    continue
                pruned_cache[key] = {
                    "size_mb": size_val,
                    "install_location": location,
                    "updated_at": str(value.get("updated_at") or ""),
                }
            state.size_cache = pruned_cache
        overrides = payload.get("related_overrides", {})
        if isinstance(overrides, dict):
            pruned_overrides: Dict[str, str] = {}
            for key, value in overrides.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, str):
                    continue
                pruned_overrides[key] = value
            state.related_overrides = pruned_overrides
        manual = payload.get("related_manual", {})
        if isinstance(manual, dict):
            pruned_manual: Dict[str, List[Dict[str, str]]] = {}
            for key, value in manual.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, list):
                    continue
                items: List[Dict[str, str]] = []
                for entry in value:
                    if not isinstance(entry, dict):
                        continue
                    path = entry.get("path")
                    kind = entry.get("kind")
                    if not isinstance(path, str) or not path.strip():
                        continue
                    if not isinstance(kind, str) or not kind.strip():
                        kind = "file"
                    items.append({"path": path.strip(), "kind": kind.strip()})
                if items:
                    pruned_manual[key] = items
            state.related_manual = pruned_manual
        ignore = payload.get("related_ignore", {})
        if isinstance(ignore, dict):
            pruned_ignore: Dict[str, List[str]] = {}
            for key, value in ignore.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, list):
                    continue
                paths: List[str] = []
                for item in value:
                    if not isinstance(item, str):
                        continue
                    item = item.strip()
                    if not item:
                        continue
                    paths.append(item)
                if paths:
                    pruned_ignore[key] = paths
            state.related_ignore = pruned_ignore
        unassigned = payload.get("related_unassigned", {})
        if isinstance(unassigned, dict):
            pruned_unassigned: Dict[str, List[str]] = {}
            for key, value in unassigned.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, list):
                    continue
                paths: List[str] = []
                for item in value:
                    if not isinstance(item, str):
                        continue
                    item = item.strip()
                    if not item:
                        continue
                    paths.append(item)
                if paths:
                    pruned_unassigned[key] = paths
            state.related_unassigned = pruned_unassigned
        overrides = payload.get("install_location_overrides", {})
        if isinstance(overrides, dict):
            pruned_overrides: Dict[str, str] = {}
            for key, value in overrides.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, str):
                    continue
                value = value.strip()
                if not value:
                    continue
                pruned_overrides[key] = value
            state.install_location_overrides = pruned_overrides
        version_overrides = payload.get("version_overrides", {})
        if isinstance(version_overrides, dict):
            pruned_versions: Dict[str, str] = {}
            for key, value in version_overrides.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, str):
                    continue
                value = value.strip()
                if not value:
                    continue
                pruned_versions[key] = value
            state.version_overrides = pruned_versions
        date_overrides = payload.get("install_date_overrides", {})
        if isinstance(date_overrides, dict):
            pruned_dates: Dict[str, str] = {}
            for key, value in date_overrides.items():
                if not isinstance(key, str):
                    continue
                if not isinstance(value, str):
                    continue
                value = value.strip()
                if not value:
                    continue
                pruned_dates[key] = value
            state.install_date_overrides = pruned_dates
    return state


def save_state(path: str, state: StoredState) -> None:
    valid_groups = set(state.groups)
    app_groups = {key: value for key, value in state.app_groups.items() if value in valid_groups}
    size_cache: Dict[str, Dict[str, object]] = {}
    for key, value in state.size_cache.items():
        if not isinstance(key, str):
            continue
        if not isinstance(value, dict):
            continue
        size = value.get("size_mb")
        location = value.get("install_location")
        if not isinstance(location, str):
            continue
        if isinstance(size, bool):
            continue
        if isinstance(size, (int, float)):
            size_val = int(size)
            if size_val < 0:
                continue
        else:
            continue
        size_cache[key] = {
            "size_mb": size_val,
            "install_location": location,
            "updated_at": str(value.get("updated_at") or ""),
        }
    payload = {
        "geometry": state.geometry,
        "sort_column": state.sort_column,
        "sort_reverse": state.sort_reverse,
        "gui_settings": state.gui_settings,
        "groups": state.groups,
        "app_groups": app_groups,
        "scan_drives": state.scan_drives,
        "group_colors": {key: value for key, value in state.group_colors.items() if key in valid_groups},
        "size_cache": size_cache,
        "related_overrides": state.related_overrides,
        "related_manual": state.related_manual,
        "related_ignore": state.related_ignore,
        "related_unassigned": state.related_unassigned,
        "install_location_overrides": state.install_location_overrides,
        "version_overrides": state.version_overrides,
        "install_date_overrides": state.install_date_overrides,
    }
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)
    except OSError:
        pass
