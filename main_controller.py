import locale
import os
import sys
import queue
import threading
import webbrowser
import shutil
from dataclasses import dataclass, replace
from typing import Dict, List, Optional, Set, Tuple

import tkinter as tk
import tkinter.simpledialog as simpledialog
import tkinter.font as tkfont
import datetime as _dt
from tkinter import filedialog, messagebox, ttk

from compare import build_scan_name_set, is_installed
from import_export import export_xlsx, import_csv, load_json, save_json
from main_view import MainView
from models import AppEntry, RelatedFile
from related_scanner import DeepScanLimits, RelatedFileScanner
from scanner import AppScanner, SizeScanLimits
from settings_view import SettingsView
from store import StoredState, default_state_path, load_state, save_state, set_configured_data_dir
from utils import normalize_date, normalize_url, unique_casefold

try:
    locale.setlocale(locale.LC_COLLATE, "")
except locale.Error:
    pass


STATE_FILE = "arc_poc_state.json"


def resource_path(*parts: str) -> str:
    # Support running from source and from a PyInstaller bundle.
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    return os.path.join(base_dir, *parts)


ASSET_DIR = resource_path("assets")
WINDOW_ICON_PATH = resource_path("assets", "app.ico")


@dataclass(frozen=True)
class RelatedRow:
    row_id: str
    app_key: str
    index: int
    related: RelatedFile
    search: str


class MainController:
    # Customize table columns/labels here; keep column keys in sync with models/import_export.
    columns = [
        ("name", "Name", tk.W, True),
        ("installed", "Installed", tk.CENTER, False),
        ("group", "Group", tk.W, False),
        ("version", "Version", tk.CENTER, False),
        ("install_date", "Install Date", tk.CENTER, False),
        ("size_mb", "Size (MB)", tk.E, False),
        ("publisher", "Publisher", tk.W, False),
        ("install_location", "Install Location", tk.W, True),
        ("website", "Website", tk.W, True),
    ]
    # Adjust default column widths for the main table.
    column_widths = [200, 90, 140, 120, 120, 100, 220, 300, 220]
    related_columns = [
        ("app_name", "App", tk.W, True),
        ("path", "Path", tk.W, True),
        ("kind", "Type", tk.CENTER, False),
        ("source", "Source", tk.W, False),
        ("confidence", "Confidence", tk.CENTER, False),
        ("marked", "Marked", tk.CENTER, False),
    ]
    related_column_widths = [200, 420, 80, 140, 110, 90]
    # Change the sort options shown in the toolbar drop-down.
    sort_options = [
        ("Name", "name"),
        ("Installed", "installed"),
        ("Group", "group"),
        ("Install Date", "install_date"),
        ("Size (MB)", "size_mb"),
    ]
    no_group_label = "(No group)"
    installed_marker = "\u2713"
    missing_marker = "\u2717"
    related_mark_cycle = ["", "keep", "ignore"]
    related_source_labels = {
        "install_location": "Install Location",
        "appdata": "AppData",
        "localappdata": "Local AppData",
        "localappdata_low": "Local AppData Low",
        "programdata": "ProgramData",
        "documents": "Documents",
        "saved_games": "Saved Games",
        "config_file": "Config File",
        "drive": "Drive",
        "manual": "Manual",
    }
    group_color_palette = [
        "#1f77b4",
        "#ff7f0e",
        "#2ca02c",
        "#d62728",
        "#9467bd",
        "#8c564b",
        "#e377c2",
        "#7f7f7f",
        "#bcbd22",
        "#17becf",
    ]
    # Typing delay before filter applies; lower feels snappier, higher reduces churn.
    FILTER_DEBOUNCE_MS = 250
    # Tradeoff between size accuracy and UI responsiveness.
    SIZE_SCAN_LIMITS = SizeScanLimits(max_files=20000, max_depth=6, max_seconds=2.0)
    # Debounce map redraws to avoid expensive redraws during rapid updates.
    MAP_REFRESH_DEBOUNCE_MS = 200

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        # Update the window title shown in the title bar/taskbar.
        self.root.title("ARC \u29BF Windows EcoSystem Mapper")
        self._apply_window_icon()
        self.style = ttk.Style(self.root)
        self.scanner = AppScanner()
        self.related_scanner = RelatedFileScanner()
        self.scan_names: Set[str] = set()
        self.app_index: Dict[str, AppEntry] = {}
        self.size_cache: Dict[str, Dict[str, object]] = {}
        self.related_overrides: Dict[str, str] = {}
        self.related_manual: Dict[str, List[Dict[str, str]]] = {}
        self.related_ignore: Dict[str, List[str]] = {}
        self.related_unassigned: Dict[str, List[str]] = {}
        self.install_location_overrides: Dict[str, str] = {}
        self.version_overrides: Dict[str, str] = {}
        self.install_date_overrides: Dict[str, str] = {}
        self._filter_job: Optional[str] = None
        self._bg_queue: queue.Queue = queue.Queue()
        self._bg_poll_job: Optional[str] = None
        self._scan_in_progress = False
        self._related_in_progress = False
        self._size_in_progress = False
        self._scan_job_id = 0
        self._related_job_id = 0
        self._size_job_id = 0
        self._size_pending: Dict[str, AppEntry] = {}
        self._size_pending_keys: Set[str] = set()
        self._size_unavailable: Set[str] = set()
        self._related_pending: Dict[str, AppEntry] = {}
        self._related_focus_app_key: Optional[str] = None
        self._collapse_related_on_enter = False
        self._deep_scan_job_id = 0
        self._deep_scan_in_progress = False
        self._deep_scan_app_key: str = ""
        self._deep_scan_results: List[RelatedFile] = []
        self._deep_scan_row_map: Dict[str, RelatedFile] = {}
        self._progress_active = False
        self._path_norm_cache: Dict[str, str] = {}
        self._map_node_cache: Dict[str, Dict[str, object]] = {}
        self._map_refresh_job: Optional[str] = None
        self._related_index_dirty = True
        self._related_index_rows: Dict[str, RelatedRow] = {}
        self._related_index_app_rows: Dict[str, List[RelatedRow]] = {}
        self._related_index_app_search: Dict[str, str] = {}
        self._related_gram_index: Dict[str, Set[str]] = {}

        self.all_apps: List[AppEntry] = []
        self.current_scan: List[AppEntry] = []
        self.reference_apps: List[AppEntry] = []
        self.reference_source: str = ""
        self.reference_dirty = False
        self.reference_saved_path: str = ""
        self.display_mode = "scan"
        self.view_mode = "system"
        self.filtered_apps: List[AppEntry] = []
        self.displayed_apps: List[AppEntry] = []
        self.related_row_map: Dict[str, Tuple[AppEntry, RelatedFile]] = {}
        self.related_parent_map: Dict[str, AppEntry] = {}
        self.sort_column = "name"
        self.sort_reverse = False
        self.sort_label_map = {label: column for label, column in self.sort_options}
        self.sort_column_map = {column: label for label, column in self.sort_options}
        self.filter_query = ""
        self.group_filter_mode = "all"

        self.default_gui_settings = self._default_gui_settings()
        self.gui_settings = dict(self.default_gui_settings)
        self.groups: List[str] = []
        self.app_groups: Dict[str, str] = {}
        self.group_colors: Dict[str, str] = {}
        self.scan_drives: List[str] = []
        self.available_drives: List[str] = []

        prompt_for_dir = None if self._should_skip_state_prompt() else self._prompt_state_dir
        self.state_path = default_state_path(STATE_FILE, prompt_for_dir)
        self._handle_missing_state_file()
        self._load_state()
        self._ensure_group_colors()
        self.available_drives = self._detect_drives()
        self.scan_drives = self._normalize_scan_drives(self.scan_drives, self.available_drives)

        # Default window size and minimum resize bounds.
        self.root.geometry(self.state_geometry or "1000x600")
        self.root.minsize(900, 400)

        self.view = MainView(
            self.root,
            columns=self.columns,
            widths=self.column_widths,
            related_columns=self.related_columns,
            related_widths=self.related_column_widths,
            sort_labels=[label for label, _ in self.sort_options],
            callbacks={
                "on_scan": self.trigger_scan,
                "on_export": self.export_csv,
                "on_import": self.import_csv,
                "on_open_json": self.open_json,
                "on_save_json": self.save_json,
                "on_open_settings": self.open_settings,
                "on_clear_scan": self.clear_scan,
                "on_close_reference": self.close_reference,
                "on_view_change": self.on_view_change,
                "on_sort_change": self.on_sort_option_change,
                "on_sort_toggle": self.toggle_sort,
                "on_sort_column": self.sort_by_column,
                "on_filter_change": self.on_filter_change,
                "on_clear_filter": self.clear_filter,
                "on_group_filter_change": self.on_group_filter_change,
                "on_group_double_click": self.on_group_double_click,
                "on_related_double_click": self.on_related_double_click,
                "on_related_add_files": self.on_related_add_files,
                "on_related_add_folder": self.on_related_add_folder,
                "on_related_deep_scan": self.on_related_deep_scan,
                "on_related_remove_manual": self.on_related_remove_manual,
                "on_related_unassign": self.on_related_unassign,
                "on_related_reassign": self.on_related_reassign,
                "on_website_click": self.on_website_click,
                "on_open_install_location": self.open_install_location,
                "on_set_install_location": self.on_set_install_location,
                "on_clear_install_location": self.on_clear_install_location,
                "on_set_version": self.on_set_version,
                "on_set_install_date": self.on_set_install_date,
                "on_view_related_for_app": self.view_related_for_app,
                "on_row_select": self.on_row_select,
                "on_close": self.on_close,
            },
        )

        self.settings_view = SettingsView(
            self.root,
            callbacks={
                "on_apply": self.apply_settings_from_view,
                "on_restore_defaults": self.reset_gui_settings,
                "on_close": self.close_settings,
                "on_add_group": self.add_group,
                "on_rename_group": self.rename_group,
                "on_delete_group": self.delete_group,
                "on_set_group_color": self.set_group_color,
            },
        )

        self._apply_gui_settings()
        self._apply_window_state()

    def _load_state(self) -> None:
        state = load_state(self.state_path, self.default_gui_settings)
        self.state_geometry = state.geometry
        self.sort_column = state.sort_column or self.sort_column
        self.sort_reverse = bool(state.sort_reverse)
        self.gui_settings = dict(state.gui_settings)
        self.groups = list(state.groups)
        self.app_groups = dict(state.app_groups)
        self.group_colors = dict(getattr(state, "group_colors", {}))
        self.scan_drives = list(getattr(state, "scan_drives", []))
        self.size_cache = dict(state.size_cache)
        self.related_overrides = dict(getattr(state, "related_overrides", {}))
        self.related_manual = dict(getattr(state, "related_manual", {}))
        self.related_ignore = dict(getattr(state, "related_ignore", {}))
        self.related_unassigned = dict(getattr(state, "related_unassigned", {}))
        self.install_location_overrides = dict(getattr(state, "install_location_overrides", {}))
        self.version_overrides = dict(getattr(state, "version_overrides", {}))
        self.install_date_overrides = dict(getattr(state, "install_date_overrides", {}))

    def _apply_window_state(self) -> None:
        self.view.set_sort_desc(self.sort_reverse)
        self.view.set_sort_label(self._label_for_column(self.sort_column))
        self.view.set_view_mode(self.view_mode)
        self.view.set_sort_enabled(self.view_mode == "system")
        self._update_export_state()
        self.view.set_save_json_enabled(bool(self.current_scan))
        self.view.set_clear_scan_enabled(bool(self.current_scan))
        self.view.set_close_reference_enabled(bool(self.reference_apps))
        self._update_related_view_state()
        self._update_map_view_state()
        self.view.set_map_group_colors(self.group_colors)
        self.view.set_map_style(self.gui_settings)

    def _apply_window_icon(self) -> None:
        if not os.path.isfile(WINDOW_ICON_PATH):
            return
        try:
            self.root.iconbitmap(default=WINDOW_ICON_PATH)
        except tk.TclError:
            pass

    def _should_skip_state_prompt(self) -> bool:
        if os.getenv("ARC_SKIP_DATA_PROMPT") == "1":
            return True
        if os.getenv("PYTEST_CURRENT_TEST"):
            return True
        return False

    def _prompt_state_dir(self) -> str:
        message = (
            "Choose where ARC should store its persistent data (settings, groups, scan cache).\n\n"
            "Tip: pick a folder on a secondary drive if you want it to survive a primary drive failure.\n\n"
            "Click Yes to choose a folder, or No to use the default AppData location."
        )
        try:
            choose = messagebox.askyesno("Choose data location", message, parent=self.root)
        except tk.TclError:
            return ""
        if not choose:
            return ""
        folder = filedialog.askdirectory(parent=self.root, title="Choose data folder for ARC")
        return folder or ""

    def _handle_missing_state_file(self) -> None:
        if os.path.exists(self.state_path):
            return
        if self._should_skip_state_prompt():
            return
        message = (
            "ARC can't find its persistent state file:\n\n"
            f"{self.state_path}\n\n"
            "Was it moved or deleted?\n\n"
            "Yes = locate the existing file\n"
            "No = choose a new data folder for a replacement file\n"
            "Cancel = keep the current location and start fresh"
        )
        try:
            choice = messagebox.askyesnocancel("State file missing", message, parent=self.root)
        except tk.TclError:
            return
        if choice is None:
            return
        if choice:
            self._locate_existing_state_file()
            return
        self._choose_new_data_dir()

    def _locate_existing_state_file(self) -> None:
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Locate ARC state file",
        )
        if not file_path:
            return
        self._use_state_file(file_path)

    def _choose_new_data_dir(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="Choose new ARC data folder")
        if not folder:
            return
        self.state_path = os.path.join(folder, STATE_FILE)
        set_configured_data_dir(folder)

    def _use_state_file(self, selected_path: str) -> None:
        folder = os.path.dirname(selected_path)
        if not folder:
            return
        target_path = os.path.join(folder, STATE_FILE)
        if os.path.normcase(os.path.abspath(selected_path)) != os.path.normcase(os.path.abspath(target_path)):
            try:
                os.makedirs(folder, exist_ok=True)
                shutil.copyfile(selected_path, target_path)
            except OSError as exc:
                messagebox.showwarning(
                    "State file copy failed",
                    f"Could not copy the selected file to:\n\n{target_path}\n\n{exc}",
                    parent=self.root,
                )
                return
        self.state_path = target_path
        set_configured_data_dir(folder)

    def trigger_scan(self) -> None:
        if self._scan_in_progress:
            return
        self._invalidate_background_jobs()
        self._size_unavailable.clear()
        self.view.set_scan_enabled(False)
        self.view.set_status("Scanning system...")
        self._scan_in_progress = True
        self._update_progress_indicator()
        self._scan_job_id += 1
        job_id = self._scan_job_id
        thread = threading.Thread(target=self._scan_worker, args=(job_id,), daemon=True)
        thread.start()
        self._ensure_polling()

    def _scan_worker(self, job_id: int) -> None:
        try:
            apps = self.scanner.scan(include_sizes=False)
            self._bg_queue.put(("scan_complete", job_id, apps))
        except Exception as exc:  # pragma: no cover - log to UI
            self._bg_queue.put(("scan_error", job_id, exc))

    def _ensure_polling(self) -> None:
        if self._bg_poll_job is None:
            self._bg_poll_job = self.root.after(100, self._poll_bg_queue)

    def _poll_bg_queue(self) -> None:
        while True:
            try:
                event = self._bg_queue.get_nowait()
            except queue.Empty:
                break
            self._handle_bg_event(event)
        if (
            self._scan_in_progress
            or self._related_in_progress
            or self._size_in_progress
            or self._size_pending
            or self._related_pending
            or self._deep_scan_in_progress
        ):
            self._bg_poll_job = self.root.after(100, self._poll_bg_queue)
        else:
            self._bg_poll_job = None

    def _update_progress_indicator(self) -> None:
        running = self._scan_in_progress or self._related_in_progress or self._deep_scan_in_progress
        if running == self._progress_active:
            return
        self._progress_active = running
        self.view.set_progress_running(running)

    def _handle_bg_event(self, event: Tuple) -> None:
        kind = event[0]
        if kind == "scan_complete":
            _kind, job_id, apps = event
            if job_id != self._scan_job_id:
                return
            self._scan_in_progress = False
            self._apply_scan_results(apps)
            self.view.set_scan_enabled(True)
            self._update_progress_indicator()
            return
        if kind == "scan_error":
            _kind, job_id, exc = event
            if job_id != self._scan_job_id:
                return
            self._scan_in_progress = False
            messagebox.showerror("Scan failed", str(exc))
            self.view.set_scan_enabled(True)
            self._update_progress_indicator()
            return
        if kind == "related_complete":
            _kind, job_id, keys, token = event
            if job_id != self._related_job_id:
                return
            self._related_in_progress = False
            current_token = self._related_scan_token()
            if token != current_token:
                for key in keys:
                    app = self.app_index.get(key)
                    if app:
                        app.related_files = []
                        app.related_scanned = False
                        app.related_scan_token = ""
            else:
                for key in keys:
                    app = self.app_index.get(key)
                    if app:
                        app.related_scanned = True
                        app.related_scan_token = token
            self._invalidate_related_index()
            self._update_related_view_state()
            self._apply_manual_related(self.all_apps)
            self._apply_related_overrides(self.all_apps)
            self._apply_related_ignores(self.all_apps)
            self._apply_related_unassigned(self.all_apps)
            self._clear_map_cache()
            if self.view_mode in {"related", "map"}:
                self.apply_filter()
            if self._related_pending:
                self._ensure_related_scan([])
            self._update_progress_indicator()
            return
        if kind == "related_error":
            _kind, job_id, exc = event
            if job_id != self._related_job_id:
                return
            self._related_in_progress = False
            messagebox.showwarning("Related file scan failed", str(exc))
            if self.view_mode in {"related", "map"}:
                self.apply_filter()
            if self._related_pending:
                self._ensure_related_scan([])
            self._update_progress_indicator()
            return
        if kind == "deep_scan_complete":
            _kind, job_id, app_key, results = event
            if job_id != self._deep_scan_job_id:
                return
            if not isinstance(results, list):
                self._deep_scan_in_progress = False
                messagebox.showwarning("Deep scan failed", "Unexpected deep scan result payload.")
                self._update_progress_indicator()
                return
            self._deep_scan_in_progress = False
            self._show_deep_scan_results(app_key, results)
            self._update_progress_indicator()
            return
        if kind == "deep_scan_error":
            _kind, job_id, exc = event
            if job_id != self._deep_scan_job_id:
                return
            self._deep_scan_in_progress = False
            messagebox.showwarning("Deep scan failed", str(exc))
            self._update_progress_indicator()
            return
        if kind == "size_complete":
            _kind, job_id, results = event
            if job_id != self._size_job_id:
                return
            self._size_in_progress = False
            self._apply_size_results(results)
            if self._size_pending:
                self._start_size_scan()
            return
        if kind == "size_error":
            _kind, job_id, exc = event
            if job_id != self._size_job_id:
                return
            self._size_in_progress = False
            messagebox.showwarning("Size calculation failed", str(exc))
            if self._size_pending:
                self._start_size_scan()
            return

    def _invalidate_background_jobs(self) -> None:
        self._related_job_id += 1
        self._related_in_progress = False
        self._related_pending.clear()
        self._size_job_id += 1
        self._size_in_progress = False
        self._size_pending.clear()
        self._size_pending_keys.clear()
        self._update_progress_indicator()

    def _apply_scan_results(self, apps: List[AppEntry]) -> None:
        self._size_pending.clear()
        self._size_pending_keys.clear()
        self._size_unavailable.clear()
        self._clear_path_cache()
        self._clear_map_cache()
        self.current_scan = apps
        self.scan_names = build_scan_name_set(apps)
        self._apply_size_cache(apps)
        self._reset_related_state(apps)
        if self.reference_apps:
            self.display_mode = "reference"
            self.all_apps = list(self.reference_apps)
        else:
            self.display_mode = "scan"
            self.all_apps = apps
        self._apply_manual_related(self.all_apps)
        self._apply_related_overrides(self.all_apps)
        self._apply_related_ignores(self.all_apps)
        self._apply_related_unassigned(self.all_apps)
        self._refresh_app_index()
        self.view.clear_system_tree()
        self.apply_filter()
        self._update_export_state()
        self.view.set_save_json_enabled(bool(self.current_scan))
        self.view.set_clear_scan_enabled(bool(self.current_scan))
        self._update_related_view_state()
        if self.view_mode == "related":
            self._ensure_related_scan(self.all_apps)

    def _reset_related_state(self, apps: List[AppEntry]) -> None:
        self._clear_map_cache()
        self._invalidate_related_index()
        for app in apps:
            app.related_files = []
            app.related_scanned = False
            app.related_scan_token = ""

    def _sync_reference_related_tokens(self) -> None:
        token = self._related_scan_token()
        for entry in self.reference_apps:
            if entry.related_files:
                entry.related_scanned = True
                entry.related_scan_token = token
            else:
                entry.related_scanned = False
                entry.related_scan_token = ""

    def _apply_size_cache(self, apps: List[AppEntry]) -> None:
        for app in apps:
            key = app.key()
            legacy = app.legacy_key()
            cached = self.size_cache.get(key) or self.size_cache.get(legacy)
            if not cached:
                continue
            location = cached.get("install_location")
            size = cached.get("size_mb")
            if location != self._install_location_for_app(app):
                continue
            if isinstance(size, bool):
                continue
            if not isinstance(size, (int, float)):
                continue
            size_val = int(size)
            if size_val < 0:
                continue
            app.size_mb = size_val
            if legacy in self.size_cache and key not in self.size_cache:
                self.size_cache[key] = cached

    def _refresh_app_index(self) -> None:
        self.app_index = {app.key(): app for app in self.all_apps}
        self._migrate_legacy_keys(self.all_apps)
        self._invalidate_related_index()

    def _migrate_legacy_keys(self, apps: List[AppEntry]) -> None:
        for app in apps:
            new_key = app.key()
            legacy = app.legacy_key()
            if legacy in self.app_groups and new_key not in self.app_groups:
                self.app_groups[new_key] = self.app_groups.pop(legacy)
            if legacy in self.size_cache and new_key not in self.size_cache:
                self.size_cache[new_key] = self.size_cache.pop(legacy)

    def _deep_scan_enabled(self) -> bool:
        return bool(self.gui_settings.get("deep_scan", False))

    def _related_scan_token(self) -> str:
        drives_token = ",".join(self.scan_drives)
        return f"deep={int(self._deep_scan_enabled())}|drives={drives_token}"

    def _ensure_related_scan(self, apps: List[AppEntry]) -> None:
        token = self._related_scan_token()
        for app in apps:
            if app.related_scan_token == token:
                continue
            self._related_pending[app.key()] = app
        if self._related_in_progress or not self._related_pending:
            return
        self._related_in_progress = True
        self._update_progress_indicator()
        self._related_job_id += 1
        job_id = self._related_job_id
        deep_scan = self._deep_scan_enabled()
        targets = list(self._related_pending.values())
        self._related_pending.clear()
        thread = threading.Thread(
            target=self._related_scan_worker,
            args=(job_id, targets, token, deep_scan),
            daemon=True,
        )
        thread.start()
        self._ensure_polling()

    def _related_scan_worker(self, job_id: int, apps: List[AppEntry], token: str, deep_scan: bool) -> None:
        try:
            self.related_scanner.scan(
                apps,
                deep_scan=deep_scan,
                include_files=True,
                extra_roots=self._scan_drive_roots(),
            )
            keys = [app.key() for app in apps]
            self._bg_queue.put(("related_complete", job_id, keys, token))
        except Exception as exc:  # pragma: no cover - log to UI
            self._bg_queue.put(("related_error", job_id, exc))

    def _schedule_size_scan(self, apps: List[AppEntry]) -> None:
        for app in apps:
            if app.size_mb is not None:
                continue
            key = app.key()
            if key in self._size_unavailable:
                continue
            if key in self._size_pending_keys:
                continue
            location = self._install_location_for_app(app)
            if not location:
                continue
            self._size_pending[key] = app
            self._size_pending_keys.add(key)
        if self._size_in_progress or not self._size_pending:
            return
        self._start_size_scan()

    def _start_size_scan(self) -> None:
        if self._size_in_progress or not self._size_pending:
            return
        self._size_in_progress = True
        self._size_job_id += 1
        job_id = self._size_job_id
        targets = list(self._size_pending.values())
        self._size_pending.clear()
        thread = threading.Thread(
            target=self._size_scan_worker,
            args=(job_id, targets, self.SIZE_SCAN_LIMITS),
            daemon=True,
        )
        thread.start()
        self._ensure_polling()

    def _size_scan_worker(self, job_id: int, apps: List[AppEntry], limits: SizeScanLimits) -> None:
        try:
            results: List[Tuple[str, Optional[int], str]] = []
            for app in apps:
                location = self._install_location_for_app(app)
                size = self.scanner.compute_install_size_mb(location, limits)
                results.append((app.key(), size, location))
            self._bg_queue.put(("size_complete", job_id, results))
        except Exception as exc:  # pragma: no cover - log to UI
            self._bg_queue.put(("size_error", job_id, exc))

    def _apply_size_results(self, results: List[Tuple[str, Optional[int], str]]) -> None:
        updated = False
        for key, size, location in results:
            self._size_pending_keys.discard(key)
            if size is None:
                self._size_unavailable.add(key)
                continue
            app = self.app_index.get(key)
            if not app:
                continue
            app.size_mb = size
            self.size_cache[key] = {
                "size_mb": size,
                "install_location": location,
                "updated_at": _dt.datetime.now().isoformat(timespec="seconds"),
            }
            self._update_app_row(app)
            updated = True
        if updated and self.sort_column == "size_mb" and self.view_mode == "system":
            self.apply_sort()

    def on_filter_change(self, value: str) -> None:
        self.filter_query = value.strip().lower()
        if self._filter_job is not None:
            self.root.after_cancel(self._filter_job)
        self._filter_job = self.root.after(self.FILTER_DEBOUNCE_MS, self.apply_filter)

    def clear_filter(self) -> None:
        self.view.set_filter("")

    def on_group_filter_change(self, mode: str) -> None:
        next_mode = "grouped" if mode == "grouped" else "all"
        if next_mode == self.group_filter_mode:
            return
        self.group_filter_mode = next_mode
        self.apply_filter()

    def apply_filter(self) -> None:
        self._filter_job = None
        if self.view_mode == "related":
            self._apply_related_filter()
            return
        query = self.filter_query
        base_apps = self._apps_for_group_filter()
        if not query:
            self.filtered_apps = list(base_apps)
        else:
            def match(app: AppEntry) -> bool:
                return query in self._app_search_blob(app)
            self.filtered_apps = [app for app in base_apps if match(app)]
        if self.view_mode == "map":
            self._apply_map_filter()
        else:
            self.apply_sort()

    def sort_by_column(self, column: str) -> None:
        if self.view_mode != "system":
            return
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            if column in self.sort_column_map:
                self.view.set_sort_label(self.sort_column_map[column])
            self.sort_reverse = False
        self.view.set_sort_desc(self.sort_reverse)
        self.apply_sort()

    def toggle_sort(self) -> None:
        if self.view_mode != "system":
            return
        self.sort_reverse = bool(self.view.sort_desc_var.get())
        self.apply_sort()

    def on_sort_option_change(self) -> None:
        if self.view_mode != "system":
            return
        column = self.sort_label_map.get(self.view.sort_var.get())
        if not column:
            return
        self.sort_column = column
        self.apply_sort()

    def apply_sort(self) -> None:
        if self.view_mode != "system":
            return
        column = self.sort_column
        reverse = self.sort_reverse
        if column == "size_mb":
            self._schedule_size_scan(self.filtered_apps)
        key_func = self._sort_key(column)
        blank_sensitive = column in {"install_date", "size_mb"}
        if blank_sensitive:
            sorted_apps = sorted(self.filtered_apps, key=key_func)
            if reverse:
                non_blank = [a for a in sorted_apps if not self._is_blank(column, a)]
                blanks = [a for a in sorted_apps if self._is_blank(column, a)]
                non_blank.reverse()
                sorted_apps = non_blank + blanks
        else:
            sorted_apps = sorted(self.filtered_apps, key=key_func, reverse=reverse)
        self._populate_tree(sorted_apps)

    def _sort_key(self, column: str):
        if column == "group":
            def group_key(app: AppEntry):
                group = self._group_for_app(app)
                text = group.casefold()
                return locale.strxfrm(text)
            return group_key
        if column == "installed":
            scan_names = self.scan_names
            has_reference = bool(self.reference_apps)
            def installed_key(app: AppEntry):
                installed = is_installed(app, scan_names, has_reference)
                return (not installed, app.name_key())
            return installed_key
        if column == "install_date":
            return lambda app: (app.install_date_value() is None, app.install_date_value() or _dt.date.min)
        if column == "size_mb":
            return lambda app: (app.size_mb is None, app.size_mb or 0)
        def default(app: AppEntry):
            text = (getattr(app, column, "") or "").casefold()
            return locale.strxfrm(text)
        return default

    def _is_blank(self, column: str, app: AppEntry) -> bool:
        if column == "install_date":
            return not (app.install_date or "")
        if column == "size_mb":
            return app.size_mb is None
        return False

    def _label_for_column(self, column: str) -> str:
        return self.sort_column_map.get(column, self.sort_options[0][0])

    def _group_for_app(self, app: AppEntry) -> str:
        key = app.key()
        if key in self.app_groups:
            return self.app_groups[key]
        legacy = app.legacy_key()
        if legacy in self.app_groups:
            return self.app_groups[legacy]
        return app.group or ""

    def _install_location_for_app(self, app: AppEntry) -> str:
        key = app.key()
        override = self.install_location_overrides.get(key)
        if override:
            return override
        legacy = app.legacy_key()
        override = self.install_location_overrides.get(legacy)
        if override:
            return override
        return app.install_location or ""

    def _version_for_app(self, app: AppEntry) -> str:
        key = app.key()
        override = self.version_overrides.get(key)
        if override:
            return override
        legacy = app.legacy_key()
        override = self.version_overrides.get(legacy)
        if override:
            return override
        return app.version or ""

    def _install_date_for_app(self, app: AppEntry) -> str:
        key = app.key()
        override = self.install_date_overrides.get(key)
        if override:
            return override
        legacy = app.legacy_key()
        override = self.install_date_overrides.get(legacy)
        if override:
            return override
        return app.install_date or ""

    def _normalize_install_override(self, raw_path: str) -> str:
        if not raw_path:
            return ""
        path = raw_path.strip().strip('"')
        if not path:
            return ""
        lower = path.casefold()
        if ".exe" in lower:
            idx = lower.find(".exe")
            path = path[: idx + 4].strip()
        path = os.path.expandvars(os.path.expanduser(path))
        if os.path.isfile(path):
            path = os.path.dirname(path)
        return path.strip().rstrip("\\/")

    def _populate_tree(self, items: List[AppEntry]) -> None:
        rows: List[Tuple[str, List[str], str]] = []
        scan_names = self.scan_names if self.reference_apps else set()
        has_reference = bool(self.reference_apps)
        for app in items:
            rows.append(self._row_for_app(app, scan_names, has_reference))
        self.view.populate_tree(rows)
        status = f"{len(items)} apps shown"
        if self.display_mode == "reference":
            status += " (reference"
            if self.reference_dirty:
                status += "*"
            status += ")"
            if self.reference_source:
                status += f": {self.reference_source}"
        elif self.display_mode == "scan":
            status += " (scan)"
        self.view.set_status(status)
        self.displayed_apps = list(items)

    def _row_for_app(self, app: AppEntry, scan_names: Set[str], has_reference: bool) -> Tuple[str, List[str], str]:
        app_key = app.key()
        group = self._group_for_app(app)
        installed = is_installed(app, scan_names, has_reference)
        installed_marker = self.installed_marker if installed else self.missing_marker
        website = self._display_website(app.website or "")
        location = self._install_location_for_app(app)
        version = self._version_for_app(app)
        install_date = self._install_date_for_app(app)
        row = [
            app.name or "",
            installed_marker,
            group,
            version,
            install_date,
            "" if app.size_mb is None else str(app.size_mb),
            app.publisher or "",
            location,
            website,
        ]
        tag = "installed" if installed else "missing"
        return (app_key, row, tag)

    def _update_app_row(self, app: AppEntry) -> None:
        scan_names = self.scan_names if self.reference_apps else set()
        has_reference = bool(self.reference_apps)
        _row_id, values, tag = self._row_for_app(app, scan_names, has_reference)
        self.view.update_system_row(app.key(), values, tag)

    def on_view_change(self, mode: str) -> None:
        if mode == self.view_mode:
            return
        if mode == "related" and not self._related_view_available():
            self.view.set_view_mode("system")
            return
        if mode == "map" and not self._related_view_available():
            self.view.set_view_mode("system")
            return
        self.view_mode = mode if mode in {"system", "related", "map"} else "system"
        if self.view_mode == "related":
            self._collapse_related_on_enter = self.view.consume_view_request("related")
        else:
            self._collapse_related_on_enter = False
        self.view.set_view_mode(self.view_mode)
        self.view.set_sort_enabled(self.view_mode == "system")
        if self.view_mode == "related":
            if self._related_focus_app_key:
                app = self.app_index.get(self._related_focus_app_key)
                if app:
                    self._ensure_related_scan([app])
                self._related_focus_app_key = None
            else:
                self._ensure_related_scan(self._apps_for_group_filter())
        elif self.view_mode == "map":
            target_apps = self.filtered_apps or self._apps_for_group_filter()
            if target_apps:
                self._ensure_related_scan(target_apps)
        self.apply_filter()

    def view_related_for_app(self, app_name: str) -> None:
        name = (app_name or "").strip()
        if not name:
            return
        if not self._related_view_available():
            messagebox.showinfo("Related files unavailable", "Run a system scan to view related files.")
            return
        target = None
        folded = name.casefold()
        for app in self.all_apps:
            if app.name and app.name.casefold() == folded:
                target = app
                break
        if target:
            if self.view_mode == "related":
                self._ensure_related_scan([target])
            else:
                self._related_focus_app_key = target.key()
        self.view.set_filter(name)
        self.on_view_change("related")

    def _related_view_available(self) -> bool:
        return bool(self.all_apps)

    def _update_related_view_state(self) -> None:
        enabled = self._related_view_available()
        self.view.set_related_view_enabled(enabled)
        if not enabled and self.view_mode == "related":
            self.view_mode = "system"
            self.view.set_view_mode("system")
            self.view.set_sort_enabled(True)
            self.apply_filter()
        self._update_map_view_state()

    def _update_map_view_state(self) -> None:
        enabled = bool(self.all_apps)
        self.view.set_map_view_enabled(enabled)
        if not enabled and self.view_mode == "map":
            self.view_mode = "system"
            self.view.set_view_mode("system")
            self.view.set_sort_enabled(True)
            self.apply_filter()

    def _apply_related_filter(self) -> None:
        query = self.filter_query.lower()
        base_apps = self._apps_for_group_filter()
        if self._related_in_progress:
            self.view.populate_related_tree([])
            self.displayed_apps = list(base_apps)
            self.view.set_status("Scanning related files...")
            return
        if not query and self._needs_related_scan(base_apps):
            self._ensure_related_scan(base_apps)
            if self._related_in_progress:
                self.view.populate_related_tree([])
                self.displayed_apps = list(base_apps)
                self.view.set_status("Scanning related files...")
                return
        groups, file_count = self._build_related_groups(base_apps, query)
        self.view.populate_related_tree(groups, preserve_expansion=not self._collapse_related_on_enter)
        if self._collapse_related_on_enter:
            self._collapse_related_on_enter = False
        self.displayed_apps = list(base_apps)
        status = f"{file_count} related files shown"
        if self.display_mode == "reference":
            status += " (reference"
            if self.reference_dirty:
                status += "*"
            status += ")"
            if self.reference_source:
                status += f": {self.reference_source}"
        elif self.display_mode == "scan":
            status += " (scan)"
        self.view.set_status(status)

    def _apps_for_group_filter(self) -> List[AppEntry]:
        if self.group_filter_mode != "grouped":
            return list(self.all_apps)
        return [app for app in self.all_apps if self._group_for_app(app)]

    def _apply_map_filter(self) -> None:
        apps = list(self.filtered_apps)
        if apps and self._needs_related_scan(apps):
            self._ensure_related_scan(apps)
        payload = self._build_system_map_payload(apps)
        self.view.set_map_group_colors(self.group_colors)
        self.view.populate_system_map(payload)
        self.displayed_apps = apps
        status = f"{len(apps)} apps mapped"
        if self._related_in_progress:
            status = "Scanning related files..."
        if self.display_mode == "reference":
            status += " (reference"
            if self.reference_dirty:
                status += "*"
            status += ")"
            if self.reference_source:
                status += f": {self.reference_source}"
        elif self.display_mode == "scan":
            status += " (scan)"
        self.view.set_status(status)

    def _build_system_map_payload(self, apps: List[AppEntry]) -> Dict[str, object]:
        drives: Set[str] = set(self.scan_drives)
        payload_apps: List[Dict[str, object]] = []
        max_related = self._map_max_related()
        for app in apps:
            group = self._group_for_app(app)
            drive = self._drive_for_path(self._install_location_for_app(app))
            if drive:
                drives.add(drive)
            related_nodes = self._map_related_nodes(app, max_related, drives)
            payload_apps.append(
                {
                    "id": app.key(),
                    "name": app.name or "",
                    "group": group,
                    "drive": drive or "Unknown",
                    "related": related_nodes,
                }
            )
        if not drives:
            drives.add("Unknown")
        return {"apps": payload_apps, "drives": sorted(drives, key=lambda d: d.casefold())}

    def _map_related_nodes(self, app: AppEntry, limit: int, drives: Set[str]) -> List[Dict[str, str]]:
        if limit <= 0:
            return []
        cached = self._map_node_cache.get(app.key())
        if cached and cached.get("limit") == limit:
            for drive in cached.get("drives", []):
                if drive:
                    drives.add(str(drive))
            return cached.get("nodes", [])
        related_nodes: List[Dict[str, str]] = []
        seen: Set[str] = set()
        related_drives: Set[str] = set()
        related_files = sorted(
            app.related_files or [],
            key=lambda item: (item.kind != "dir", (item.path or "").casefold()),
        )
        for idx, related in enumerate(related_files):
            if len(related_nodes) >= limit:
                break
            path = related.path or ""
            if not path:
                continue
            norm = self._normalize_path_cached(path)
            if norm in seen:
                continue
            seen.add(norm)
            drive = self._drive_for_path(path)
            if drive:
                drives.add(drive)
                related_drives.add(drive)
            label = self._map_label_for_path(path)
            related_nodes.append(
                {
                    "id": f"{app.key()}::{idx}",
                    "label": label,
                    "path": path,
                    "drive": drive or "Unknown",
                    "kind": related.kind or "",
                }
            )
        self._map_node_cache[app.key()] = {
            "limit": limit,
            "nodes": related_nodes,
            "drives": sorted(related_drives),
        }
        return related_nodes

    def _map_label_for_path(self, path: str) -> str:
        cleaned = (path or "").strip().rstrip("\\/").strip('"')
        if not cleaned:
            return ""
        name = os.path.basename(cleaned)
        return name or cleaned

    def _drive_for_path(self, path: str) -> str:
        cleaned = (path or "").strip().strip('"')
        if not cleaned:
            return "Unknown"
        drive, _rest = os.path.splitdrive(cleaned)
        if drive:
            return drive.upper()
        if cleaned.startswith("\\\\"):
            return "Network"
        return "Unknown"

    def _map_max_related(self) -> int:
        raw = self.gui_settings.get("map_max_related", 6)
        try:
            value = int(raw)
        except (TypeError, ValueError):
            return 6
        return max(0, min(50, value))

    def _refresh_map_view(self) -> None:
        self._map_refresh_job = None
        if self.view_mode != "map":
            return
        self._apply_map_filter()

    def _schedule_map_refresh(self) -> None:
        if self.view_mode != "map":
            return
        if self._map_refresh_job is not None:
            self.root.after_cancel(self._map_refresh_job)
        self._map_refresh_job = self.root.after(self.MAP_REFRESH_DEBOUNCE_MS, self._refresh_map_view)

    def _clear_map_cache(self) -> None:
        self._map_node_cache.clear()

    def _invalidate_related_index(self) -> None:
        self._related_index_dirty = True
        self._related_index_rows.clear()
        self._related_index_app_rows.clear()
        self._related_index_app_search.clear()
        self._related_gram_index.clear()

    @staticmethod
    def _related_search_grams(text: str, size: int) -> Set[str]:
        if not text or size <= 0 or len(text) < size:
            return set()
        return {text[idx : idx + size] for idx in range(len(text) - size + 1)}

    def _ensure_related_index(self) -> None:
        if not self._related_index_dirty and self._related_index_rows:
            return
        self._related_index_rows.clear()
        self._related_index_app_rows.clear()
        self._related_index_app_search.clear()
        self._related_gram_index.clear()
        for app in self.all_apps:
            if not app.related_files:
                continue
            app_key = app.key()
            app_search = (self._app_search_blob(app) or "").casefold().strip()
            self._related_index_app_search[app_key] = app_search
            rows: List[RelatedRow] = []
            for idx, related in enumerate(app.related_files):
                row_id = self._related_row_id(app, related, idx)
                related_search = related.search_blob()
                if related_search:
                    row_search = f"{app_search} {related_search}".strip() if app_search else related_search
                else:
                    row_search = app_search
                row = RelatedRow(
                    row_id=row_id,
                    app_key=app_key,
                    index=idx,
                    related=related,
                    search=row_search,
                )
                rows.append(row)
                self._related_index_rows[row_id] = row
                for gram in self._related_search_grams(row_search, 3):
                    bucket = self._related_gram_index.get(gram)
                    if bucket is None:
                        self._related_gram_index[gram] = {row_id}
                    else:
                        bucket.add(row_id)
            if rows:
                self._related_index_app_rows[app_key] = rows
        self._related_index_dirty = False

    def _needs_related_scan(self, apps: List[AppEntry]) -> bool:
        token = self._related_scan_token()
        return any(app.related_scan_token != token for app in apps)

    def _build_related_groups(
        self,
        apps: List[AppEntry],
        query: str,
    ) -> Tuple[List[Tuple[str, List[str], List[Tuple[str, List[str], str]]]], int]:
        self._ensure_related_index()
        groups: List[Tuple[str, List[str], List[Tuple[str, List[str], str]]]] = []
        self.related_row_map = {}
        self.related_parent_map = {}
        file_count = 0
        query_cf = (query or "").casefold().strip()
        app_hits: Set[str] = set()
        matched_by_app: Dict[str, List[RelatedRow]] = {}

        if query_cf:
            for app_key, app_search in self._related_index_app_search.items():
                if query_cf in app_search:
                    app_hits.add(app_key)
            if len(query_cf) >= 3 and self._related_gram_index:
                grams = self._related_search_grams(query_cf, 3)
                if grams:
                    buckets = [self._related_gram_index.get(gram) for gram in grams]
                    if all(buckets):
                        candidate_ids = set.intersection(*buckets)
                    else:
                        candidate_ids = set()
                else:
                    candidate_ids = set()
                for row_id in candidate_ids:
                    row = self._related_index_rows.get(row_id)
                    if not row or row.app_key in app_hits:
                        continue
                    if query_cf in row.search:
                        matched_by_app.setdefault(row.app_key, []).append(row)
            else:
                for row in self._related_index_rows.values():
                    if row.app_key in app_hits:
                        continue
                    if query_cf in row.search:
                        matched_by_app.setdefault(row.app_key, []).append(row)

        for app in apps:
            app_key = app.key()
            rows = self._related_index_app_rows.get(app_key)
            if not rows:
                continue
            if not query_cf:
                selected_rows = rows
            elif app_key in app_hits:
                selected_rows = rows
            else:
                matches = matched_by_app.get(app_key)
                if not matches:
                    continue
                matches.sort(key=lambda item: item.index)
                selected_rows = matches
            children: List[Tuple[str, List[str], str]] = []
            for row in selected_rows:
                related = row.related
                values = [
                    "",
                    related.path,
                    self._display_related_kind(related.kind),
                    self._display_related_source(related.source),
                    related.confidence or "",
                    self._display_related_marked(related.marked),
                ]
                children.append((row.row_id, values, ""))
                self.related_row_map[row.row_id] = (app, related)
            if not children:
                continue
            parent_id = self._related_parent_id(app)
            parent_values = [
                app.name or "",
                "",
                "",
                "",
                "",
                "",
            ]
            groups.append((parent_id, parent_values, children))
            self.related_parent_map[parent_id] = app
            file_count += len(children)
        groups.sort(key=lambda item: (item[1][0].casefold()))
        return groups, file_count

    def _normalize_related_path(self, path: str) -> str:
        return self._normalize_path_cached(path)

    def _normalize_path_cached(self, path: str) -> str:
        cleaned = (path or "").strip().strip('"')
        if not cleaned:
            return ""
        cached = self._path_norm_cache.get(cleaned)
        if cached is not None:
            return cached
        try:
            norm = os.path.normcase(os.path.abspath(cleaned))
        except OSError:
            norm = cleaned.casefold()
        self._path_norm_cache[cleaned] = norm
        return norm

    def _clear_path_cache(self) -> None:
        self._path_norm_cache.clear()

    def _apply_manual_related(self, apps: List[AppEntry]) -> None:
        if not self.related_manual:
            return
        app_by_key = {app.key(): app for app in apps}
        if not app_by_key:
            return
        manual_by_path: Dict[str, Tuple[str, str, str]] = {}
        for app_key, items in self.related_manual.items():
            if app_key not in app_by_key:
                continue
            for entry in items:
                path = str(entry.get("path") or "").strip()
                if not path:
                    continue
                kind = str(entry.get("kind") or "file").strip() or "file"
                norm = self._normalize_related_path(path)
                if not norm:
                    continue
                manual_by_path[norm] = (app_key, kind, path)
        if not manual_by_path:
            return
        for app in apps:
            kept: List[RelatedFile] = []
            seen: Set[str] = set()
            for related in app.related_files or []:
                norm = self._normalize_related_path(related.path)
                if not norm:
                    continue
                owner = manual_by_path.get(norm)
                if owner:
                    if owner[0] != app.key():
                        continue
                    if related.source != "manual":
                        continue
                if norm in seen:
                    continue
                seen.add(norm)
                kept.append(related)
            app.related_files = kept
        for norm, (app_key, kind, raw_path) in manual_by_path.items():
            app = app_by_key.get(app_key)
            if not app:
                continue
            if any(self._normalize_related_path(r.path) == norm for r in app.related_files):
                continue
            app.related_files.append(
                RelatedFile(
                    path=raw_path,
                    kind=kind,
                    source="manual",
                    confidence="High",
                    marked="keep",
                )
            )
        self._invalidate_related_index()

    def _apply_related_ignores(self, apps: List[AppEntry]) -> None:
        if not self.related_ignore:
            return
        ignore_norms: Dict[str, Set[str]] = {}
        for app_key, items in self.related_ignore.items():
            norms = {self._normalize_related_path(path) for path in items if path}
            if norms:
                ignore_norms[app_key] = norms
        if not ignore_norms:
            return
        for app in apps:
            norms = ignore_norms.get(app.key())
            if not norms:
                continue
            kept: List[RelatedFile] = []
            for related in app.related_files or []:
                norm = self._normalize_related_path(related.path)
                if norm and norm in norms and related.source != "manual":
                    continue
                kept.append(related)
            app.related_files = kept
        self._invalidate_related_index()

    def _apply_related_unassigned(self, apps: List[AppEntry]) -> None:
        if not self.related_unassigned:
            return
        unassigned_norms: Dict[str, Set[str]] = {}
        for app_key, items in self.related_unassigned.items():
            norms = {self._normalize_related_path(path) for path in items if path}
            if norms:
                unassigned_norms[app_key] = norms
        if not unassigned_norms:
            return
        for app in apps:
            norms = unassigned_norms.get(app.key())
            if not norms:
                continue
            kept: List[RelatedFile] = []
            for related in app.related_files or []:
                norm = self._normalize_related_path(related.path)
                if norm and norm in norms and related.source != "manual":
                    continue
                kept.append(related)
            app.related_files = kept
        self._invalidate_related_index()

    def _apply_related_overrides(self, apps: List[AppEntry]) -> None:
        if not self.related_overrides:
            return
        app_by_key = {app.key(): app for app in apps}
        if not app_by_key:
            return
        path_lookup: Dict[str, RelatedFile] = {}
        for app in apps:
            for related in app.related_files or []:
                norm = self._normalize_related_path(related.path)
                if not norm:
                    continue
                if norm not in path_lookup:
                    path_lookup[norm] = related
        new_lists: Dict[str, List[RelatedFile]] = {key: [] for key in app_by_key}
        seen_by_app: Dict[str, Set[str]] = {key: set() for key in app_by_key}
        for app in apps:
            app_key = app.key()
            for related in app.related_files or []:
                norm = self._normalize_related_path(related.path)
                if not norm:
                    continue
                target_key = self.related_overrides.get(norm)
                if target_key and target_key != app_key:
                    continue
                if norm in seen_by_app[app_key]:
                    continue
                seen_by_app[app_key].add(norm)
                new_lists[app_key].append(related)
        for norm, target_key in self.related_overrides.items():
            target_app = app_by_key.get(target_key)
            if not target_app:
                continue
            related = path_lookup.get(norm)
            if not related:
                continue
            if norm in seen_by_app[target_key]:
                continue
            seen_by_app[target_key].add(norm)
            new_lists[target_key].append(related)
        for app in apps:
            app.related_files = new_lists.get(app.key(), [])
        self._invalidate_related_index()

    def _related_matches(self, app_search: str, related: RelatedFile, query: str) -> bool:
        if query in app_search:
            return True
        return query in related.search_blob()

    def _app_search_blob(self, app: AppEntry) -> str:
        base = app.search_blob()
        group = self._group_for_app(app)
        if not group:
            return base
        group_cf = group.casefold()
        if group_cf and group_cf not in base:
            return f"{base} {group_cf}".strip()
        return base

    def _related_row_id(self, app: AppEntry, related: RelatedFile, index: int) -> str:
        return f"file::{app.key()}::{index}"

    def _related_parent_id(self, app: AppEntry) -> str:
        return f"app::{app.key()}"

    def _display_related_kind(self, kind: str) -> str:
        return "Folder" if kind == "dir" else "File"

    def _display_related_source(self, source: str) -> str:
        return self.related_source_labels.get(source, source)

    @staticmethod
    def _display_related_marked(value: str) -> str:
        if value == "keep":
            return "Keep"
        if value == "ignore":
            return "Ignore"
        return ""

    def on_related_double_click(self, row_id: str) -> None:
        entry = self.related_row_map.get(row_id)
        if not entry:
            return
        _app, related = entry
        related.marked = self._next_marked_value(related.marked)
        self._invalidate_related_index()
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        self.view.set_related_row_marked(row_id, self._display_related_marked(related.marked))

    def on_related_add_files(self, parent_id: str) -> None:
        app = self.related_parent_map.get(parent_id)
        if not app:
            return
        paths = filedialog.askopenfilenames(parent=self.root, title=f"Add related files for {app.name}")
        if not paths:
            return
        added = self._record_manual_related(app, list(paths), kind="file")
        if added:
            self._apply_manual_related(self.all_apps)
            self._apply_related_overrides(self.all_apps)
            self._apply_related_ignores(self.all_apps)
            self._apply_related_unassigned(self.all_apps)
            if self.view_mode in {"related", "map"}:
                self._clear_map_cache()
                self.apply_filter()

    def on_related_add_folder(self, parent_id: str) -> None:
        app = self.related_parent_map.get(parent_id)
        if not app:
            return
        folder = filedialog.askdirectory(parent=self.root, title=f"Add related folder for {app.name}")
        if not folder:
            return
        file_info = self._collect_folder_files(folder, max_files=5000, max_depth=10)
        if file_info is None:
            messagebox.showwarning("Folder scan failed", "Unable to read the selected folder.", parent=self.root)
            return
        files, hit_limit = file_info
        count_label = f"{len(files)}" + ("+" if hit_limit else "")
        include_all = messagebox.askyesno(
            "Add folder contents?",
            "Add all contents of this folder as related files?\n\n"
            f"This will add {count_label} files.\n"
            "Choose 'No' to add just the folder itself.",
            parent=self.root,
        )
        added_paths: List[str] = []
        if include_all:
            added_paths.append(folder)
            if hit_limit:
                messagebox.showinfo(
                    "Folder scan limit reached",
                    "The folder contains many files. Only the first 5,000 were added.",
                    parent=self.root,
                )
            added_paths.extend(files)
        else:
            added_paths.append(folder)
        added = self._record_manual_related(app, added_paths, kind="file", folder_path=folder)
        if added:
            self._apply_manual_related(self.all_apps)
            self._apply_related_overrides(self.all_apps)
            self._apply_related_ignores(self.all_apps)
            self._apply_related_unassigned(self.all_apps)
            if self.view_mode in {"related", "map"}:
                self._clear_map_cache()
                self.apply_filter()

    def on_related_deep_scan(self, parent_id: str) -> None:
        app = self.related_parent_map.get(parent_id)
        if not app:
            return
        if self._deep_scan_in_progress:
            messagebox.showinfo("Deep scan running", "A deep scan is already in progress.", parent=self.root)
            return
        roots = self._scan_drive_roots()
        if not roots:
            messagebox.showinfo("No drives selected", "Select drives in Settings to run a deep scan.", parent=self.root)
            return
        self._deep_scan_job_id += 1
        job_id = self._deep_scan_job_id
        self._deep_scan_in_progress = True
        self._update_progress_indicator()
        self._deep_scan_app_key = app.key()
        limits = DeepScanLimits()
        thread = threading.Thread(
            target=self._deep_scan_worker,
            args=(job_id, app.key(), roots, limits),
            daemon=True,
        )
        thread.start()
        self._ensure_polling()

    def _deep_scan_worker(
        self,
        job_id: int,
        app_key: str,
        roots: List[str],
        limits: DeepScanLimits,
    ) -> None:
        try:
            app = self.app_index.get(app_key)
            if not app:
                raise RuntimeError("App not found for deep scan.")
            ignored = list(self.related_ignore.get(app_key, []))
            results = self.related_scanner.deep_scan_for_app(app, roots, limits, ignored=ignored)
            self._bg_queue.put(("deep_scan_complete", job_id, app_key, results))
        except Exception as exc:  # pragma: no cover - log to UI
            self._bg_queue.put(("deep_scan_error", job_id, exc))

    def _show_deep_scan_results(self, app_key: str, results: List[RelatedFile]) -> None:
        app = self.app_index.get(app_key)
        if not app:
            return
        existing = {self._normalize_related_path(r.path) for r in app.related_files or [] if r.path}
        manual_items = self.related_manual.get(app_key, [])
        for item in manual_items:
            existing.add(self._normalize_related_path(item.get("path", "")))
        ignored = {self._normalize_related_path(path) for path in self.related_ignore.get(app_key, [])}
        filtered: List[RelatedFile] = []
        for related in results:
            norm = self._normalize_related_path(related.path)
            if not norm or norm in existing or norm in ignored:
                continue
            filtered.append(related)
        filtered.sort(key=lambda item: (-(item.score or 0), (item.path or "").casefold()))
        self._deep_scan_results = filtered
        self._deep_scan_app_key = app_key
        self._deep_scan_row_map = {}
        rows: List[Tuple[str, List[str]]] = []
        for idx, related in enumerate(filtered):
            row_id = f"deep::{idx}"
            self._deep_scan_row_map[row_id] = related
            rows.append(
                (
                    row_id,
                    [
                        related.path,
                        "Folder" if related.kind == "dir" else "File",
                        related.confidence or "",
                        "" if related.score is None else str(related.score),
                    ],
                )
            )
        label = app.name or app_key
        self.view.open_deep_scan_window(
            label,
            rows,
            on_add=self.on_deep_scan_add,
            on_ignore=self.on_deep_scan_ignore,
            on_close=self.on_deep_scan_close,
        )

    def on_deep_scan_add(self, row_ids: List[str]) -> None:
        if not row_ids or not self._deep_scan_app_key:
            return
        app = self.app_index.get(self._deep_scan_app_key)
        if not app:
            return
        paths = []
        for row_id in row_ids:
            related = self._deep_scan_row_map.get(row_id)
            if not related:
                continue
            paths.append(related.path)
        if not paths:
            return
        self._record_manual_related(app, paths, kind="file")
        self._apply_manual_related(self.all_apps)
        self._apply_related_overrides(self.all_apps)
        self._apply_related_ignores(self.all_apps)
        self._apply_related_unassigned(self.all_apps)
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        if self.view_mode in {"related", "map"}:
            self._clear_map_cache()
            self.apply_filter()
        self._remove_deep_scan_rows(row_ids)

    def on_deep_scan_ignore(self, row_ids: List[str]) -> None:
        if not row_ids or not self._deep_scan_app_key:
            return
        app_key = self._deep_scan_app_key
        ignored = self.related_ignore.get(app_key, [])
        ignored_norms = {self._normalize_related_path(path) for path in ignored}
        changed = False
        for row_id in row_ids:
            related = self._deep_scan_row_map.get(row_id)
            if not related:
                continue
            norm = self._normalize_related_path(related.path)
            if not norm or norm in ignored_norms:
                continue
            ignored.append(related.path)
            ignored_norms.add(norm)
            changed = True
        if changed:
            self.related_ignore[app_key] = ignored
            if self.display_mode == "reference":
                self._mark_reference_dirty()
            self._apply_related_ignores(self.all_apps)
            self._apply_related_unassigned(self.all_apps)
            if self.view_mode in {"related", "map"}:
                self._clear_map_cache()
                self.apply_filter()
        self._remove_deep_scan_rows(row_ids)

    def _remove_deep_scan_rows(self, row_ids: List[str]) -> None:
        if not row_ids:
            return
        remaining: List[RelatedFile] = []
        removed = set(row_ids)
        for row_id, related in list(self._deep_scan_row_map.items()):
            if row_id in removed:
                continue
            remaining.append(related)
        self._deep_scan_results = remaining
        self._deep_scan_row_map = {}
        rows: List[Tuple[str, List[str]]] = []
        for idx, related in enumerate(remaining):
            row_id = f"deep::{idx}"
            self._deep_scan_row_map[row_id] = related
            rows.append(
                (
                    row_id,
                    [
                        related.path,
                        "Folder" if related.kind == "dir" else "File",
                        related.confidence or "",
                        "" if related.score is None else str(related.score),
                    ],
                )
            )
        self.view.update_deep_scan_rows(rows)

    def on_deep_scan_close(self) -> None:
        self._deep_scan_results = []
        self._deep_scan_row_map = {}
        self._deep_scan_app_key = ""

    def on_related_remove_manual(self, row_ids: List[str]) -> None:
        if not row_ids:
            return
        removals: Dict[str, Set[str]] = {}
        for row_id in row_ids:
            entry = self.related_row_map.get(row_id)
            if not entry:
                continue
            app, related = entry
            if related.source != "manual":
                continue
            norm = self._normalize_related_path(related.path)
            if not norm:
                continue
            removals.setdefault(app.key(), set()).add(norm)
        if not removals:
            return
        changed = False
        for app_key, norms in removals.items():
            items = self.related_manual.get(app_key, [])
            if not items:
                continue
            kept: List[Dict[str, str]] = []
            for item in items:
                norm = self._normalize_related_path(item.get("path", ""))
                if norm and norm in norms:
                    changed = True
                    continue
                kept.append(item)
            if kept:
                self.related_manual[app_key] = kept
            else:
                self.related_manual.pop(app_key, None)
        if not changed:
            return
        self._apply_manual_related(self.all_apps)
        self._apply_related_overrides(self.all_apps)
        self._apply_related_ignores(self.all_apps)
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        if self.view_mode in {"related", "map"}:
            self._clear_map_cache()
            self.apply_filter()

    def on_related_unassign(self, row_ids: List[str]) -> None:
        if not row_ids:
            return
        updates: Dict[str, Set[str]] = {}
        for row_id in row_ids:
            entry = self.related_row_map.get(row_id)
            if not entry:
                continue
            app, related = entry
            if related.source == "manual":
                continue
            norm = self._normalize_related_path(related.path)
            if not norm:
                continue
            updates.setdefault(app.key(), set()).add(norm)
        if not updates:
            return
        changed = False
        for app_key, norms in updates.items():
            items = self.related_unassigned.get(app_key, [])
            item_norms = {self._normalize_related_path(path) for path in items}
            for norm in norms:
                if norm in self.related_overrides and self.related_overrides.get(norm) == app_key:
                    del self.related_overrides[norm]
                    changed = True
                if norm in item_norms:
                    continue
                items.append(norm)
                item_norms.add(norm)
                changed = True
            if items:
                self.related_unassigned[app_key] = items
        if not changed:
            return
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        self._apply_related_unassigned(self.all_apps)
        if self.view_mode in {"related", "map"}:
            self._clear_map_cache()
            self.apply_filter()

    def on_related_reassign(self, row_ids: List[str]) -> None:
        if not row_ids:
            return
        selection = []
        seen = set()
        for row_id in row_ids:
            entry = self.related_row_map.get(row_id)
            if not entry:
                continue
            _app, related = entry
            norm = self._normalize_related_path(related.path)
            if not norm or norm in seen:
                continue
            seen.add(norm)
            selection.append(norm)
        if not selection:
            return
        options, label_to_key = self._reassign_app_options()
        if not options:
            return

        def on_confirm(label: str) -> None:
            target_key = label_to_key.get(label, "")
            if not target_key:
                return
            for norm in selection:
                self.related_overrides[norm] = target_key
            self._apply_related_overrides(self.all_apps)
            self._apply_related_ignores(self.all_apps)
            self._apply_related_unassigned(self.all_apps)
            if self.display_mode == "reference":
                self._mark_reference_dirty()
            if self.view_mode in {"related", "map"}:
                self._clear_map_cache()
                self.apply_filter()

        self.view.open_reassign_dialog(options, on_confirm, lambda: None)

    def _reassign_app_options(self) -> Tuple[List[str], Dict[str, str]]:
        labels: List[str] = []
        label_to_key: Dict[str, str] = {}
        if not self.all_apps:
            return labels, label_to_key
        base_labels: Dict[str, int] = {}
        for app in self.all_apps:
            name = app.name or ""
            version = app.version or ""
            label = f"{name} ({version})" if version else name
            count = base_labels.get(label, 0)
            base_labels[label] = count + 1
        for app in self.all_apps:
            name = app.name or ""
            version = app.version or ""
            label = f"{name} ({version})" if version else name
            if base_labels.get(label, 0) > 1:
                pub = app.publisher or ""
                if pub:
                    label = f"{label} - {pub}"
            suffix = 1
            final_label = label
            while final_label in label_to_key:
                suffix += 1
                final_label = f"{label} #{suffix}"
            label_to_key[final_label] = app.key()
            labels.append(final_label)
        labels.sort(key=lambda item: item.casefold())
        return labels, label_to_key

    def _collect_folder_files(self, root: str, max_files: int, max_depth: int) -> Optional[Tuple[List[str], bool]]:
        if not root or not os.path.isdir(root):
            return None
        results: List[str] = []
        hit_limit = False
        try:
            root_norm = os.path.normcase(os.path.abspath(root))
        except OSError:
            return None
        root_depth = root_norm.count(os.sep)
        for current, dirs, files in os.walk(root, topdown=True):
            try:
                current_norm = os.path.normcase(os.path.abspath(current))
            except OSError:
                continue
            depth = current_norm.count(os.sep) - root_depth
            if depth >= max_depth:
                dirs[:] = []
            for filename in files:
                if len(results) >= max_files:
                    hit_limit = True
                    return results, hit_limit
                results.append(os.path.join(current, filename))
        return results, hit_limit

    def _record_manual_related(
        self,
        app: AppEntry,
        paths: List[str],
        kind: str,
        folder_path: str = "",
    ) -> int:
        app_key = app.key()
        existing = self.related_manual.get(app_key, [])
        existing_norms = {self._normalize_related_path(item.get("path", "")) for item in existing}
        added = 0
        folder_norm = self._normalize_related_path(folder_path) if folder_path else ""
        for path in paths:
            raw = str(path or "").strip().strip('"')
            if not raw:
                continue
            norm = self._normalize_related_path(raw)
            if not norm or norm in existing_norms:
                continue
            entry_kind = kind
            if folder_norm and norm == folder_norm:
                entry_kind = "dir"
            elif os.path.isdir(raw):
                entry_kind = "dir"
            existing.append({"path": raw, "kind": entry_kind})
            existing_norms.add(norm)
            added += 1
        if added:
            self.related_manual[app_key] = existing
            if self.display_mode == "reference":
                self._mark_reference_dirty()
        return added

    def on_row_select(self, row_id: str) -> None:
        if self.view_mode != "system":
            return
        app = self.app_index.get(row_id)
        if not app:
            return
        if app.size_mb is not None:
            return
        self._schedule_size_scan([app])

    def _next_marked_value(self, current: str) -> str:
        try:
            index = self.related_mark_cycle.index(current)
        except ValueError:
            return self.related_mark_cycle[1]
        return self.related_mark_cycle[(index + 1) % len(self.related_mark_cycle)]

    def export_csv(self) -> None:
        if not self.displayed_apps:
            messagebox.showinfo("Nothing to export", "There are no apps to export.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Export applications to XLSX",
        )
        if not file_path:
            return
        try:
            export_xlsx(file_path, self._entries_with_groups(self.displayed_apps))
        except (OSError, RuntimeError) as exc:
            messagebox.showerror("Export failed", str(exc))

    def import_csv(self) -> None:
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Import applications from CSV",
        )
        if not file_path:
            return
        try:
            apps = import_csv(file_path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Import failed", str(exc))
            return
        if not apps:
            messagebox.showinfo("No apps found", "No applications were found in the CSV file.")
            return
        self._set_reference_apps(apps, os.path.basename(file_path), dirty=True)

    def save_json(self) -> None:
        if not self.current_scan:
            messagebox.showinfo("Nothing to save", "Scan the system before saving a JSON file.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save scan to JSON",
        )
        if not file_path:
            return
        try:
            save_json(file_path, self._entries_with_groups(self.current_scan))
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc))

    def save_reference_json(self) -> bool:
        if not self.reference_apps:
            return False
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save reference scan to JSON",
        )
        if not file_path:
            return False
        try:
            save_json(file_path, self._entries_with_groups(self.reference_apps))
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc))
            return False
        self.reference_dirty = False
        self.reference_saved_path = file_path
        self.reference_source = os.path.basename(file_path)
        self.apply_sort()
        return True

    def open_json(self) -> None:
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Open scan JSON",
        )
        if not file_path:
            return
        try:
            apps = load_json(file_path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Open failed", str(exc))
            return
        if not apps:
            messagebox.showinfo("No apps found", "No applications were found in the JSON file.")
            return
        self._set_reference_apps(apps, os.path.basename(file_path), dirty=False)

    def _set_reference_apps(self, apps: List[AppEntry], source_label: str, dirty: bool) -> None:
        self._invalidate_background_jobs()
        self.reference_source = source_label
        self.reference_apps = list(apps)
        self.display_mode = "reference"
        self.reference_dirty = dirty
        self.reference_saved_path = ""
        incoming_groups = [entry.group for entry in apps if entry.group]
        merged_groups = self._merge_groups(incoming_groups)
        self.groups = merged_groups
        self._ensure_group_colors()
        incoming_map = {entry.key(): entry.group for entry in apps if entry.group}
        normalized_map = self._normalize_group_assignments(incoming_map, merged_groups)
        self.app_groups.update(normalized_map)
        for idx, entry in enumerate(self.reference_apps):
            group = self._group_for_app(entry)
            self.reference_apps[idx] = replace(entry, group=group)
        self._sync_reference_related_tokens()
        self.all_apps = list(self.reference_apps)
        self._refresh_app_index()
        self._size_unavailable.clear()
        self._clear_path_cache()
        self._clear_map_cache()
        self.view.clear_system_tree()
        self.apply_filter()
        self._update_export_state()
        self._update_related_view_state()
        self.view.set_close_reference_enabled(bool(self.reference_apps))
        if self.settings_view.window and self.settings_view.window.winfo_exists():
            self.settings_view.refresh_groups(self.groups, self.group_colors)
        self.view.update_group_editor_values(self.groups, self.no_group_label)

    def _entries_with_groups(self, entries: List[AppEntry]) -> List[AppEntry]:
        with_groups: List[AppEntry] = []
        for entry in entries:
            group = self._group_for_app(entry)
            with_groups.append(replace(entry, group=group))
        return with_groups

    def _update_export_state(self) -> None:
        self.view.set_export_enabled(bool(self.all_apps))

    def clear_scan(self) -> None:
        if not self.current_scan:
            return
        self._invalidate_background_jobs()
        self.current_scan = []
        self.scan_names = set()
        self.view.set_save_json_enabled(False)
        self.view.set_clear_scan_enabled(False)
        if self.reference_apps:
            self.display_mode = "reference"
            self.all_apps = list(self.reference_apps)
        else:
            self.display_mode = "scan"
            self.all_apps = []
        self._refresh_app_index()
        self._size_unavailable.clear()
        self._clear_path_cache()
        self._clear_map_cache()
        self.view.clear_system_tree()
        self.apply_filter()
        self._update_export_state()
        self._update_related_view_state()

    def close_reference(self) -> None:
        if not self.reference_apps:
            return
        if not self._confirm_save_dirty_reference("before closing?"):
            return
        self._invalidate_background_jobs()
        self.reference_apps = []
        self.reference_source = ""
        self.reference_dirty = False
        self.reference_saved_path = ""
        self.display_mode = "scan"
        self.all_apps = list(self.current_scan)
        self._refresh_app_index()
        self._size_unavailable.clear()
        self._clear_path_cache()
        self._clear_map_cache()
        self.view.clear_system_tree()
        self.apply_filter()
        self._update_export_state()
        self.view.set_close_reference_enabled(False)
        self.view.set_save_json_enabled(bool(self.current_scan))
        self.view.set_clear_scan_enabled(bool(self.current_scan))
        self._update_related_view_state()

    def on_group_double_click(self, row_id: str) -> None:
        if not self.groups:
            messagebox.showinfo(
                "No groups",
                "Create a group in Settings > Groups before assigning one.",
                parent=self.root,
            )
            return
        current_group = self.app_groups.get(row_id, "")
        self.view.open_group_editor(
            row_id,
            current_group,
            self.groups,
            self.no_group_label,
            on_save=lambda selection: self._save_group_selection(row_id, selection),
            on_cancel=self.view.cancel_group_editor,
        )

    def _save_group_selection(self, row_id: str, selection: str) -> None:
        group = "" if selection == self.no_group_label else selection
        if group and group not in self.groups:
            group = ""
        if group:
            self.app_groups[row_id] = group
        else:
            self.app_groups.pop(row_id, None)
        self._invalidate_related_index()
        self._mark_reference_dirty()
        self.view.set_row_group(row_id, group)
        self.view.cancel_group_editor()

    def _refresh_groups(self) -> None:
        self.view.update_group_editor_values(self.groups, self.no_group_label)
        if self.settings_view.window and self.settings_view.window.winfo_exists():
            self.settings_view.refresh_groups(self.groups, self.group_colors)

    def _ensure_group_colors(self) -> None:
        updated = False
        for name in self.groups:
            if name not in self.group_colors:
                self.group_colors[name] = self._next_group_color()
                updated = True
        for name in list(self.group_colors.keys()):
            if name not in self.groups:
                del self.group_colors[name]
                updated = True
        if updated and self.view_mode == "map":
            self._schedule_map_refresh()

    def _next_group_color(self) -> str:
        used = {color.casefold() for color in self.group_colors.values() if isinstance(color, str)}
        for color in self.group_color_palette:
            if color.casefold() not in used:
                return color
        return self.group_color_palette[len(used) % len(self.group_color_palette)]

    def _detect_drives(self) -> List[str]:
        if os.name != "nt":
            return [os.path.abspath(os.sep)]
        drives: List[str] = []
        try:
            import string
            import ctypes

            bitmask = ctypes.windll.kernel32.GetLogicalDrives()
            for idx, letter in enumerate(string.ascii_uppercase):
                if bitmask & (1 << idx):
                    path = f"{letter}:"
                    if os.path.exists(path + "\\"):
                        drives.append(path)
        except Exception:
            import string

            for letter in string.ascii_uppercase:
                path = f"{letter}:"
                if os.path.exists(path + "\\"):
                    drives.append(path)
        return drives

    def _normalize_scan_drives(self, drives: List[str], available: List[str], allow_empty: bool = False) -> List[str]:
        available_set = {drive.casefold() for drive in available}
        normalized: List[str] = []
        for drive in drives:
            if not drive:
                continue
            if drive.casefold() in available_set and drive not in normalized:
                normalized.append(drive)
        if not normalized and not allow_empty:
            return list(available)
        return normalized

    def _scan_drive_roots(self) -> List[str]:
        roots: List[str] = []
        for drive in self.scan_drives:
            text = (drive or "").strip()
            if not text:
                continue
            if len(text) == 2 and text[1] == ":":
                root = text + "\\"
            else:
                root = text
            if os.path.isdir(root):
                roots.append(root)
        return roots

    def _merge_groups(self, incoming: List[str]) -> List[str]:
        return unique_casefold(list(self.groups) + list(incoming))

    def _normalize_group_assignments(self, app_groups: Dict[str, str], groups: List[str]) -> Dict[str, str]:
        canonical = {name.casefold(): name for name in groups}
        normalized: Dict[str, str] = {}
        for key, value in app_groups.items():
            canonical_name = canonical.get(value.casefold())
            if not canonical_name:
                continue
            normalized[key] = canonical_name
        return normalized

    def add_group(self) -> None:
        name = self.settings_view.get_group_name()
        if not name:
            messagebox.showinfo("Missing name", "Enter a group name to add.", parent=self.settings_view.window)
            return
        if self._group_exists(name):
            messagebox.showinfo("Duplicate group", "That group already exists.", parent=self.settings_view.window)
            return
        self.groups.append(name)
        self._ensure_group_colors()
        self._mark_reference_dirty()
        self._refresh_groups()
        self.settings_view.clear_group_name()
        if self.view_mode == "map":
            self._schedule_map_refresh()

    def rename_group(self) -> None:
        selected = self.settings_view.get_selected_group()
        if not selected:
            messagebox.showinfo("No selection", "Select a group to rename.", parent=self.settings_view.window)
            return
        new_name = self.settings_view.get_group_name()
        if not new_name:
            messagebox.showinfo("Missing name", "Enter a new name to rename the group.", parent=self.settings_view.window)
            return
        if new_name.casefold() != selected.casefold() and self._group_exists(new_name):
            messagebox.showinfo("Duplicate group", "That group already exists.", parent=self.settings_view.window)
            return
        try:
            index = self.groups.index(selected)
        except ValueError:
            return
        self.groups[index] = new_name
        if selected in self.group_colors:
            self.group_colors[new_name] = self.group_colors.pop(selected)
        for key, value in list(self.app_groups.items()):
            if value == selected:
                self.app_groups[key] = new_name
        self._invalidate_related_index()
        self._mark_reference_dirty()
        self._refresh_groups()
        self.settings_view.select_group_index(index)
        self.settings_view.set_group_name(new_name)
        self.apply_sort()
        if self.view_mode == "map":
            self._schedule_map_refresh()

    def delete_group(self) -> None:
        selected = self.settings_view.get_selected_group()
        if not selected:
            messagebox.showinfo("No selection", "Select a group to delete.", parent=self.settings_view.window)
            return
        if not messagebox.askyesno("Delete group", f"Delete the group '{selected}'?", parent=self.settings_view.window):
            return
        try:
            self.groups.remove(selected)
        except ValueError:
            return
        self.group_colors.pop(selected, None)
        for key, value in list(self.app_groups.items()):
            if value == selected:
                del self.app_groups[key]
        self._invalidate_related_index()
        self._mark_reference_dirty()
        self._refresh_groups()
        self.settings_view.clear_group_name()
        self.apply_sort()
        if self.view_mode == "map":
            self._schedule_map_refresh()

    def set_group_color(self) -> None:
        selected = self.settings_view.get_selected_group()
        if not selected:
            messagebox.showinfo("No selection", "Select a group to set its color.", parent=self.settings_view.window)
            return
        color = self.settings_view.get_group_color()
        if not color:
            messagebox.showinfo("Missing color", "Pick a color for the selected group.", parent=self.settings_view.window)
            return
        self.group_colors[selected] = color
        self._mark_reference_dirty()
        self._refresh_groups()
        if self.view_mode == "map":
            self._schedule_map_refresh()

    def _group_exists(self, name: str) -> bool:
        folded = name.casefold()
        return any(existing.casefold() == folded for existing in self.groups)

    def _mark_reference_dirty(self) -> None:
        if self.reference_apps:
            self.reference_dirty = True

    def _confirm_save_dirty_reference(self, suffix: str) -> bool:
        if not self.reference_dirty:
            return True
        response = messagebox.askyesnocancel(
            "Save reference scan",
            f"The reference scan has unsaved changes. Save it as JSON {suffix}",
            parent=self.root,
        )
        if response is None:
            return False
        if response:
            return self.save_reference_json()
        return True

    def on_website_click(self, value: str) -> None:
        if value == "NOT FOUND":
            return
        url = normalize_url(value)
        if url:
            webbrowser.open(url)
        else:
            messagebox.showinfo("Invalid URL", "No valid website URL found for this entry.")

    def on_set_install_location(self, row_id: str) -> None:
        if not row_id:
            return
        app = self.app_index.get(row_id)
        if not app:
            return
        folder = filedialog.askdirectory(parent=self.root, title=f"Set install location for {app.name}")
        if not folder:
            return
        normalized = self._normalize_install_override(folder)
        if not normalized:
            return
        if not os.path.exists(normalized):
            messagebox.showwarning(
                "Location not found",
                "The selected path does not exist. It will be saved anyway.",
                parent=self.root,
            )
        self.install_location_overrides[app.key()] = normalized
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        self._update_app_row(app)
        self._size_unavailable.discard(app.key())
        self._schedule_size_scan([app])

    def on_clear_install_location(self, row_id: str) -> None:
        if not row_id:
            return
        app = self.app_index.get(row_id)
        if not app:
            return
        removed = False
        for key in (app.key(), app.legacy_key()):
            if key in self.install_location_overrides:
                del self.install_location_overrides[key]
                removed = True
        if not removed:
            return
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        self._update_app_row(app)
        self._size_unavailable.discard(app.key())
        self._schedule_size_scan([app])

    def on_set_version(self, row_id: str) -> None:
        if not row_id:
            return
        app = self.app_index.get(row_id)
        if not app:
            return
        has_override = app.key() in self.version_overrides or app.legacy_key() in self.version_overrides
        if app.version and not has_override:
            messagebox.showinfo(
                "Edit not allowed",
                "Version was provided by the scan and cannot be edited.",
                parent=self.root,
            )
            return
        current = self._version_for_app(app)
        value = simpledialog.askstring(
            "Set Version",
            "Enter version (leave blank to clear):",
            parent=self.root,
            initialvalue=current or "",
        )
        if value is None:
            return
        value = value.strip()
        if not value:
            for key in (app.key(), app.legacy_key()):
                self.version_overrides.pop(key, None)
        else:
            self.version_overrides[app.key()] = value
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        self._update_app_row(app)

    def on_set_install_date(self, row_id: str) -> None:
        if not row_id:
            return
        app = self.app_index.get(row_id)
        if not app:
            return
        has_override = app.key() in self.install_date_overrides or app.legacy_key() in self.install_date_overrides
        if app.install_date and not has_override:
            messagebox.showinfo(
                "Edit not allowed",
                "Install date was provided by the scan and cannot be edited.",
                parent=self.root,
            )
            return
        current = self._install_date_for_app(app)
        value = simpledialog.askstring(
            "Set Install Date",
            "Enter install date (YYYY-MM-DD) or leave blank to clear:",
            parent=self.root,
            initialvalue=current or "",
        )
        if value is None:
            return
        value = value.strip()
        if not value:
            for key in (app.key(), app.legacy_key()):
                self.install_date_overrides.pop(key, None)
        else:
            normalized = normalize_date(value)
            if not normalized:
                messagebox.showerror(
                    "Invalid date",
                    "Enter a valid date in YYYY-MM-DD format.",
                    parent=self.root,
                )
                return
            self.install_date_overrides[app.key()] = normalized
        if self.display_mode == "reference":
            self._mark_reference_dirty()
        self._update_app_row(app)

    def open_install_location(self, raw_path: str) -> None:
        if not raw_path:
            return
        path = raw_path.strip().strip('"')
        lower = path.lower()
        if ".exe" in lower:
            idx = lower.find(".exe")
            path = path[: idx + 4].strip()
        path = os.path.expandvars(os.path.expanduser(path))
        if not os.path.exists(path):
            messagebox.showinfo("Location not found", "The install location path does not exist.")
            return
        target = path
        if os.path.isfile(path):
            target = os.path.dirname(path)
        try:
            os.startfile(target)
        except OSError as exc:
            messagebox.showerror("Open failed", str(exc))

    def open_settings(self) -> None:
        self.available_drives = self._detect_drives()
        self.scan_drives = self._normalize_scan_drives(self.scan_drives, self.available_drives)
        self.settings_view.show(
            self.gui_settings,
            self.groups,
            self.group_colors,
            self.available_drives,
            self.scan_drives,
        )

    def close_settings(self) -> None:
        self.settings_view.destroy()

    def reset_gui_settings(self) -> None:
        previous_deep = bool(self.gui_settings.get("deep_scan", False))
        previous_drives = list(self.scan_drives)
        previous_map_max = self.gui_settings.get("map_max_related")
        self.gui_settings = dict(self.default_gui_settings)
        self.scan_drives = list(self.available_drives)
        self._apply_gui_settings()
        if previous_map_max != self.gui_settings.get("map_max_related"):
            self._clear_map_cache()
        if self.view_mode == "map":
            self._schedule_map_refresh()
        drives_changed = previous_drives != self.scan_drives
        if previous_deep != bool(self.gui_settings.get("deep_scan", False)) or drives_changed:
            self.related_scanner.reset_cache()
            self._reset_related_state(self.current_scan)
            self._sync_reference_related_tokens()
            self._update_related_view_state()
            if self.view_mode == "related":
                self._ensure_related_scan(self.all_apps)
                self.apply_filter()
            if self.view_mode == "map":
                self._ensure_related_scan(self.all_apps)
                self.apply_filter()
        self.settings_view.show(
            self.gui_settings,
            self.groups,
            self.group_colors,
            self.available_drives,
            self.scan_drives,
        )

    def apply_settings_from_view(self) -> None:
        if not self.settings_view.window:
            return
        settings = self.settings_view.get_settings()
        previous = dict(self.gui_settings)
        new_settings = dict(self.gui_settings)
        for key in (
            "window_bg",
            "text_color",
            "table_bg",
            "table_fg",
            "accent",
            "installed_text",
            "missing_text",
            "map_bg",
            "map_text",
            "map_edge",
            "map_drive_bg",
            "map_drive_outline",
            "map_node_outline",
            "map_unknown_group",
            "map_highlight",
        ):
            value = str(settings.get(key, "")).strip()
            if value:
                new_settings[key] = value
        font_family = str(settings.get("font_family", "")).strip()
        if font_family:
            new_settings["font_family"] = font_family
        size_raw = str(settings.get("font_size", "")).strip()
        if size_raw:
            try:
                size = int(size_raw)
            except ValueError:
                messagebox.showerror("Invalid font size", "Font size must be a number.", parent=self.settings_view.window)
                return
            new_settings["font_size"] = max(6, min(72, size))
        map_max_raw = str(settings.get("map_max_related", "")).strip()
        if map_max_raw:
            try:
                map_max = int(map_max_raw)
            except ValueError:
                messagebox.showerror(
                    "Invalid map setting",
                    "Map max related items must be a whole number.",
                    parent=self.settings_view.window,
                )
                return
            new_settings["map_max_related"] = max(0, min(50, map_max))
        new_settings["deep_scan"] = bool(settings.get("deep_scan"))
        selected_drives = self.settings_view.get_selected_drives()
        new_scan_drives = self._normalize_scan_drives(selected_drives, self.available_drives, allow_empty=True)

        self.gui_settings = new_settings
        try:
            self._apply_gui_settings()
        except tk.TclError as exc:
            self.gui_settings = previous
            self._apply_gui_settings()
            messagebox.showerror("Invalid setting", str(exc), parent=self.settings_view.window)
            return
        if previous.get("map_max_related") != new_settings.get("map_max_related"):
            self._clear_map_cache()
        if self.view_mode == "map":
            self._schedule_map_refresh()
        drives_changed = new_scan_drives != self.scan_drives
        if drives_changed:
            self.scan_drives = list(new_scan_drives)
        if previous.get("deep_scan") != new_settings.get("deep_scan") or drives_changed:
            self.related_scanner.reset_cache()
            self._reset_related_state(self.current_scan)
            self._sync_reference_related_tokens()
            self._update_related_view_state()
            if self.view_mode in {"related", "map"}:
                self._ensure_related_scan(self.all_apps)
                self.apply_filter()

    def _style_map_value(self, style_name: str, option: str, state: str) -> str:
        for statespec, value in self.style.map(style_name, option):
            if state in statespec:
                return value
        return ""

    def _default_gui_settings(self) -> Dict[str, object]:
        default_font = tkfont.nametofont("TkDefaultFont")
        available_fonts = set(tkfont.families())
        default_family = "Tahoma" if "Tahoma" in available_fonts else default_font.actual("family")
        return {
            # Adjust default fonts/colors used across the UI.
            "font_family": default_family,
            "font_size": 14,
            "window_bg": self.root.cget("bg"),
            "text_color": self.style.lookup("TLabel", "foreground") or "black",
            "table_bg": self.style.lookup("Treeview", "background") or "black",
            "table_fg": self.style.lookup("Treeview", "foreground") or "black",
            "accent": self._style_map_value("Treeview", "background", "selected") or "#646c74",
            "installed_text": "#1a7f37",
            "missing_text": "#b00020",
            "map_bg": "#f4f6f9",
            "map_text": "#1f2933",
            "map_edge": "#9aa5b1",
            "map_drive_bg": "#d9e2ec",
            "map_drive_outline": "#334e68",
            "map_node_outline": "#52606d",
            "map_unknown_group": "#cbd2d9",
            "map_highlight": "#0b69ff",
            "map_max_related": 6,
            "deep_scan": False,
        }

    def _apply_gui_settings(self) -> None:
        font_family = str(self.gui_settings.get("font_family", "Segoe UI"))
        font_size = int(self.gui_settings.get("font_size", 10))
        window_bg = str(self.gui_settings.get("window_bg", self.root.cget("bg")))
        text_color = str(self.gui_settings.get("text_color", "black"))
        table_bg = str(self.gui_settings.get("table_bg", "white"))
        table_fg = str(self.gui_settings.get("table_fg", "black"))
        accent = str(self.gui_settings.get("accent", "#4a6984"))
        installed_text = str(self.gui_settings.get("installed_text", table_fg))
        missing_text = str(self.gui_settings.get("missing_text", "red"))

        self.base_font = tkfont.Font(family=font_family, size=font_size)
        self.heading_font = tkfont.Font(family=font_family, size=font_size, weight="bold")
        self.root.option_add("*Font", self.base_font)

        # Adjust row height/padding here if the table feels cramped.
        rowheight = max(22, self.base_font.metrics("linespace") + 8)
        self.style.configure(".", font=self.base_font)
        self.style.configure("TFrame", background=window_bg)
        self.style.configure("TLabelframe", background=window_bg)
        self.style.configure("TLabelframe.Label", background=window_bg, foreground=text_color)
        self.style.configure("TLabel", background=window_bg, foreground=text_color)
        self.style.configure("TCheckbutton", background=window_bg, foreground=text_color)
        self.style.configure("TButton", foreground=text_color)
        self.style.configure("TEntry", foreground=text_color)
        self.style.configure("TCombobox", foreground=text_color)
        self.style.configure("Treeview", background=table_bg, foreground=table_fg, fieldbackground=table_bg, rowheight=rowheight)
        self.style.configure("Treeview.Heading", foreground=text_color, font=self.heading_font)
        self.style.map("Treeview", background=[("selected", accent)], foreground=[("selected", table_fg)])
        self.root.configure(bg=window_bg)
        self.view.set_tree_tag_colors(installed_text, missing_text)
        self.view.set_map_style(self.gui_settings)

    def _display_website(self, website: str) -> str:
        if not website.strip():
            return "NOT FOUND"
        return website.strip()

    def on_close(self) -> None:
        if not self._confirm_save_dirty_reference("before exiting?"):
            return
        geometry = self.root.winfo_geometry()
        state = StoredState(
            geometry=geometry,
            sort_column=self.sort_column,
            sort_reverse=self.sort_reverse,
            gui_settings=self.gui_settings,
            groups=self.groups,
            app_groups=self.app_groups,
            scan_drives=self.scan_drives,
            group_colors=self.group_colors,
            size_cache=self.size_cache,
            related_overrides=self.related_overrides,
            related_manual=self.related_manual,
            related_ignore=self.related_ignore,
            related_unassigned=self.related_unassigned,
            install_location_overrides=self.install_location_overrides,
            version_overrides=self.version_overrides,
            install_date_overrides=self.install_date_overrides,
        )
        save_state(self.state_path, state)
        self.close_settings()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    MainController(root)
    root.mainloop()
