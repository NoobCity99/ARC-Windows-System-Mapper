import tkinter as tk
import tkinter.font as tkfont
from tkinter import colorchooser, ttk
from typing import Callable, Dict, List, Optional

# DEFAULT UI VALUES (first-run / app start)
# To change the *default* colors, fonts, and sizes the app starts with,
# edit the _default_gui_settings() method in `main_controller.py`.
# Keys you can adjust there:
# - "font_family", "font_size"
# - "window_bg", "text_color", "table_bg", "table_fg", "accent"
# - "installed_text", "missing_text"
# - "map_bg", "map_text", "map_edge", "map_drive_bg"
# - "map_drive_outline", "map_node_outline", "map_unknown_group", "map_highlight"
# - "map_max_related", "deep_scan"


class SettingsView:
    def __init__(self, root: tk.Tk, callbacks: dict) -> None:
        self.root = root
        self.callbacks = callbacks
        self.window: Optional[tk.Toplevel] = None
        self.settings_vars: Dict[str, tk.StringVar] = {}
        self.groups_list: Optional[tk.Listbox] = None
        self.group_name_var: Optional[tk.StringVar] = None
        self.group_color_var: Optional[tk.StringVar] = None
        self.group_colors: Dict[str, str] = {}
        self.drive_vars: Dict[str, tk.BooleanVar] = {}
        self.drive_order: List[str] = []
        self.drive_container: Optional[ttk.Frame] = None

    def show(
        self,
        gui_settings: Dict[str, object],
        groups: List[str],
        group_colors: Dict[str, str],
        drives: List[str],
        selected_drives: List[str],
    ) -> None:
        if self.window and self.window.winfo_exists():
            self.set_settings(gui_settings)
            self.refresh_groups(groups, group_colors)
            self.set_drive_options(drives, selected_drives)
            self.window.lift()
            self.window.focus_set()
            return
        self.group_colors = dict(group_colors)

        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.transient(self.root)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", self._dispatch("on_close"))
        self.window = win

        container = ttk.Frame(win, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        style = ttk.Style(win)
        style.configure("Settings.TNotebook", borderwidth=2, relief="ridge", tabmargins=(2, 2, 2, 0))
        style.configure("Settings.TNotebook.Tab", padding=(12, 6), borderwidth=2)
        style.map(
            "Settings.TNotebook.Tab",
            relief=[("selected", "ridge"), ("!selected", "groove")],
            background=[("selected", "#e6e6e6"), ("!selected", "#d0d0d0")],
        )

        notebook = ttk.Notebook(container, style="Settings.TNotebook")
        notebook.grid(row=0, column=0, sticky="nsew")

        gui_tab = ttk.Frame(notebook, padding=6)
        notebook.add(gui_tab, text="GUI Customization")
        gui_tab.columnconfigure(0, weight=1)

        gui_frame = ttk.Labelframe(gui_tab, text="GUI Customization", padding=10)
        gui_frame.grid(row=0, column=0, sticky="nsew")
        gui_frame.columnconfigure(1, weight=1)

        color_fields = [
            ("Window background", "window_bg"),
            ("Text color", "text_color"),
            ("Table background", "table_bg"),
            ("Table text", "table_fg"),
            ("Installed text", "installed_text"),
            ("Not installed text", "missing_text"),
            ("Selection accent", "accent"),
        ]
        map_fields = [
            ("Map background", "map_bg"),
            ("Map text", "map_text"),
            ("Map edge", "map_edge"),
            ("Map drive", "map_drive_bg"),
            ("Map drive outline", "map_drive_outline"),
            ("Map node outline", "map_node_outline"),
            ("Map ungrouped", "map_unknown_group"),
            ("Map highlight", "map_highlight"),
        ]

        self.settings_vars = {
            "font_family": tk.StringVar(value=str(gui_settings.get("font_family", ""))),
            "font_size": tk.StringVar(value=str(gui_settings.get("font_size", ""))),
            "deep_scan": tk.BooleanVar(value=bool(gui_settings.get("deep_scan", False))),
        }
        for _, key in color_fields:
            self.settings_vars[key] = tk.StringVar(value=str(gui_settings.get(key, "")))
        for _, key in map_fields:
            self.settings_vars[key] = tk.StringVar(value=str(gui_settings.get(key, "")))
        self.settings_vars["map_max_related"] = tk.StringVar(value=str(gui_settings.get("map_max_related", "")))

        row = 0
        for label, key in color_fields:
            ttk.Label(gui_frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(gui_frame, textvariable=self.settings_vars[key], width=16).grid(row=row, column=1, sticky="w")
            ttk.Button(
                gui_frame,
                text="Pick",
                command=lambda v=self.settings_vars[key]: self._pick_color(win, v),
            ).grid(row=row, column=2, padx=(6, 0))
            row += 1

        ttk.Separator(gui_frame, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=3, sticky="ew", pady=8)
        row += 1
        ttk.Label(gui_frame, text="System Map").grid(row=row, column=0, sticky="w", pady=(0, 6))
        row += 1

        for label, key in map_fields:
            ttk.Label(gui_frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(gui_frame, textvariable=self.settings_vars[key], width=16).grid(row=row, column=1, sticky="w")
            ttk.Button(
                gui_frame,
                text="Pick",
                command=lambda v=self.settings_vars[key]: self._pick_color(win, v),
            ).grid(row=row, column=2, padx=(6, 0))
            row += 1

        ttk.Separator(gui_frame, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=3, sticky="ew", pady=8)
        row += 1
        ttk.Label(gui_frame, text="Map max related").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Spinbox(
            gui_frame,
            from_=0,
            to=50,
            textvariable=self.settings_vars["map_max_related"],
            width=8,
        ).grid(row=row, column=1, sticky="w")
        row += 1

        ttk.Separator(gui_frame, orient=tk.HORIZONTAL).grid(row=row, column=0, columnspan=3, sticky="ew", pady=8)
        row += 1

        ttk.Label(gui_frame, text="Font family").grid(row=row, column=0, sticky="w", pady=4)
        fonts = sorted(tkfont.families())
        ttk.Combobox(
            gui_frame,
            textvariable=self.settings_vars["font_family"],
            values=fonts,
            width=22,
            state="readonly",
        ).grid(row=row, column=1, columnspan=2, sticky="w")
        row += 1

        ttk.Label(gui_frame, text="Font size").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Spinbox(
            gui_frame,
            from_=6,
            to=72,
            textvariable=self.settings_vars["font_size"],
            width=8,
        ).grid(row=row, column=1, sticky="w")

        scan_tab = ttk.Frame(notebook, padding=6)
        notebook.add(scan_tab, text="Scan Options")
        scan_tab.columnconfigure(0, weight=1)

        ttk.Label(scan_tab, text="Related file scanning is on-demand.").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Checkbutton(
            scan_tab,
            text="Deep scan related files (AppData, ProgramData, Documents)",
            variable=self.settings_vars["deep_scan"],
        ).grid(row=1, column=0, sticky="w")

        drive_frame = ttk.Labelframe(scan_tab, text="Drives to include", padding=6)
        drive_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        drive_frame.columnconfigure(0, weight=1)
        self.drive_container = ttk.Frame(drive_frame)
        self.drive_container.grid(row=0, column=0, sticky="ew")
        self.set_drive_options(drives, selected_drives)

        groups_tab = ttk.Frame(notebook, padding=6)
        notebook.add(groups_tab, text="Groups")
        groups_tab.columnconfigure(0, weight=1)
        groups_tab.rowconfigure(1, weight=1)

        ttk.Label(groups_tab, text="Create custom groups to categorize software.").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )

        list_frame = ttk.Frame(groups_tab)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        self.groups_list = tk.Listbox(list_frame, height=8, exportselection=False)
        groups_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.groups_list.yview)
        self.groups_list.configure(yscrollcommand=groups_scroll.set)
        self.groups_list.grid(row=0, column=0, sticky="nsew")
        groups_scroll.grid(row=0, column=1, sticky="ns")

        edit_frame = ttk.Frame(groups_tab)
        edit_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        edit_frame.columnconfigure(1, weight=1)

        ttk.Label(edit_frame, text="Group name").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.group_name_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.group_name_var, width=24).grid(row=0, column=1, sticky="w")

        buttons_frame = ttk.Frame(edit_frame)
        buttons_frame.grid(row=0, column=2, sticky="e", padx=(8, 0))
        ttk.Button(buttons_frame, text="Add", command=self._dispatch("on_add_group")).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(buttons_frame, text="Rename", command=self._dispatch("on_rename_group")).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(buttons_frame, text="Delete", command=self._dispatch("on_delete_group")).grid(row=0, column=2)

        ttk.Label(edit_frame, text="Group color").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(8, 0))
        self.group_color_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.group_color_var, width=12).grid(row=1, column=1, sticky="w", pady=(8, 0))
        color_buttons = ttk.Frame(edit_frame)
        color_buttons.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(8, 0))
        ttk.Button(
            color_buttons,
            text="Pick",
            command=lambda v=self.group_color_var: self._pick_color(win, v),
        ).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(
            color_buttons,
            text="Set Color",
            command=self._dispatch("on_set_group_color"),
        ).grid(row=0, column=1)

        self.groups_list.bind("<<ListboxSelect>>", self._on_group_select)
        self.refresh_groups(groups, group_colors)

        actions = ttk.Frame(container)
        actions.grid(row=1, column=0, sticky="e", pady=(10, 0))
        ttk.Button(actions, text="Restore Defaults", command=self._dispatch("on_restore_defaults")).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(actions, text="Apply", command=self._dispatch("on_apply")).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(actions, text="Close", command=self._dispatch("on_close")).grid(row=0, column=2)

    def set_settings(self, gui_settings: Dict[str, object]) -> None:
        if not self.settings_vars:
            return
        for key, var in self.settings_vars.items():
            if key in gui_settings:
                value = gui_settings.get(key, "")
                if isinstance(var, tk.BooleanVar):
                    var.set(bool(value))
                else:
                    var.set(str(value))

    def _dispatch(self, name: str) -> Callable:
        return self.callbacks.get(name, lambda *args, **kwargs: None)

    def _pick_color(self, parent: tk.Toplevel, target: tk.StringVar) -> None:
        color = colorchooser.askcolor(initialcolor=target.get(), parent=parent)
        if color and color[1]:
            target.set(color[1])

    def refresh_groups(self, groups: List[str], group_colors: Optional[Dict[str, str]] = None) -> None:
        if group_colors is not None:
            self.group_colors = dict(group_colors)
        if not self.groups_list:
            return
        self.groups_list.delete(0, tk.END)
        for name in groups:
            self.groups_list.insert(tk.END, name)
        self._sync_group_color()

    def set_drive_options(self, drives: List[str], selected_drives: List[str]) -> None:
        if not self.drive_container:
            return
        self.drive_vars = {}
        self.drive_order = list(drives)
        selected = {drive for drive in selected_drives}
        for child in self.drive_container.winfo_children():
            child.destroy()
        if not drives:
            ttk.Label(self.drive_container, text="No drives detected.").grid(row=0, column=0, sticky="w")
            return
        for idx, drive in enumerate(drives):
            var = tk.BooleanVar(value=drive in selected)
            self.drive_vars[drive] = var
            ttk.Checkbutton(self.drive_container, text=drive, variable=var).grid(
                row=idx // 6,
                column=idx % 6,
                sticky="w",
                padx=(0, 12),
                pady=2,
            )

    def get_selected_drives(self) -> List[str]:
        if not self.drive_vars:
            return []
        selected: List[str] = []
        for drive in self.drive_order:
            var = self.drive_vars.get(drive)
            if var and var.get():
                selected.append(drive)
        return selected

    def get_group_name(self) -> str:
        return self.group_name_var.get().strip() if self.group_name_var else ""

    def get_selected_group(self) -> str:
        if not self.groups_list:
            return ""
        selection = self.groups_list.curselection()
        if not selection:
            return ""
        return self.groups_list.get(selection[0])

    def select_group_index(self, index: int) -> None:
        if not self.groups_list:
            return
        self.groups_list.selection_clear(0, tk.END)
        self.groups_list.selection_set(index)

    def set_group_name(self, name: str) -> None:
        if self.group_name_var:
            self.group_name_var.set(name)

    def clear_group_name(self) -> None:
        if self.group_name_var:
            self.group_name_var.set("")

    def _on_group_select(self, _event=None) -> None:
        if not self.groups_list or not self.group_name_var:
            return
        selection = self.groups_list.curselection()
        if not selection:
            return
        name = self.groups_list.get(selection[0])
        self.group_name_var.set(name)
        self._sync_group_color()

    def _sync_group_color(self) -> None:
        if not self.group_color_var:
            return
        selected = self.get_selected_group()
        if not selected:
            self.group_color_var.set("")
            return
        color = self.group_colors.get(selected, "")
        self.group_color_var.set(color)

    def get_group_color(self) -> str:
        return self.group_color_var.get().strip() if self.group_color_var else ""

    def get_settings(self) -> Dict[str, str]:
        return {key: var.get() for key, var in self.settings_vars.items()}

    def destroy(self) -> None:
        if self.window and self.window.winfo_exists():
            self.window.destroy()
        self.window = None
        self.settings_vars = {}
        self.groups_list = None
        self.group_name_var = None
        self.group_color_var = None
        self.group_colors = {}
        self.drive_vars = {}
        self.drive_order = []
        self.drive_container = None
