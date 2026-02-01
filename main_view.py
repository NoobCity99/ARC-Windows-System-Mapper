import re
import tkinter as tk
import webbrowser
from tkinter import ttk
from typing import Callable, Dict, List, Optional, Tuple


class MainView:
    def __init__(
        self,
        root: tk.Tk,
        columns: List[Tuple[str, str, str, bool]],
        widths: List[int],
        related_columns: List[Tuple[str, str, str, bool]],
        related_widths: List[int],
        sort_labels: List[str],
        callbacks: dict,
    ) -> None:
        self.root = root
        self.columns = columns
        self.column_keys = [col for col, *_ in columns]
        self.related_columns = related_columns
        self.related_column_keys = [col for col, *_ in related_columns]
        self.callbacks = callbacks
        self.group_editor: Optional[ttk.Combobox] = None
        self.group_editor_var: Optional[tk.StringVar] = None
        self._context_path: str = ""
        self._context_app_name: str = ""
        self._context_related_rows: List[str] = []
        self._context_related_parent: str = ""
        self._context_related_manual_rows: List[str] = []
        self._user_view_request: str = ""
        self._context_row_id: str = ""
        self._system_item_ids: set = set()
        self._map_payload: Optional[Dict[str, object]] = None
        self._map_style: Dict[str, object] = {}
        self._map_group_colors: Dict[str, str] = {}
        self._map_item_styles: Dict[int, Dict[str, object]] = {}
        self._map_highlight_tag: str = ""
        self._drive_band_colors: Dict[str, str] = {}
        self._drive_tag_map: Dict[str, str] = {}
        self._map_context_drive: str = ""
        self._drive_tint_palette = [
            "#b4cbf2",
            "#f6b7ea",
            "#f7d4b5",
            "#e0b8f5",
            "#aceaf2",
            "#f6c1c1",
        ]
        self._tint_icons: Dict[str, tk.PhotoImage] = {}
        self._reassign_window: Optional[tk.Toplevel] = None
        self._deep_scan_window: Optional[tk.Toplevel] = None
        self._deep_scan_tree: Optional[ttk.Treeview] = None
        self._deep_scan_count_var: Optional[tk.StringVar] = None
        self._deep_scan_add_btn: Optional[ttk.Button] = None
        self._deep_scan_ignore_btn: Optional[ttk.Button] = None
        self._about_window: Optional[tk.Toplevel] = None
        self._manual_window: Optional[tk.Toplevel] = None
        self._manual_page_var: Optional[tk.IntVar] = None
        self._manual_page_text: Optional[tk.StringVar] = None
        self._manual_body: Optional[tk.Text] = None
        self.group_filter_var: Optional[tk.StringVar] = None

        self._build_menubar()

        main = ttk.Frame(root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        root.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)

        toolbar = ttk.Frame(main)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        toolbar.columnconfigure(8, weight=1)

        view_frame = ttk.Frame(toolbar)
        view_frame.grid(row=0, column=0, padx=(0, 12))

        self.view_var = tk.StringVar(value="system")
        self.system_view_btn = ttk.Radiobutton(
            view_frame,
            text="SYSTEM VIEW",
            value="system",
            variable=self.view_var,
            command=lambda: self._on_view_request("system"),
        )
        self.system_view_btn.grid(row=0, column=0, padx=(0, 6))
        self.related_view_btn = ttk.Radiobutton(
            view_frame,
            text="RELATED FILES",
            value="related",
            variable=self.view_var,
            command=lambda: self._on_view_request("related"),
        )
        self.related_view_btn.grid(row=0, column=1, padx=(0, 6))
        self.map_view_btn = ttk.Radiobutton(
            view_frame,
            text="SYSTEM MAP",
            value="map",
            variable=self.view_var,
            command=lambda: self._on_view_request("map"),
        )
        self.map_view_btn.grid(row=0, column=2)

        self.scan_btn = ttk.Button(toolbar, text="Scan", command=self._dispatch("on_scan"))
        self.scan_btn.grid(row=0, column=1, padx=(0, 8))

        self.clear_scan_btn = ttk.Button(toolbar, text="Clear Scan", command=self._dispatch("on_clear_scan"), state=tk.DISABLED)
        self.clear_scan_btn.grid(row=0, column=2, padx=(0, 8))

        self.close_ref_btn = ttk.Button(toolbar, text="Close Reference", command=self._dispatch("on_close_reference"), state=tk.DISABLED)
        self.close_ref_btn.grid(row=0, column=3, padx=(0, 12))

        ttk.Label(toolbar, text="Sort by").grid(row=0, column=4, padx=(0, 4))
        self.sort_var = tk.StringVar(value=sort_labels[0] if sort_labels else "")
        self.sort_select = ttk.Combobox(
            toolbar, textvariable=self.sort_var, values=sort_labels, width=14, state="readonly"
        )
        self.sort_select.grid(row=0, column=5, padx=(0, 8))
        self.sort_select.bind("<<ComboboxSelected>>", lambda _e: self._dispatch("on_sort_change")())

        self.sort_desc_var = tk.BooleanVar(value=False)
        self.sort_toggle = ttk.Checkbutton(
            toolbar, text="Descending", variable=self.sort_desc_var, command=self._dispatch("on_sort_toggle")
        )
        self.sort_toggle.state(["!alternate"])
        self.sort_toggle.grid(row=0, column=6, padx=(0, 8))

        ttk.Label(toolbar, text="Filter").grid(row=0, column=7, padx=(0, 4))
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self._dispatch("on_filter_change")(self.filter_var.get()))
        self.filter_entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=24)
        self.filter_entry.grid(row=0, column=8, sticky="ew")

        self.clear_filter_btn = ttk.Button(toolbar, text="Clear Filter", command=self._dispatch("on_clear_filter"))
        self.clear_filter_btn.grid(row=0, column=9, padx=(8, 8))

        self.settings_btn = ttk.Button(toolbar, text="\u2699", width=3, command=self._dispatch("on_open_settings"))
        self.settings_btn.grid(row=0, column=10, padx=(0, 0))

        tree_frame = ttk.Frame(main)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.system_tree_frame = ttk.Frame(tree_frame)
        self.system_tree_frame.grid(row=0, column=0, sticky="nsew")
        self.related_tree_frame = ttk.Frame(tree_frame)
        self.related_tree_frame.grid(row=0, column=0, sticky="nsew")
        self.related_tree_frame.grid_remove()
        self.system_map_frame = ttk.Frame(tree_frame)
        self.system_map_frame.grid(row=0, column=0, sticky="nsew")
        self.system_map_frame.grid_remove()

        self.system_tree = ttk.Treeview(self.system_tree_frame, columns=self.column_keys, show="headings", selectmode="browse")
        system_scrollbar = ttk.Scrollbar(self.system_tree_frame, orient=tk.VERTICAL, command=self.system_tree.yview)
        self.system_tree.configure(yscrollcommand=system_scrollbar.set)
        self.system_tree.grid(row=0, column=0, sticky="nsew")
        system_scrollbar.grid(row=0, column=1, sticky="ns")
        self.system_tree_frame.rowconfigure(0, weight=1)
        self.system_tree_frame.columnconfigure(0, weight=1)

        for (col, heading, anchor, stretch), width in zip(columns, widths):
            self.system_tree.heading(col, text=heading, anchor=anchor, command=lambda c=col: self._dispatch("on_sort_column")(c))
            self.system_tree.column(col, width=width, anchor=anchor, stretch=stretch)

        self.related_tree = ttk.Treeview(self.related_tree_frame, columns=self.related_column_keys, show="tree headings", selectmode="extended")
        related_scrollbar = ttk.Scrollbar(self.related_tree_frame, orient=tk.VERTICAL, command=self.related_tree.yview)
        self.related_tree.configure(yscrollcommand=related_scrollbar.set)
        self.related_tree.grid(row=0, column=0, sticky="nsew")
        related_scrollbar.grid(row=0, column=1, sticky="ns")
        self.related_tree_frame.rowconfigure(0, weight=1)
        self.related_tree_frame.columnconfigure(0, weight=1)

        self.system_map_frame.rowconfigure(0, weight=1)
        self.system_map_frame.columnconfigure(0, weight=1)
        self.map_canvas = tk.Canvas(self.system_map_frame, highlightthickness=0, background="white")
        map_scroll_y = ttk.Scrollbar(self.system_map_frame, orient=tk.VERTICAL, command=self.map_canvas.yview)
        map_scroll_x = ttk.Scrollbar(self.system_map_frame, orient=tk.HORIZONTAL, command=self.map_canvas.xview)
        self.map_canvas.configure(yscrollcommand=map_scroll_y.set, xscrollcommand=map_scroll_x.set)
        self.map_canvas.grid(row=0, column=0, sticky="nsew")
        map_scroll_y.grid(row=0, column=1, sticky="ns")
        map_scroll_x.grid(row=1, column=0, sticky="ew")

        self.related_tree.heading("#0", text="")
        self.related_tree.column("#0", width=24, stretch=False)

        for (col, heading, anchor, stretch), width in zip(related_columns, related_widths):
            self.related_tree.heading(col, text=heading, anchor=anchor)
            self.related_tree.column(col, width=width, anchor=anchor, stretch=stretch)

        self.status_var = tk.StringVar(value="0 apps shown")
        status_frame = ttk.Frame(main)
        status_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        status_frame.columnconfigure(0, weight=1)
        self.status = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
        self.status.grid(row=0, column=0, sticky="ew")
        self.progress = ttk.Progressbar(status_frame, mode="indeterminate", length=120)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.progress.grid_remove()

        self.system_tree.bind("<Button-1>", self._on_tree_click, add=True)
        self.system_tree.bind("<Motion>", self._on_tree_motion, add=True)
        self.system_tree.bind("<Double-1>", self._on_tree_double_click, add=True)
        self.system_tree.bind("<<TreeviewSelect>>", self._on_tree_select, add=True)
        self.system_tree.bind("<MouseWheel>", lambda _e: self.cancel_group_editor(), add=True)
        self.system_tree.bind("<Button-3>", self._on_right_click, add=True)

        self.related_tree.bind("<Double-1>", self._on_related_double_click, add=True)
        self.related_tree.bind("<Button-3>", self._on_related_right_click, add=True)
        self.map_canvas.bind("<Button-1>", self._on_map_click, add=True)
        self.map_canvas.bind("<Button-3>", self._on_map_right_click, add=True)
        self.map_canvas.bind("<MouseWheel>", self._on_map_mousewheel, add=True)
        self.map_canvas.bind("<Shift-MouseWheel>", self._on_map_mousewheel, add=True)
        self.map_canvas.bind("<Button-4>", self._on_map_mousewheel, add=True)
        self.map_canvas.bind("<Button-5>", self._on_map_mousewheel, add=True)
        self.root.protocol("WM_DELETE_WINDOW", self._dispatch("on_close"))
        self.map_context_menu = tk.Menu(self.root, tearoff=0)

    def _dispatch(self, name: str) -> Callable:
        return self.callbacks.get(name, lambda *args, **kwargs: None)

    def _on_view_request(self, mode: str) -> None:
        self._user_view_request = mode
        self._dispatch("on_view_change")(mode)

    def consume_view_request(self, mode: str) -> bool:
        if self._user_view_request == mode:
            self._user_view_request = ""
            return True
        return False

    def set_sort_label(self, label: str) -> None:
        self.sort_var.set(label)

    def set_sort_desc(self, value: bool) -> None:
        self.sort_desc_var.set(value)

    def set_status(self, text: str) -> None:
        self.status_var.set(text)

    def set_progress_running(self, running: bool) -> None:
        if running:
            if not self.progress.winfo_ismapped():
                self.progress.grid()
            self.progress.start(10)
        else:
            self.progress.stop()
            if self.progress.winfo_ismapped():
                self.progress.grid_remove()

    def set_filter(self, value: str) -> None:
        self.filter_var.set(value)

    def set_export_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.file_menu.entryconfig(self.export_label, state=state)

    def set_scan_enabled(self, enabled: bool) -> None:
        self.scan_btn.state(["!disabled"] if enabled else ["disabled"])

    def set_save_json_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.file_menu.entryconfig(self.save_json_label, state=state)

    def set_clear_scan_enabled(self, enabled: bool) -> None:
        self.clear_scan_btn.state(["!disabled"] if enabled else ["disabled"])

    def set_close_reference_enabled(self, enabled: bool) -> None:
        self.close_ref_btn.state(["!disabled"] if enabled else ["disabled"])

    def set_view_mode(self, mode: str) -> None:
        if mode not in {"system", "related", "map"}:
            mode = "system"
        if mode in {"related", "map"}:
            self.cancel_group_editor()
        self.view_var.set(mode)
        if mode == "related":
            self.related_tree_frame.grid()
            self.system_tree_frame.grid_remove()
            self.system_map_frame.grid_remove()
        elif mode == "map":
            self.system_map_frame.grid()
            self.system_tree_frame.grid_remove()
            self.related_tree_frame.grid_remove()
        else:
            self.system_tree_frame.grid()
            self.related_tree_frame.grid_remove()
            self.system_map_frame.grid_remove()

    def set_related_view_enabled(self, enabled: bool) -> None:
        self.related_view_btn.state(["!disabled"] if enabled else ["disabled"])

    def set_map_view_enabled(self, enabled: bool) -> None:
        self.map_view_btn.state(["!disabled"] if enabled else ["disabled"])

    def set_sort_enabled(self, enabled: bool) -> None:
        state = "readonly" if enabled else "disabled"
        self.sort_select.configure(state=state)
        if enabled:
            self.sort_toggle.state(["!disabled"])
        else:
            self.sort_toggle.state(["disabled"])

    def _build_menubar(self) -> None:
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        self.import_label = "Import CSV..."
        self.export_label = "Export XLSX..."
        self.open_json_label = "Open JSON..."
        self.save_json_label = "Save JSON..."
        file_menu.add_command(label=self.import_label, command=self._dispatch("on_import"))
        file_menu.add_command(label=self.export_label, command=self._dispatch("on_export"))
        file_menu.add_separator()
        file_menu.add_command(label=self.open_json_label, command=self._dispatch("on_open_json"))
        file_menu.add_command(label=self.save_json_label, command=self._dispatch("on_save_json"))
        menubar.add_cascade(label="File", menu=file_menu)
        view_menu = tk.Menu(menubar, tearoff=0)
        self.group_filter_var = tk.StringVar(value="all")
        view_menu.add_radiobutton(
            label="All Results",
            value="all",
            variable=self.group_filter_var,
            command=lambda: self._dispatch("on_group_filter_change")(self.group_filter_var.get()),
        )
        view_menu.add_radiobutton(
            label="Grouped Only",
            value="grouped",
            variable=self.group_filter_var,
            command=lambda: self._dispatch("on_group_filter_change")(self.group_filter_var.get()),
        )
        menubar.add_cascade(label="View", menu=view_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="How To", command=self._show_manual)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)
        self.file_menu = file_menu
        self.set_export_enabled(False)
        self.set_save_json_enabled(False)

        self.context_menu = tk.Menu(self.root, tearoff=0)

    def _show_manual(self) -> None:
        if self._manual_window is not None and self._manual_window.winfo_exists():
            self._manual_window.lift()
            self._manual_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title("ARC User Manual")
        window.resizable(True, True)
        window.transient(self.root)
        window.geometry("760x560")

        self._manual_page_var = tk.IntVar(value=1)
        self._manual_page_text = tk.StringVar(value="Page 1 of 3")

        header = ttk.Frame(window, padding=(12, 10, 12, 6))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        prev_btn = ttk.Button(header, text="\u25c0", width=3, command=lambda: self._manual_step(-1))
        prev_btn.grid(row=0, column=0, sticky="w")
        page_label = ttk.Label(header, textvariable=self._manual_page_text)
        page_label.grid(row=0, column=1, sticky="ew")
        next_btn = ttk.Button(header, text="\u25b6", width=3, command=lambda: self._manual_step(1))
        next_btn.grid(row=0, column=2, sticky="e")

        body_frame = ttk.Frame(window, padding=(12, 0, 12, 12))
        body_frame.grid(row=1, column=0, sticky="nsew")
        body_frame.rowconfigure(0, weight=1)
        body_frame.columnconfigure(0, weight=1)

        body = tk.Text(body_frame, wrap="word", height=20, padx=8, pady=6)
        body.configure(state="disabled")
        body.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=body.yview)
        body.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")
        self._manual_body = body

        self._manual_window = window
        window.protocol("WM_DELETE_WINDOW", self._close_manual_window)

        self._manual_update()

    def _manual_step(self, delta: int) -> None:
        if not self._manual_page_var:
            return
        page = self._manual_page_var.get()
        page = max(1, min(3, page + delta))
        self._manual_page_var.set(page)
        self._manual_update()

    def _manual_update(self) -> None:
        if not self._manual_body or not self._manual_page_var:
            return
        page = self._manual_page_var.get()
        if self._manual_page_text:
            self._manual_page_text.set(f"Page {page} of 3")
        content = self._manual_page_content(page)
        self._manual_body.configure(state="normal")
        self._manual_body.delete("1.0", tk.END)
        self._manual_body.insert("1.0", content)
        self._manual_body.configure(state="disabled")

    def _manual_page_content(self, page: int) -> str:
        # DEV NOTE (manual formatting):
        # Tkinter Text does NOT render markdown. To add bold/underline, use Text tags.
        # Suggested approach:
        # 1) Change _manual_page_content to return a list of (text, tag) segments instead of one string.
        # 2) In _show_manual, after creating self._manual_body, define tags once, e.g.:
        #    self._manual_body.tag_configure("manual_bold", font=("TkDefaultFont", 9, "bold"))
        #    self._manual_body.tag_configure("manual_underline", font=("TkDefaultFont", 9, "underline"))
        # 3) In _manual_update, iterate segments and insert with tag:
        #    for text, tag in segments: self._manual_body.insert("end", text, tag or "")
        # 4) Keep headings as tagged segments (e.g., ("System View\n\n", "manual_bold")).
        pages = {
            1: (
                "\u29BF SYSTEM VIEW \u29BF \n\n"
                "\u2726Access: select SYSTEM VIEW in the top toggle; this view is the primary table for scanned apps and reference data.\n"
                "_____\n"
                "\u2726Scan controls: Scan starts a system scan, Clear Scan removes the current scan results, and Close Reference exits a loaded reference dataset (you will be prompted to save if it was changed).\n"
                "_____\n"
                "\u2726File menu (global): Import CSV… and Open JSON… load a reference dataset; Export XLSX… exports the currently displayed apps (including filtered results) with a Related Files sheet; Save JSON… saves the current scan only.\n"
                "_____\n"
                "\u2726View menu (global): All Results shows all apps; Grouped Only shows only apps that have a group assigned.\n"
                "_____\n"
                "\u2726Sorting/filtering: use the Sort by dropdown and Descending toggle; click any column header to sort by that column; the Filter box matches app name and publisher (use Clear Filter to reset).\n"
                "_____\n"
                "\u2726Table interactions: click a website cell to open the URL (cursor becomes a hand); right-click Name → VIEW RELATED FILES; right-click Install Location → Open, Set Install Location…, or Clear Install Location Override; right-click Version → Set Version… (only if the scan did not supply one); right-click Install Date → Set Install Date… (only if the scan did not supply one).\n"
                "_____\n"
                "\u2726Groups: double-click a Group cell to assign a group (groups are created in Settings); group colors carry into the System Map.\n"
                "_____\n"
                "\u2726Size (MB): selecting a row or sorting by Size (MB) triggers background size scans for missing values.\n"
                "_____\n"
                "\u2726Status bar: shows count and whether you are viewing (scan) or (reference); a * indicates unsaved reference changes, and the source filename is shown when applicable.\n"
                "_____\n"
                "\u2726Settings (gear): opens GUI customization (colors/fonts), System Map styling, Map max related limit (0–50), deep scan toggle, drive selection, and group management.\n"
            ),
            2: (
                "\u29BF RELATED FILES \u29BF \n\n"
                "\u2726Access: select RELATED FILES or right-click an app name in System View → VIEW RELATED FILES (this also filters to that app).\n"
                "_____\n"
                "\u2726Availability: this view is disabled until a system scan exists; it runs related-file scanning on demand.\n"
                "_____\n"
                "\u2726Layout: parent rows are apps; child rows list related paths with columns Path, Type, Source, Confidence, and Marked.\n"
                "_____\n"
                "\u2726Filtering: the main Filter box searches app name/publisher plus related file details (path/source/confidence/marked); Grouped Only applies here too.\n"
                "_____\n"
                "\u2726Marking: double-click a child row to cycle Marked through (blank → Keep → Ignore → blank).\n"
                "_____\n"
                "\u2726Parent row context menu (right-click app header): Add Files…, Add Folder…, Deeper Scan….\n"
                "_____\n"
                "\u2726Add Files/Folder: adds manual related items; for folders you can add just the folder or the folder plus its contents (up to 5,000 files, depth 10).\n"
                "_____\n"
                "\u2726Deeper Scan: scans the selected drives (from Settings) for additional candidates; results open in a window where you can Add Selected or Ignore Selected.\n"
                "_____\n"
                "\u2726Child row context menu (right-click Path cell): Open, Reassign to…, Unassign/Unassign Items, Remove Manual Item(s) (manual items only).\n"
                "_____\n"
                "\u2726Reassign: opens a searchable list of apps so selected related items can be re-attached to another app.\n"
                "_____\n"
                "\u2726Status bar: shows related file count and whether you are viewing (scan) or (reference).\n"
            ),
            3: (
                "\u29BF SYSTEM MAP \u29BF \n\n"
                "\u2726Access: select SYSTEM MAP; requires a scan (view is disabled without one).\n"
                "_____\n"
                "\u2726What you see: drives (diamonds), apps (rounded rectangles), and related nodes (rectangles) laid out by drive; app/related colors reflect group colors, ungrouped apps use the Map ungrouped color.\n"
                "_____\n"
                "\u2726Filtering: the main Filter and Grouped Only options limit which apps are mapped.\n"
                "_____\n"
                "\u2726Highlighting: click an app node or drive diamond to highlight its nodes/edges; click empty space to clear the highlight.\n"
                "_____\n"
                "\u2726Navigation: mouse wheel scrolls vertically; Shift + wheel scrolls horizontally.\n"
                "_____\n"
                "\u2726Drive tinting: right-click a drive diamond to apply a lane tint; choose Default to remove the custom tint.\n"
                "_____\n"
                "\u2726Related node limit: the number of related items per app is capped by Settings → Map max related (0 hides related nodes).\n"
                "_____\n"
                "\u2726Status bar: shows apps mapped or a scanning message if related files are still being gathered.\n"
                "_____\n"
                "\u2726Settings (gear): customize map background/text/edge/outline/highlight colors and drive selection used by deeper scans.\n"
            ),
        }
        return pages.get(page, "")

    def _close_manual_window(self) -> None:
        if self._manual_window and self._manual_window.winfo_exists():
            self._manual_window.destroy()
        self._manual_window = None
        self._manual_page_var = None
        self._manual_page_text = None
        self._manual_body = None

    def _show_about(self) -> None:
        if self._about_window is not None and self._about_window.winfo_exists():
            self._about_window.lift()
            self._about_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title("About ARC")
        window.resizable(False, False)
        window.transient(self.root)
        window.configure(padx=16, pady=12)

        about_text = (
            "ARC - Architecture Rebuilding Cartographer\n\n"
            "A Windows EcoSystem Mapping Tool\n\n"
            "Developed by: NoobCity99 (2026)\n\n"
            "Provide feedback or collab at:"
        )
        message = ttk.Label(window, text=about_text, justify="left")
        message.grid(row=0, column=0, sticky="w")

        url = "https://github.com/NoobCity99/ARC-Windows-System-Mapper"
        link_font = ("TkDefaultFont", 9, "underline")
        link = tk.Label(window, text=url, fg="#1a0dab", cursor="hand2", font=link_font)
        link.grid(row=1, column=0, sticky="w", pady=(4, 0))
        link.bind("<Button-1>", lambda _e: webbrowser.open(url))

        self._about_window = window
        window.protocol("WM_DELETE_WINDOW", self._close_about_window)

    def _close_about_window(self) -> None:
        if self._about_window is None:
            return
        if self._about_window.winfo_exists():
            self._about_window.destroy()
        self._about_window = None

    def _open_context_path(self) -> None:
        if self._context_path:
            self._dispatch("on_open_install_location")(self._context_path)

    def _view_related_files(self) -> None:
        if self._context_app_name:
            self._dispatch("on_view_related_for_app")(self._context_app_name)

    def _show_context_menu(self, event: tk.Event, entries: List[Tuple[str, Callable[[], None]]]) -> None:
        if not entries:
            return
        self.context_menu.delete(0, tk.END)
        for label, command in entries:
            self.context_menu.add_command(label=label, command=command)
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _on_right_click(self, event: tk.Event) -> None:
        region = self.system_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.system_tree.identify_column(event.x)
        row_id = self.system_tree.identify_row(event.y)
        if not row_id:
            return
        self._context_row_id = row_id
        col_index = int(column.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.column_keys):
            return
        column_key = self.column_keys[col_index]
        if column_key == "install_location":
            value = self.system_tree.set(row_id, "install_location")
            self._context_path = value or ""
            self._context_app_name = ""
            self.system_tree.selection_set(row_id)
            entries = []
            if value:
                entries.append(("Open", self._open_context_path))
            entries.append(("Set Install Location...", self._set_install_location))
            if value:
                entries.append(("Clear Install Location Override", self._clear_install_location))
            self._show_context_menu(event, entries)
        elif column_key == "version":
            self._context_app_name = ""
            self._context_path = ""
            self.system_tree.selection_set(row_id)
            self._show_context_menu(event, [("Set Version...", self._set_version)])
        elif column_key == "install_date":
            self._context_app_name = ""
            self._context_path = ""
            self.system_tree.selection_set(row_id)
            self._show_context_menu(event, [("Set Install Date...", self._set_install_date)])
        elif column_key == "name":
            value = self.system_tree.set(row_id, "name")
            if not value:
                return
            self._context_app_name = value
            self._context_path = ""
            self.system_tree.selection_set(row_id)
            self._show_context_menu(event, [("VIEW RELATED FILES", self._view_related_files)])
        else:
            return

    def _on_related_right_click(self, event: tk.Event) -> None:
        region = self.related_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.related_tree.identify_column(event.x)
        if column == "#0":
            return
        row_id = self.related_tree.identify_row(event.y)
        if not row_id:
            return
        if self.related_tree.parent(row_id) == "":
            self._context_related_parent = row_id
            self._context_related_rows = []
            self._context_related_manual_rows = []
            self._context_path = ""
            self._context_app_name = ""
            self.related_tree.selection_set(row_id)
            self._show_context_menu(
                event,
                [
                    ("Add Files...", self._add_related_files),
                    ("Add Folder...", self._add_related_folder),
                    ("Deeper Scan...", self._deep_scan_related),
                ],
            )
            return
        selection = list(self.related_tree.selection())
        if row_id not in selection:
            self.related_tree.selection_set(row_id)
            selection = [row_id]
        selected_rows = [item for item in selection if self.related_tree.parent(item) != ""]
        if not selected_rows:
            return
        manual_rows = [
            item for item in selected_rows if self.related_tree.set(item, "source").casefold() == "manual"
        ]
        unassign_rows = [
            item for item in selected_rows if self.related_tree.set(item, "source").casefold() != "manual"
        ]
        col_index = int(column.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.related_column_keys):
            return
        if self.related_column_keys[col_index] != "path":
            return
        value = self.related_tree.set(row_id, "path")
        if not value:
            return
        self._context_path = value
        self._context_app_name = ""
        self._context_related_rows = selected_rows
        self._context_related_manual_rows = manual_rows
        entries = [("Open", self._open_context_path)]
        if selected_rows:
            entries.append(("Reassign to...", self._reassign_related_rows))
        if unassign_rows:
            label = "Unassign" if len(unassign_rows) == 1 else "Unassign Items"
            entries.append((label, self._unassign_related_rows))
        if manual_rows:
            label = "Remove Manual Item" if len(manual_rows) == 1 else "Remove Manual Items"
            entries.append((label, self._remove_manual_related_rows))
        self._show_context_menu(event, entries)

    def _add_related_files(self) -> None:
        if self._context_related_parent:
            self._dispatch("on_related_add_files")(self._context_related_parent)

    def _add_related_folder(self) -> None:
        if self._context_related_parent:
            self._dispatch("on_related_add_folder")(self._context_related_parent)

    def _deep_scan_related(self) -> None:
        if self._context_related_parent:
            self._dispatch("on_related_deep_scan")(self._context_related_parent)

    def _remove_manual_related_rows(self) -> None:
        if self._context_related_manual_rows:
            self._dispatch("on_related_remove_manual")(list(self._context_related_manual_rows))

    def _unassign_related_rows(self) -> None:
        if self._context_related_rows:
            self._dispatch("on_related_unassign")(list(self._context_related_rows))

    def _reassign_related_rows(self) -> None:
        if self._context_related_rows:
            self._dispatch("on_related_reassign")(list(self._context_related_rows))

    def _set_install_location(self) -> None:
        self._dispatch("on_set_install_location")(self._context_row_id)

    def _clear_install_location(self) -> None:
        self._dispatch("on_clear_install_location")(self._context_row_id)

    def _set_version(self) -> None:
        self._dispatch("on_set_version")(self._context_row_id)

    def _set_install_date(self) -> None:
        self._dispatch("on_set_install_date")(self._context_row_id)

    def populate_tree(self, rows: List[Tuple[str, List[str], str]]) -> None:
        self.populate_system_tree(rows)

    def populate_system_tree(self, rows: List[Tuple[str, List[str], str]]) -> None:
        self.cancel_group_editor()
        incoming_ids = [row_id for row_id, _values, _tag in rows]
        incoming_set = set(incoming_ids)
        for row_id in self.system_tree.get_children():
            if row_id not in incoming_set:
                self.system_tree.detach(row_id)
        for index, (row_id, values, tag) in enumerate(rows):
            if self.system_tree.exists(row_id):
                self.system_tree.item(row_id, values=values, tags=(tag,))
            else:
                self.system_tree.insert("", tk.END, iid=row_id, values=values, tags=(tag,))
                self._system_item_ids.add(row_id)
            self.system_tree.move(row_id, "", index)

    def clear_system_tree(self) -> None:
        self.cancel_group_editor()
        if self._system_item_ids:
            self.system_tree.delete(*self._system_item_ids)
            self._system_item_ids.clear()

    def populate_related_tree(
        self,
        groups: List[Tuple[str, List[str], List[Tuple[str, List[str], str]]]],
        preserve_expansion: bool = False,
    ) -> None:
        expanded: set = set()
        if preserve_expansion:
            for item_id in self.related_tree.get_children():
                try:
                    if self.related_tree.item(item_id, "open"):
                        expanded.add(item_id)
                except tk.TclError:
                    continue
        selection = self.related_tree.selection()
        if selection:
            self.related_tree.selection_remove(selection)
        incoming_parents = [parent_id for parent_id, _values, _children in groups]
        incoming_set = set(incoming_parents)
        for item_id in self.related_tree.get_children():
            if item_id not in incoming_set:
                self.related_tree.delete(item_id)
        for parent_index, (parent_id, parent_values, children) in enumerate(groups):
            if self.related_tree.exists(parent_id):
                self.related_tree.item(parent_id, values=parent_values)
            else:
                self.related_tree.insert("", tk.END, iid=parent_id, text="", values=parent_values, open=False)
            self.related_tree.move(parent_id, "", parent_index)
            incoming_children = [row_id for row_id, _values, _tag in children]
            incoming_children_set = set(incoming_children)
            for child_id in self.related_tree.get_children(parent_id):
                if child_id not in incoming_children_set:
                    self.related_tree.delete(child_id)
            for child_index, (row_id, values, tag) in enumerate(children):
                if self.related_tree.exists(row_id):
                    self.related_tree.item(row_id, values=values, tags=(tag,))
                else:
                    self.related_tree.insert(parent_id, tk.END, iid=row_id, values=values, tags=(tag,))
                self.related_tree.move(row_id, parent_id, child_index)
            if preserve_expansion:
                self.related_tree.item(parent_id, open=parent_id in expanded)
            else:
                self.related_tree.item(parent_id, open=False)

    def open_reassign_dialog(
        self,
        options: List[str],
        on_confirm: Callable[[str], None],
        on_cancel: Callable[[], None],
    ) -> None:
        if self._reassign_window and self._reassign_window.winfo_exists():
            self._reassign_window.lift()
            return
        win = tk.Toplevel(self.root)
        self._reassign_window = win
        win.title("Reassign related files")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        frame = ttk.Frame(win, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Filter apps").grid(row=0, column=0, sticky="w")
        filter_var = tk.StringVar()
        filter_entry = ttk.Entry(frame, textvariable=filter_var, width=52)
        filter_entry.grid(row=1, column=0, sticky="ew", pady=(4, 4))
        filter_row = ttk.Frame(frame)
        filter_row.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        filter_row.columnconfigure(0, weight=1)
        match_var = tk.StringVar(value="Matches: 0")
        match_label = ttk.Label(filter_row, textvariable=match_var)
        match_label.grid(row=0, column=0, sticky="w")
        clear_filter_btn = ttk.Button(filter_row, text="Clear Filter", command=lambda: filter_var.set(""))
        clear_filter_btn.grid(row=0, column=1, sticky="e")

        ttk.Label(frame, text="Assign selected files to:").grid(row=3, column=0, sticky="w")
        choice_var = tk.StringVar(value=options[0] if options else "")
        combo = ttk.Combobox(frame, textvariable=choice_var, values=options, state="readonly", width=50)
        combo.grid(row=4, column=0, sticky="ew", pady=(6, 12))

        def apply_filter(*_args: object) -> None:
            text = filter_var.get().strip().casefold()
            if text:
                filtered = [opt for opt in options if text in opt.casefold()]
            else:
                filtered = list(options)
            combo.configure(values=filtered)
            match_var.set(f"Matches: {len(filtered)}")
            if not filtered:
                choice_var.set("")
                apply_btn.state(["disabled"])
                return
            current = choice_var.get()
            if current not in filtered:
                choice_var.set(filtered[0])
            apply_btn.state(["!disabled"])

        filter_var.trace_add("write", apply_filter)

        btns = ttk.Frame(frame)
        btns.grid(row=5, column=0, sticky="e")
        apply_btn = ttk.Button(btns, text="Apply", command=lambda: self._confirm_reassign(win, choice_var, on_confirm))
        cancel_btn = ttk.Button(btns, text="Cancel", command=lambda: self._close_reassign(win, on_cancel))
        apply_btn.grid(row=0, column=0, padx=(0, 8))
        cancel_btn.grid(row=0, column=1)
        apply_filter()

        def on_close() -> None:
            self._close_reassign(win, on_cancel)

        win.protocol("WM_DELETE_WINDOW", on_close)
        if options:
            filter_entry.focus_set()
    def _confirm_reassign(self, window: tk.Toplevel, choice_var: tk.StringVar, on_confirm: Callable[[str], None]) -> None:
        choice = choice_var.get()
        if not choice:
            return
        self._close_reassign(window, lambda: None)
        on_confirm(choice)

    def _close_reassign(self, window: tk.Toplevel, on_cancel: Callable[[], None]) -> None:
        if window and window.winfo_exists():
            window.grab_release()
            window.destroy()
        if self._reassign_window is window:
            self._reassign_window = None
        on_cancel()

    def open_deep_scan_window(
        self,
        app_label: str,
        rows: List[Tuple[str, List[str]]],
        on_add: Callable[[List[str]], None],
        on_ignore: Callable[[List[str]], None],
        on_close: Callable[[], None],
    ) -> None:
        if self._deep_scan_window and self._deep_scan_window.winfo_exists():
            self._deep_scan_window.destroy()
        win = tk.Toplevel(self.root)
        self._deep_scan_window = win
        win.title(f"Deep Scan Results - {app_label}")
        win.transient(self.root)
        win.grab_set()
        win.geometry("800x500")

        frame = ttk.Frame(win, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)

        count_var = tk.StringVar(value=f"Results: {len(rows)}")
        self._deep_scan_count_var = count_var
        ttk.Label(frame, textvariable=count_var).grid(row=0, column=0, sticky="w")

        tree = ttk.Treeview(
            frame,
            columns=("path", "kind", "confidence", "score"),
            show="headings",
            selectmode="extended",
        )
        self._deep_scan_tree = tree
        tree.heading("path", text="Path")
        tree.heading("kind", text="Type")
        tree.heading("confidence", text="Confidence")
        tree.heading("score", text="Score")
        tree.column("path", width=520, anchor=tk.W)
        tree.column("kind", width=80, anchor=tk.CENTER, stretch=False)
        tree.column("confidence", width=100, anchor=tk.CENTER, stretch=False)
        tree.column("score", width=80, anchor=tk.CENTER, stretch=False)
        tree.grid(row=1, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.grid(row=1, column=1, sticky="ns")

        for row_id, values in rows:
            tree.insert("", tk.END, iid=row_id, values=values)

        btns = ttk.Frame(frame)
        btns.grid(row=2, column=0, sticky="e", pady=(10, 0))
        add_btn = ttk.Button(btns, text="Add Selected", command=lambda: on_add(list(tree.selection())))
        ignore_btn = ttk.Button(btns, text="Ignore Selected", command=lambda: on_ignore(list(tree.selection())))
        close_btn = ttk.Button(btns, text="Close", command=lambda: self._close_deep_scan_window(on_close))
        add_btn.grid(row=0, column=0, padx=(0, 8))
        ignore_btn.grid(row=0, column=1, padx=(0, 8))
        close_btn.grid(row=0, column=2)
        self._deep_scan_add_btn = add_btn
        self._deep_scan_ignore_btn = ignore_btn

        def update_buttons(_event: Optional[tk.Event] = None) -> None:
            has_selection = bool(tree.selection())
            state = ["!disabled"] if has_selection else ["disabled"]
            add_btn.state(state)
            ignore_btn.state(state)

        tree.bind("<<TreeviewSelect>>", update_buttons, add=True)
        update_buttons()

        def on_close_window() -> None:
            self._close_deep_scan_window(on_close)

        win.protocol("WM_DELETE_WINDOW", on_close_window)

    def update_deep_scan_rows(self, rows: List[Tuple[str, List[str]]]) -> None:
        if not self._deep_scan_tree or not self._deep_scan_tree.winfo_exists():
            return
        tree = self._deep_scan_tree
        tree.delete(*tree.get_children())
        for row_id, values in rows:
            tree.insert("", tk.END, iid=row_id, values=values)
        if self._deep_scan_count_var:
            self._deep_scan_count_var.set(f"Results: {len(rows)}")
        if self._deep_scan_add_btn and self._deep_scan_ignore_btn:
            self._deep_scan_add_btn.state(["disabled"])
            self._deep_scan_ignore_btn.state(["disabled"])

    def _close_deep_scan_window(self, on_close: Callable[[], None]) -> None:
        win = self._deep_scan_window
        if win and win.winfo_exists():
            win.grab_release()
            win.destroy()
        self._deep_scan_window = None
        self._deep_scan_tree = None
        self._deep_scan_count_var = None
        self._deep_scan_add_btn = None
        self._deep_scan_ignore_btn = None
        on_close()

    def set_related_row_marked(self, row_id: str, value: str) -> None:
        self.related_tree.set(row_id, "marked", value)

    def set_tree_tag_colors(self, installed_color: str, missing_color: str) -> None:
        self.system_tree.tag_configure("installed", foreground=installed_color)
        self.system_tree.tag_configure("missing", foreground=missing_color)

    def set_row_group(self, row_id: str, group: str) -> None:
        self.system_tree.set(row_id, "group", group)

    def update_system_row(self, row_id: str, values: List[str], tag: str) -> None:
        if not self.system_tree.exists(row_id):
            return
        self.system_tree.item(row_id, values=values, tags=(tag,))

    def open_group_editor(
        self,
        row_id: str,
        current_group: str,
        groups: List[str],
        no_group_label: str,
        on_save: Callable[[str], None],
        on_cancel: Callable[[], None],
    ) -> None:
        bbox = self.system_tree.bbox(row_id, "group")
        if not bbox:
            return
        self.cancel_group_editor()
        x, y, width, height = bbox
        display_value = current_group if current_group else no_group_label
        self.group_editor_var = tk.StringVar(value=display_value)
        self.group_editor = ttk.Combobox(
            self.system_tree,
            textvariable=self.group_editor_var,
            values=[no_group_label] + groups,
            state="readonly",
        )
        self.group_editor.place(x=x, y=y, width=width, height=height)
        self.group_editor.focus_set()
        self.group_editor.bind("<<ComboboxSelected>>", lambda _e: on_save(self.group_editor_var.get()))
        self.group_editor.bind("<FocusOut>", lambda _e: on_save(self.group_editor_var.get()))
        self.group_editor.bind("<Return>", lambda _e: on_save(self.group_editor_var.get()))
        self.group_editor.bind("<Escape>", lambda _e: on_cancel())

    def update_group_editor_values(self, groups: List[str], no_group_label: str) -> None:
        if not self.group_editor or not self.group_editor.winfo_exists():
            return
        values = [no_group_label] + groups
        self.group_editor.configure(values=values)
        current = self.group_editor_var.get() if self.group_editor_var else ""
        if current and current not in values:
            if self.group_editor_var:
                self.group_editor_var.set(no_group_label)

    def cancel_group_editor(self) -> None:
        if self.group_editor and self.group_editor.winfo_exists():
            self.group_editor.destroy()
        self.group_editor = None
        self.group_editor_var = None

    def set_map_style(self, style: Dict[str, object]) -> None:
        self._map_style = dict(style)
        bg = str(self._map_style.get("map_bg", "white"))
        self.map_canvas.configure(background=bg)

    def set_map_group_colors(self, group_colors: Dict[str, str]) -> None:
        self._map_group_colors = dict(group_colors)

    def populate_system_map(self, payload: Dict[str, object]) -> None:
        self._map_payload = payload
        self._draw_system_map()

    def _draw_system_map(self) -> None:
        canvas = self.map_canvas
        canvas.delete("all")
        self._map_item_styles = {}
        self._map_highlight_tag = ""
        payload = self._map_payload or {}
        apps = list(payload.get("apps") or [])
        if not apps:
            text_color = str(self._map_style.get("map_text", "black"))
            canvas.create_text(20, 20, text="Run a scan to build the system map.", anchor="nw", fill=text_color)
            canvas.configure(scrollregion=(0, 0, 400, 200))
            return

        drives = list(payload.get("drives") or [])
        drive_apps: Dict[str, List[Dict[str, object]]] = {drive: [] for drive in drives}
        drive_related: Dict[str, List[Tuple[Dict[str, object], Dict[str, object]]]] = {drive: [] for drive in drives}
        for app in apps:
            drive = str(app.get("drive", "Unknown"))
            drive_apps.setdefault(drive, []).append(app)
            for related in app.get("related", []) or []:
                rel_drive = str(related.get("drive", "Unknown"))
                drive_related.setdefault(rel_drive, []).append((app, related))
        drives = sorted(drive_apps.keys(), key=lambda d: d.casefold())
        for drive in drives:
            drive_apps[drive].sort(key=lambda a: str(a.get("name", "")).casefold())
            drive_related[drive].sort(key=lambda pair: str(pair[1].get("label", "")).casefold())

        # --- Node sizing controls (manual tuning) ---
        # Adjust these to change node geometry and spacing:
        # - drive_w/drive_h: size of DRIVE diamonds
        # - app_w/app_h: size of APP rounded rectangles
        # - rel_w/rel_h: size of file/folder nodes (left as plain rectangles)
        # - app_corner_radius: corner roundness for APP nodes
        margin_x = 20
        margin_y = 20
        lane_gap = 24
        drive_w, drive_h = 170, 44
        app_w, app_h = 180, 38
        rel_w, rel_h = 160, 28
        app_corner_radius = 50
        app_gap = 12
        rel_gap = 10
        lane_width = max(app_w, rel_w * 2 + rel_gap, drive_w) + lane_gap
        related_cols = 2 if lane_width >= rel_w * 2 + rel_gap + 20 else 1

        max_apps = max((len(drive_apps.get(drive, [])) for drive in drives), default=0)
        drive_y = margin_y
        apps_start_y = drive_y + drive_h + 30
        related_start_y = apps_start_y + max_apps * (app_h + app_gap) + 24

        drive_positions: Dict[str, Tuple[int, int, int, int]] = {}
        app_positions: Dict[str, Tuple[int, int, int, int]] = {}
        related_positions: Dict[str, Tuple[int, int, int, int]] = {}
        app_tags: Dict[str, str] = {}
        drive_tags: Dict[str, str] = {}

        for idx, drive in enumerate(drives):
            lane_x = margin_x + idx * lane_width
            cx = lane_x + lane_width // 2
            x1 = cx - drive_w // 2
            x2 = cx + drive_w // 2
            y1 = drive_y
            y2 = drive_y + drive_h
            drive_positions[drive] = (x1, y1, x2, y2)
            drive_tags[drive] = f"drive:{self._safe_tag(drive)}"

            apps_for_drive = drive_apps.get(drive, [])
            for a_idx, app in enumerate(apps_for_drive):
                app_id = str(app.get("id", ""))
                tag = f"app:{self._safe_tag(app_id)}"
                app_tags[app_id] = tag
                ay1 = apps_start_y + a_idx * (app_h + app_gap)
                ay2 = ay1 + app_h
                ax1 = cx - app_w // 2
                ax2 = cx + app_w // 2
                app_positions[app_id] = (ax1, ay1, ax2, ay2)

            related_for_drive = drive_related.get(drive, [])
            for r_idx, (app, related) in enumerate(related_for_drive):
                rel_id = str(related.get("id", ""))
                row = r_idx // related_cols
                col = r_idx % related_cols
                rx1 = lane_x + 10 + col * (rel_w + rel_gap)
                rx2 = rx1 + rel_w
                ry1 = related_start_y + row * (rel_h + rel_gap)
                ry2 = ry1 + rel_h
                related_positions[rel_id] = (rx1, ry1, rx2, ry2)

        max_related_rows = 0
        for drive in drives:
            count = len(drive_related.get(drive, []))
            rows = (count + related_cols - 1) // related_cols if related_cols else 0
            if rows > max_related_rows:
                max_related_rows = rows
        total_width = margin_x * 2 + (len(drives) - 1) * lane_width + lane_width
        total_height = related_start_y + max_related_rows * (rel_h + rel_gap) + 60
        self._drive_tag_map = {tag: drive for drive, tag in drive_tags.items()}
        self._draw_map_background(total_width, total_height, drives, lane_width, margin_x)

        edge_color = str(self._map_style.get("map_edge", "#9aa5b1"))
        drive_outline = str(self._map_style.get("map_drive_outline", "#334e68"))
        drive_fill = str(self._map_style.get("map_drive_bg", "#d9e2ec"))
        node_outline = str(self._map_style.get("map_node_outline", "#52606d"))
        text_color = str(self._map_style.get("map_text", "#1f2933"))
        unknown_group = str(self._map_style.get("map_unknown_group", "#cbd2d9"))

        # Draw edges first so nodes sit above the lines.
        for drive in drives:
            drive_tag = drive_tags.get(drive, "")
            for app in drive_apps.get(drive, []):
                app_id = str(app.get("id", ""))
                app_tag = app_tags.get(app_id, "")
                if app_id not in app_positions:
                    continue
                x1, y1, x2, y2 = drive_positions.get(drive, (0, 0, 0, 0))
                ax1, ay1, ax2, ay2 = app_positions[app_id]
                line_id = canvas.create_line(
                    (x1 + x2) / 2,
                    y2,
                    (ax1 + ax2) / 2,
                    ay1,
                    fill=edge_color,
                    width=1,
                    tags=("map", "edge", drive_tag, app_tag),
                )
                self._map_item_styles[line_id] = {"fill": edge_color, "width": 1}

        for app in apps:
            app_id = str(app.get("id", ""))
            app_tag = app_tags.get(app_id, "")
            if app_id not in app_positions:
                continue
            ax1, ay1, ax2, ay2 = app_positions[app_id]
            for related in app.get("related", []) or []:
                rel_id = str(related.get("id", ""))
                if rel_id not in related_positions:
                    continue
                rx1, ry1, rx2, ry2 = related_positions[rel_id]
                line_id = canvas.create_line(
                    (ax1 + ax2) / 2,
                    ay2,
                    (rx1 + rx2) / 2,
                    ry1,
                    fill=edge_color,
                    width=1,
                    tags=("map", "edge", app_tag),
                )
                self._map_item_styles[line_id] = {"fill": edge_color, "width": 1}

        # Draw drive nodes.
        for drive in drives:
            x1, y1, x2, y2 = drive_positions[drive]
            drive_tag = drive_tags.get(drive, "")
            rect_id = self._create_drive_diamond(
                x1,
                y1,
                x2,
                y2,
                fill=drive_fill,
                outline=drive_outline,
                width=1,
                tags=("map", "node", "node:drive", drive_tag),
            )
            self._map_item_styles[rect_id] = {"fill": drive_fill, "outline": drive_outline, "width": 1}
            label = f"{drive}\\" if drive.endswith(":") else drive
            text_id = canvas.create_text(
                (x1 + x2) / 2,
                (y1 + y2) / 2,
                text=label,
                fill=text_color,
                tags=("map", "node", "node:drive", drive_tag),
            )
            self._map_item_styles[text_id] = {"fill": text_color}

        # Draw app nodes.
        for app in apps:
            app_id = str(app.get("id", ""))
            if app_id not in app_positions:
                continue
            app_tag = app_tags.get(app_id, "")
            group = str(app.get("group", "") or "")
            group_color = self._map_group_colors.get(group, unknown_group)
            drive_tag = drive_tags.get(str(app.get("drive", "")), "")
            x1, y1, x2, y2 = app_positions[app_id]
            rect_id = self._create_rounded_rect(
                x1,
                y1,
                x2,
                y2,
                radius=app_corner_radius,
                fill=group_color,
                outline=node_outline,
                width=1,
                tags=("map", "node", "node:app", app_tag, drive_tag),
            )
            self._map_item_styles[rect_id] = {"fill": group_color, "outline": node_outline, "width": 1}
            text_id = canvas.create_text(
                (x1 + x2) / 2,
                (y1 + y2) / 2,
                text=str(app.get("name", "")),
                fill=text_color,
                tags=("map", "node", "node:app", app_tag, drive_tag),
            )
            self._map_item_styles[text_id] = {"fill": text_color}

        # Draw related nodes.
        for app in apps:
            app_id = str(app.get("id", ""))
            app_tag = app_tags.get(app_id, "")
            group = str(app.get("group", "") or "")
            group_color = self._map_group_colors.get(group, unknown_group)
            related_fill = self._tint_color(group_color, 0.35)
            for related in app.get("related", []) or []:
                rel_id = str(related.get("id", ""))
                if rel_id not in related_positions:
                    continue
                drive_tag = drive_tags.get(str(related.get("drive", "")), "")
                x1, y1, x2, y2 = related_positions[rel_id]
                rect_id = canvas.create_rectangle(
                    x1,
                    y1,
                    x2,
                    y2,
                    fill=related_fill,
                    outline=node_outline,
                    width=1,
                    tags=("map", "node", "node:file", app_tag, drive_tag),
                )
                self._map_item_styles[rect_id] = {"fill": related_fill, "outline": node_outline, "width": 1}
                text_id = canvas.create_text(
                    (x1 + x2) / 2,
                    (y1 + y2) / 2,
                    text=str(related.get("label", "")),
                    fill=text_color,
                    tags=("map", "node", "node:file", app_tag, drive_tag),
                )
                self._map_item_styles[text_id] = {"fill": text_color}

        canvas.configure(scrollregion=(0, 0, total_width, total_height))

    def _on_map_click(self, _event: tk.Event) -> None:
        tags = self.map_canvas.gettags("current")
        if not tags:
            self._clear_map_highlight()
            return
        selection = ""
        for tag in tags:
            if tag.startswith("app:"):
                selection = tag
                break
        if not selection:
            for tag in tags:
                if tag.startswith("drive:"):
                    selection = tag
                    break
        if not selection:
            self._clear_map_highlight()
            return
        self._highlight_map_items(selection)

    def _on_map_mousewheel(self, event: tk.Event) -> None:
        if not self.map_canvas.winfo_exists():
            return
        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = int(event.delta / 120)
        elif hasattr(event, "num"):
            if event.num == 4:
                delta = 1
            elif event.num == 5:
                delta = -1
        if delta == 0:
            return
        if getattr(event, "state", 0) & 0x0001:
            self.map_canvas.xview_scroll(-delta, "units")
        else:
            self.map_canvas.yview_scroll(-delta, "units")

    def _clear_map_highlight(self) -> None:
        if not self._map_item_styles:
            return
        for item_id, style in self._map_item_styles.items():
            try:
                self.map_canvas.itemconfigure(item_id, **style)
            except tk.TclError:
                continue
        self._map_highlight_tag = ""

    def _highlight_map_items(self, selection: str) -> None:
        self._clear_map_highlight()
        highlight = str(self._map_style.get("map_highlight", "#0b69ff"))
        for item_id in self.map_canvas.find_withtag(selection):
            kind = self.map_canvas.type(item_id)
            if kind == "line":
                self.map_canvas.itemconfigure(item_id, fill=highlight, width=2)
            elif kind in {"rectangle", "oval", "polygon"}:
                self.map_canvas.itemconfigure(item_id, outline=highlight, width=2)
            elif kind == "text":
                self.map_canvas.itemconfigure(item_id, fill=highlight)
        self._map_highlight_tag = selection

    def _on_map_right_click(self, event: tk.Event) -> None:
        tags = self.map_canvas.gettags("current")
        drive = ""
        for tag in tags:
            if tag.startswith("drive:"):
                drive = self._drive_tag_map.get(tag, "")
                break
        if not drive:
            return
        self._map_context_drive = drive
        self._build_drive_tint_menu(drive)
        try:
            self.map_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.map_context_menu.grab_release()

    def _build_drive_tint_menu(self, drive: str) -> None:
        self.map_context_menu.delete(0, tk.END)
        default_color = str(self._map_style.get("map_bg", "#f4f6f9"))
        self.map_context_menu.add_command(
            label="Default",
            image=self._get_tint_icon(default_color),
            compound="left",
            command=lambda: self._set_drive_tint(""),
        )
        for color in self._drive_tint_palette:
            self.map_context_menu.add_command(
                label="",
                image=self._get_tint_icon(color),
                compound="left",
                command=lambda c=color: self._set_drive_tint(c),
            )

    def _set_drive_tint(self, color: str) -> None:
        drive = self._map_context_drive
        if not drive:
            return
        if color:
            self._drive_band_colors[drive] = color
        else:
            self._drive_band_colors.pop(drive, None)
        self._draw_system_map()

    def _draw_map_background(self, width: int, height: int, drives: List[str], lane_width: int, margin_x: int) -> None:
        if width <= 0 or height <= 0:
            return
        bg = str(self._map_style.get("map_bg", "#f4f6f9"))
        noise_color = self._tint_color(bg, 0.05)
        noise_color_alt = self._tint_color(bg, 0.1)
        specks = max(200, min(2000, (width * height) // 8000))
        seed = width * 131 + height * 17 + len(drives) * 19
        x = seed % max(width, 1)
        y = (seed * 3) % max(height, 1)
        for i in range(int(specks)):
            x = (x * 1664525 + 1013904223) % max(width, 1)
            y = (y * 22695477 + 1) % max(height, 1)
            size = 2 if (i % 7 == 0) else 1
            color = noise_color if (i % 2 == 0) else noise_color_alt
            self.map_canvas.create_rectangle(x, y, x + size, y + size, outline="", fill=color, tags=("map", "noise"))

        for idx, drive in enumerate(drives):
            tint = self._drive_band_color(drive, idx)
            if not tint:
                continue
            x1 = margin_x + idx * lane_width
            x2 = x1 + lane_width - 8
            self.map_canvas.create_rectangle(
                x1,
                0,
                x2,
                height,
                fill=tint,
                outline="",
                tags=("map", "band"),
            )

    def _drive_band_color(self, drive: str, index: int) -> str:
        if drive in self._drive_band_colors:
            return self._drive_band_colors[drive]
        if not self._drive_tint_palette:
            return ""
        return self._drive_tint_palette[index % len(self._drive_tint_palette)]

    def _create_rounded_rect(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        radius: int,
        **kwargs,
    ) -> int:
        radius = max(0, min(radius, int((x2 - x1) / 2), int((y2 - y1) / 2)))
        points = [
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
        ]
        return self.map_canvas.create_polygon(points, smooth=True, splinesteps=12, **kwargs)

    def _create_drive_diamond(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        **kwargs,
    ) -> int:
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        points = [
            x1,
            cy,
            cx,
            y1,
            x2,
            cy,
            cx,
            y2,
        ]
        return self.map_canvas.create_polygon(points, **kwargs)

    def _get_tint_icon(self, color: str) -> tk.PhotoImage:
        swatch = self._swatch_color(color)
        key = f"{color.casefold()}::{swatch.casefold()}"
        icon = self._tint_icons.get(key)
        if icon:
            return icon
        size = 14
        icon = tk.PhotoImage(width=size, height=size)
        border = "#8a8a8a"
        icon.put(border, to=(0, 0, size, size))
        icon.put(swatch, to=(1, 1, size - 1, size - 1))
        self._tint_icons[key] = icon
        return icon

    def _swatch_color(self, color: str) -> str:
        text = color.strip()
        if not text.startswith("#"):
            return text
        base = text.lstrip("#")
        if len(base) != 6:
            return text
        try:
            r = int(base[0:2], 16)
            g = int(base[2:4], 16)
            b = int(base[4:6], 16)
        except ValueError:
            return text
        luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0
        if luminance > 0.88:
            return self._shade_color(text, 0.82)
        return text

    @staticmethod
    def _shade_color(value: str, factor: float) -> str:
        text = value.strip().lstrip("#")
        if len(text) != 6:
            return value
        try:
            r = int(text[0:2], 16)
            g = int(text[2:4], 16)
            b = int(text[4:6], 16)
        except ValueError:
            return value
        factor = max(0.0, min(1.0, factor))
        r = int(r * factor)
        g = int(g * factor)
        b = int(b * factor)
        return f"#{r:02x}{g:02x}{b:02x}"

    @staticmethod
    def _safe_tag(raw: str) -> str:
        return re.sub(r"[^a-zA-Z0-9_-]+", "_", raw or "")

    @staticmethod
    def _tint_color(value: str, amount: float) -> str:
        text = value.strip().lstrip("#")
        if len(text) != 6:
            return value
        try:
            r = int(text[0:2], 16)
            g = int(text[2:4], 16)
            b = int(text[4:6], 16)
        except ValueError:
            return value
        amount = max(0.0, min(1.0, amount))
        r = int(r + (255 - r) * amount)
        g = int(g + (255 - g) * amount)
        b = int(b + (255 - b) * amount)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _on_tree_click(self, event: tk.Event) -> None:
        self.cancel_group_editor()
        region = self.system_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.system_tree.identify_column(event.x)
        row_id = self.system_tree.identify_row(event.y)
        if not row_id:
            return
        col_index = int(column.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.column_keys):
            return
        if self.column_keys[col_index] != "website":
            return
        value = self.system_tree.set(row_id, "website")
        self._dispatch("on_website_click")(value)

    def _on_tree_motion(self, event: tk.Event) -> None:
        region = self.system_tree.identify_region(event.x, event.y)
        if region != "cell":
            self.system_tree.configure(cursor="")
            return
        column = self.system_tree.identify_column(event.x)
        row_id = self.system_tree.identify_row(event.y)
        if not row_id:
            self.system_tree.configure(cursor="")
            return
        col_index = int(column.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.column_keys):
            self.system_tree.configure(cursor="")
            return
        if self.column_keys[col_index] != "website":
            self.system_tree.configure(cursor="")
            return
        value = self.system_tree.set(row_id, "website")
        if value and value != "NOT FOUND":
            self.system_tree.configure(cursor="hand2")
        else:
            self.system_tree.configure(cursor="")

    def _on_tree_double_click(self, event: tk.Event) -> None:
        region = self.system_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.system_tree.identify_column(event.x)
        row_id = self.system_tree.identify_row(event.y)
        if not row_id:
            return
        col_index = int(column.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.column_keys):
            return
        if self.column_keys[col_index] != "group":
            return
        self._dispatch("on_group_double_click")(row_id)

    def _on_tree_select(self, _event: tk.Event) -> None:
        selection = self.system_tree.selection()
        if not selection:
            return
        self._dispatch("on_row_select")(selection[0])

    def _on_related_double_click(self, event: tk.Event) -> None:
        region = self.related_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        row_id = self.related_tree.identify_row(event.y)
        if not row_id:
            return
        if self.related_tree.parent(row_id) == "":
            return
        self._dispatch("on_related_double_click")(row_id)
