ARC App Scanner (PoC) — README

Purpose
-------
This app scans a Windows machine for installed software and presents the results
in a sortable, filterable table. It can import a previous scan (CSV/JSON) to
serve as a reference list, then compares the current system against that
reference to show what is installed vs missing. It also lets you create custom
"Groups" for software, persist those assignments, and export or save scans for
reuse.

How It Achieves the Goal
------------------------
1) Registry Scan (Current System)
   - Reads uninstall entries from Windows registry (HKLM/HKCU, 32/64-bit views).
   - Normalizes install dates, sizes, and basic fields (name, version, etc.).

2) Reference Scan
   - Imports CSV files exported by the app.
   - Opens JSON scan files saved by the app.
   - Uses the reference list to compare with the current scan.

3) Installed / Not Installed
   - Builds a set of installed app names from the current scan.
   - Marks each reference item as installed (green check) or missing (red X).
   - Missing rows are shown in a customizable "Not installed" color.

4) Groups and Customization
   - Users can create/rename/delete groups and assign them per app.
   - GUI settings (fonts/colors) are configurable and stored in state.

Run It
------
- Windows only.
- Start with: python arc_poc.py

Packaging (PyInstaller)
-----------------------
- Install: pip install pyinstaller
- Build (Git Bash): bash scripts/package.sh
- Output: dist/ARC App Scanner.exe
- Icons:
  - Window/taskbar icon at runtime: assets/app.ico
    - Update path in main_controller.py (WINDOW_ICON_PATH) if you rename/move it.
  - Packaged exe icon: assets/icons/app.ico
    - Update path in scripts/package.sh (EXE_ICON) if you rename/move it.

Codebase Organization (MVC-ish)
-------------------------------
Entry Point
- arc_poc.py
  Thin entrypoint that launches the main controller.

Models
- models.py
  Data structures for AppEntry and AppGroup.

Services / Helpers
- scanner.py
  Registry scanning and conversion into AppEntry objects.
- import_export.py
  CSV/JSON import/export logic and parsing.
- compare.py
  Installed/missing comparison helpers.
- store.py
  Persistence for window geometry, GUI settings, and group assignments.
- utils.py
  Shared utilities (date normalization, size parsing, URL normalization, etc.).

Views
- main_view.py
  Tkinter layout for the main window: toolbar, table, status bar, and handlers
  for user interaction (delegated to controller).
- settings_view.py
  Tkinter layout for the settings window: GUI customization + Groups tab.

Controller
- main_controller.py
  Orchestrates state, user actions, and view updates. This is where most
  application logic lives (scan, import, compare, sort, filter, group changes).

How the Pieces Work Together
-----------------------------
- main_controller.py creates the views and wires callbacks.
- scanner.py provides the current system scan.
- import_export.py loads/saves reference lists and scan data.
- compare.py calculates installed status for each reference entry.
- store.py persists settings and group assignments between runs.
- views update the UI based on controller instructions (no direct business logic).

Notes for Collaborators
-----------------------
- Keep business logic in services or the controller (not in view classes).
- Maintain AppEntry as the central data shape for installed and reference apps.
- If adding new fields, update models.py, import_export.py, and view columns.
- When changing GUI settings, update default settings and store.py validation.
- Prefer small, testable helpers in services over adding more logic to views.

**ASSETS folder**

Place your .ico files here.

Suggested files:
- assets/app.ico      (window/taskbar icon used by Tk)
- assets/icons/app.ico (reserve for packaging/installer icon)

Update file names/paths in main_controller.py if you change them.