**ASSETS folder**

Place your .ico files here.

Suggested files:
- assets/app.ico      (window/taskbar icon used by Tk)
- assets/icons/app.ico (reserve for packaging/installer icon)

Update file names/paths in main_controller.py if you change them.



ARC : The Windows Ecosystem Cartographer - Project Summary
--------------------------------------

Purpose:
Ecosystem Cartographer is a lightweight Windows utility designed to scan a computer, catalog all installed applications, their associated plugins, drivers, and saved user profiles, and produce a clear map of the system’s software ecosystem. The goal is to give users a complete, interactive reference of their current setup to simplify rebuilding it after a fresh Windows installation—without cloning drives or reinstalling blindly.

Core Functionality:
1. System Scan
   - Enumerates installed applications, versions, install paths, and installation sources using Windows registry keys, WMI, and Winget.
   - Records additional metadata such as publisher, install date, and uninstall string.

2. Mapping Relationships
   - Identifies plugins and extensions related to parent applications by checking known folders and recipe-based detection rules.
   - Maps device drivers to their associated applications when possible, using driver provider and manufacturer information.

3. Profile and Configuration Capture
   - Detects and optionally copies user-created configuration or profile files (e.g., OBS scenes, Stream Deck profiles, VS Code settings).
   - Captures only user data (JSON, INI, or DB files), never passwords or licenses.
   - Exports chosen profiles into a structured payload folder for restoration later.

4. Manifest and Guide Generation
   - Generates a single manifest.json file summarizing the system map: applications, plugins, drivers, and profiles.
   - Creates a “Rebuild Guide” (Markdown or HTML) that lists each application with download links, reinstall notes, and related profile files.
   - Optionally zips selected profiles and configuration files for backup.

5. Lightweight UI
   - Simple PySide6-based interface showing a tree view of applications and related items.
   - Allows users to check which profile/config files to include in export.
   - No background services or system hooks—on-demand scanning only.

Design Philosophy:
- **Read-only by default:** never modifies or removes anything on the system.
- **Privacy-first:** no passwords, license keys, or credentials are collected.
- **Human-readable output:** plain JSON and Markdown files, easy to inspect or version control.
- **Extensible:** new app detection rules (“recipes”) can be added as YAML files.
- **Local-first:** no internet connection required beyond optional link lookups.

Intended Use:
After a clean Windows install, users open their saved manifest and Rebuild Guide to quickly reinstall essential software, restore saved configurations, and reapply profile data. The tool acts as a personal ecosystem map and rebuild checklist—not a backup or cloning utility.


