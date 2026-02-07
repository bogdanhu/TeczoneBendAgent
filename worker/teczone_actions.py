import os
import re
import subprocess
import time
from pathlib import Path

from pywinauto import Desktop
from pywinauto.application import Application

from ui_utils import (
    describe_controls,
    debug_open_dialog_search,
    dump_window_titles,
    find_control,
    find_unexpected_dialog,
    handle_possible_dialogs,
    open_dialog_present,
    press_open,
    set_file_name,
    wait_for_open_dialog,
    wait_for_window,
)


class NeedsHelpError(RuntimeError):
    pass


class TecZoneSession:
    def __init__(self, logger, screenshotter=None, teczone_exe=None, teczone_title_re=None):
        self.logger = logger
        self.screenshotter = screenshotter
        self.app = None
        self.main = None
        self.teczone_exe = teczone_exe
        self.teczone_title_re = teczone_title_re

        self.main_title_re = os.getenv("TECZONE_MAIN_TITLE_RE", r".*TecZone.*Bend.*")
        self.material_menu_titles = ["Material", "Material..."]
        self.material_dialog_title_re = os.getenv("TECZONE_MATERIAL_DIALOG_RE", r".*Material.*")
        self.export_menu_paths = [
            "File->Export",
            "File->Export...",
            "File->Save As",
            "File->Save As...",
        ]

    def _title_patterns(self):
        patterns = []
        if self.teczone_title_re:
            patterns.append(self.teczone_title_re)
        if self.main_title_re:
            patterns.append(self.main_title_re)
        patterns.extend([r".*Flux.*", r".*TecZone.*"])

        valid = []
        seen = set()
        for p in patterns:
            if p in seen:
                continue
            seen.add(p)
            try:
                re.compile(p)
                valid.append(p)
            except re.error as e:
                self.logger.warning("Ignoring invalid title regex '%s': %s", p, e)
        return valid

    def _find_main_window(self, pattern):
        for w in Desktop(backend="uia").windows():
            try:
                title = w.window_text() or ""
            except Exception:
                title = ""
            if re.search(pattern, title, flags=re.IGNORECASE):
                return w
        return None

    def _wait_for_main_window(self, timeout=60):
        patterns = self._title_patterns()
        deadline = time.time() + timeout
        while time.time() < deadline:
            for pattern in patterns:
                win = self._find_main_window(pattern)
                if win:
                    return win
            time.sleep(0.4)
        return None

    def _is_flux_running(self):
        try:
            proc = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq Flux.exe", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
            )
            return "flux.exe" in (proc.stdout or "").lower()
        except Exception as e:
            self.logger.warning("Failed process check for Flux.exe: %s", e)
            return False

    def _registry_flux_path(self):
        try:
            import winreg
        except Exception:
            return None

        keys = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Flux.exe"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Flux.exe"),
        ]
        for root, key_path in keys:
            try:
                with winreg.OpenKey(root, key_path) as k:
                    value, _ = winreg.QueryValueEx(k, None)
                    if value and Path(value).exists():
                        return value
            except Exception:
                continue
        return None

    def _program_files_flux_path(self):
        bases = [os.getenv("ProgramFiles", r"C:\Program Files"), os.getenv("ProgramFiles(x86)", r"C:\Program Files (x86)")]

        direct = []
        for base in bases:
            if not base:
                continue
            direct.extend(
                [
                    Path(base) / "TecZone Bend" / "Flux.exe",
                    Path(base) / "TecZone" / "Flux.exe",
                    Path(base) / "Flux" / "Flux.exe",
                ]
            )

        for p in direct:
            if p.exists():
                return str(p)

        for base in bases:
            if not base or not Path(base).exists():
                continue
            try:
                for root, _, files in os.walk(base):
                    root_l = root.lower()
                    if "teczone" not in root_l and "flux" not in root_l:
                        continue
                    for file_name in files:
                        if file_name.lower() == "flux.exe":
                            return str(Path(root) / file_name)
            except Exception:
                continue

        return None

    def _resolve_launch_path(self):
        candidates = [
            self.teczone_exe,
            os.getenv("TECZONE_EXE"),
            self._registry_flux_path(),
            self._program_files_flux_path(),
        ]
        for c in candidates:
            if c and Path(c).exists():
                return str(Path(c))
        return None

    def connect(self, timeout=60):
        flux_running = self._is_flux_running()
        launch_path = None

        if not flux_running:
            launch_path = self._resolve_launch_path()
            if not launch_path:
                raise NeedsHelpError(
                    "Flux.exe not running; could not resolve launch path. "
                    "Provide --teczone-exe \"C:\\Path\\Flux.exe\"."
                )
            try:
                self.logger.info("Flux.exe not running; launching TecZone at: %s", launch_path)
                self.app = Application(backend="uia").start(f"\"{launch_path}\"")
            except Exception as e:
                raise NeedsHelpError(
                    f"Flux.exe not running; attempted to launch at {launch_path}; failed because {e}"
                )

        win = self._wait_for_main_window(timeout=timeout)
        if not win:
            patterns = self._title_patterns()
            candidates = dump_window_titles()[:30]
            if flux_running:
                raise NeedsHelpError(
                    "Flux.exe running but main window not found; "
                    f"title_patterns={patterns}; candidate windows: {candidates}"
                )
            raise NeedsHelpError(
                "Flux.exe not running; attempted to launch at "
                f"{launch_path}; main window not found within {timeout}s; "
                f"title_patterns={patterns}; candidate windows: {candidates}"
            )

        try:
            self.main = win.wrapper_object()
            self.logger.info("Connected to TecZone window: %s", self.main.window_text())
        except Exception as e:
            raise NeedsHelpError(f"TecZone main window wrapper failed: {e}")

    def open_file(self, path):
        if not Path(path).exists():
            raise NeedsHelpError(f"Input file not found: {path}")
        if Path(path).suffix.lower() not in [".stp", ".step"]:
            raise NeedsHelpError(f"Unsupported input extension: {path}")

        self.main.set_focus()
        self.main.type_keys("^o")

        try:
            dialog = wait_for_open_dialog(timeout=5, parent=self.main)
        except Exception as e:
            # Some TecZone builds ignore Ctrl+O; retry via File->Open.
            opened = False
            for menu_path in ["File->Open", "File->Open..."]:
                try:
                    self.main.menu_select(menu_path)
                    opened = True
                    break
                except Exception:
                    continue
            if not opened:
                dbg = debug_open_dialog_search(self.main)
                raise NeedsHelpError(
                    f"Open dialog not found and menu fallback failed: {e}; "
                    f"searched={dbg['searched']}; found_windows={dbg['found_windows']}"
                )
            try:
                dialog = wait_for_open_dialog(timeout=5, parent=self.main)
            except Exception as e2:
                dbg = debug_open_dialog_search(self.main)
                raise NeedsHelpError(
                    f"Open dialog not found: {e2}; "
                    f"searched={dbg['searched']}; found_windows={dbg['found_windows']}"
                )

        try:
            set_file_name(dialog, path)
            press_open(dialog)
        except Exception as e:
            controls = describe_controls(dialog, limit=35)
            raise NeedsHelpError(
                "Open dialog interaction failed; "
                "searched=file_name_edit(auto_id=1148/label File name) + open_button(auto_id=1/title Open); "
                f"found_controls={controls}; error={e}"
            )

        deadline = time.time() + 90
        while time.time() < deadline:
            unexpected = find_unexpected_dialog()
            if unexpected:
                raise NeedsHelpError(f"Unexpected dialog while opening file: {unexpected}")
            if not open_dialog_present(parent=self.main):
                return
            time.sleep(0.3)

        dbg = debug_open_dialog_search(self.main)
        raise NeedsHelpError(
            "Open dialog did not close within 90 seconds; "
            f"searched={dbg['searched']}; found_windows={dbg['found_windows']}"
        )

    def set_material(self, material):
        if not material:
            return None, "material not provided; skipped"

        self.main.set_focus()
        menu_clicked = False
        for title in self.material_menu_titles:
            ctrl = find_control(self.main, title=title, control_type="MenuItem")
            if ctrl:
                ctrl.click_input()
                menu_clicked = True
                break
        if not menu_clicked:
            raise NeedsHelpError("Material menu item not found")

        dialog = wait_for_window(self.material_dialog_title_re, timeout=10)
        list_ctrl = find_control(dialog, control_type="List")
        if not list_ctrl:
            list_ctrl = find_control(dialog, control_type="ListItem")
        if not list_ctrl:
            raise NeedsHelpError("Material list not found")

        used_material = None
        note = ""
        items = dialog.descendants(control_type="ListItem")
        if not items:
            raise NeedsHelpError("Material list items not found")

        material_lower = material.lower()
        for item in items:
            if material_lower in item.window_text().lower():
                item.click_input()
                used_material = item.window_text()
                break
        if not used_material:
            items[0].click_input()
            used_material = items[0].window_text()
            note = f"material fallback to: {used_material}; requested: {material}"

        ok_btn = find_control(dialog, title_re=r"OK|Ok|O(K|k)", control_type="Button")
        if ok_btn:
            ok_btn.click_input()
        else:
            dialog.type_keys("%{F4}")

        handle_possible_dialogs()
        return used_material, note

    def export_geo(self, export_path):
        self.main.set_focus()
        opened = False
        for path in self.export_menu_paths:
            try:
                self.main.menu_select(path)
                opened = True
                break
            except Exception:
                continue
        if not opened:
            raise NeedsHelpError("Export menu path not found")

        dialog = wait_for_window(r".*Save.*|.*Export.*", timeout=10)
        edit = find_control(dialog, title_re=r"File name.*", control_type="Edit")
        if not edit:
            edit = find_control(dialog, control_type="Edit")
        if not edit:
            raise NeedsHelpError("Export dialog file name edit not found")

        edit.set_edit_text(export_path)
        save_btn = find_control(dialog, title_re=r"Save|Export|OK", control_type="Button")
        if not save_btn:
            raise NeedsHelpError("Export dialog Save button not found")
        save_btn.click_input()

        handle_possible_dialogs()
        time.sleep(1.0)

        if not Path(export_path).exists():
            self.logger.warning("Export path not found after save: %s", export_path)

    def get_thickness_mm(self):
        return None
