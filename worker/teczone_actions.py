import copy
import csv
import json
import os
import re
import subprocess
import time
from pathlib import Path

from pywinauto import Desktop
from pywinauto.application import Application

from ui_utils import (
    click_menu_item_anywhere,
    describe_controls,
    debug_open_dialog_search,
    dump_window_titles,
    find_control,
    find_unexpected_dialog,
    handle_possible_dialogs,
    open_dialog_present,
    press_open,
    press_save,
    save_dialog_present,
    set_file_name,
    wait_for_open_dialog,
    wait_for_save_dialog,
)


DEFAULT_WORKFLOW = {
    "enterBendHotkey": "b",
    "useEnterBend": True,
    "exportMethod": "menu",
    "menuExportPath": ["File", "Export", "2D Geometry"],
    "timeouts": {
        "enterBendSeconds": 120,
        "exportMenuSeconds": 10,
        "saveAsSeconds": 20,
        "exportCompleteSeconds": 30,
    },
}


class NeedsHelpError(RuntimeError):
    pass


def deep_merge(base, override):
    out = copy.deepcopy(base)
    if not isinstance(override, dict):
        return out
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = deep_merge(out[key], value)
        else:
            out[key] = value
    return out


class TecZoneSession:
    def __init__(
        self,
        logger,
        screenshotter=None,
        teczone_exe=None,
        teczone_title_re=None,
        workflow_config_path=None,
    ):
        self.logger = logger
        self.screenshotter = screenshotter
        self.app = None
        self.main = None
        self.teczone_exe = teczone_exe
        self.teczone_title_re = teczone_title_re
        self.workflow_config_path = workflow_config_path

        self.main_title_re = os.getenv("TECZONE_MAIN_TITLE_RE", r".*TecZone.*Bend.*")
        self.material_menu_titles = ["Material", "Material..."]
        self.material_dialog_title_re = os.getenv("TECZONE_MATERIAL_DIALOG_RE", r".*Material.*")
        self.workflow = self._load_workflow()

    def _load_workflow(self):
        default_path = Path(__file__).resolve().parent / "teczone_workflow.json"
        path = Path(self.workflow_config_path) if self.workflow_config_path else default_path
        if not path.exists():
            self.logger.warning("Workflow config not found at %s; using defaults", path)
            return copy.deepcopy(DEFAULT_WORKFLOW)
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = deep_merge(DEFAULT_WORKFLOW, data)
            self.logger.info("Loaded workflow config: %s", path)
            return merged
        except Exception as e:
            self.logger.warning("Failed to load workflow config %s: %s. Using defaults.", path, e)
            return copy.deepcopy(DEFAULT_WORKFLOW)

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
        flux_pids = self._flux_process_ids()
        for w in Desktop(backend="uia").windows():
            try:
                pid = int(w.element_info.process_id)
            except Exception:
                pid = None
            if flux_pids and pid not in flux_pids:
                continue
            try:
                title = w.window_text() or ""
            except Exception:
                title = ""
            if re.search(pattern, title, flags=re.IGNORECASE):
                return w
        return None

    def _flux_process_ids(self):
        try:
            proc = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq Flux.exe", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
            )
            rows = (proc.stdout or "").strip().splitlines()
            pids = set()
            for row in rows:
                if not row or "No tasks are running" in row:
                    continue
                parsed = next(csv.reader([row]))
                if len(parsed) < 2:
                    continue
                if parsed[0].strip().lower() != "flux.exe":
                    continue
                try:
                    pids.add(int(parsed[1].strip()))
                except Exception:
                    continue
            return pids
        except Exception:
            return set()

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
        pids = self._flux_process_ids()
        return len(pids) > 0

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
                    Path(base) / "TRUMPF" / "Flux" / "Bin" / "Flux.exe",
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
                    if "teczone" not in root_l and "flux" not in root_l and "trumpf" not in root_l:
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
            if hasattr(win, "wrapper_object"):
                self.main = win.wrapper_object()
            else:
                self.main = win
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
            raise NeedsHelpError("Material from Xometry missing; material is required")

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

        deadline = time.time() + 15
        dialog = None
        while time.time() < deadline:
            try:
                dialog = Desktop(backend="uia").window(title_re=self.material_dialog_title_re)
                if dialog.exists(timeout=0.2):
                    break
            except Exception:
                pass
            time.sleep(0.2)
        if not dialog or not dialog.exists(timeout=0.2):
            raise NeedsHelpError("Material selection dialog not found")

        items = dialog.descendants(control_type="ListItem")
        if not items:
            raise NeedsHelpError("Material list items not found")

        used_material = None
        note = ""
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
            dialog.type_keys("{ENTER}")

        unexpected = find_unexpected_dialog()
        if unexpected:
            raise NeedsHelpError(f"Unexpected dialog after setting material: {unexpected}")

        return used_material, note

    def _hotkey_to_type_keys(self, hotkey):
        token_map = {"ctrl": "^", "control": "^", "alt": "%", "shift": "+"}
        parts = [p.strip().lower() for p in str(hotkey).split("+") if p.strip()]
        if not parts:
            return "b"
        mods = ""
        key = parts[-1]
        for p in parts[:-1]:
            mods += token_map.get(p, "")
        return f"{mods}{key}"

    def _enter_bend_mode_if_needed(self):
        if not self.workflow.get("useEnterBend", True):
            return
        hotkey = self.workflow.get("enterBendHotkey", "b")
        key_seq = self._hotkey_to_type_keys(hotkey)
        self.main.set_focus()
        self.main.type_keys(key_seq, set_foreground=True)
        if self.screenshotter:
            self.screenshotter.snap("enter_bend")

        deadline = time.time() + float(self.workflow.get("timeouts", {}).get("enterBendSeconds", 120))
        while time.time() < deadline:
            unexpected = find_unexpected_dialog()
            if unexpected:
                raise NeedsHelpError(f"Unexpected dialog while entering bend mode: {unexpected}")
            # No reliable public indicator for bend state in current build; continue after short settle.
            time.sleep(0.5)
            return

    def _open_export_menu(self):
        menu_path_items = self.workflow.get("menuExportPath", ["File", "Export", "2D Geometry"])
        if not isinstance(menu_path_items, list) or len(menu_path_items) < 3:
            raise NeedsHelpError(f"Invalid workflow menuExportPath: {menu_path_items}")

        joined = "->".join(menu_path_items)
        try:
            self.main.menu_select(joined)
            return
        except Exception:
            pass

        # UIA fallback without random coordinates.
        first = menu_path_items[0]
        first_ctrl = find_control(self.main, title_re=rf"(?i){re.escape(first)}", control_type="MenuItem")
        if first_ctrl:
            first_ctrl.click_input()
        else:
            self.main.type_keys("%f", set_foreground=True)

        timeout = float(self.workflow.get("timeouts", {}).get("exportMenuSeconds", 10))
        if not click_menu_item_anywhere(menu_path_items[1], timeout=timeout):
            raise NeedsHelpError(
                f"Export menu item not found via fallback; menu path={menu_path_items}"
            )
        if not click_menu_item_anywhere(menu_path_items[2], timeout=timeout):
            raise NeedsHelpError(
                f"Export target item not found via fallback; menu path={menu_path_items}"
            )

    def export_geo(self, export_path):
        self.main.set_focus()
        self._enter_bend_mode_if_needed()

        try:
            self._open_export_menu()
            if self.screenshotter:
                self.screenshotter.snap("export_menu_open")
        except Exception as e:
            if self.screenshotter:
                self.screenshotter.snap("export_failed")
            if isinstance(e, NeedsHelpError):
                raise
            raise NeedsHelpError(f"Export menu open failed: {e}")

        save_timeout = float(self.workflow.get("timeouts", {}).get("saveAsSeconds", 20))
        try:
            dialog = wait_for_save_dialog(timeout=save_timeout, parent=self.main)
            if self.screenshotter:
                self.screenshotter.snap("saveas_dialog")
        except Exception as e:
            if self.screenshotter:
                self.screenshotter.snap("export_failed")
            raise NeedsHelpError(f"Save As dialog not found: {e}")

        try:
            Path(export_path).parent.mkdir(parents=True, exist_ok=True)
            set_file_name(dialog, export_path)
            press_save(dialog)
        except Exception as e:
            controls = describe_controls(dialog, limit=35)
            if self.screenshotter:
                self.screenshotter.snap("export_failed")
            raise NeedsHelpError(
                "Save As interaction failed; "
                "searched=file_name_edit(auto_id=1148/label File name) + save_button(auto_id=1/title Save); "
                f"found_controls={controls}; error={e}"
            )

        deadline = time.time() + float(self.workflow.get("timeouts", {}).get("exportCompleteSeconds", 30))
        while time.time() < deadline:
            unexpected = find_unexpected_dialog()
            if unexpected:
                if self.screenshotter:
                    self.screenshotter.snap("export_failed")
                raise NeedsHelpError(f"Unexpected dialog during export: {unexpected}")

            if not save_dialog_present(parent=self.main):
                if Path(export_path).exists() and Path(export_path).stat().st_size > 0:
                    if self.screenshotter:
                        self.screenshotter.snap("export_done")
                    return

            if Path(export_path).exists() and Path(export_path).stat().st_size > 0:
                if self.screenshotter:
                    self.screenshotter.snap("export_done")
                return
            time.sleep(0.4)

        if self.screenshotter:
            self.screenshotter.snap("export_failed")
        raise NeedsHelpError(
            f"Export GEO failed: file missing or empty after timeout; export_path={export_path}"
        )

    def get_thickness_mm(self):
        return None
