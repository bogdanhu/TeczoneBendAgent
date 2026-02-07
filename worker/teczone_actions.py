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
from pywinauto.keyboard import send_keys

from ui_utils import (
    click_menu_item_anywhere,
    confirm_overwrite_dialog,
    describe_controls,
    debug_open_dialog_search,
    dump_window_titles,
    ensure_dialog_focus,
    find_control,
    find_overwrite_dialog,
    find_unexpected_dialog,
    handle_possible_dialogs,
    normalize_windows_path,
    press_open,
    press_save,
    save_dialog_present,
    set_file_name,
    wait_window_closed,
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
    "materialRequired": False,
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
        self.material_menu_titles = ["Material", "Material...", "Stock"]
        self.material_dialog_title_re = os.getenv("TECZONE_MATERIAL_DIALOG_RE", r".*Material.*")
        self.workflow = self._load_workflow()

    def _main_process_id(self):
        try:
            return int(self.main.element_info.process_id)
        except Exception:
            return None

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
        path = normalize_windows_path(path)
        if not Path(path).exists():
            raise NeedsHelpError(f"Input file not found: {path}")
        if Path(path).suffix.lower() not in [".stp", ".step"]:
            raise NeedsHelpError(f"Unsupported input extension: {path}")

        open_start = time.perf_counter()
        self.main.set_focus()
        main_pid = self._main_process_id()
        try:
            send_keys("^o")
        except Exception:
            self.main.type_keys("^o", set_foreground=True)

        try:
            # Fast path: Open dialog usually appears quickly; poll at high frequency.
            dialog = wait_for_open_dialog(
                timeout=1.4,
                parent=self.main,
                process_id=main_pid,
                poll_interval=0.06,
            )
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
                dbg = debug_open_dialog_search(self.main, process_id=main_pid)
                raise NeedsHelpError(
                    f"Open dialog not found and menu fallback failed: {e}; "
                    f"searched={dbg['searched']}; found_windows={dbg['found_windows']}"
                )
            try:
                dialog = wait_for_open_dialog(
                    timeout=3.0,
                    parent=self.main,
                    process_id=main_pid,
                    poll_interval=0.06,
                )
            except Exception as e2:
                dbg = debug_open_dialog_search(self.main, process_id=main_pid)
                raise NeedsHelpError(
                    f"Open dialog not found: {e2}; "
                    f"searched={dbg['searched']}; found_windows={dbg['found_windows']}"
                )

        try:
            detect_elapsed = time.perf_counter() - open_start
            strategy = self._fill_open_dialog_file(dialog, path)
            fill_elapsed = time.perf_counter() - open_start - detect_elapsed
            press_open(dialog)
            press_elapsed = time.perf_counter() - open_start
            self.logger.info(
                "OPEN timings: detect=%.3fs, fill=%.3fs, to_press_open=%.3fs, strategy=%s",
                detect_elapsed,
                max(fill_elapsed, 0.0),
                press_elapsed,
                strategy,
            )
        except Exception as e:
            controls = describe_controls(dialog, limit=35)
            raise NeedsHelpError(
                "Open dialog interaction failed; "
                "searched=file_name_edit(auto_id=1148/label File name) + open_button(auto_id=1/title Open); "
                f"found_controls={controls}; error={e}"
            )

        timeout_s = 90
        deadline = time.time() + timeout_s
        while time.time() < deadline:
            unexpected = find_unexpected_dialog(process_id=main_pid)
            if unexpected:
                raise NeedsHelpError(f"Unexpected dialog while opening file: {unexpected}")
            if wait_window_closed(dialog, timeout=0.2):
                return
            time.sleep(0.3)

        dbg = debug_open_dialog_search(self.main, process_id=main_pid)
        raise NeedsHelpError(
            "Open dialog did not close within 90 seconds; "
            f"searched={dbg['searched']}; found_windows={dbg['found_windows']}"
        )

    def _fill_open_dialog_file(self, dialog, full_path):
        full_path = normalize_windows_path(full_path)
        file_name = Path(full_path).name
        directory = normalize_windows_path(str(Path(full_path).parent))

        # Strategy A: full path directly in File name.
        try:
            edit = set_file_name(dialog, full_path)
            typed = ""
            try:
                typed = (edit.window_text() or "").strip()
            except Exception:
                typed = ""
            if typed and (":" in typed or file_name.lower() in typed.lower()):
                return "full_path"
        except Exception:
            pass

        # Strategy B (requested): F4 -> directory -> Enter -> file name -> Enter/Open.
        if not ensure_dialog_focus(dialog, timeout=1.5):
            try:
                dialog.click_input()
            except Exception:
                pass
        send_keys("{F4}")
        time.sleep(0.08)
        send_keys("^a{BACKSPACE}")
        send_keys(directory, with_spaces=True)
        send_keys("{ENTER}")
        time.sleep(0.2)
        set_file_name(dialog, file_name)
        return "f4_directory_then_filename"

    def set_material(self, material):
        if not material:
            raise NeedsHelpError("Material from Xometry missing; material is required")

        self.main.set_focus()
        material_required = bool(self.workflow.get("materialRequired", False))
        main_pid = self._main_process_id()

        def maybe_fail_or_skip(message):
            if material_required:
                raise NeedsHelpError(message)
            self.logger.warning("Material skipped: %s", message)
            return None, f"material not set: {message}"

        # First attempt: direct menu item if currently visible.
        for title in self.material_menu_titles:
            ctrl = find_control(self.main, title=title, control_type="MenuItem")
            if ctrl:
                try:
                    ctrl.click_input()
                    break
                except Exception:
                    continue
        else:
            # Fallback: open top menus and search material-like entries for this process only.
            candidates = []
            for hotkey in ["%f", "%e", "%v", "%h"]:
                try:
                    self.main.set_focus()
                    self.main.type_keys(hotkey, set_foreground=True)
                    time.sleep(0.35)
                except Exception:
                    continue

                for w in Desktop(backend="uia").windows():
                    try:
                        if main_pid is not None and int(w.element_info.process_id) != int(main_pid):
                            continue
                    except Exception:
                        continue
                    try:
                        items = w.descendants(control_type="MenuItem")
                    except Exception:
                        continue
                    for item in items:
                        try:
                            txt = (item.window_text() or "").strip()
                        except Exception:
                            txt = ""
                        if not txt:
                            continue
                        if re.search(r"(?i)material|stock|werkstoff|materiau|inox|steel", txt):
                            candidates.append(item)
                try:
                    self.main.type_keys("{ESC}", set_foreground=True)
                except Exception:
                    pass

                if candidates:
                    break

            if not candidates:
                return maybe_fail_or_skip("material menu item not found")

            clicked = False
            for item in candidates:
                try:
                    item.click_input()
                    clicked = True
                    break
                except Exception:
                    continue
            if not clicked:
                return maybe_fail_or_skip("material menu item could not be clicked")

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
            return maybe_fail_or_skip("material selection dialog not found")

        items = dialog.descendants(control_type="ListItem")
        if not items:
            return maybe_fail_or_skip("material list items not found")

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

        unexpected = find_unexpected_dialog(process_id=main_pid)
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
        main_pid = self._main_process_id()
        while time.time() < deadline:
            unexpected = find_unexpected_dialog(process_id=main_pid)
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

        # Keyboard fallback requested by operator:
        # Alt+F -> release Alt -> E -> 2 -> Enter
        self.main.set_focus()
        try:
            send_keys("%f")
            time.sleep(0.12)
            send_keys("{VK_MENU up}")
            time.sleep(0.08)
            send_keys("e2{ENTER}")
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
        pid = self._main_process_id()
        if not click_menu_item_anywhere(menu_path_items[1], timeout=timeout, process_id=pid):
            raise NeedsHelpError(
                f"Export menu item not found via fallback; menu path={menu_path_items}"
            )
        if not click_menu_item_anywhere(menu_path_items[2], timeout=timeout, process_id=pid):
            raise NeedsHelpError(
                f"Export target item not found via fallback; menu path={menu_path_items}"
            )

    def export_geo(self, export_path):
        export_path = normalize_windows_path(export_path)
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
        main_pid = self._main_process_id()
        try:
            dialog = wait_for_save_dialog(
                timeout=save_timeout,
                parent=self.main,
                process_id=main_pid,
                poll_interval=0.06,
            )
            if self.screenshotter:
                self.screenshotter.snap("saveas_dialog")
        except Exception as e:
            if self.screenshotter:
                self.screenshotter.snap("export_failed")
            raise NeedsHelpError(f"Save As dialog not found: {e}")

        try:
            Path(export_path).parent.mkdir(parents=True, exist_ok=True)
            if not ensure_dialog_focus(dialog, timeout=2.5):
                self.logger.warning("Save As focus check failed before typing; attempting recovery")
                try:
                    dialog.click_input()
                except Exception:
                    pass
                time.sleep(0.2)

            file_edit = set_file_name(dialog, export_path)
            # Requested flow: move focus with TAB after typing file name.
            try:
                dialog.type_keys("{TAB}", set_foreground=True)
            except Exception:
                pass

            typed = ""
            try:
                typed = (file_edit.window_text() or "").strip()
            except Exception:
                typed = ""
            expected_name = Path(export_path).name.lower()
            if typed and expected_name not in typed.lower():
                self.logger.warning(
                    "Save As file name mismatch after typing: expected~%s; got=%s",
                    expected_name,
                    typed,
                )

            if not ensure_dialog_focus(dialog, timeout=1.8):
                self.logger.warning("Save As focus check failed after typing; retrying focus")
                try:
                    file_edit.set_focus()
                except Exception:
                    pass

            # Requested sequence: ALT+S for save action.
            try:
                send_keys("%s")
                time.sleep(0.08)
                send_keys("{VK_MENU up}")
            except Exception:
                press_save(dialog)

            # Requested overwrite path: confirm with Y if overwrite dialog appears.
            deadline_y = time.time() + 2.0
            while time.time() < deadline_y:
                overwrite_dlg = find_overwrite_dialog(process_id=main_pid)
                if not overwrite_dlg:
                    time.sleep(0.15)
                    continue
                if self.screenshotter:
                    self.screenshotter.snap("overwrite_prompt")
                try:
                    send_keys("y")
                    time.sleep(0.25)
                except Exception:
                    pass
                if find_overwrite_dialog(process_id=main_pid):
                    if not confirm_overwrite_dialog(overwrite_dlg):
                        raise RuntimeError("Overwrite dialog detected but could not confirm with Y/Yes")
                if self.screenshotter:
                    self.screenshotter.snap("overwrite_confirmed")
                break
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
            overwrite_dlg = find_overwrite_dialog(process_id=main_pid)
            if overwrite_dlg:
                if self.screenshotter:
                    self.screenshotter.snap("overwrite_prompt")
                if not confirm_overwrite_dialog(overwrite_dlg):
                    raise NeedsHelpError("Overwrite dialog detected but could not confirm overwrite")
                if self.screenshotter:
                    self.screenshotter.snap("overwrite_confirmed")
                time.sleep(0.3)
                continue

            unexpected = find_unexpected_dialog(process_id=main_pid, ignore_overwrite=True)
            if unexpected:
                if self.screenshotter:
                    self.screenshotter.snap("export_failed")
                raise NeedsHelpError(f"Unexpected dialog during export: {unexpected}")

            if not save_dialog_present(parent=self.main, process_id=main_pid):
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

    def close_active_file(self):
        self.main.set_focus()
        self.main.type_keys("^w", set_foreground=True)
        time.sleep(0.4)
