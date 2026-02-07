import os
import time
from pathlib import Path

from pywinauto.application import Application

from ui_utils import (
    find_child,
    find_control,
    find_unexpected_dialog,
    handle_possible_dialogs,
    press_open,
    set_file_name,
    wait_for_open_dialog,
    wait_for_window,
)


class NeedsHelpError(RuntimeError):
    pass


class TecZoneSession:
    def __init__(self, logger, screenshotter=None):
        self.logger = logger
        self.screenshotter = screenshotter
        self.app = None
        self.main = None

        self.main_title_re = os.getenv("TECZONE_MAIN_TITLE_RE", r".*TecZone.*Bend.*")
        self.material_menu_titles = ["Material", "Material..."]
        self.material_dialog_title_re = os.getenv("TECZONE_MATERIAL_DIALOG_RE", r".*Material.*")
        self.export_menu_paths = [
            "File->Export",
            "File->Export...",
            "File->Save As",
            "File->Save As...",
        ]

    def connect(self, timeout=20):
        try:
            main_spec = wait_for_window(self.main_title_re, timeout=timeout)
            self.main = main_spec.wrapper_object()
            self.logger.info("Connected to TecZone window: %s", self.main.window_text())
        except Exception as e:
            exe = os.getenv("TECZONE_EXE")
            if exe and os.path.exists(exe):
                try:
                    self.logger.info("TecZone window not found, launching: %s", exe)
                    self.app = Application(backend="uia").start(exe)
                    self.main = wait_for_window(self.main_title_re, timeout=timeout).wrapper_object()
                    self.logger.info("Connected to TecZone window after launch: %s", self.main.window_text())
                    return
                except Exception as e2:
                    raise NeedsHelpError(f"TecZone launch failed: {e2}")
            raise NeedsHelpError(f"TecZone main window not found: {e}")

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
                raise NeedsHelpError(f"Open dialog not found and menu fallback failed: {e}")
            try:
                dialog = wait_for_open_dialog(timeout=5, parent=self.main)
            except Exception as e2:
                raise NeedsHelpError(f"Open dialog not found: {e2}")

        try:
            set_file_name(dialog, path)
            press_open(dialog)
        except Exception as e:
            raise NeedsHelpError(f"Open dialog interaction failed: {e}")

        deadline = time.time() + 90
        while time.time() < deadline:
            unexpected = find_unexpected_dialog()
            if unexpected:
                raise NeedsHelpError(f"Unexpected dialog while opening file: {unexpected}")
            if not find_child(self.main, control_type="Window", title="Open"):
                return
            time.sleep(0.3)

        raise NeedsHelpError("Open dialog did not close within 90 seconds")

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
