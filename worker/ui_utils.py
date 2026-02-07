import time
import re

from pywinauto import Desktop


def wait_for_window(title_re, timeout=10, backend="uia"):
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            win = Desktop(backend=backend).window(title_re=title_re)
            if win.exists():
                return win
        except Exception as e:
            last_error = e
        time.sleep(0.3)
    raise RuntimeError(f"Window not found: {title_re}; last_error={last_error}")


def find_control(parent, **criteria):
    try:
        ctrl = parent.child_window(**criteria)
        if ctrl.exists():
            return ctrl
    except Exception:
        return None
    return None


def find_child(parent, control_type=None, title_re=None, auto_id=None, title=None):
    try:
        descendants = parent.descendants()
    except Exception:
        descendants = []
    for ctrl in descendants:
        try:
            ctrl_type = ctrl.element_info.control_type
        except Exception:
            ctrl_type = None
        try:
            ctrl_title = ctrl.window_text() or ""
        except Exception:
            ctrl_title = ""
        try:
            ctrl_auto_id = ctrl.element_info.automation_id
        except Exception:
            ctrl_auto_id = None

        if control_type and ctrl_type != control_type:
            continue
        if auto_id and ctrl_auto_id != auto_id:
            continue
        if title and ctrl_title != title:
            continue
        if title_re and not re.search(title_re, ctrl_title):
            continue
        return ctrl
    return None


def set_file_name(dialog, value):
    edit = find_child(dialog, control_type="Edit", auto_id="1148")
    if not edit:
        label = find_child(dialog, control_type="Text", title_re=r"(?i)file\s*name")
        if label:
            try:
                edit = label.parent().child_window(control_type="Edit")
                if not edit.exists():
                    edit = None
            except Exception:
                edit = None
    if not edit:
        edits = dialog.descendants(control_type="Edit")
        edit = edits[0] if edits else None
    if not edit:
        raise RuntimeError("Open dialog file name edit not found")
    edit.set_focus()
    try:
        edit.set_edit_text(value)
    except Exception:
        pass
    # Force the exact path through keystrokes as fallback for Win11 common dialog.
    try:
        edit.type_keys("^a{BACKSPACE}", set_foreground=True)
        edit.type_keys(value, with_spaces=True, set_foreground=True)
    except Exception:
        pass
    return edit


def press_open(dialog):
    button = find_child(dialog, control_type="Button", auto_id="1")
    if not button:
        button = find_child(dialog, control_type="Button", title_re=r"(?i)^open$")
    if button:
        button.click_input()
        return
    dialog.type_keys("{ENTER}")


def wait_window_closed(window, timeout=90):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if not window.exists(timeout=0.1):
                return True
        except Exception:
            return True
        time.sleep(0.3)
    return False


def wait_for_open_dialog(timeout=5, parent=None):
    deadline = time.time() + timeout
    while time.time() < deadline:
        candidates = []
        if parent is not None:
            try:
                candidates.append(parent)
                candidates.extend(parent.descendants(control_type="Window"))
            except Exception:
                pass
        try:
            candidates.extend(Desktop(backend="uia").windows())
        except Exception:
            pass

        for dlg in candidates:
            # Common file dialog has a File name edit and an Open button.
            try:
                title = (dlg.window_text() or "").strip()
            except Exception:
                title = ""
            if title and "open" not in title.lower():
                continue

            edit = find_child(dlg, control_type="Edit", auto_id="1148")
            if not edit:
                edit = find_child(dlg, control_type="Edit", title_re=r"(?i)file\s*name")
            if not edit:
                try:
                    edits = dlg.descendants(control_type="Edit")
                except Exception:
                    edits = []
                edit = edits[0] if edits else None
            if not edit:
                continue
            open_btn = find_child(dlg, control_type="Button", auto_id="1")
            if not open_btn:
                open_btn = find_child(dlg, control_type="Button", title_re=r"(?i)^open$")
            if open_btn:
                return dlg
        time.sleep(0.25)
    raise RuntimeError("Open file dialog not found")


def find_unexpected_dialog():
    desktop = Desktop(backend="uia")
    for dlg in desktop.windows():
        title = (dlg.window_text() or "").strip()
        if not title:
            continue
        title_l = title.lower()
        if any(k in title_l for k in ["warning", "confirm", "overwrite", "error", "attention", "question"]):
            return title
    return None


def handle_possible_dialogs():
    desktop = Desktop(backend="uia")
    for dlg in desktop.windows():
        title = dlg.window_text()
        if not title:
            continue
        if "warning" in title.lower() or "confirm" in title.lower() or "overwrite" in title.lower():
            btn = find_control(dlg, title_re=r"Yes|OK|Overwrite|Replace", control_type="Button")
            if btn:
                btn.click_input()
                time.sleep(0.2)


def get_active_window_title():
    try:
        return Desktop(backend="uia").get_active().window_text()
    except Exception:
        return ""


def dump_window_titles():
    titles = []
    for w in Desktop(backend="uia").windows():
        t = w.window_text()
        if t:
            titles.append(t)
    return titles
