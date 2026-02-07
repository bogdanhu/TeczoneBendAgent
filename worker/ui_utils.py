import re
import time

from pywinauto import Desktop

OPEN_TITLE_RE = r"(?i)\bopen\b|deschide|oeffnen|offnen|ouvrir"
SAVE_TITLE_RE = r"(?i)\bsave\b|\bexport\b|salveaza|speichern|enregistrer"
FILE_NAME_LABEL_RE = r"(?i)file\s*name|nume\s*fisier|nume\s*fi.sier|dateiname"
UNEXPECTED_DIALOG_KEYWORDS = [
    "warning",
    "confirm",
    "overwrite",
    "error",
    "attention",
    "question",
    "invalid",
    "not valid",
    "cannot",
    "failed",
    "does not exist",
]


def normalize_windows_path(value):
    text = str(value or "").strip().strip('"')
    text = text.replace("/", "\\")
    if text.startswith("\\\\"):
        prefix = "\\\\"
        rest = text[2:]
        rest = re.sub(r"\\{2,}", r"\\", rest)
        return prefix + rest
    return re.sub(r"\\{2,}", r"\\", text)


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


def describe_controls(parent, limit=40):
    items = []
    try:
        descendants = parent.descendants()
    except Exception:
        descendants = []
    for ctrl in descendants[:limit]:
        try:
            ctrl_type = ctrl.element_info.control_type
        except Exception:
            ctrl_type = ""
        try:
            ctrl_title = ctrl.window_text() or ""
        except Exception:
            ctrl_title = ""
        try:
            ctrl_auto_id = ctrl.element_info.automation_id or ""
        except Exception:
            ctrl_auto_id = ""
        items.append(f"{ctrl_type}|title={ctrl_title}|auto_id={ctrl_auto_id}")
    return items


def _open_search_roots(parent=None):
    roots = []
    if parent is not None:
        roots.append(parent)
        try:
            roots.extend(parent.descendants(control_type="Window"))
        except Exception:
            pass
    try:
        roots.extend(Desktop(backend="uia").windows())
    except Exception:
        pass
    return roots


def _is_common_file_dialog(dlg):
    try:
        if (dlg.element_info.control_type or "") != "Window":
            return False
    except Exception:
        return False

    try:
        class_name = (dlg.element_info.class_name or "").strip()
    except Exception:
        class_name = ""

    has_file_name = find_child(dlg, auto_id="1148") is not None
    if not has_file_name:
        label = find_child(dlg, control_type="Text", title_re=FILE_NAME_LABEL_RE)
        if label:
            try:
                edits = dlg.descendants(control_type="Edit")
            except Exception:
                edits = []
            has_file_name = len(edits) > 0
    if not has_file_name:
        try:
            combos = dlg.descendants(control_type="ComboBox")
        except Exception:
            combos = []
        for combo in combos:
            try:
                combo_title = (combo.window_text() or "").strip()
            except Exception:
                combo_title = ""
            if combo_title and re.search(FILE_NAME_LABEL_RE, combo_title):
                has_file_name = True
                break

    has_primary = find_child(dlg, control_type="Button", auto_id="1") is not None
    if not has_primary:
        has_primary = find_child(dlg, control_type="Button", title_re=r"(?i)open|save|export|ok") is not None
    has_cancel = find_child(dlg, control_type="Button", auto_id="2") is not None

    if class_name == "#32770":
        return has_file_name and has_primary
    return has_file_name and has_primary and has_cancel


def find_open_dialog(parent=None, title_re=OPEN_TITLE_RE):
    for dlg in _open_search_roots(parent):
        if not _is_common_file_dialog(dlg):
            continue
        try:
            title = (dlg.window_text() or "").strip()
        except Exception:
            title = ""
        if title and re.search(title_re, title):
            return dlg
    return None


def find_save_dialog(parent=None, title_re=SAVE_TITLE_RE):
    for dlg in _open_search_roots(parent):
        if not _is_common_file_dialog(dlg):
            continue
        try:
            title = (dlg.window_text() or "").strip()
        except Exception:
            title = ""
        if title and re.search(title_re, title):
            return dlg
        # Structure fallback for builds that expose non-standard title text.
        primary = find_child(dlg, control_type="Button", auto_id="1")
        if not primary:
            primary = find_child(dlg, control_type="Button", title_re=r"(?i)^save$|^export$|^ok$")
        if primary:
            return dlg
    return None


def debug_open_dialog_search(parent=None):
    found_windows = []
    for dlg in _open_search_roots(parent):
        try:
            title = (dlg.window_text() or "").strip()
        except Exception:
            title = ""
        if title:
            found_windows.append(title)
    return {
        "searched": "open dialog by title~OPEN + file_name_edit(auto_id=1148/label File name) + open_button(auto_id=1/title Open)",
        "found_windows": found_windows[:20],
    }


def set_file_name(dialog, value):
    value = normalize_windows_path(value)
    candidates = []

    # Primary path for modern common dialogs.
    combo_1148 = find_child(dialog, control_type="ComboBox", auto_id="1148")
    if combo_1148:
        try:
            combo_edit = combo_1148.child_window(control_type="Edit")
            if combo_edit.exists(timeout=0.2):
                candidates.append(combo_edit)
        except Exception:
            pass
        candidates.append(combo_1148)

    edit_1148 = find_child(dialog, control_type="Edit", auto_id="1148")
    if edit_1148:
        candidates.append(edit_1148)

    generic_1148 = find_child(dialog, auto_id="1148")
    if generic_1148:
        candidates.append(generic_1148)

    # Secondary path: infer from "File name:" label row.
    label = find_child(dialog, control_type="Text", title_re=FILE_NAME_LABEL_RE)
    if label:
        try:
            parent = label.parent()
            if parent:
                for edit in parent.descendants(control_type="Edit"):
                    candidates.append(edit)
        except Exception:
            pass

    # Deduplicate and remove obvious non-target edits (search/rename inline fields).
    deduped = []
    seen = set()
    for candidate in candidates:
        if candidate is None:
            continue
        try:
            key = tuple(candidate.element_info.runtime_id)
        except Exception:
            key = id(candidate)
        if key in seen:
            continue
        seen.add(key)
        try:
            title = (candidate.window_text() or "").strip().lower()
        except Exception:
            title = ""
        if "search" in title:
            continue
        deduped.append(candidate)

    last_error = None
    for candidate in deduped:
        try:
            candidate.set_focus()
            candidate.set_edit_text(value)
            return candidate
        except Exception as e:
            last_error = e
            continue

    raise RuntimeError(f"File name edit not found/settable (strict mode): {last_error}")


def press_open(dialog):
    button = find_child(dialog, control_type="Button", auto_id="1")
    if not button:
        button = find_child(dialog, control_type="Button", title_re=OPEN_TITLE_RE)
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
        dlg = find_open_dialog(parent=parent)
        if dlg:
            return dlg
        time.sleep(0.25)
    raise RuntimeError("Open file dialog not found")


def wait_for_save_dialog(timeout=20, parent=None):
    deadline = time.time() + timeout
    while time.time() < deadline:
        dlg = find_save_dialog(parent=parent)
        if dlg:
            return dlg
        time.sleep(0.25)
    raise RuntimeError("Save/Export dialog not found")


def open_dialog_present(parent=None):
    return find_open_dialog(parent=parent) is not None


def save_dialog_present(parent=None):
    return find_save_dialog(parent=parent) is not None


def press_save(dialog):
    button = find_child(dialog, control_type="Button", auto_id="1")
    if not button:
        button = find_child(dialog, control_type="Button", title_re=r"(?i)^save$|^export$|^ok$")
    if button:
        button.click_input()
        return
    dialog.type_keys("{ENTER}")


def click_menu_item_anywhere(label, timeout=4, process_id=None):
    pattern = re.compile(rf"(?i)^{re.escape(label)}$")
    deadline = time.time() + timeout
    while time.time() < deadline:
        for w in Desktop(backend="uia").windows():
            try:
                if process_id is not None and int(w.element_info.process_id) != int(process_id):
                    continue
            except Exception:
                if process_id is not None:
                    continue
            try:
                menu_items = w.descendants(control_type="MenuItem")
            except Exception:
                continue
            for item in menu_items:
                try:
                    text = item.window_text() or ""
                except Exception:
                    text = ""
                if not text:
                    continue
                if pattern.search(text):
                    try:
                        item.click_input()
                        return True
                    except Exception:
                        continue
        time.sleep(0.2)
    return False


def find_unexpected_dialog(process_id=None):
    desktop = Desktop(backend="uia")
    for dlg in desktop.windows():
        try:
            if process_id is not None and int(dlg.element_info.process_id) != int(process_id):
                continue
        except Exception:
            if process_id is not None:
                continue
        title = (dlg.window_text() or "").strip()
        if not title:
            continue

        # Ignore regular Open/Save dialogs that are part of normal flow.
        try:
            if find_child(dlg, auto_id="1148") and (
                re.search(OPEN_TITLE_RE, title) or re.search(SAVE_TITLE_RE, title)
            ):
                continue
        except Exception:
            pass

        title_l = title.lower()
        text_blobs = [title_l]
        try:
            for txt in dlg.descendants(control_type="Text"):
                t = (txt.window_text() or "").strip().lower()
                if t:
                    text_blobs.append(t)
        except Exception:
            pass
        whole = " | ".join(text_blobs)
        if any(k in whole for k in UNEXPECTED_DIALOG_KEYWORDS):
            return f"{title}: {whole[:300]}"
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
