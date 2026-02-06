import time

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
