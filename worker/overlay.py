import json
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path


def _run_overlay_child():
    import tkinter as tk

    initial_text = os.environ.get("OVERLAY_INITIAL_TEXT", "WORKER")
    command_queue = queue.Queue()

    def reader():
        for raw_line in sys.stdin:
            line = raw_line.strip()
            if not line:
                continue
            try:
                payload = json.loads(line)
            except Exception:
                continue
            command_queue.put(payload)

    reader_thread = threading.Thread(target=reader, daemon=True)
    reader_thread.start()

    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.attributes("-alpha", 0.96)
    root.geometry("1700x92+8+8")

    label = tk.Label(
        root,
        text=initial_text,
        bg="#111111",
        fg="#ffffff",
        font=("Segoe UI", 12, "bold"),
        justify="left",
        anchor="w",
        padx=12,
        pady=8,
        wraplength=1680,
    )
    label.pack(fill="both", expand=True)

    def poll_queue():
        should_stop = False
        while True:
            try:
                payload = command_queue.get_nowait()
            except queue.Empty:
                break
            cmd = payload.get("cmd")
            if cmd == "set_text":
                text = payload.get("text", "")
                try:
                    label.configure(text=text)
                except Exception:
                    pass
            elif cmd == "stop":
                should_stop = True
        if should_stop:
            try:
                root.destroy()
            except Exception:
                pass
            return
        root.after(80, poll_queue)

    root.after(80, poll_queue)
    root.mainloop()


class Overlay:
    def __init__(self, text):
        self._text = text
        self._proc = None

    def start(self):
        if self._proc and self._proc.poll() is None:
            return

        env = os.environ.copy()
        env["OVERLAY_INITIAL_TEXT"] = self._text

        self._proc = subprocess.Popen(
            [sys.executable, str(Path(__file__).resolve()), "--overlay-child"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            env=env,
        )

    def set_text(self, text):
        self._text = text
        if not self._proc or self._proc.poll() is not None or not self._proc.stdin:
            return
        try:
            self._proc.stdin.write(json.dumps({"cmd": "set_text", "text": text}) + "\n")
            self._proc.stdin.flush()
        except Exception:
            pass

    def stop(self):
        if not self._proc:
            return
        if self._proc.poll() is None and self._proc.stdin:
            try:
                self._proc.stdin.write(json.dumps({"cmd": "stop"}) + "\n")
                self._proc.stdin.flush()
            except Exception:
                pass
        try:
            self._proc.wait(timeout=5.0)
        except Exception:
            try:
                self._proc.kill()
            except Exception:
                pass
        self._proc = None


if __name__ == "__main__":
    if "--overlay-child" in sys.argv:
        _run_overlay_child()
