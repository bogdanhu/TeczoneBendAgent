import queue
import threading
import tkinter as tk


class Overlay:
    def __init__(self, text):
        self._text = text
        self._thread = None
        self._queue = queue.Queue()
        self._ready = threading.Event()

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(timeout=2.0)

    def set_text(self, text):
        self._text = text
        if self._thread and self._thread.is_alive():
            self._queue.put(("set_text", text))

    def stop(self):
        if not self._thread:
            return
        if self._thread.is_alive():
            self._queue.put(("stop", None))
            self._thread.join(timeout=3.0)
        self._thread = None

    def _run(self):
        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.96)
        root.geometry("1700x64+8+8")

        label = tk.Label(
            root,
            text=self._text,
            bg="#111111",
            fg="#ffffff",
            font=("Segoe UI", 12, "bold"),
            justify="left",
            anchor="w",
            padx=12,
            pady=6,
            wraplength=1680,
        )
        label.pack(fill="both", expand=True)

        def poll_queue():
            should_stop = False
            while True:
                try:
                    command, value = self._queue.get_nowait()
                except queue.Empty:
                    break
                if command == "set_text":
                    try:
                        label.configure(text=value)
                    except Exception:
                        pass
                elif command == "stop":
                    should_stop = True
            if should_stop:
                root.quit()
                return
            root.after(80, poll_queue)

        self._ready.set()
        root.after(80, poll_queue)
        root.mainloop()
        try:
            root.destroy()
        except Exception:
            pass
