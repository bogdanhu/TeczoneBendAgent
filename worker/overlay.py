import threading
import tkinter as tk


class Overlay:
    def __init__(self, text):
        self.text = text
        self._thread = None
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self._ready = threading.Event()
        self._root = None

    def start(self):
        if self._thread:
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def set_text(self, text):
        with self._lock:
            self.text = text

    def stop(self):
        self._stop.set()
        if self._ready.wait(timeout=1.0) and self._root is not None:
            try:
                self._root.after(0, self._root.quit)
            except Exception:
                pass
        if self._thread:
            self._thread.join(timeout=2.0)

    def _run(self):
        root = tk.Tk()
        self._root = root
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.96)
        root.geometry("1700x64+8+8")

        label = tk.Label(
            root,
            text=self.text,
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

        def tick():
            if self._stop.is_set():
                root.quit()
                return
            with self._lock:
                label.configure(text=self.text)
            root.after(200, tick)

        self._ready.set()
        tick()
        root.mainloop()
        try:
            root.destroy()
        except Exception:
            pass
