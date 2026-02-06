import threading
import tkinter as tk


class Overlay:
    def __init__(self, text):
        self.text = text
        self._thread = None
        self._stop = threading.Event()
        self._lock = threading.Lock()

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

    def _run(self):
        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.geometry("800x30+10+10")

        label = tk.Label(root, text=self.text, bg="#222222", fg="#ffffff", font=("Segoe UI", 12))
        label.pack(fill="both", expand=True)

        def tick():
            if self._stop.is_set():
                root.destroy()
                return
            with self._lock:
                label.configure(text=self.text)
            root.after(200, tick)

        tick()
        root.mainloop()
