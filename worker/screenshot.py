import os
import time
from datetime import datetime


class Screenshotter:
    def __init__(self, output_dir):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        try:
            import mss  # noqa: F401
            self._use_mss = True
        except Exception:
            self._use_mss = False

    def snap(self, name):
        ts = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        path = os.path.join(self.output_dir, f"{ts}_{name}.png")
        try:
            if self._use_mss:
                import mss
                with mss.mss() as sct:
                    sct.shot(output=path)
            else:
                from PIL import ImageGrab
                img = ImageGrab.grab()
                img.save(path)
        except Exception:
            pass
        time.sleep(0.1)
        return path
