import argparse
import json
import logging
import os
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

from overlay import Overlay
from screenshot import Screenshotter
from teczone_actions import NeedsHelpError, TecZoneSession
from ui_utils import dump_window_titles, get_active_window_title
from xometry_parser import load_xometry_map

DEFAULT_POLL_SECONDS = 2
STEP_ORDER = ["OPEN_FILE", "SET_MATERIAL", "EXPORT_GEO"]


class PauseController:
    def __init__(self, logger):
        self.logger = logger
        self._paused = False
        self._lock = threading.Lock()
        self._listener = None

    def start(self, hotkey_spec):
        try:
            from pynput import keyboard
        except Exception as e:
            self.logger.warning("Hotkeys disabled (pynput unavailable): %s", e)
            return

        def on_toggle():
            with self._lock:
                self._paused = not self._paused
                state = "PAUSED" if self._paused else "RESUMED"
                self.logger.info(state)

        normalized = normalize_hotkey(hotkey_spec)
        self._listener = keyboard.GlobalHotKeys({normalized: on_toggle})
        self._listener.daemon = True
        self._listener.start()
        self.logger.info("Pause hotkey enabled: %s", hotkey_spec)

    def is_paused(self):
        with self._lock:
            return self._paused

    def stop(self):
        if self._listener:
            self._listener.stop()


def normalize_hotkey(value):
    tokens = [x.strip().lower() for x in value.split("+") if x.strip()]
    mapped = []
    for token in tokens:
        if token in ["ctrl", "control"]:
            mapped.append("<ctrl>")
        elif token == "alt":
            mapped.append("<alt>")
        elif token == "shift":
            mapped.append("<shift>")
        elif token in ["win", "windows", "cmd"]:
            mapped.append("<cmd>")
        else:
            mapped.append(token)
    return "+".join(mapped)


def play_sound_pattern(pattern, logger):
    try:
        import winsound

        for freq, duration_ms in pattern:
            winsound.Beep(freq, duration_ms)
            time.sleep(0.03)
    except Exception as e:
        logger.debug("Sound playback skipped: %s", e)


def play_job_start_sound(logger):
    # Short and friendly ascending tone.
    play_sound_pattern([(880, 110), (1175, 130)], logger)


def play_job_end_sound(logger, status):
    # Distinct endings by status.
    if status == "DONE":
        play_sound_pattern([(1318, 120), (1760, 140)], logger)
    elif status == "PARTIAL":
        play_sound_pattern([(988, 120), (784, 160)], logger)
    else:
        play_sound_pattern([(740, 150), (587, 180)], logger)


def read_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def write_json(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def configure_logger(log_path):
    logger = logging.getLogger("worker")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


def claim_job(job_id, state_dir):
    Path(state_dir).mkdir(parents=True, exist_ok=True)
    marker = Path(state_dir) / f"{job_id}.processing"
    try:
        with open(marker, "x", encoding="utf-8") as f:
            f.write(datetime.utcnow().isoformat())
        return True, marker
    except FileExistsError:
        return False, marker


def release_job(marker_path, status):
    try:
        done_marker = str(marker_path).replace(".processing", f".{status.lower()}")
        os.replace(marker_path, done_marker)
    except Exception:
        pass


def write_needs_help(path, step, found):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    titles = dump_window_titles()
    with open(path, "w", encoding="utf-8") as f:
        f.write("NEEDS_HELP\n")
        f.write(f"active_window: {get_active_window_title()}\n")
        f.write(f"step: {step}\n")
        f.write(f"found: {found}\n")
        f.write("window_titles:\n")
        for t in titles:
            f.write(f"- {t}\n")
    windows_json_path = str(Path(path).parent / "windows.json")
    write_json(windows_json_path, {"windows": titles})


def build_step_labels(parts):
    labels = []
    for part in parts:
        part_name = part.get("partName") or str(part.get("partId") or "unknown_part")
        for step in STEP_ORDER:
            labels.append(f"{step} {part_name}")
    return labels


def format_overlay_text(job_id, done_steps, total_steps, current_label, next_label, hotkey_hint, paused=False):
    state = "PAUSED" if paused else "RUNNING"
    line1 = f"WORKER {state}: {job_id} | steps {done_steps}/{total_steps} | current: {current_label}"
    line2 = f"next: {next_label} | hint: {hotkey_hint} = pause/resume"
    return f"{line1}\n{line2}"


def process_job(job_path, hotkey_pause="ctrl+alt+p", disable_hotkeys=False, disable_sounds=False):
    job = read_json(job_path)
    job_id = job["jobId"]
    project_root = job["projectRoot"]
    settings = job.get("settings", {})

    log_dir = Path(project_root) / "WORK" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = str(log_dir / f"{job_id}.log")
    logger = configure_logger(log_path)
    if not disable_sounds and not settings.get("disableSounds", False):
        play_job_start_sound(logger)

    screenshots_dir = Path(project_root) / "WORK" / "screenshots" / job_id
    screenshots_dir.mkdir(parents=True, exist_ok=True)

    export_dir = settings.get("exportDir") or str(Path(project_root) / "WORK" / "out" / "flat")
    Path(export_dir).mkdir(parents=True, exist_ok=True)

    xometry_map = {}
    if job.get("xometryJson"):
        try:
            xometry_map = load_xometry_map(job["xometryJson"], logger)
        except Exception as e:
            if settings.get("dryRun"):
                raise NeedsHelpError(f"Failed to parse xometry json: {e}")
            logger.warning("Failed to parse xometry json: %s", e)

    input_files = job.get("inputFiles", [])
    step_labels = build_step_labels(input_files)
    total_steps = len(step_labels)

    hotkey_enabled = (not disable_hotkeys) and (not settings.get("disableHotkeys", False))
    effective_hotkey = settings.get("hotkeyPause", hotkey_pause)
    hotkey_hint = effective_hotkey if hotkey_enabled else "hotkeys disabled"

    overlay = Overlay(
        format_overlay_text(
            job_id,
            0,
            total_steps,
            "INIT",
            "CONNECT_TECZONE",
            hotkey_hint,
            paused=False,
        )
    )
    overlay.start()

    screenshotter = Screenshotter(str(screenshots_dir))

    periodic_seconds = settings.get("screenshotsEverySeconds", 0)
    stop_event = threading.Event()
    pause_controller = PauseController(logger)
    if hotkey_enabled:
        pause_controller.start(effective_hotkey)

    if periodic_seconds and periodic_seconds > 0:

        def periodic():
            while not stop_event.is_set():
                screenshotter.snap("periodic")
                stop_event.wait(periodic_seconds)

        threading.Thread(target=periodic, daemon=True).start()

    result = {
        "jobId": job_id,
        "status": "DONE",
        "parts": [],
        "screenshotsDir": str(screenshots_dir),
        "logPath": log_path,
    }

    overall_status = "DONE"
    done_steps = 0

    def next_label_at(index):
        if index + 1 < len(step_labels):
            return step_labels[index + 1]
        return "DONE"

    def set_overlay(current_label, next_label, paused=False):
        overlay.set_text(
            format_overlay_text(
                job_id,
                done_steps,
                total_steps,
                current_label,
                next_label,
                hotkey_hint,
                paused=paused,
            )
        )

    try:
        tz = TecZoneSession(logger, screenshotter)
        try:
            set_overlay("CONNECT_TECZONE", "DRY_RUN" if settings.get("dryRun") else (step_labels[0] if step_labels else "DONE"))
            tz.connect()
        except NeedsHelpError as e:
            overall_status = "NEEDS_HELP"
            screenshotter.snap("needs_help")
            needs_help_path = str(Path(project_root) / "WORK" / "logs" / f"{job_id}_NEEDS_HELP.txt")
            write_needs_help(needs_help_path, "CONNECT_TECZONE", str(e))
            result["status"] = overall_status
            logs_dir = Path(project_root) / "WORK" / "logs"
            result_path = str(logs_dir / f"{job_id}.result.json")
            write_json(result_path, result)
            write_json(str(logs_dir / "result.json"), result)
            if not disable_sounds and not settings.get("disableSounds", False):
                play_job_end_sound(logger, overall_status)
            return result_path, overall_status

        if settings.get("dryRun"):
            set_overlay("DRY_RUN", "DONE")
            screenshotter.snap("dryrun_connected")
            logger.info("Dry run completed: connected to TecZone and parsed xometry json")
            result["status"] = "DONE"
            result_path = str(Path(project_root) / "WORK" / "logs" / f"{job_id}.result.json")
            write_json(result_path, result)
            write_json(str(Path(project_root) / "WORK" / "logs" / "result.json"), result)
            if not disable_sounds and not settings.get("disableSounds", False):
                play_job_end_sound(logger, "DONE")
            return result_path, "DONE"

        pause_shot_taken = False

        for index, part in enumerate(input_files):
            part_index = index + 1
            part_id = part.get("partId")
            part_name = part.get("partName") or str(part_id or "unknown_part")
            input_path = part.get("path")

            if input_path and os.path.splitext(input_path)[1].lower() not in [".stp", ".step"]:
                raise NeedsHelpError(f"Unsupported input extension: {input_path}")

            material = None
            if part_id in xometry_map:
                material = xometry_map[part_id].get("material")

            part_result = {
                "partId": part_id,
                "inputPath": input_path,
                "materialFromXometry": material,
                "materialUsedInTecZone": None,
                "thicknessMmDetected": None,
                "geoPath": None,
                "status": "DONE",
                "notes": "",
            }

            step = ""

            def wait_if_paused(current_label, next_label):
                nonlocal pause_shot_taken
                while pause_controller.is_paused():
                    set_overlay(current_label, next_label, paused=True)
                    if not pause_shot_taken:
                        screenshotter.snap("paused")
                        pause_shot_taken = True
                    time.sleep(0.25)
                if pause_shot_taken:
                    pause_shot_taken = False

            try:
                step_idx = (part_index - 1) * len(STEP_ORDER)
                step = f"OPEN_FILE {part_name}"
                wait_if_paused(step, next_label_at(step_idx))
                set_overlay(step, next_label_at(step_idx))
                screenshotter.snap("open_file_start")
                tz.open_file(input_path)
                screenshotter.snap("open_file_done")
                done_steps += 1

                step = f"SET_MATERIAL {part_name}"
                wait_if_paused(step, next_label_at(step_idx + 1))
                set_overlay(step, next_label_at(step_idx + 1))
                screenshotter.snap("material_start")
                used_material, note = tz.set_material(material)
                part_result["materialUsedInTecZone"] = used_material
                if note:
                    part_result["notes"] += note
                screenshotter.snap("material_done")
                done_steps += 1

                step = f"EXPORT_GEO {part_name}"
                wait_if_paused(step, next_label_at(step_idx + 2))
                set_overlay(step, next_label_at(step_idx + 2))
                screenshotter.snap("export_start")
                export_name_template = settings.get("exportNameTemplate", "<partName>.geo")
                export_name = export_name_template.replace("<partName>", part_name)
                export_path = str(Path(export_dir) / export_name)
                tz.export_geo(export_path)
                part_result["geoPath"] = export_path
                screenshotter.snap("export_done")

                thickness = tz.get_thickness_mm()
                part_result["thicknessMmDetected"] = thickness
                done_steps += 1

            except NeedsHelpError as e:
                part_result["status"] = "NEEDS_HELP"
                part_result["notes"] += str(e)
                overall_status = "NEEDS_HELP"
                screenshotter.snap("needs_help")
                needs_help_path = str(Path(project_root) / "WORK" / "logs" / f"{job_id}_NEEDS_HELP.txt")
                write_needs_help(needs_help_path, step, str(e))
                result["parts"].append(part_result)
                break
            except Exception as e:
                part_result["status"] = "FAILED"
                part_result["notes"] += f"Exception: {e}"
                logger.error("Exception on part %s: %s", part_id, traceback.format_exc())
                screenshotter.snap("failed")
                if overall_status == "DONE":
                    overall_status = "PARTIAL"

            result["parts"].append(part_result)

    finally:
        stop_event.set()
        pause_controller.stop()
        overlay.stop()

    if overall_status == "DONE":
        if any(p["status"] == "FAILED" for p in result["parts"]):
            overall_status = "PARTIAL" if any(p["status"] == "DONE" for p in result["parts"]) else "FAILED"

    result["status"] = overall_status
    logs_dir = Path(project_root) / "WORK" / "logs"
    result_path = str(logs_dir / f"{job_id}.result.json")
    write_json(result_path, result)
    write_json(str(logs_dir / "result.json"), result)
    if not disable_sounds and not settings.get("disableSounds", False):
        play_job_end_sound(logger, overall_status)

    return result_path, overall_status


def run_loop(jobs_dir, hotkey_pause="ctrl+alt+p", disable_hotkeys=False, disable_sounds=False):
    jobs_dir = Path(jobs_dir)
    state_dir = jobs_dir.parent / "state"

    while True:
        for job_path in jobs_dir.glob("*.json"):
            try:
                job = read_json(job_path)
                job_id = job.get("jobId") or job_path.stem
            except Exception:
                continue

            ok, marker = claim_job(job_id, state_dir)
            if not ok:
                continue

            status = "FAILED"
            try:
                _, status = process_job(
                    job_path,
                    hotkey_pause=hotkey_pause,
                    disable_hotkeys=disable_hotkeys,
                    disable_sounds=disable_sounds,
                )
            except Exception:
                status = "FAILED"
            finally:
                release_job(marker, status)

        time.sleep(DEFAULT_POLL_SECONDS)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--jobs-dir", required=True, help="Path to WORK\\jobs")
    parser.add_argument("--hotkey-pause", default="ctrl+alt+p", help="Global pause hotkey")
    parser.add_argument("--disable-hotkeys", action="store_true", help="Disable global hotkeys")
    parser.add_argument("--disable-sounds", action="store_true", help="Disable start/end sounds")
    args = parser.parse_args()
    run_loop(
        args.jobs_dir,
        hotkey_pause=args.hotkey_pause,
        disable_hotkeys=args.disable_hotkeys,
        disable_sounds=args.disable_sounds,
    )


if __name__ == "__main__":
    main()
