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
MAJOR_STEPS = ["OPEN_FILE", "SET_MATERIAL", "EXPORT_GEO"]

try:
    import sentry_sdk
except Exception:
    sentry_sdk = None

_TELEMETRY_INIT_DONE = False
_TELEMETRY_ENABLED = False


def resolve_glitchtip_dsn():
    return os.getenv("GLITCHTIP_DSN") or os.getenv("SENTRY_DSN")


def init_glitchtip(logger=None):
    global _TELEMETRY_INIT_DONE, _TELEMETRY_ENABLED
    if _TELEMETRY_INIT_DONE:
        return _TELEMETRY_ENABLED

    dsn = resolve_glitchtip_dsn()
    _TELEMETRY_INIT_DONE = True
    if not dsn:
        _TELEMETRY_ENABLED = False
        if logger:
            logger.info("GlitchTip disabled: GLITCHTIP_DSN/SENTRY_DSN not set")
        return False

    if sentry_sdk is None:
        _TELEMETRY_ENABLED = False
        if logger:
            logger.warning("GlitchTip DSN provided but sentry_sdk is unavailable")
        return False

    try:
        sentry_sdk.init(dsn=dsn, traces_sample_rate=0.0)
        _TELEMETRY_ENABLED = True
        if logger:
            logger.info("GlitchTip enabled")
        return True
    except Exception as e:
        _TELEMETRY_ENABLED = False
        if logger:
            logger.warning("GlitchTip init failed: %s", e)
        return False


def capture_glitchtip_event(
    level,
    message,
    *,
    job_id=None,
    xometry_ref=None,
    status=None,
    step=None,
    part_id=None,
    project_root=None,
    input_path=None,
    export_path=None,
    log_path=None,
    screenshots_dir=None,
    reason=None,
    exc=None,
):
    if not _TELEMETRY_ENABLED or sentry_sdk is None:
        return

    with sentry_sdk.new_scope() as scope:
        scope.set_tag("app", "teczonebend-worker")
        if job_id:
            scope.set_tag("jobId", str(job_id))
        if xometry_ref:
            scope.set_tag("xometryRef", str(xometry_ref))
        if status:
            scope.set_tag("status", str(status))
        if step:
            scope.set_tag("step", str(step))
        if part_id is not None:
            scope.set_tag("partId", str(part_id))

        if project_root:
            scope.set_extra("projectRoot", str(project_root))
        if input_path:
            scope.set_extra("inputPath", str(input_path))
        if export_path:
            scope.set_extra("exportPath", str(export_path))
        if log_path:
            scope.set_extra("logPath", str(log_path))
        if screenshots_dir:
            scope.set_extra("screenshotsDir", str(screenshots_dir))
        if reason:
            scope.set_extra("reason", str(reason))

        if exc is not None:
            sentry_sdk.capture_exception(exc)
        else:
            sentry_sdk.capture_message(message, level=level)


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
            self.logger.warning(
                "Hotkeys disabled because pynput is unavailable: %s. "
                "Install dependencies with: pip install -r requirements.txt",
                e,
            )
            return

        def on_toggle():
            with self._lock:
                self._paused = not self._paused
                state = "PAUSED by hotkey" if self._paused else "RESUMED by hotkey"
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


def format_overlay_text(job_id, done_steps, total_steps, current_action, next_action, hotkey_hint, paused=False):
    state = "PAUSED" if paused else "RUNNING"
    line1 = f"WORKER {state}: {job_id} | steps {done_steps}/{total_steps} | current: {current_action}"
    line2 = f"next: {next_action} | hint: {hotkey_hint} = pause/resume"
    return f"{line1}\n{line2}"


def process_job(
    job_path,
    hotkey_pause="ctrl+alt+p",
    disable_hotkeys=False,
    disable_sounds=False,
    no_overlay=False,
    teczone_exe=None,
    teczone_title_re=None,
):
    job = read_json(job_path)
    job_id = job["jobId"]
    project_root = job["projectRoot"]
    xometry_ref = job.get("xometryRef")
    settings = job.get("settings", {})

    log_dir = Path(project_root) / "WORK" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = str(log_dir / f"{job_id}.log")
    logger = configure_logger(log_path)
    init_glitchtip(logger)
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
    total_parts = len(input_files)
    total_steps = total_parts * len(MAJOR_STEPS)
    hotkey_enabled = (not disable_hotkeys) and (not settings.get("disableHotkeys", False))
    effective_hotkey = settings.get("hotkeyPause", hotkey_pause)
    hotkey_hint = effective_hotkey if hotkey_enabled else "hotkeys disabled"

    overlay = None
    initial_next = "OPEN_FILE" if total_parts > 0 else "WAIT_JOB"
    if not no_overlay:
        try:
            overlay = Overlay(format_overlay_text(job_id, 0, total_steps, "INIT", initial_next, hotkey_hint))
            overlay.start()
        except Exception as e:
            logger.warning("Overlay disabled due startup error: %s", e)
            overlay = None
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
    last_step = "INIT"
    last_part_id = None
    last_input_path = None
    last_export_path = None

    def _part_name(i):
        if i < 1 or i > total_parts:
            return "-"
        candidate = input_files[i - 1].get("partName")
        if candidate:
            return candidate
        pid = input_files[i - 1].get("partId")
        return str(pid or "unknown_part")

    def _step_index(step_name):
        try:
            return MAJOR_STEPS.index(step_name) + 1
        except ValueError:
            return None

    def _overlay_progress(part_index, step_name):
        idx = _step_index(step_name)
        if idx is None:
            if step_name == "CONNECT_TECZONE":
                return 0, f"OPEN_FILE {_part_name(1)}" if total_parts else "DRY_RUN"
            if step_name == "DRY_RUN":
                return 0, "FINISH_JOB"
            if step_name == "CLOSE_FILE":
                done_steps = min(total_steps, part_index * len(MAJOR_STEPS))
                next_action = f"OPEN_FILE {_part_name(part_index + 1)}" if part_index < total_parts else "FINISH_JOB"
                return done_steps, next_action
            return 0, "WAIT"

        done_steps = ((part_index - 1) * len(MAJOR_STEPS)) + (idx - 1)
        if idx < len(MAJOR_STEPS):
            next_action = f"{MAJOR_STEPS[idx]} {_part_name(part_index)}"
        elif part_index < total_parts:
            next_action = f"OPEN_FILE {_part_name(part_index + 1)}"
        else:
            next_action = "FINISH_JOB"
        return done_steps, next_action

    def set_overlay(part_index, step, part_name, paused=False):
        if overlay is not None:
            done_steps, next_action = _overlay_progress(part_index, step)
            current_action = f"{step} {part_name}".strip()
            overlay.set_text(
                format_overlay_text(
                    job_id,
                    done_steps,
                    total_steps,
                    current_action,
                    next_action,
                    hotkey_hint,
                    paused=paused,
                )
            )

    try:
        tz = TecZoneSession(
            logger,
            screenshotter,
            teczone_exe=teczone_exe,
            teczone_title_re=teczone_title_re,
            workflow_config_path=settings.get("teczoneWorkflowConfig"),
        )
        try:
            set_overlay(0, "CONNECT_TECZONE", "-")
            last_step = "CONNECT_TECZONE"
            tz.connect()
        except NeedsHelpError as e:
            overall_status = "NEEDS_HELP"
            screenshotter.snap("needs_help")
            needs_help_path = str(Path(project_root) / "WORK" / "logs" / f"{job_id}_NEEDS_HELP.txt")
            write_needs_help(needs_help_path, "CONNECT_TECZONE", str(e))
            capture_glitchtip_event(
                "error",
                "worker NEEDS_HELP at connect",
                job_id=job_id,
                xometry_ref=xometry_ref,
                status=overall_status,
                step="CONNECT_TECZONE",
                part_id=None,
                project_root=project_root,
                input_path=None,
                export_path=None,
                log_path=log_path,
                screenshots_dir=str(screenshots_dir),
                reason=str(e),
            )
            result["status"] = overall_status
            result_path = str(log_dir / f"{job_id}.result.json")
            write_json(result_path, result)
            write_json(str(log_dir / "result.json"), result)
            if not disable_sounds and not settings.get("disableSounds", False):
                play_job_end_sound(logger, overall_status)
            return result_path, overall_status

        if settings.get("dryRun"):
            set_overlay(0, "DRY_RUN", "-")
            last_step = "DRY_RUN"
            screenshotter.snap("dryrun_connected")
            logger.info("Dry run completed: connected to TecZone and parsed xometry json")
            result["status"] = "DONE"
            result_path = str(log_dir / f"{job_id}.result.json")
            write_json(result_path, result)
            write_json(str(log_dir / "result.json"), result)
            capture_glitchtip_event(
                "info",
                "worker DONE (dryRun)",
                job_id=job_id,
                xometry_ref=xometry_ref,
                status="DONE",
                step="DRY_RUN",
                part_id=None,
                project_root=project_root,
                input_path=None,
                export_path=None,
                log_path=log_path,
                screenshots_dir=str(screenshots_dir),
                reason="dry run completed",
            )
            if not disable_sounds and not settings.get("disableSounds", False):
                play_job_end_sound(logger, "DONE")
            return result_path, "DONE"

        pause_shot_taken = False
        for index, part in enumerate(input_files):
            part_index = index + 1
            part_id = part.get("partId")
            part_name = part.get("partName") or str(part_id or "unknown_part")
            input_path = part.get("path")
            last_part_id = part_id
            last_input_path = input_path
            last_export_path = None

            if input_path and os.path.splitext(input_path)[1].lower() not in [".stp", ".step"]:
                raise NeedsHelpError(f"Unsupported input extension: {input_path}")

            material = xometry_map.get(part_id, {}).get("material")
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
            opened_document = False

            def wait_if_paused(current_step):
                nonlocal pause_shot_taken
                while pause_controller.is_paused():
                    set_overlay(part_index, current_step, part_name, paused=True)
                    if not pause_shot_taken:
                        screenshotter.snap("paused")
                        pause_shot_taken = True
                    time.sleep(0.25)
                if pause_shot_taken:
                    pause_shot_taken = False

            try:
                step = "OPEN_FILE"
                last_step = step
                wait_if_paused(step)
                set_overlay(part_index, step, part_name)
                screenshotter.snap("open_file_start")
                tz.open_file(input_path)
                screenshotter.snap("open_file_done")
                opened_document = True

                step = "SET_MATERIAL"
                last_step = step
                wait_if_paused(step)
                set_overlay(part_index, step, part_name)
                screenshotter.snap("material_start")
                used_material, note = tz.set_material(material)
                part_result["materialUsedInTecZone"] = used_material
                if note:
                    part_result["notes"] += note
                screenshotter.snap("material_done")

                step = "EXPORT_GEO"
                last_step = step
                wait_if_paused(step)
                set_overlay(part_index, step, part_name)
                screenshotter.snap("export_start")
                export_name_template = settings.get("exportNameTemplate", "<partName>.geo")
                export_name = export_name_template.replace("<partName>", part_name)
                export_path = str(Path(export_dir) / export_name)
                last_export_path = export_path
                tz.export_geo(export_path)
                part_result["geoPath"] = export_path
                screenshotter.snap("export_done")
                part_result["thicknessMmDetected"] = tz.get_thickness_mm()

            except NeedsHelpError as e:
                part_result["status"] = "NEEDS_HELP"
                part_result["notes"] += str(e)
                overall_status = "NEEDS_HELP"
                screenshotter.snap("needs_help")
                needs_help_path = str(Path(project_root) / "WORK" / "logs" / f"{job_id}_NEEDS_HELP.txt")
                write_needs_help(needs_help_path, f"{step} {part_name}", str(e))
                capture_glitchtip_event(
                    "error",
                    "worker NEEDS_HELP",
                    job_id=job_id,
                    xometry_ref=xometry_ref,
                    status="NEEDS_HELP",
                    step=step,
                    part_id=part_id,
                    project_root=project_root,
                    input_path=input_path,
                    export_path=last_export_path,
                    log_path=log_path,
                    screenshots_dir=str(screenshots_dir),
                    reason=str(e),
                )
                result["parts"].append(part_result)
                break
            except Exception as e:
                part_result["status"] = "FAILED"
                part_result["notes"] += f"Exception: {e}"
                logger.error("Exception on part %s: %s", part_id, traceback.format_exc())
                capture_glitchtip_event(
                    "error",
                    "worker FAILED on part",
                    job_id=job_id,
                    xometry_ref=xometry_ref,
                    status="FAILED",
                    step=step or last_step,
                    part_id=part_id,
                    project_root=project_root,
                    input_path=input_path,
                    export_path=last_export_path,
                    log_path=log_path,
                    screenshots_dir=str(screenshots_dir),
                    reason=str(e),
                    exc=e,
                )
                screenshotter.snap("failed")
                if overall_status == "DONE":
                    overall_status = "PARTIAL"
            finally:
                try:
                    if opened_document:
                        wait_if_paused("CLOSE_FILE")
                        set_overlay(part_index, "CLOSE_FILE", part_name)
                        tz.close_active_file()
                        screenshotter.snap("close_file_done")
                except Exception as e:
                    logger.warning("Failed to close active file with Ctrl+W: %s", e)

            result["parts"].append(part_result)
    finally:
        stop_event.set()
        pause_controller.stop()
        if overlay is not None:
            overlay.stop()

    if overall_status == "DONE" and any(p["status"] == "FAILED" for p in result["parts"]):
        overall_status = "PARTIAL" if any(p["status"] == "DONE" for p in result["parts"]) else "FAILED"

    result["status"] = overall_status
    result_path = str(log_dir / f"{job_id}.result.json")
    write_json(result_path, result)
    write_json(str(log_dir / "result.json"), result)
    if overall_status == "DONE":
        capture_glitchtip_event(
            "info",
            "worker DONE",
            job_id=job_id,
            xometry_ref=xometry_ref,
            status="DONE",
            step=last_step,
            part_id=last_part_id,
            project_root=project_root,
            input_path=last_input_path,
            export_path=last_export_path,
            log_path=log_path,
            screenshots_dir=str(screenshots_dir),
            reason="job completed",
        )
    elif overall_status in ["NEEDS_HELP", "FAILED", "PARTIAL"]:
        capture_glitchtip_event(
            "error",
            f"worker {overall_status}",
            job_id=job_id,
            xometry_ref=xometry_ref,
            status=overall_status,
            step=last_step,
            part_id=last_part_id,
            project_root=project_root,
            input_path=last_input_path,
            export_path=last_export_path,
            log_path=log_path,
            screenshots_dir=str(screenshots_dir),
            reason=f"job ended with status={overall_status}",
        )
    if not disable_sounds and not settings.get("disableSounds", False):
        play_job_end_sound(logger, overall_status)
    return result_path, overall_status


def run_loop(
    jobs_dir,
    hotkey_pause="ctrl+alt+p",
    disable_hotkeys=False,
    disable_sounds=False,
    once=False,
    no_overlay=False,
    teczone_exe=None,
    teczone_title_re=None,
):
    jobs_dir = Path(jobs_dir)
    state_dir = jobs_dir.parent / "state"
    processed_any = False

    while True:
        for job_path in sorted(jobs_dir.glob("*.json")):
            try:
                job = read_json(job_path)
                job_id = job.get("jobId") or job_path.stem
            except Exception:
                continue

            ok, marker = claim_job(job_id, state_dir)
            if not ok:
                continue

            processed_any = True
            status = "FAILED"
            try:
                _, status = process_job(
                    job_path,
                    hotkey_pause=hotkey_pause,
                    disable_hotkeys=disable_hotkeys,
                    disable_sounds=disable_sounds,
                    no_overlay=no_overlay,
                    teczone_exe=teczone_exe,
                    teczone_title_re=teczone_title_re,
                )
            except Exception as e:
                capture_glitchtip_event(
                    "error",
                    "worker FAILED in run_loop",
                    job_id=job_id,
                    xometry_ref=job.get("xometryRef"),
                    status="FAILED",
                    step="RUN_LOOP",
                    part_id=None,
                    project_root=job.get("projectRoot"),
                    input_path=None,
                    export_path=None,
                    log_path=None,
                    screenshots_dir=None,
                    reason=str(e),
                    exc=e,
                )
                status = "FAILED"
            finally:
                release_job(marker, status)

            if once:
                return processed_any

        if once:
            return processed_any
        time.sleep(DEFAULT_POLL_SECONDS)


def main():
    parser = argparse.ArgumentParser()
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument("--jobs-dir", help="Path to WORK\\jobs")
    input_group.add_argument("--project-root", help="Path to project root (uses <projectRoot>\\WORK\\jobs)")
    parser.add_argument("--once", action="store_true", help="Process only one available job and exit")
    parser.add_argument("--hotkey-pause", default="ctrl+alt+p", help="Global pause hotkey")
    parser.add_argument("--disable-hotkeys", action="store_true", help="Disable global hotkeys")
    parser.add_argument("--disable-sounds", action="store_true", help="Disable start/end sounds")
    parser.add_argument("--no-overlay", action="store_true", help="Disable Tk overlay")
    parser.add_argument("--teczone-exe", help="Explicit path to Flux.exe for auto-start")
    parser.add_argument("--teczone-title-re", help="Regex for TecZone main window title matching")
    parser.add_argument("--glitchtip-test", action="store_true", help="Send test event to GlitchTip and exit")
    args = parser.parse_args()

    init_glitchtip()
    if args.glitchtip_test:
        capture_glitchtip_event(
            "info",
            "dorina glitchtip test",
            status="TEST",
            step="GLITCHTIP_TEST",
            reason="manual test event",
        )
        print("GlitchTip test event sent (if DSN configured).")
        return

    jobs_dir = args.jobs_dir or str(Path(args.project_root) / "WORK" / "jobs")
    try:
        run_loop(
            jobs_dir,
            hotkey_pause=args.hotkey_pause,
            disable_hotkeys=args.disable_hotkeys,
            disable_sounds=args.disable_sounds,
            once=args.once,
            no_overlay=args.no_overlay,
            teczone_exe=args.teczone_exe,
            teczone_title_re=args.teczone_title_re,
        )
    except Exception as e:
        capture_glitchtip_event(
            "error",
            "worker uncaught exception",
            status="FAILED",
            step="MAIN",
            reason=str(e),
            exc=e,
        )
        raise


if __name__ == "__main__":
    main()
