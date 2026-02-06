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
from teczone_actions import TecZoneSession, NeedsHelpError
from ui_utils import get_active_window_title
from xometry_parser import load_xometry_map

DEFAULT_POLL_SECONDS = 2


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
    with open(path, "w", encoding="utf-8") as f:
        f.write("NEEDS_HELP\n")
        f.write(f"active_window: {get_active_window_title()}\n")
        f.write(f"step: {step}\n")
        f.write(f"found: {found}\n")


def process_job(job_path):
    job = read_json(job_path)
    job_id = job["jobId"]
    project_root = job["projectRoot"]
    settings = job.get("settings", {})

    log_dir = Path(project_root) / "WORK" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = str(log_dir / f"{job_id}.log")
    logger = configure_logger(log_path)

    screenshots_dir = Path(project_root) / "WORK" / "screenshots" / job_id
    screenshots_dir.mkdir(parents=True, exist_ok=True)

    export_dir = settings.get("exportDir") or str(Path(project_root) / "WORK" / "out" / "flat")
    Path(export_dir).mkdir(parents=True, exist_ok=True)

    xometry_map = {}
    if job.get("xometryJson"):
        try:
            xometry_map = load_xometry_map(job["xometryJson"], logger)
        except Exception as e:
            logger.warning("Failed to parse xometry json: %s", e)

    overlay = Overlay(f"WORKER RUNNING: {job_id} / START")
    overlay.start()

    screenshotter = Screenshotter(str(screenshots_dir))

    periodic_seconds = settings.get("screenshotsEverySeconds", 0)
    stop_event = threading.Event()
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
    try:
        tz = TecZoneSession(logger, screenshotter)
        tz.connect()

        for part in job.get("inputFiles", []):
            part_id = part.get("partId")
            part_name = part.get("partName")
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
            try:
                step = f"OPEN_FILE {part_name}"
                overlay.set_text(f"WORKER RUNNING: {job_id} / {step}")
                screenshotter.snap("open_file_start")
                tz.open_file(input_path)
                screenshotter.snap("open_file_done")

                step = f"SET_MATERIAL {part_name}"
                overlay.set_text(f"WORKER RUNNING: {job_id} / {step}")
                screenshotter.snap("material_start")
                used_material, note = tz.set_material(material)
                part_result["materialUsedInTecZone"] = used_material
                if note:
                    part_result["notes"] += note
                screenshotter.snap("material_done")

                step = f"EXPORT_GEO {part_name}"
                overlay.set_text(f"WORKER RUNNING: {job_id} / {step}")
                screenshotter.snap("export_start")
                export_name_template = settings.get("exportNameTemplate", "<partName>.geo")
                export_name = export_name_template.replace("<partName>", part_name)
                export_path = str(Path(export_dir) / export_name)
                tz.export_geo(export_path)
                part_result["geoPath"] = export_path
                screenshotter.snap("export_done")

                thickness = tz.get_thickness_mm()
                part_result["thicknessMmDetected"] = thickness

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
        overlay.stop()

    if overall_status == "DONE":
        if any(p["status"] == "FAILED" for p in result["parts"]):
            overall_status = "PARTIAL" if any(p["status"] == "DONE" for p in result["parts"]) else "FAILED"

    result["status"] = overall_status
    logs_dir = Path(project_root) / "WORK" / "logs"
    result_path = str(logs_dir / f"{job_id}.result.json")
    write_json(result_path, result)
    write_json(str(logs_dir / "result.json"), result)

    return result_path, overall_status


def run_loop(jobs_dir):
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
                _, status = process_job(job_path)
            except Exception:
                status = "FAILED"
            finally:
                release_job(marker, status)

        time.sleep(DEFAULT_POLL_SECONDS)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--jobs-dir", required=True, help="Path to WORK\\jobs")
    args = parser.parse_args()
    run_loop(args.jobs_dir)


if __name__ == "__main__":
    main()
