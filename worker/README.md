# TecZone Bend Worker (Windows)

## What this does
This worker watches `WORK\jobs` for `job.json` files, drives TecZone Bend via UI Automation, exports flat `.geo`, and writes `result.json` plus screenshots and logs. It stops with `NEEDS_HELP` when UI elements are missing or unexpected dialogs appear.

## Prereqs
- Windows 11
- Python 3.10+
- TecZone Bend installed and already running
- Display scaling 100%, fixed resolution
- TecZone UI language kept consistent (English recommended)

## Install
```powershell
cd worker
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run
```powershell
.\.venv\Scripts\Activate.ps1
python worker.py --jobs-dir X:\\33259_TEST_OC_20260206-210632\\WORK\\jobs
```

Optional:
```powershell
python worker.py --jobs-dir X:\\33259_TEST_OC_20260206-210632\\WORK\\jobs --disable-sounds
```

## Example job (realistic paths)
```json
{
  "jobId": "33259_J-1797130-318293_teczone",
  "projectRoot": "X:\\33259_TEST_OC_20260206-210632",
  "docDir": "X:\\33259_TEST_OC_20260206-210632\\DOC",
  "xometryRef": "J-1797130-318293",
  "xometryJson": "X:\\33259_TEST_OC_20260206-210632\\DOC\\J-1797130-318293.xometry.json",
  "inputFiles": [
    {
      "path": "X:\\33259_TEST_OC_20260206-210632\\DOC\\part735256\\Main.STEP",
      "partId": 735256,
      "partName": "part735256_Main"
    }
  ],
  "settings": {
    "exportDir": "X:\\33259_TEST_OC_20260206-210632\\WORK\\out\\flat",
    "exportNameTemplate": "<partName>.geo",
    "screenshotsEverySeconds": 10,
    "dryRun": false
  }
}
```

## Quick test today (no UI automation yet)
1. Set `dryRun` to `true` in your job.
2. Run:
```powershell
.\.venv\Scripts\Activate.ps1
python worker.py --jobs-dir X:\\33259_TEST_OC_20260206-210632\\WORK\\jobs
```
3. The worker will:
- Find the TecZone window (or launch it if `TECZONE_EXE` is set).
- Parse the `xometry.json` (fails to NEEDS_HELP if parsing fails).
- Write a screenshot and logs.

## Notes
- If a control/menu is not found, the worker writes `NEEDS_HELP` and stops. Check screenshots in `WORK\screenshots\<jobId>` and the log at `WORK\logs\<jobId>.log`.
- `result.json` is written to `WORK\logs\<jobId>.result.json`.
- Update UI selectors in `worker\teczone_actions.py` if your TecZone build uses different menu names.
- Input files must be `.stp` or `.step` (case-insensitive) and are read from `job.json`.
- Persistent behavior specs are stored in `worker\WORKER_SPEC.md`.
- Worker plays a short sound at job start and job end (can be disabled with `--disable-sounds` or job setting `disableSounds`).
