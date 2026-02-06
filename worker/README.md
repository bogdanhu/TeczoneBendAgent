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

## Notes
- If a control/menu is not found, the worker writes `NEEDS_HELP` and stops. Check screenshots in `WORK\screenshots\<jobId>` and the log at `WORK\logs\<jobId>.log`.
- `result.json` is written to `WORK\logs\<jobId>.result.json`.
- Update UI selectors in `worker\teczone_actions.py` if your TecZone build uses different menu names.
- Input files must be `.stp` or `.step` (case-insensitive) and are read from `job.json`.
