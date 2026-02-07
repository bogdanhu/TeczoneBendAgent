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

Disable overlay (useful if Tk crashes in your environment):
```powershell
python worker.py --jobs-dir X:\\33259_TEST_OC_20260206-210632\\WORK\\jobs --no-overlay
```

Run once from project root:
```powershell
python worker.py --project-root X:\\33259_TEST_OC_20260206-210632 --once
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

## TEST QUICK
1. Setup env and deps:
```powershell
cd worker
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```
2. Run once:
```powershell
python worker.py --project-root X:\\33259_TEST_OC_20260206-210632 --once
```
3. Verify:
- Worker trece de dialogul `Open`.
- Fișierul `.geo` apare în `WORK\out\flat`.
- Log și result JSON sunt create în `WORK\logs`.
4. Debug locations:
- Screenshots: `WORK\screenshots\<jobId>\`
- Log: `WORK\logs\<jobId>.log`
- Needs help: `WORK\logs\<jobId>_NEEDS_HELP.txt`
- Window dump: `WORK\logs\windows.json`

## Notes
- If a control/menu is not found, the worker writes `NEEDS_HELP` and stops. Check screenshots in `WORK\screenshots\<jobId>` and the log at `WORK\logs\<jobId>.log`.
- `result.json` is written to `WORK\logs\<jobId>.result.json`.
- Update UI selectors in `worker\teczone_actions.py` if your TecZone build uses different menu names.
- Input files must be `.stp` or `.step` (case-insensitive) and are read from `job.json`.
- Persistent behavior specs are stored in `worker\WORKER_SPEC.md`.
- Worker plays a short sound at job start and job end (can be disabled with `--disable-sounds` or job setting `disableSounds`).
- Overlay format during run: `WORKER: <jobId> [i/n] <STEP> <partName>` and on pause `WORKER: <jobId> [paused] [i/n] ...`.

## How to rollback to last working tag
Create `working` tags only after real Dorina validation (`OPEN_FILE` passes, at least one `.geo` exported, and `WORK\logs\<jobId>.result.json` exists).

1. Fetch latest tags:
```powershell
git fetch --tags
```
2. See newest tags:
```powershell
git tag --sort=-creatordate
```
3. Create a rollback branch from a known working tag (example `v0.1.0`):
```powershell
git switch -c rollback/v0.1.0 v0.1.0
```
4. Run worker from that branch and validate behavior before any merge back to `main`.
