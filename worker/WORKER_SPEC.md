# Worker Specs (persistent)

## Overlay UX
- Overlay must always show:
  - current action
  - next action
  - steps done/total
  - pause/resume hint (`ctrl+alt+p` by default)
- Format (2 lines):
  - `WORKER RUNNING|PAUSED: <jobId> | steps <done>/<total> | current: <action>`
  - `next: <action> | hint: <hotkey> = pause/resume`
- While paused, overlay state changes to `PAUSED` and one `paused` screenshot is taken (no spam).

## Step model
- Each part has 3 major steps:
  - `OPEN_FILE`
  - `SET_MATERIAL`
  - `EXPORT_GEO`
- Total steps = `parts_count * 3`.
- `next` action is computed from this sequence.

## Hotkeys
- Default global hotkey: `ctrl+alt+p`.
- CLI:
  - `--hotkey-pause "ctrl+alt+p"`
  - `--disable-hotkeys`
- If hotkeys are disabled, overlay hint shows `hotkeys disabled`.

## Sounds
- Worker plays a short distinctive sound when a job starts.
- Worker plays a short distinctive sound when a job ends.
- End sound differs by status (`DONE`, `PARTIAL`, `NEEDS_HELP/FAILED`).
- CLI option:
  - `--disable-sounds`
- Job setting:
  - `"disableSounds": true`

## TecZone availability
- Worker tries to connect to TecZone main window.
- If not found and `TECZONE_EXE` env var exists, worker launches TecZone and retries.
- If still not available, worker writes:
  - `WORK\logs\<jobId>_NEEDS_HELP.txt`
  - `WORK\logs\windows.json`
  - screenshot in `WORK\screenshots\<jobId>`
  - result with `status = NEEDS_HELP`

## Safety
- No random coordinate clicks.
- Unknown dialogs during open/export flow trigger `NeedsHelpError`.
