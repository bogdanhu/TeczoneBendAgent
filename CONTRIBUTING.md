# Contributing

Minimal workflow for fast but safe development on `main`.

## 1. Before Editing
Run:

```bash
git status
```

Working tree must be clean. If not clean, commit or stash first.

Run:

```bash
git pull --rebase
```

Keep local `main` up to date before any edits.

## 2. Make Changes
- Keep commits small and atomic.
- Use clear commit messages, for example: `Fix open_file dialog handling`.

## 3. Sanity Check Before Push
Run:

```bash
python -m compileall worker
```

If you add dependencies, update:
- `worker/requirements.txt`

## 4. Push
Run `git push` only after sanity checks pass.

## 5. Working Checkpoint Tags (Rollback)
Create annotated `working` tags only after a real Dorina test confirms all:
- `OPEN_FILE` passes
- at least one `.geo` is exported
- `WORK\logs\<jobId>.result.json` is written

First stable milestone tag:
- `v0.1.0`

Then version by impact:
- `PATCH` (`0.1.X`): bugfixes
- `MINOR` (`0.X.0`): backward-compatible features
- `MAJOR` (`X.0.0`): breaking changes (job schema, output schema, CLI)

Commands to create and push a working tag:

```bash
git status
git tag -a v0.1.0 -m "working: open+export+pause"
git push --tags
```

List tags (newest first):

```bash
git fetch --tags
git tag --sort=-creatordate
```

Quick rollback to a tag:

```bash
git fetch --tags
git switch -c rollback/v0.1.0 v0.1.0
```

Important:
- Do not create tags automatically after every commit.
- Create tags only after real Dorina validation.
- Do not change `job.json` schema unless explicitly requested.
