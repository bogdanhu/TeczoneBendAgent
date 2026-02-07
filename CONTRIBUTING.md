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

## 5. Checkpoint Tags (Rollback)
When you reach a version that works on Dorina (for example: passes `OPEN_FILE` and exports at least one `.geo`), create a checkpoint tag:

```bash
git tag -a v0.1.0 -m "working: open+export+pause"
git push --tags
```

Versioning rules:
- `PATCH` (`0.1.X`): bugfixes
- `MINOR` (`0.X.0`): backward-compatible features
- `MAJOR` (`X.0.0`): breaking changes (job schema, output schema, CLI)

## Compatibility Rule
Do not change `job.json` schema unless explicitly requested.
