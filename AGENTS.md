# AGENTS.md

## Cursor Cloud specific instructions

### Branch note

The runnable **Lifecycle Tracker** application lives on branch `cursor/lifecycle-tracker-mvp-6da3`. The default `main` branch currently holds Cortave workbook assets only (no `app.py` or `requirements.txt`). Check out the MVP branch before developing or running the app unless `main` has been merged.

### Services

| Service | Required | Notes |
|---------|----------|--------|
| Flask web app (`python app.py`) | Yes | Single process on port **5000** (`PORT` env var) |
| SQLite | Yes (embedded) | File `lifecycle_tracker.db`; override with `LIFECYCLE_DB` |

There is no separate database server, frontend build, Docker Compose, or worker process.

### One-time VM prerequisites

Ubuntu images may lack `python3-venv`. If `python3 -m venv .venv` fails, install once (outside the update script):

```bash
sudo apt-get install -y python3.12-venv
```

### Dependencies and tests

See `README.md` for the canonical local run flow. Quick reference:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m unittest discover -s tests -v
python app.py
```

Open http://localhost:5000. The app auto-seeds SQLite and assigns a default Admin session (MVP role switcher on `/login`).

### Lint

No linter or formatter is configured in this repository. Use `python -m unittest discover -s tests -v` as the primary quality gate.

### Non-obvious runtime notes

- **Debug mode:** set `FLASK_DEBUG=1` when running `python app.py`.
- **Tests:** use Flask `test_client` with a temp DB; a live server is not required for unit tests.
- **Hello-world flow:** Master Dashboard → create or open an Innovator/Direct Customer Account Plan → confirm lifecycle items in the plan workspace (`/plans/<id>`).
- Optional Excel fixtures (`Lifecycle Workbook.xlsx`, `Partner Dashboard .xlsx`) are only on `main` for import E2E; not required for unit tests.
