# Lifecycle Tracker

Lifecycle Tracker is a Flask + SQLite web application that replaces the Lifecycle Workbook with a browser-based internal business application.

## MVP features

- Overview dashboard with lifecycle totals, status and stage summaries, overdue work, due-this-week work, and recent updates
- Separate I20 and D20 tracker pages
- Add, edit, delete, search, sort, and filter tracker records
- Controlled Status and Stage dropdowns backed by application settings
- Professional status badges replacing Excel conditional formatting
- CSV and Excel import/export
- Detail drawer with notes, dates, owners, next action, additional imported workbook fields, and change history
- Settings page for statuses, stages, owners, priorities, and status colours
- Simple Admin, Editor, and Viewer role preview

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000.

The app creates `lifecycle_tracker.db` automatically on first request and seeds the required workbook dropdown values:

- Status: Not Started, In Progress, On Hold, Completed
- Stage: I20, I50, D20, Qualified Out, I0, D0
- Priority: Low, Medium, High, Critical

## Import notes

CSV and Excel imports map common workbook headers such as `Opportunity Name`, `Account Name`, `Owner`, `Stage`, `Status`, `Due Date`, and `Next Action` to app fields. Unrecognised workbook columns are preserved in the record's additional imported fields JSON so workbook-specific columns are not lost.

## Production hardening ideas

- Replace the MVP role switcher with SSO or your identity provider
- Add per-field audit detail views and record comments
- Add saved dashboard filters and cross-tracker drill-down views
- Add migration tooling if the schema evolves beyond the MVP
