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

## Branding

The UI follows the `Logo Ref for Team-2.pptx` guide from the repository:

- Logo: extracted Cortave logo asset used in the application shell
- Font stack: `Neue Haas Grotesk Text Pro`, `New Haas Grotesk Text Pro`, then system sans-serif fallbacks
- Primary blue: RGB 0, 74, 136 / `#004A88`
- Brand grey: RGB 152, 162, 170 / `#98A2AA`
- Positive green: RGB 120, 190, 32 / `#78BE20`

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
- Stage: I20, I50, D20, D50, Qualified Out, I0, D0
- Owners from the Lifecycle Workbook validation sheet, including Wade / Sales, Melissa / Marketing, Mark / Innovator, Bobby / Experience, Nick / Sales, Customer, Legal, Finance, Innovator, and Innovator / cortave
- Priority: Low, Medium, High, Critical

## Import notes

CSV and Excel imports map common workbook headers such as `Opportunity Name`, `Account Name`, `Owner`, `Stage`, `Status`, `Due Date`, and `Next Action` to app fields. The uploaded Lifecycle Workbook is also supported directly: when importing an `.xlsx`, the app selects the matching `I20 Tracker` or `D20 Tracker` sheet and maps workbook columns including `Activity`, `who @ cortave`, `Start Day`, `Actual Start Date`, `Target Due Day`, `cortave Owner`, `Innovator Owner`, `Link`, and `Notes`. Unrecognised workbook columns are preserved in the record's additional imported fields JSON so workbook-specific columns are not lost.

## Master Partner Dashboard integration

The app includes an add-on module for `Partner Dashboard.xlsx`:

- New `Partner` master data table populated from the workbook's `Partners` sheet
- Import mapping for Partner, Owner, Territory, Active stage, Date Of Stage Change, Age of Stage, Days Overdue, Workbook Link, SF Account, Next Steps/notes, Partner Type, and CSM Involved
- Partner imports upsert by partner name so re-importing refreshes master data without touching I20/D20 tracker records
- A `lifecycle_workbooks` link table associates this Lifecycle Tracker workspace with one imported partner now and can support more workbooks later
- The Overview and tracker pages show a Linked Master Partner panel with search, link, unlink, and live partner summary details
- The Master Partner Dashboard page provides searchable/filterable partner records, summary cards, and lifecycle link status
- Import / Export includes Master Partner Dashboard re-import controls and import history

## Production hardening ideas

- Replace the MVP role switcher with SSO or your identity provider
- Add per-field audit detail views and record comments
- Add saved dashboard filters and cross-tracker drill-down views
- Add migration tooling if the schema evolves beyond the MVP
