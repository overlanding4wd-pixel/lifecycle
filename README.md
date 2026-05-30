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

## Account Plan terminology and lifecycle templates

The app now treats the workbook journeys as editable lifecycle templates rather than spreadsheet tabs:

- **Account**: either an Innovator/partner or a Direct Customer
- **Account Type**: `Innovator` or `Direct Customer`
- **Account Plan**: the project plan for an account
- **Lifecycle Template**: default journey template used to create a plan
- **Lifecycle Item**: an individual task/activity/step copied into an Account Plan

Default templates are seeded in the database:

- **Innovator Journey**: I0, I20, and I50 stages from the Lifecycle Workbook `I20 Tracker` sheet
- **Direct Customer Journey**: D0 and D20 stages from the Lifecycle Workbook `D20 Tracker` sheet

When an Account Plan is created, the selected account type determines the template. Template item offsets calculate dates from the kick-off date:

- `actualStartDate = kickOffDate + startDayOffset`
- `dueDate = kickOffDate + targetDueDayOffset`

Admins can view and edit lifecycle template items from Data Validation / Settings.


### Refined Master Dashboard workflow

The landing page is the portfolio-level **Master Dashboard**. It focuses on active Innovator and Direct Customer plans, with summary cards for active plans, Innovators, Direct Customers, In Progress, On Hold, Overdue, Live, and Qualified Out.

From the Master Dashboard users can:

- Create an Innovator Plan from the default Innovator Journey / I20 lifecycle template
- Create a Direct Customer Plan from the default Direct Customer Journey / D20 lifecycle template
- Filter and sort active plans by type, stage, health/status, owner, overdue state, live state, next due date, days overdue, stage, owner, and type
- Open an Account Plan workspace to view, edit, add, or delete Lifecycle Items

Current Stage is calculated from the next active/incomplete Lifecycle Item unless manually overridden on the Account Plan. Next Step is the nearest-due incomplete Lifecycle Item. Days Overdue is calculated from incomplete items with due dates before today. Live is true when the final live Lifecycle Item is completed.

## Master Partner Dashboard integration

The app includes an add-on module for `Partner Dashboard.xlsx`:

- New `Partner` master data table populated from the workbook's `Partners` sheet
- Import mapping for Partner, Owner, Territory, Active stage, Date Of Stage Change, Age of Stage, Days Overdue, Workbook Link, SF Account, Next Steps/notes, Partner Type, and CSM Involved
- Org Impact sheet import for parameters, base assumptions, estimated averages, and summary impact metrics
- Partner imports upsert by partner name so re-importing refreshes master data without touching I20/D20 tracker records
- A `lifecycle_workbooks` link table associates this Lifecycle Tracker workspace with one imported partner now and can support more workbooks later
- The Overview and tracker pages show a Linked Master Partner panel with search, link, unlink, and live partner summary details
- The Master Partner Dashboard page provides searchable/filterable partner records, partner summary cards, lifecycle link status, and an Org Impact section
- Import / Export includes Master Partner Dashboard re-import controls and import history

## Production hardening ideas

- Replace the MVP role switcher with SSO or your identity provider
- Add per-field audit detail views and record comments
- Add saved dashboard filters and cross-tracker drill-down views
- Add migration tooling if the schema evolves beyond the MVP
