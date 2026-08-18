# Employment Center — Client Intake & Check-In

A simple, touch-friendly local web app that replaces manual sign-in/intake entry
while continuing to populate the existing monthly Excel reporting workbook.

## What it does

- **Kiosk check-in** clients use themselves. They enter First Name, Last Name,
  and Date of Birth. The app then decides:
  - **New client** → full monthly intake.
  - **Existing client, no intake this calendar month** → *Monthly Intake
    Required*, prefilled from their most recent intake to review/update.
  - **Existing client, intake already done this month** → *Welcome back* → one
    big **CHECK IN** button.
- **Every physical visit** (including the first) creates a separate Visit record.
- **One Monthly Intake per client per calendar month** — enforced in the DB on
  `client_id + reporting_year + reporting_month`. Historical months are never
  altered when a client updates info in a later month.
- **Excel export** into the existing workbook, preserving its exact column
  structure and `1`/blank coding (see below).
- **Admin dashboard** for staff: daily/monthly stats, client search & history,
  correcting data-entry mistakes, and exporting to Excel — including a
  **drag-and-drop** uploader: drop the month's workbook, the app writes the
  data into the correct sheets and hands it back to download.

## Excel coding (discovered from the workbook, not assumed)

Inspecting `Feb. Numbers Final(1).xlsx` showed:

| Concept | Representation in the workbook |
|---|---|
| Demographic / category columns (Veteran, Hispanic, race, household type, population, Female Head, Disabled) | `1` when it applies, **blank** otherwise |
| House Size | actual integer |
| Large numbers at the bottom of a sheet | `=SUM()` **totals rows**, not client data |
| Race / Household type / Population | single-select (exactly one `1` per client) |

There is **no** `Yes=1 / No=2` scheme. Header spellings are preserved verbatim
(`Veteren`, `Two Parent ` with a trailing space, `American India`,
`Fam. With Minor`, `Female Head of`). The canonical new-month layout follows the
workbook's newest, cleanest sheet, **June Complete** (the only one with a UID
column). All mapping lives in [`app/mapping.py`](app/mapping.py) as
`form answer → reporting category → Excel column/value`.

## How Excel writes stay duplicate-free

The database is the source of truth. For the current reporting month the app
owns two sheets — `"<Month> <Year>"` and `"<Month> <Year> Sign Ins"` — and
**rebuilds them from the database** on each sync. Rebuilding is idempotent, so
it can never create duplicate intake rows or duplicate visits, and it leaves
every historical sheet completely untouched. A timestamped backup is written to
`backups/` before every save.

## Run it

```bash
pip install -r requirements.txt
python run.py
# Kiosk:  http://127.0.0.1:5000/
# Admin:  http://127.0.0.1:5000/admin
```

The app keeps a working copy of the workbook at `data/reporting.xlsx` and the
SQLite database at `data/app.db`.

## Tests

```bash
python -m pytest -q
```

Covers the six required scenarios: new client full intake, same client returning
in the same month (visit only, no duplicate demographic row), same client next
month (intake required, prefilled, history unchanged), exact demographic coding,
double-click duplicate protection, and workbook backup.

## Project layout

```
app/
  main.py                 Flask routes (kiosk + admin)
  database.py             SQLite schema & connection
  models.py               normalization / parsing helpers
  mapping.py              form → reporting category → Excel column/value
  services/
    client_service.py     identity: match / create / update / search
    intake_service.py     one monthly intake per client per month
    checkin_service.py     visit records (+ double-click dedupe)
    excel_service.py      backup, sheet resolution, coded writes
  templates/  static/     kiosk + admin UI
data/    app.db, reporting.xlsx, uploads/
backups/  timestamped workbook backups
tests/   test_workflow.py
```
