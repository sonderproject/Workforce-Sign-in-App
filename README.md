# Employment Center — Sign-In Sheet

A dead-simple, touch-friendly sign-in page. **Nothing is stored on any server.**
All data lives in your Excel workbook, which is read and written entirely in the
browser and kept on the device so sign-ins continue through the day.

## How it works

1. **Staff** opens the page and loads the month's Excel workbook (drag & drop).
   It stays on the device — it is never uploaded anywhere.
2. A client enters **First Name, Last Name, Date of Birth**:
   - **New client** → answers the sign-up questions once. Their answers are
     written to the **Intakes** sheet (using the workbook's exact demographic
     columns and `1`/blank coding), and today's visit is logged.
   - **Returning client** → sees *Welcome back*, ticks **"Yes, I'm here today"**,
     and taps **Sign In**. Only a visit row is added — no repeat questions.
3. Every visit is appended to the **Sign Ins** sheet (Date, Time, Name, DOB,
   Visitor Type).
4. Staff clicks **Download updated workbook** anytime to save it back to their
   files.

A client counts as "signed up" once they have a row in the **Intakes** sheet.
Matching is case- and whitespace-insensitive on name + date of birth.

## Excel coding (taken from the existing workbook, not assumed)

- Demographic/category columns are `1` when they apply and **blank** otherwise
  (there is no Yes=1/No=2 scheme).
- `House Size` is a real integer.
- A `TOTALS` row with `SUM()` formulas is written under the data.
- Header spellings are preserved verbatim (`Veteren`, `Two Parent ` with a
  trailing space, `American India`, `Fam. With Minor`, `Female Head of`).
- All existing/historical sheets in the workbook are preserved untouched.

## Deploying to Vercel

This is a **static site** — no server, no database, no build step.

- **Framework Preset: `Other`** (not Flask). Build Command: none.
  Output Directory: `.` (repo root). A `vercel.json` sets this for you.
- Just import the repo and deploy. That's it.

## Run locally

Any static file server works, e.g.:

```bash
python3 -m http.server 8000
# open http://localhost:8000
```

## Files

```
index.html            all screens (kiosk + staff)
assets/style.css      styling
assets/app.js         all logic (reads/writes the workbook in-browser)
vendor/xlsx.full.min.js  SheetJS (bundled — no CDN, works offline)
vercel.json           static-site config
```

## Notes & limits

- Data persists on the device via the browser (IndexedDB). Use **Download
  updated workbook** to keep a real copy; **Clear this device** wipes the local
  copy. Because it's per-device, run sign-in on one kiosk device (or download/
  re-load the workbook to move between devices).
- The bundled spreadsheet library preserves cell **values and formulas** on all
  sheets, but heavy cell **styling** (colors/borders) on historical sheets may
  not survive a round-trip. The reporting structure, headers, coding, and totals
  are preserved.
- Sign-up questions are asked once per client (not re-collected monthly). If you
  later want the monthly "Numbers" tabs auto-populated per month, that's a small
  addition.
