# Staff Attendance System

An employee attendance and work-hour tracking system built on **Google Apps Script** and **Google Sheets**. Handles clock-in/out recording, automatic overtime calculation, meal allowances, and monthly payroll summaries.

Employees can log attendance directly via the Google Sheet menu or through a browser-based **Web App**. Rows are automatically locked 30 minutes after the recorded departure time to prevent unauthorized edits.

---

## Features

- **Daily auto-append** — new attendance rows for all active staff are created every day at 06:00
- **Clock-in/out stamping** — staff record times via menu or Web App; email validation ensures each person can only edit their own row
- **Automatic hour calculations** — effective hours, regular hours, OT Tier 1 & OT Tier 2 computed via spreadsheet formulas
- **Sunday & public holiday support** — DOUBLE, SWAP, and HALF DAY SUNDAY modes
- **Meal allowances** — automatically determined based on overtime hours and day type
- **Payroll recap** — per-staff summary using SUMIFS formulas linked directly to the source data sheet
- **Auto-lock** — rows lock 30 minutes after departure via a time-based trigger
- **Evening reminder** — notifies HR if any staff have not logged a departure time (default: 17:00)
- **Per-row protection** — NOTE and SUNDAY/RED DAY columns are admin-only

---

## Architecture

```
absent-staff/
├── Config.js          # Global config (sheet names, column indices, default settings)
├── Utils.js           # Helpers: time parsing, date math, admin email detection
├── Triggers.js        # onEdit(), onOpen(), setupTrigger(), setupAwal()
├── Setup.js           # Sheet creation, protections, dropdown validation, settings spreadsheet
├── Append.js          # appendHariIni(), full-month generation, row grouping
├── Stamp.js           # Staff actions: navigate to row, record times, view personal recap
├── Lock.js            # Evening reminder (cekBelumIsiPulang()), auto-lock rows
├── Rekap.js           # Payroll recap: hitungRekap(), generateTemplateRekap(), buatSheetRentang()
├── WebApp.html        # Web App UI (mobile-first HTML + CSS + JS)
├── WebAppServer.js    # doGet(), getDataHariIni(), submitAbsensi()
├── appsscript.json    # Manifest: timezone, OAuth scopes, Web App execution config
├── .clasp.json        # clasp config (target script ID)
└── deploy.sh          # Push codebase to multiple spreadsheets at once
```

### Data Flow

```
06:00 → appendHariIni()      → new row per staff → formulas L–O → per-email protection
Staff → onEdit / Web App     → email validation  → write time   → formula auto-recalculates
17:00 → cekBelumIsiPulang()  → scan today's rows → notify HR if departure missing
Hourly → lockBarisWeb…()     → check 30-min threshold → lock row if passed
Admin  → generateTemplateRekap() → SUMIFS to source sheet → payroll summary
```

---

## Column Structure

| Column | Content | Access |
|--------|---------|--------|
| A | Date | Auto (locked) |
| B | Name | Auto (locked) |
| C | Email | Auto (locked) |
| D | Division | Auto (locked) |
| E | Status (Present / Sick / Leave / Absent) | Staff |
| F | Clock-in time | Staff |
| G | Break out 1 | Staff |
| H | Break in 1 | Staff |
| I | Break out 2 | Staff |
| J | Break in 2 | Staff |
| K | Clock-out time | Staff |
| L | Effective hours | Formula (locked) |
| M | Regular hours | Formula (locked) |
| N | OT Tier 1 | Formula (locked) |
| O | OT Tier 2 | Formula (locked) |
| P | Note | Admin only |
| Q | Day type (SUNDAY / RED DAY) | Admin only |
| R | Shift | Staff |
| S | Device | Staff |
| T | Absence reason | Staff |
| U | Late arrival reason | Staff |
| V | Early departure reason | Staff |

---

## Roles & Permissions

| Role | Capabilities |
|------|-------------|
| **Staff** | Edit own row (columns E–V), stamp via menu or Web App |
| **Admin / HR** | Edit all rows, access columns P & Q, run all menu functions |
| **Owner** | One-time initial setup, trigger registration, full access |

Access control is enforced in `onEdit()` — the active user's email is compared against column C of the edited row. Emails listed in `ADMIN_EMAILS` bypass this restriction.

---

## Configuration

### Structural Config (`Config.js`)

Changes here require a redeploy via `clasp`.

| Constant | Description |
|----------|-------------|
| `NAMA_INSTANSI` | Institution / company name |
| `SHEET_MASTER` | Name of the Master_Data sheet |
| `DIVISI` | Array of division names, e.g. `['WEB', 'HR']` |
| `TIMEZONE` | Timezone, default `Asia/Makassar` |
| `DAYS_HOUR` | Regular working hours per workday |

### Operational Config (Settings Spreadsheet)

Editable by admins at any time without redeploying — stored in a separate Settings spreadsheet.

| Key | Description | Default |
|-----|-------------|---------|
| `ADMIN_EMAILS` | Comma-separated list of admin emails | — |
| `JAM_REMINDER` | Evening reminder hour (0–23) | `17` |
| `SELISIH_MENIT_LOCK` | Minutes after departure before the row locks | `30` |
| `PLAN_JAM` | Shift options, comma-separated | `"08:00 - 17:00"` |

---

## Initial Setup

### Prerequisites

- A Google account with access to Google Sheets
- [clasp](https://github.com/google/clasp) installed: `npm install -g @google/clasp`
- Logged in to clasp: `clasp login`

### Installation Steps

1. **Create a new spreadsheet** in Google Sheets for each division.

2. **Open Apps Script** — Extensions → Apps Script inside that spreadsheet.

3. **Push the code** to the target spreadsheet:
   ```bash
   # Edit deploy.sh — set TARGETS to each division's script ID
   bash deploy.sh
   ```

4. **Run initial setup** — in the spreadsheet, open the menu:
   ```
   🔑 Owner → Setup Awal
   ```
   This will:
   - Create the Master_Data sheet and the current month's sheet
   - Register all time-based triggers
   - Create the external Settings spreadsheet
   - Apply admin-only column protections

5. **Populate Master_Data** — add active staff (name, email, division).

6. **Share the Web App URL** (optional) — Deploy → New deployment → Web App → Anyone with Google Account.

### Deploying to Multiple Divisions

Edit the `TARGETS` array in `deploy.sh`:

```bash
TARGETS=(
  "Script_ID_for_WEB_division"
  "Script_ID_for_HR_division"
)
```

Then run:

```bash
bash deploy.sh
```

The script automatically swaps `.clasp.json` and pushes the codebase to each target.

---

## Automated Triggers

| Trigger | Schedule | Function |
|---------|----------|----------|
| Daily append | Every day at 06:00 | `appendHariIni()` |
| Evening reminder | Every day at 17:00 (configurable) | `cekBelumIsiPulang()` |
| Auto-lock | Every hour | `lockBarisWebSudahPulang()` |

---

## Payroll Recap

Four recap methods are available via the **Admin → Rekap** menu:

| Function | Description |
|----------|-------------|
| `hitungRekap()` | On-demand calculation from source data |
| `buatSheetRekap()` | Static aggregation written to a Rekap sheet |
| `generateTemplateRekap()` | Dynamic SUMIFS formulas — auto-updates when source data changes |
| `buatSheetRentang()` | Extracts data for a specific date range via QUERY formula |

---

## Tech Stack

- **Google Apps Script** — backend logic (JavaScript)
- **Google Sheets** — primary data store
- **HTML / CSS / JS** — Web App UI (embedded in `.html`)
- **clasp** — Apps Script deployment CLI
