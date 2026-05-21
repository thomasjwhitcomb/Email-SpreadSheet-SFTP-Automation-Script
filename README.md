# AHA Registration Bot
### CPR Lifeline - Sac State - CSC131 Team 08

An automated desktop application that handles student registration for AHA (American Heart Association) CPR classes. It monitors incoming Outlook emails, accepts pending student requests in AHA Atlas, updates Google Sheets, sends confirmation and reminder emails via Microsoft Graph, and uploads records to the RQI server over SFTP.

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Setup & Installation](#setup--installation)
3. [Usage Examples](#usage-examples)
4. [Configuration](#configuration)
5. [Project Structure](#project-structure)
6. [Testing](#testing)
7. [Troubleshooting](#troubleshooting)
8. [Building the Standalone .exe (Advanced)](#building-the-standalone-exe-advanced)
9. [Development and Safety](#development-and-safety)
10. [File Reference](#file-reference)

---

## Project Overview

Each scan cycle performs the following steps in order:

| Step | What Happens |
|------|-------------|
| 1 | Reads new AHA registration notification emails from Outlook (Playwright, persistent browser session) |
| 2 | Reads new Acuity scheduling notification emails from Outlook |
| 3 | Logs into [atlas.heart.org](https://atlas.heart.org) and accepts pending student class requests |
| 4 | Sends a confirmation email to each newly registered student |
| 5 | Records student information in the **AHA Student Registration** Google Sheet |
| 6 | Records appointment details in the **RQI Registration Sheet** Google Sheet |
| 7 | Schedules and sends reminder emails to students with upcoming appointments (3 days, 1 day, and 1 hour before) via Microsoft Graph API |
| 8 | Exports new/changed records to CSV and uploads to the RQI SFTP server *(on its own 15-minute schedule)* |

In **Auto Mode** the bot repeats steps 1–7 on a configurable interval (default: every 2 minutes) with no user interaction.

**Key technical characteristics:**

- GUI built with [customtkinter](https://github.com/TomSchimansky/CustomTkinter)
- Browser automation via [Playwright](https://playwright.dev/python/) with a persistent Chromium session (cookies survive restarts)
- Google Sheets read/write via [gspread](https://docs.gspread.org/) + OAuth 2.0 (token cached in `token.json`)
- Microsoft Graph email sending via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python) device-code flow (token cached in `graph_token.json`)
- Reminder scheduling via [APScheduler](https://apscheduler.readthedocs.io/)
- SFTP uploads via [paramiko](https://www.paramiko.org/)
- Passwords stored in the OS keychain via [keyring](https://github.com/jaraco/keyring) — never written to disk
- Exponential back-off with jitter on all Google Sheets API calls

---

## Setup & Installation

### Prerequisites

- A Windows PC (Windows 10 or 11)
- The Outlook email account used to receive AHA and Acuity notification emails
- An Atlas account at [atlas.heart.org](https://atlas.heart.org)
- A Google account with edit access to both Google Sheets
- The RQI SFTP server credentials (provided by RQI)
- A Google Cloud project with the Sheets API and Drive API enabled (one-time setup — see Step 4)

---

### Step 1 — Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/)
2. Download **Python 3.12** (minimum 3.10)
3. Run the installer and check **"Add Python to PATH"** before clicking Install

Verify:
```
python --version
# Python 3.12.x
```

---

### Step 2 — Download the Project

1. On the GitHub repository page click **Code → Download ZIP**
2. Extract the ZIP to a folder, e.g. `C:\AHA Bot\`

---

### Step 3 — Install Dependencies

```
cd C:\AHA Bot
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

This only needs to be done once.

---

### Step 4 — Set Up Google Sheets Access

#### 4a — Enable the APIs

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create or select a project
3. Navigate to **APIs & Services → Library** and enable:
   - **Google Sheets API**
   - **Google Drive API**

#### 4b — Create OAuth Credentials

1. Go to **APIs & Services → Credentials**
2. Click **+ Create Credentials → OAuth client ID**
3. Configure the consent screen (External, any app name), then select **Desktop app**
4. Download the JSON file, rename it `credentials.json`, and place it in the project folder

#### 4c — First Login

On first run, a browser window opens for Google sign-in. After approval, `token.json` is written and re-used automatically until it expires (re-authentication button appears in the GUI when that happens).

---

### Step 5 — Configure Your Settings

```
copy .env.example .env
```

Open `.env` in a text editor and fill in the required values (see [Configuration](#configuration) for a full reference).

---

### Step 6 — Set Up Passwords

Passwords are stored in **Windows Credential Manager**, not in `.env`. After launching the app:

1. **Settings → Outlook → Set Password** — Outlook email password
2. **Settings → Atlas → Set Password** — atlas.heart.org password
3. **RQI Upload page → Set Up Password** — SFTP server password

---

### Step 7 — Launch the App

```
.venv\Scripts\activate
python gui.py
```

Or run the pre-built executable:

```
dist\AHA Bot\AHA Bot.exe
```

---

## Usage Examples

### Starting Auto Mode (recommended for daily use)

1. Launch the app: `python gui.py`
2. On the **Home** tab, click **Start Auto Mode**
3. The bot scans every 2 minutes (configurable via `SCAN_INTERVAL_SECONDS`) and shows live progress in the status bar and Activity Log
4. Click **Stop Auto Mode** when done for the day

### Running a Single Scan

Click **Run Once** on the **Home** tab. The bot completes one full cycle and stops.

### Triggering an Immediate RQI Upload

Go to the **RQI Upload** tab and click **Upload Now**. The bot rebuilds the delta sheet and uploads over SFTP without waiting for the next 15-minute window.

### Checking Connections

| Tab | Button | What it checks |
|-----|--------|---------------|
| Email Monitor | Test Connection | Outlook login via Playwright |
| Google Sheets | Test Access | gspread OAuth + sheet read |
| RQI Upload | Test Connection | SFTP reachability (paramiko) |

### Re-authenticating Google Sheets

When `token.json` expires an orange banner appears at the top of the window. Click **Re-authenticate** and sign in through the browser window that opens.

### Viewing Analytics

Open the **Analytics** tab for total, daily, and weekly metrics. Baselines reset each day/week automatically (persisted in `.analytics_cache.json`).

### Editing Settings at Runtime

Open the **Settings** tab, change any field, and click **Save**. Restart the app for changes to take effect.

---

## Configuration

All configuration lives in `.env` (copy from `.env.example`). Passwords are **not** stored here — use the GUI's "Set Password" buttons instead.

### Required Settings

| Variable | Description |
|----------|-------------|
| `OUTLOOK_EMAIL` | Outlook address that receives AHA and Acuity notification emails |
| `ATLAS_EMAIL` | Login email for [atlas.heart.org](https://atlas.heart.org) |
| `ORGANIZATION_NAME` | Organization name as it appears in Atlas (e.g. `Sac State`) |
| `GOOGLE_SHEET_NAME` | Exact title of the AHA Student Registration Google Sheet |
| `ACUITY_GOOGLE_SHEET_NAME` | Exact title of the RQI/Acuity Appointments Google Sheet |
| `GOOGLE_CREDENTIALS_FILE` | Path to the OAuth credentials JSON (default: `credentials.json`) |
| `SFTP_HOST` | RQI SFTP server hostname (provided by RQI) |
| `SFTP_PORT` | RQI SFTP server port (provided by RQI) |
| `SFTP_USERNAME` | RQI account username |
| `SFTP_REMOTE_DIR` | Upload directory on the RQI server |
| `AZURE_CLIENT_ID` | Azure AD app client ID for Microsoft Graph mail sending |
| `AZURE_TENANT_ID` | Azure AD tenant ID (use `common` for personal/work accounts) |

### Optional / Advanced Settings

| Variable | Default | Description |
|----------|---------|-------------|
| `SCAN_INTERVAL_SECONDS` | `120` | Seconds between Auto Mode scans |
| `EMAIL_LOOKBACK_DAYS` | `7` | How many days back to search Outlook for emails |
| `HEADLESS` | `true` | Run the browser invisibly (`false` shows the browser window) |
| `DRY_RUN` | `false` | Log actions without writing to Sheets, sending email, or uploading |
| `BROWSER` | `chromium` | Playwright browser engine: `chromium`, `firefox`, or `webkit` |
| `ACUITY_SENDER_EMAIL` | — | Exact sender address of Acuity scheduling emails |
| `SFTP_LOCAL_DIR` | — | Local directory for the CSV file before upload (e.g. `C:\temp`) |
| `SFTP_FILENAME` | — | CSV filename written to `SFTP_LOCAL_DIR` |
| `SFTP_KEYRING_SERVICE` | — | Keyring service name used to look up the SFTP password |
| `DELTA_SHEET_NAME` | — | Name of the RQI Upload Delta sheet (managed automatically) |

### Keyring-backed Passwords

These are stored in Windows Credential Manager via `keyring` and are **never** written to `.env`:

| Credential | Set via |
|------------|---------|
| Outlook password | Settings → Outlook → Set Password |
| Atlas password | Settings → Atlas → Set Password |
| SFTP password | RQI Upload page → Set Up Password |

---

## Project Structure

```
PythonAHAScript/
│
├── gui.py                    # Main entry point — run this to launch the app
├── aha_registration_bot.py   # Core automation logic (imported by gui.py)
├── reminder_email.py         # Microsoft Graph mail client + APScheduler reminders
│
├── tests/
│   └── test_core_helpers.py  # pytest unit tests for parsers and helpers
│
├── test_acuity_parsing.py    # Standalone utility: test Acuity email parsing on .msg files
├── make_icon.py              # Utility: regenerate icon.ico from logo.png
│
├── requirements.txt          # Python dependencies
├── AHA Bot.spec              # PyInstaller build configuration
│
├── .env                      # Runtime secrets — never commit this file
├── .env.example              # Template for .env (safe to share)
├── credentials.json          # Google OAuth client secret — never commit this file
├── token.json                # Cached Google OAuth token (auto-generated)
├── graph_token.json          # Cached Microsoft Graph token (auto-generated)
│
├── logo.png                  # CPR Lifeline logo (source for icon)
├── icon.ico                  # Window icon (multi-resolution: 16–256 px)
│
├── aha_bot.log               # Runtime log (auto-generated)
├── .analytics_cache.json     # Analytics baseline cache (auto-generated)
├── .browser_session/         # Playwright persistent browser context (auto-generated)
└── .venv/                    # Python virtual environment (local, not committed)
```

### Module Responsibilities

**`gui.py`** (`App` class, 1 731 lines)

Six-tab customtkinter window:

| Tab | Contents |
|-----|----------|
| Home | Stat cards, Activity Log, Auto Mode controls, Run Once, scan progress |
| Analytics | Total / daily / weekly metrics with baseline deltas |
| Email Monitor | Outlook scan stats, Test Connection |
| Google Sheets | AHA sheet and Acuity sheet stats, Test Access |
| RQI Upload | SFTP stats, Upload Now, Test Connection, SFTP host-key management |
| Settings | All `.env` fields editable in-app plus keyring password buttons |

**`aha_registration_bot.py`** (core library)

Primary public API consumed by `gui.py`:

| Symbol | Purpose |
|--------|---------|
| `run_scan(stop_event)` | Execute one complete scan cycle |
| `auto_mode(scan_interval_seconds, stop_event)` | Loop `run_scan` on an interval |
| `sftp_upload_sheet()` | Rebuild delta sheet and upload via SFTP |
| `get_analytics()` | Return dashboard metric dict |
| `reauthenticate()` | Open browser for Google OAuth re-flow |
| `get_scan_step()` | Current step label for status bar |
| `get_next_scan_time()` | Datetime of next scheduled Auto Mode scan |
| `get_last_scan_result()` | Outcome of the most recent scan |
| `test_outlook_connection()` | Verify Outlook login |
| `test_sheets_connection()` | Verify Google Sheets access |
| `test_sftp_connection()` | Verify SFTP reachability |
| `parse_aha_email_body(body)` | Extract course name + date from AHA email text |
| `parse_acuity_email_body(body)` | Extract 17-field appointment record from Acuity email |
| `StudentRecord` | Dataclass for an AHA-registered student |
| `AcuityRecord` | Dataclass for an Acuity appointment (17 fields) |

**`reminder_email.py`**

| Symbol | Purpose |
|--------|---------|
| `GraphMailClient` | Send HTML email via Microsoft Graph (MSAL device-code, token cached) |
| `ReminderEmail` | High-level scheduler: fires reminders at 3 days, 1 day, and 1 hour before each appointment via APScheduler `BackgroundScheduler` |

---

## Testing

Tests use **pytest** and live in `tests/`.

### Run the test suite

```
.venv\Scripts\activate
pytest
```

### What is tested (`tests/test_core_helpers.py`)

| Test | What it covers |
|------|---------------|
| `test_parse_aha_email_body` | Extracts course name and date from a realistic AHA notification email body |
| `test_parse_acuity_email_body_normal` | Parses a standard Acuity scheduling email into an `AcuityRecord` |
| `test_parse_acuity_email_body_rescheduled` | Handles rescheduled-appointment variant of Acuity emails |
| `test_course_to_group` (parametrized ×3) | Maps Acuity course name strings to the correct RQI group field value |
| `test_sftp_due` | Validates the SFTP upload-window logic (uploads trigger at :12, :27, :42, :57 past each hour) |

### Standalone Acuity parsing utility

To test parsing against real `.msg` files on disk:

```
python test_acuity_parsing.py
```

This reads all `.msg` files from `C:\Users\thoma\Downloads\Acuity Emails`, prints a summary table, and optionally writes results to the configured Google Sheet.

---

## Troubleshooting

### "The app opens but the scan fails immediately"
- Go to **Email Monitor → Test Connection** — if it fails, your Outlook password may need to be re-entered in Settings
- Check that `OUTLOOK_EMAIL` in Settings matches the account that receives the AHA notification emails

### "Google Sheets shows an error / orange re-authenticate banner appears"
- Click the **Re-authenticate** button in the orange banner
- A browser window will open — sign in with your Google account
- This happens automatically every few months when the login token expires

### "RQI Upload fails"
- Go to **RQI Upload → Test Connection** to check if the server is reachable
- Make sure the SFTP password is set (RQI Upload page → Set Up Password)
- If the test passes but the upload still fails, check that `SFTP_LOCAL_DIR` (e.g. `C:\temp`) exists on your computer

### "No students are being found even though emails arrived"
- Check **Email Lookback** in Settings — if set to `7`, it only looks at emails from the past 7 days
- Make sure `ACUITY_SENDER_EMAIL` in Settings matches the exact sender address of your Acuity emails

### "The bot seems frozen during a scan"
- Watch the status bar at the bottom of the app — it shows the current step (e.g. *Step 3 / 5 - Checking Atlas classes...*)
- Atlas can take 1-2 minutes to load, especially if there are many pending students — this is normal
- If it stays frozen for more than 5 minutes, click Stop Auto Mode and try Run Once to see the error in the Activity Log

### "I need to move the app to a different computer"
1. Copy the entire project folder to the new computer
2. Install Python and run Steps 3–6 from the setup guide again
3. Re-enter all three passwords (Outlook, Atlas, RQI) on the new machine — passwords stored in Windows Credential Manager do not transfer between computers

---

## Building the Standalone .exe (Advanced)

If you want to distribute the app as a single folder that runs without needing Python installed:

1. Install PyInstaller:
   ```
   pip install pyinstaller
   ```
2. Make sure `credentials.json` and `.env` are in the project folder
3. Run the build:
   ```
   pyinstaller "AHA Bot.spec"
   ```
4. The finished app will be in:
   ```
   dist\AHA Bot\AHA Bot.exe
   ```

> **Note:** The build now includes `.env.example` instead of bundling `.env` or `credentials.json`.
> Put the real `.env` and `credentials.json` beside the built app during setup so secrets are not baked into the distributable.

---

## Development and Safety

### Dry Run Mode

Set this in `.env` when testing parser, login, and scheduling behavior without making remote changes:

```
DRY_RUN=true
```

Dry run mode logs intended actions but skips student emails, Google Sheets writes, Atlas acceptance, RQI delta rebuilds, SFTP uploads, and upload stamping.

### Tests

Focused parser/helper tests live in `tests/`. Run them after installing dependencies:

```
pytest
```

---

## File Reference

| File | Purpose |
|------|---------|
| `gui.py` | The desktop application — run this to start the bot |
| `aha_registration_bot.py` | Core automation logic — do not run directly |
| `reminder_email.py` | Microsoft Graph mail client and APScheduler reminder engine |
| `requirements.txt` | List of required Python packages |
| `AHA Bot.spec` | Configuration for building the standalone .exe |
| `.env` | **Your private settings** — never share this file |
| `.env.example` | Template for `.env` — safe to share |
| `credentials.json` | **Google login key** — never share this file |
| `token.json` | Saved Google login session — do not delete |
| `graph_token.json` | Saved Microsoft Graph login session — do not delete |
| `logo.png` | CPR Lifeline logo source image |
| `icon.ico` | App window icon (generated from logo.png) |
| `make_icon.py` | Utility to regenerate `icon.ico` from `logo.png` |
| `test_acuity_parsing.py` | Utility to test Acuity email parsing on local .msg files |
| `tests/test_core_helpers.py` | pytest unit tests for parsers and helpers |

---

*Built by CSC131 Team 08 - Sacramento State University, Spring 2026*
