# AHA Registration Bot
### CPR Lifeline — Sac State · CSC131 Team 08

An automated desktop application that handles student registration for AHA (American Heart Association) CPR classes. It monitors incoming emails, logs students into the AHA Atlas system, updates Google Sheets, sends confirmation emails, and uploads records to RQI — all automatically.

---

## Table of Contents

1. [What the Bot Does](#what-the-bot-does)
2. [What You Need Before Starting](#what-you-need-before-starting)
3. [First-Time Setup](#first-time-setup)
   - [Step 1 — Install Python](#step-1--install-python)
   - [Step 2 — Download the Project](#step-2--download-the-project)
   - [Step 3 — Install Dependencies](#step-3--install-dependencies)
   - [Step 4 — Set Up Google Sheets Access](#step-4--set-up-google-sheets-access)
   - [Step 5 — Configure Your Settings](#step-5--configure-your-settings)
   - [Step 6 — Set Up Passwords](#step-6--set-up-passwords)
   - [Step 7 — Launch the App](#step-7--launch-the-app)
4. [Using the App](#using-the-app)
   - [Home — Dashboard](#home--dashboard)
   - [Email Monitor](#email-monitor)
   - [Google Sheets](#google-sheets)
   - [RQI Upload](#rqi-upload)
   - [Settings](#settings)
5. [Auto Mode vs. Manual Mode](#auto-mode-vs-manual-mode)
6. [How Each Sheet Is Used](#how-each-sheet-is-used)
7. [Troubleshooting](#troubleshooting)
8. [Building the Standalone .exe (Advanced)](#building-the-standalone-exe-advanced)

---

## What the Bot Does

Each time it runs, the bot performs these steps in order:

| Step | What Happens |
|------|-------------|
| 1 | Reads new AHA registration notification emails from Outlook |
| 2 | Reads new Acuity appointment notification emails from Outlook |
| 3 | Logs into [atlas.heart.org](https://atlas.heart.org) and accepts pending student class requests |
| 4 | Sends a confirmation email to each newly registered student |
| 5 | Records student information in the **AHA Student Registration** Google Sheet |
| 6 | Records appointment details in the **RQI Registration Sheet** Google Sheet |
| 7 | Sends reminder emails to students with upcoming appointments (3 days, 1 day, and 2 hours before) |
| 8 | Exports new/changed records to a CSV file and uploads it to the RQI server *(on its own schedule — every 15 minutes)* |

In **Auto Mode**, the bot repeats steps 1–7 every 2 minutes automatically, with no interaction needed.

---

## What You Need Before Starting

Before setting up, make sure you have access to:

- ✅ A Windows PC (Windows 10 or 11)
- ✅ The **Outlook email account** used to receive AHA and Acuity notification emails
- ✅ The **Atlas account** at [atlas.heart.org](https://atlas.heart.org)
- ✅ A **Google account** that owns or has edit access to both Google Sheets
- ✅ The **RQI SFTP server password** (provided by RQI — contact them if you don't have it)
- ✅ A Google Cloud project with the Sheets & Drive APIs enabled *(one-time setup — see Step 4)*

---

## First-Time Setup

### Step 1 — Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/)
2. Click **Download Python 3.12** (or the latest 3.x version shown)
3. Run the installer
4. **Important:** On the first screen, check the box that says **"Add Python to PATH"** before clicking Install

To verify it worked, open **Command Prompt** (search "cmd" in the Start menu) and type:
```
python --version
```
You should see something like `Python 3.12.4`.

---

### Step 2 — Download the Project

1. Go to the GitHub repository page
2. Click the green **Code** button → **Download ZIP**
3. Extract the ZIP to a folder on your computer, for example:
   ```
   C:\AHA Bot\
   ```

---

### Step 3 — Install Dependencies

The bot uses several helper packages that need to be installed once.

1. Open **Command Prompt** and navigate to the project folder:
   ```
   cd C:\AHA Bot
   ```
2. Create an isolated environment (keeps everything tidy):
   ```
   python -m venv .venv
   .venv\Scripts\activate
   ```
3. Install all required packages:
   ```
   pip install -r requirements.txt
   ```
4. Install the browser the bot uses internally:
   ```
   playwright install chromium
   ```

This may take a few minutes. You only need to do this once.

---

### Step 4 — Set Up Google Sheets Access

The bot needs permission to read and write your Google Sheets. This is done through a secure Google login — your password is never stored anywhere.

#### 4a — Enable the APIs

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or select an existing one)
3. In the left menu, go to **APIs & Services → Library**
4. Search for and enable both:
   - **Google Sheets API**
   - **Google Drive API**

#### 4b — Create Credentials

1. Go to **APIs & Services → Credentials**
2. Click **+ Create Credentials → OAuth client ID**
3. If prompted to configure the consent screen, choose **External**, fill in an app name (e.g. "AHA Bot"), and save
4. Back at Create Credentials, select **Desktop app** as the application type
5. Click **Create**, then **Download JSON**
6. Rename the downloaded file to `credentials.json` and place it in your project folder:
   ```
   C:\AHA Bot\credentials.json
   ```

#### 4c — First Login

The first time you run the app and it connects to Google Sheets, a browser window will open asking you to sign in with your Google account. After you approve access, the app saves a login token (`token.json`) and will never ask again unless the token expires.

> **Note:** If Google shows a warning that the app is "unverified," click **Advanced → Go to [App Name] (unsafe)**. This is expected for internal tools that haven't been submitted for Google's review process.

---

### Step 5 — Configure Your Settings

1. In the project folder, make a copy of `.env.example` and name it `.env`:
   ```
   copy .env.example .env
   ```
2. Open `.env` in Notepad and fill in your details:

| Setting | What to Enter |
|---------|--------------|
| `OUTLOOK_EMAIL` | The Outlook email address that receives AHA/Acuity notifications |
| `ATLAS_EMAIL` | Your login email for atlas.heart.org |
| `ORGANIZATION_NAME` | Your organization name as it appears in Atlas (e.g. `Sac State`) |
| `GOOGLE_SHEET_NAME` | Exact name of your AHA registration Google Sheet |
| `ACUITY_GOOGLE_SHEET_NAME` | Exact name of your RQI/Acuity appointments Google Sheet |
| `SFTP_HOST` | RQI server address (provided by RQI) |
| `SFTP_PORT` | RQI server port (provided by RQI) |
| `SFTP_USERNAME` | Your RQI account number/username |
| `SFTP_REMOTE_DIR` | Upload folder on the RQI server (provided by RQI) |

Leave all other settings at their default values unless told otherwise.

> **Passwords are NOT entered in this file.** They are stored securely in Windows' built-in password manager. See Step 6.

---

### Step 6 — Set Up Passwords

Passwords are stored in the **Windows Credential Manager** (the same place Windows stores your Wi-Fi and website passwords). They are never written to any file.

Launch the app first (Step 7), then:

1. Go to **Settings → Outlook → Set Password** — enter your Outlook email password
2. Go to **Settings → Atlas → Set Password** — enter your atlas.heart.org password
3. Go to the **RQI Upload** page → **Set Up Password** — enter the SFTP password provided by RQI

Each password only needs to be entered once. The app will remember it automatically.

---

### Step 7 — Launch the App

With your `.venv` activated, run:
```
python gui.py
```

Or, if you built the standalone `.exe` (see the [advanced section](#building-the-standalone-exe-advanced)):
```
dist\AHA Bot\AHA Bot.exe
```

---

## Using the App

### Home — Dashboard

The main screen you see when the app opens.

**Stat cards** across the top show at a glance:
- **Total Students** — number of unique students recorded in the AHA sheet
- **Pending RQI Upload** — records that haven't been sent to RQI yet
- **Upcoming Appointments** — Acuity appointments in the next 7 days
- **Reminders Sent** — reminder emails sent during this session

**Last Scan Result** shows the outcome of the most recent scan:
- ✅ Completed — how many emails were read, students registered, and appointments found
- ❌ Failed — a red banner appears at the top of the screen with a description of what went wrong

**Activity Log** at the bottom shows a live feed of everything the bot is doing, in plain language. Use **Open Log File** to open the full history in Notepad.

**Buttons:**

| Button | What It Does |
|--------|-------------|
| ▶ Start Auto Mode | Starts continuous scanning every 2 minutes — leave it running all day |
| ■ Stop Auto Mode | Stops the automatic scanning |
| ⟳ Run Once | Runs one complete scan right now, then stops |
| ↻ Refresh Stats | Updates the numbers on the dashboard immediately |

---

### Email Monitor

Shows statistics about email scanning:
- When the last scan happened
- How many students and appointments have been found in total
- How many scans have run this session
- Whether any errors have occurred

The **Scan Outlook Now** button triggers a scan immediately without starting Auto Mode.

The **Test Connection** button checks that the app can reach Outlook — useful for diagnosing login problems.

---

### Google Sheets

Shows statistics pulled directly from your Google Sheets:

**AHA Registration Sheet** — the master record of all registered students, including:
- Total students, most recent registration, top course
- RQI upload status (how many have been uploaded vs. pending)

**Acuity Appointments Sheet** — appointment tracking, including:
- Total appointments, upcoming count, reminder email history

Click **Test Access** to verify the app can connect to your Google Sheets.

---

### RQI Upload

Manages the automated upload of student records to the RQI server.

- **Last Upload** — when the most recent upload occurred
- **Records in Last Delta** — how many records were included in that upload
- **Records Pending** — records queued for the next upload
- **Next Upload Window** — the next scheduled upload time (uploads run every 15 minutes)

Click **Upload Now** to trigger an upload immediately without waiting for the schedule.

Click **Test Connection** to verify the app can reach the RQI server.

> The upload only sends **new or changed** records each time — it does not re-upload everything.

---

### Settings

All configuration options in one place. After making changes, click **💾 Save** and **restart the app** for changes to take effect.

**Key settings explained:**

| Setting | Plain-English Explanation |
|---------|--------------------------|
| Outlook Email | The email address the bot logs into to read notification emails |
| Atlas Email | The email address used to log into atlas.heart.org |
| Organization Name | Filters Atlas classes to only your organization |
| AHA Sheet Name | The exact title of your AHA registration Google Sheet |
| Acuity Sheet Name | The exact title of your Acuity appointments Google Sheet |
| Scan Interval | How often Auto Mode scans, in seconds (120 = every 2 minutes) |
| Email Lookback | How many days back to search for emails |
| Headless Mode | When ON, the browser runs invisibly in the background (recommended) |

> **Passwords** have a **Set Password** button — they are stored in Windows and never written to this file.

---

## Auto Mode vs. Manual Mode

| | Auto Mode | Manual (Run Once) |
|--|-----------|-------------------|
| **How to start** | Click ▶ Start Auto Mode | Click ⟳ Run Once |
| **How long it runs** | Until you click Stop, or close the app | One scan, then stops |
| **Best for** | Leaving the computer running all day | Quick one-off checks |
| **RQI uploads** | Automatically every 15 minutes | Included in each run |
| **Status bar** | Shows live step progress and countdown | Shows live step progress |

**Recommended daily use:** Start Auto Mode at the beginning of the work day and stop it when done.

---

## How Each Sheet Is Used

### AHA Student Registration Sheet
- **Purpose:** Permanent record-keeping for AHA class registrations
- **Updated by:** The bot when it processes AHA atlas.heart.org notification emails
- **Contents:** Student name, email, course, date, instructor, location, etc.
- **Do not rename** this sheet — the name must match what's in Settings exactly

### RQI Registration Sheet
- **Purpose:** Tracks Acuity appointments and RQI upload status
- **Updated by:** The bot when it processes Acuity scheduling emails
- **Contents:** 17 appointment fields (Location ID, student info, course details, etc.) plus an internal "RQI Uploaded" tracking column added automatically by the bot
- **Do not rename or delete** the "RQI Uploaded" column — the bot uses it to track what has already been sent

### RQI Upload Delta Sheet
- **Purpose:** Temporary holding sheet used just before each upload to RQI
- **Updated by:** The bot automatically before each SFTP upload
- **Contents:** Only the records that are new or changed since the last upload
- **You do not need to interact with this sheet** — it is managed entirely by the bot

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
- Watch the status bar at the bottom of the app — it shows the current step (e.g. *Step 3 / 5 — Checking Atlas classes…*)
- Atlas can take 1–2 minutes to load, especially if there are many pending students — this is normal
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

> **Note:** The built `.exe` bundles a copy of `.env` and `credentials.json` at build time. If you change your settings after building, you either need to rebuild or edit the copies inside `dist\AHA Bot\_internal\`.

---

## File Reference

| File | Purpose |
|------|---------|
| `gui.py` | The desktop application — run this to start the bot |
| `aha_registration_bot.py` | Core automation logic — do not run directly |
| `requirements.txt` | List of required Python packages |
| `AHA Bot.spec` | Configuration for building the standalone .exe |
| `.env` | **Your private settings** — never share this file |
| `.env.example` | Template for `.env` — safe to share |
| `credentials.json` | **Google login key** — never share this file |
| `token.json` | Saved Google login session — do not delete |
| `logo.png` | CPR Lifeline logo source image |
| `icon.ico` | App window icon (generated from logo.png) |
| `make_icon.py` | Utility to regenerate `icon.ico` from `logo.png` |

---

*Built by CSC131 Team 08 — Sacramento State University, Spring 2026*
