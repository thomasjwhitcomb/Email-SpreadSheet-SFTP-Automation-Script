"""
AHA Registration Automation Bot
=================================
Automates the workflow of:
  1. Reading AHA/Atlas enrollment notification emails from Outlook
  2. Logging into atlas.heart.org to accept pending student requests
  3. Scraping student contact details
  4. Writing the data to a Google Sheet
  5. Sending a confirmation email to the registered student

Dependencies:
    pip install playwright gspread google-auth-oauthlib python-dotenv
    playwright install          # downloads browser binaries (Chromium, Firefox, WebKit)

Setup:
    - Copy .env.example to .env and fill in all values
    - Enable the Google Sheets and Google Drive APIs in Google Cloud Console
    - Download the OAuth 2.0 client credentials JSON and save as 'credentials.json'
      (Cloud Console → APIs & Services → Credentials → Create OAuth client ID → Desktop app)
    - On first run a browser window will open asking you to log in and grant access;
      the token is then cached in 'token.json' and reused silently on every future run
    - Set BROWSER in .env to one of: chromium, firefox, webkit (default: chromium)
      'chromium' works with Chrome and Edge; 'webkit' covers Safari-based browsers
"""

import os
import re
import csv
import time
import errno
import random
import hashlib
import logging
from datetime import datetime, timedelta
from dataclasses import dataclass

from dotenv import load_dotenv
import gspread
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

try:
    from google.auth.exceptions import RefreshError as _GoogleRefreshError
except ImportError:                          # shouldn't happen; guard against missing dep
    _GoogleRefreshError = Exception          # type: ignore[assignment,misc]

try:
    import keyring
    _KEYRING_AVAILABLE = True
except ImportError:
    keyring = None          # type: ignore[assignment]
    _KEYRING_AVAILABLE = False

try:
    import paramiko
    _SFTP_AVAILABLE = _KEYRING_AVAILABLE
except ImportError:
    _SFTP_AVAILABLE = False

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

load_dotenv()

OUTLOOK_EMAIL    = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")

ATLAS_EMAIL    = os.getenv("ATLAS_EMAIL")
ATLAS_PASSWORD = os.getenv("ATLAS_PASSWORD")

ORGANIZATION_NAME = os.getenv("ORGANIZATION_NAME", "Sac State")

GOOGLE_SHEET_NAME    = os.getenv("GOOGLE_SHEET_NAME")
# Path to the OAuth client credentials JSON downloaded from Google Cloud Console.
# The cached token is saved alongside it as 'token.json' after first login.
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

EMAIL_LOOKBACK_DAYS = int(os.getenv("EMAIL_LOOKBACK_DAYS", "7"))

HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
# Supported values: chromium, firefox, webkit  (case-insensitive)
BROWSER = os.getenv("BROWSER", "chromium").lower().strip()

# Acuity Scheduling
ACUITY_SENDER_EMAIL      = os.getenv("ACUITY_SENDER_EMAIL", "no-reply@acuityscheduling.com")
ACUITY_GOOGLE_SHEET_NAME = os.getenv("ACUITY_GOOGLE_SHEET_NAME", "")

# SFTP upload
SFTP_HOST             = os.getenv("SFTP_HOST", "")
SFTP_PORT             = int(os.getenv("SFTP_PORT", "22"))
SFTP_USERNAME         = os.getenv("SFTP_USERNAME", "")
SFTP_REMOTE_DIR       = os.getenv("SFTP_REMOTE_DIR", "")
SFTP_LOCAL_DIR        = os.getenv("SFTP_LOCAL_DIR", "")
SFTP_FILENAME         = os.getenv("SFTP_FILENAME", "")
SFTP_KEYRING_SERVICE  = os.getenv("SFTP_KEYRING_SERVICE", "rqi-sftp")

# Keyring service names for Outlook and Atlas credentials.
# Passwords are stored here instead of .env to avoid plaintext secrets on disk.
OUTLOOK_KEYRING_SERVICE = "aha-outlook"
ATLAS_KEYRING_SERVICE   = "aha-atlas"
SFTP_VERIFY_SHA256    = os.getenv("SFTP_VERIFY_SHA256", "false").lower() == "true"
SFTP_VERIFY_SIZE      = os.getenv("SFTP_VERIFY_SIZE", "true").lower() == "true"
SFTP_AUTO_ADD_HOST_KEY = os.getenv("SFTP_AUTO_ADD_HOST_KEY", "false").lower() == "true"
# Separate Google Sheet written before each SFTP upload (delta / new-records only).
RQI_DELTA_SHEET_NAME  = os.getenv("RQI_DELTA_SHEET_NAME", "")

_LOG_FILE    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aha_bot.log")
_LOG_FORMAT  = "%(asctime)s [%(levelname)s] %(message)s"
_LOG_MAX_BYTES  = 5 * 1024 * 1024   # 5 MB per file
_LOG_BACKUP_COUNT = 5               # keep aha_bot.log + 5 rotated copies (≤ 30 MB total)

from logging.handlers import RotatingFileHandler as _RotatingFileHandler

logging.basicConfig(
    level=logging.INFO,
    format=_LOG_FORMAT,
    handlers=[
        logging.StreamHandler(),
        _RotatingFileHandler(
            _LOG_FILE,
            maxBytes=_LOG_MAX_BYTES,
            backupCount=_LOG_BACKUP_COUNT,
            encoding="utf-8",
        ),
    ],
)
log = logging.getLogger(__name__)

# Exposed to the GUI so it can display a live countdown to the next scan.
# Set by auto_mode() just before each sleep; None when not in auto mode.
_next_scan_time: "datetime | None" = None

# Populated by run_scan() after every cycle so the GUI can show a summary card.
# Keys: ok, ts, aha_emails, students, acuity, duration_s  (or 'error' on failure).
_last_scan_result: dict = {}

# Current step description written by run_scan() as it progresses.
# Empty string when no scan is running.  Read by the GUI tick to show live progress.
_scan_step: str = ""

# ---------------------------------------------------------------------------
# Google Sheets OAuth helpers
# ---------------------------------------------------------------------------

# Absolute path to the cached user token written by gspread after first login.
_TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token.json")

# Phrases in exception messages that reliably indicate an auth problem rather
# than a transient network error.
_OAUTH_PHRASES = (
    "invalid_grant",
    "token has been expired",
    "token_revoked",
    "unauthorized",
    "invalid credentials",
)


class OAuthExpiredError(Exception):
    """
    Raised when the Google OAuth token is invalid, expired, or revoked.

    Every code path that calls ``_gspread_client()`` will propagate this
    exception upward.  The GUI catches it, shows an auth banner, and offers
    a one-click 'Re-authenticate' button.
    """


def _gspread_client() -> gspread.Client:
    """
    Return an authenticated gspread client.

    Wraps ``gspread.oauth()`` and converts all token-expiry / revocation
    exceptions into ``OAuthExpiredError`` so that callers receive a single,
    typed signal instead of a raw traceback, making it easy for the GUI to
    distinguish auth failures from ordinary I/O errors.

    The underlying ``gspread.oauth()`` call:
      - On a valid token.json  → returns immediately with no network round-trip.
      - On an expired access token (but valid refresh token) → silently refreshes.
      - On a revoked / expired refresh token → raises RefreshError (caught here).
      - On a missing token.json → opens a browser login flow.
    """
    try:
        return gspread.oauth(
            credentials_filename=GOOGLE_CREDENTIALS_FILE,
            authorized_user_filename=_TOKEN_FILE,
        )
    except _GoogleRefreshError as exc:
        raise OAuthExpiredError(str(exc)) from exc
    except Exception as exc:
        if any(phrase in str(exc).lower() for phrase in _OAUTH_PHRASES):
            raise OAuthExpiredError(str(exc)) from exc
        raise


def reauthenticate() -> bool:
    """
    Delete the stale token and run a fresh OAuth browser flow.

    Intended to be called from the GUI's 'Re-authenticate' button inside a
    background thread so the GUI remains responsive while the browser is open.

    Steps
    -----
    1. Delete token.json so gspread does not attempt to refresh the bad token.
    2. Call gspread.oauth() — this opens the system browser and writes a new
       token.json once the user completes the sign-in.

    Returns True on success, False if the flow fails for any reason.
    """
    try:
        if os.path.exists(_TOKEN_FILE):
            os.remove(_TOKEN_FILE)
            log.info("Removed stale token.json — starting fresh OAuth flow …")
    except OSError as exc:
        log.warning("Could not remove token.json: %s", exc)

    try:
        gspread.oauth(
            credentials_filename=GOOGLE_CREDENTIALS_FILE,
            authorized_user_filename=_TOKEN_FILE,
        )
        log.info("Re-authentication successful — token.json has been refreshed.")
        return True
    except Exception as exc:
        log.error("Re-authentication failed: %s", exc)
        return False


# ---------------------------------------------------------------------------
# Google Sheets API rate-limit backoff
# ---------------------------------------------------------------------------

# HTTP status codes from the Sheets API that are transient and safe to retry.
#   429 – quota / rate-limit exceeded
#   500 – internal server error
#   502 – bad gateway (Sheets infra blip)
#   503 – service unavailable
_GS_RETRYABLE_CODES = frozenset({429, 500, 502, 503})
_GS_MAX_ATTEMPTS    = 6       # 1 initial try + up to 5 retries
_GS_BASE_DELAY      = 1.0     # seconds before the first retry
_GS_MAX_DELAY       = 64.0    # ceiling so waits never grow absurd


def _gs_call(fn, *args, **kwargs):
    """
    Call a gspread API method with exponential backoff and jitter.

    Retries automatically on transient HTTP errors (429 rate-limit,
    500 / 502 / 503 server errors).  All other exceptions — including
    ``OAuthExpiredError`` and ``gspread.SpreadsheetNotFound`` — propagate
    immediately so callers can handle them without interference.

    The delay sequence with default settings (base=1 s, max=64 s):
        retry 1 → ~1 s   retry 2 → ~2 s   retry 3 → ~4 s
        retry 4 → ~8 s   retry 5 → ~16 s

    Usage::
        rows = _gs_call(ws.get_all_values)
        _gs_call(ws.append_rows, data, value_input_option="USER_ENTERED")
    """
    delay = _GS_BASE_DELAY
    for attempt in range(_GS_MAX_ATTEMPTS):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as exc:
            # Extract the HTTP status code from the response attached to the exception.
            response = getattr(exc, "response", None)
            status   = getattr(response, "status_code", 0)
            if status not in _GS_RETRYABLE_CODES or attempt == _GS_MAX_ATTEMPTS - 1:
                raise
            # Full jitter: wait between [0, delay * 1.25] so concurrent callers
            # don't all retry at the same instant.
            wait = min(delay + random.uniform(0, delay * 0.25), _GS_MAX_DELAY)
            log.warning(
                "Sheets API HTTP %d — attempt %d/%d, retrying in %.1f s …",
                status, attempt + 1, _GS_MAX_ATTEMPTS - 1, wait,
            )
            time.sleep(wait)
            delay = min(delay * 2, _GS_MAX_DELAY)


# ---------------------------------------------------------------------------
# Runtime state  (in-memory, reset each time the bot process starts)
# ---------------------------------------------------------------------------

_runtime: dict = {
    "last_scan_time":        None,   # datetime | None
    "last_sftp_upload_time": None,   # datetime | None
    "consecutive_errors":    0,      # resets to 0 after any successful scan
    "total_scans":           0,
    "successful_scans":      0,
    "last_delta_count":      0,      # records included in the most recent SFTP upload
    "total_students_found":  0,      # cumulative new Atlas students across all scans
    "total_reminders_sent":  0,      # cumulative reminder emails sent this session
}

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class StudentRecord:
    email: str = ""
    first_name: str = ""
    middle_initial: str = ""
    last_name: str = ""
    phone: str = ""
    course_name: str = ""
    course_date: str = ""
    acuity_registered: str = ""  # filled externally if applicable
    aha_registered: str = "YES"
    reminder_email_sent: str = ""  # stamped after confirmation email is sent


@dataclass
class AcuityRecord:
    """
    Holds data parsed from an Acuity Scheduling notification email.
    Fields map 1-to-1 to the columns in the 'RQI Registration Sheet'.
    Fields not present in the Acuity email default to empty string.
    """
    location_id: str = ""
    location_name: str = ""
    user_id: str = ""
    first_name: str = ""
    middle_name: str = ""
    last_name: str = ""
    email: str = ""
    job_code: str = ""
    job_name: str = ""
    hire_date: str = ""
    status: str = ""
    date_of_birth: str = ""
    gender: str = ""
    years_of_experience: str = ""
    active_date: str = ""
    inactive_date: str = ""
    group: str = ""


ACUITY_SHEET_HEADERS = [
    "LocationID", "LocationName", "UserID",
    "FirstName", "MiddleName", "LastName", "Email",
    "JobCode", "JobName", "HireDate", "Status",
    "DateOfBirth", "Gender", "YearsofExperiences",
    "ActiveDate", "InactiveDate", "Group",
]

# ---------------------------------------------------------------------------
# Browser helpers
# ---------------------------------------------------------------------------

# Path where Playwright stores cookies/session so login persists across runs.
# Delete this folder to force a fresh login.
USER_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".browser_session")


def build_page(playwright) -> tuple:
    """
    Launch the browser specified by BROWSER and return (context, page).

    Uses a *persistent context* so cookies and session data are saved to
    USER_DATA_DIR between runs.  On the very first run Outlook will require
    a real login; every run after that reuses the saved session silently.

    Supported BROWSER values:
      chromium  – Chromium-based (works with Chrome & Edge installs)
      firefox   – Mozilla Firefox
      webkit    – WebKit-based (Safari engine)
    """
    valid = ("chromium", "firefox", "webkit")
    if BROWSER not in valid:
        raise ValueError(
            f"Unsupported browser '{BROWSER}'. "
            f"Set BROWSER in .env to one of: {", ".join(valid)}."
        )

    os.makedirs(USER_DATA_DIR, exist_ok=True)

    # Remove stale Chromium lock files left by a previous crashed run.
    # Without this, Chromium exits immediately (exit code 21) if the profile
    # directory still has a SingletonLock from a prior crash.
    for lock_name in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        lock_path = os.path.join(USER_DATA_DIR, lock_name)
        try:
            if os.path.exists(lock_path):
                os.remove(lock_path)
                log.info(f"Removed stale browser lock: {lock_name}")
        except OSError:
            pass  # lock held by a live process — leave it alone

    browser_type = getattr(playwright, BROWSER)
    # launch_persistent_context keeps cookies/localStorage on disk
    context = browser_type.launch_persistent_context(
        USER_DATA_DIR,
        headless=HEADLESS,
        viewport={"width": 1400, "height": 900},
    )
    page = context.new_page()
    return context, page


def click(page: Page, selector: str, timeout: int = 20_000):
    """Wait for a selector and click it."""
    page.wait_for_selector(selector, timeout=timeout)
    page.click(selector)


def fill(page: Page, selector: str, value: str, timeout: int = 20_000):
    """Wait for a selector and fill it with value."""
    page.wait_for_selector(selector, timeout=timeout)
    page.fill(selector, value)


def inner_text(page: Page, selector: str, default: str = "") -> str:
    """Return inner text of an element, or default if not found."""
    try:
        el = page.query_selector(selector)
        return el.inner_text().strip() if el else default
    except Exception:
        return default

# ---------------------------------------------------------------------------
# Part 1 – Outlook: read AHA notification emails
# ---------------------------------------------------------------------------

def parse_aha_email_body(body: str) -> dict:
    """
    Extract structured fields from an AHA Atlas enrollment notification email.

    Expected format (from no-eccreply@heart.org):
        "You have one or more incoming class enrollment requests for
         <Course Name> on MM/DD/YYYY."

    Returns a dict with keys: course_name, course_date.
    Returns an empty dict if neither field can be parsed.
    """
    data = {}

    # Single sentence carries both course name and date:
    # "...enrollment requests for BLS Provider Course on 02/23/2026."
    m = re.search(
        r"enrollment requests?\s+for\s+(.+?)\s+on\s+(\d{1,2}/\d{1,2}/\d{4})",
        body, re.IGNORECASE
    )
    if m:
        data["course_name"] = m.group(1).strip()
        data["course_date"] = m.group(2).strip()   # already MM/DD/YYYY

    return data


def parse_acuity_email_body(body: str, subject: str = "", sender_name: str = "") -> dict:
    """
    Extract structured fields from an Acuity Scheduling notification email.
    Handles both 'New Appointment' and 'Appointment Rescheduled' formats.

    NOTE: Acuity email processing is reserved for future functionality.

    Expected body table format:
        What    <Course Type> (<Location>)
        When    <Day>, <Month> DD, YYYY H:MMam/pm (duration)
        Where   <Street Address>
    """
    data = {}

    # --- Email type ---
    subj_lower = subject.lower()
    if "rescheduled" in subj_lower:
        data["email_type"] = "rescheduled"
    else:
        data["email_type"] = "new"

    # --- Student name ---
    # Sender display name (e.g. "Jayden Kearney") is the most reliable source.
    # Fall back to "for <Name>" line in the body.
    if sender_name:
        data["student_name"] = sender_name.strip()
    else:
        m = re.search(r"^for\s+(.+)$", body, re.IGNORECASE | re.MULTILINE)
        if m:
            data["student_name"] = m.group(1).strip()

    # --- Course type (What row) ---
    # e.g. "What    BLS Skills Test Only (CPR Lifeline Nashville, Music Circle)"
    m = re.search(r"What\s+(.+?)\s*\(CPR Lifeline", body, re.IGNORECASE)
    if m:
        data["course_name"] = m.group(1).strip()

    # --- Location name (parenthetical in What row) ---
    m = re.search(r"What\s+.+?\((CPR Lifeline[^)]+)\)", body, re.IGNORECASE)
    if m:
        data["course_location"] = m.group(1).strip()

    # --- Date/time (When row) ---
    # e.g. "When    Friday, February 27, 2026 9:00am (1 hour)"
    # Converts to MM/DD/YYYY for Atlas date filtering.
    m = re.search(
        r"When\s+\w+,\s+(\w+)\s+(\d{1,2}),\s+(\d{4})\s+([\d:]+(?:am|pm))",
        body, re.IGNORECASE
    )
    if m:
        try:
            dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%B %d %Y")
            data["course_date"] = dt.strftime("%m/%d/%Y")
            data["course_time"] = m.group(4)
        except ValueError:
            pass

    # --- Physical address (Where row) ---
    m = re.search(r"Where\s+(.+?)(?:\n|$)", body, re.IGNORECASE)
    if m:
        data["course_address"] = m.group(1).strip()

    # --- Rescheduled: override course_date with the new appointment time ---
    if data.get("email_type") == "rescheduled":
        m = re.search(
            r"New Time\s+\w+,\s+(\w+)\s+(\d{1,2}),\s+(\d{4})\s+([\d:]+(?:am|pm))",
            body, re.IGNORECASE
        )
        if m:
            try:
                dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%B %d %Y")
                data["course_date"] = dt.strftime("%m/%d/%Y")
                data["course_time"] = m.group(4)
            except ValueError:
                pass

    return data


# Selectors covering both classic OWA (office365.com) and new Outlook
# (outlook.cloud.microsoft).  Deliberately excludes div[role='main'] which
# also appears on the Microsoft login page.
_INBOX_READY_SELS = [
    # Classic OWA
    "div[aria-label='Mail']",
    "[aria-label='Message list']",
    # New Outlook (Monarch / outlook.cloud.microsoft)
    "[data-app-section='MailList']",
    "[aria-label='Inbox']",
    "[aria-label='Message list, Inbox']",
    "div[role='list'][aria-label]",   # message list with any label
    "[data-convid]",                  # any conversation item = inbox rendered
]


def _wait_for_inbox(page: Page, timeout: int = 60_000):
    """Block until the Outlook inbox UI is fully rendered.

    Uses time.sleep() for polling — NOT page.wait_for_timeout() — so the
    loop keeps ticking even if the Playwright connection stalls momentarily.
    Handles KMSI ('Stay signed in?') pages that appear mid-wait.
    """
    import time as _time
    log.info("Waiting for Outlook inbox to fully render …")
    deadline = _time.monotonic() + timeout / 1000

    while _time.monotonic() < deadline:
        # Dismiss KMSI if it appears (click "Yes" so the session gets saved)
        try:
            btn = page.query_selector("#idSIButton9")
            if btn and btn.is_visible():
                # Accept regardless of heading text — in this context the
                # only reason #idSIButton9 appears is the KMSI page.
                btn.click()
                log.info("KMSI 'Yes' clicked (detected during inbox wait).")
                _time.sleep(2)
                continue
        except Exception:
            pass

        # Check if inbox is ready (try each selector)
        try:
            for sel in _INBOX_READY_SELS:
                el = page.query_selector(sel)
                if el and el.is_visible():
                    log.info(f"Inbox rendered (selector: {sel}). URL: {page.url}")
                    return
        except Exception:
            pass

        _time.sleep(1)

    raise RuntimeError(
        f"Outlook inbox did not finish rendering within {timeout // 1000} s. "
        "If a login or MFA prompt is visible, complete it manually "
        "then delete .browser_session/ and re-run to save a fresh session."
    )


def _outlook_login(page: Page):
    """Run the full Microsoft 365 login sequence.

    The Microsoft MSAL login page keeps the password field in the DOM from the
    start (class 'moveOffScreen' = hidden).  We must wait for it to become
    *visible* before filling it, otherwise we write into the hidden element.

    #idSIButton9 is reused as:
      1. "Next"       — on the email step
      2. "Sign in"    — on the password step
      3. "Yes"        — on the "Stay signed in?" (KMSI) page
    We click it after each step and rely on state='visible' waits to pace us.
    """
    log.info(f"Auth redirect detected ({page.url[:80]}…). Logging in …")

    # OWA uses a client-side MSAL.js redirect, so the login page may still be
    # initialising when we arrive here.  Wait for it to settle before we probe.
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass
    log.info(f"Login page settled. URL: {page.url}")

    # --- Account picker (prompt=select_account) ---
    # Microsoft sometimes shows a "pick an account" page.  Clicking our account
    # tile skips the email step and lands directly on the password page.
    # Clicking "Use another account" / "otherTile" goes to the email step.
    ACCOUNT_TILE_SEL   = f"[data-test-id='{OUTLOOK_EMAIL}']"
    OTHER_ACCOUNT_SELS = ("[data-test-id='otherTile']", "#otherTile")

    skip_email = False
    for sel in (ACCOUNT_TILE_SEL,) + OTHER_ACCOUNT_SELS:
        try:
            page.wait_for_selector(sel, timeout=3_000, state="visible")
            log.info(f"Account picker found (selector: {sel}). Clicking …")
            page.click(sel)
            skip_email = (sel == ACCOUNT_TILE_SEL)
            break
        except PWTimeout:
            continue
    try:
        page.wait_for_load_state("networkidle", timeout=10_000)
    except PWTimeout:
        pass

    # --- Step 1: Email (only if we didn't click our account tile directly) ---
    if not skip_email:
        try:
            page.wait_for_selector("#i0116", timeout=10_000, state="visible")
            page.fill("#i0116", OUTLOOK_EMAIL)
            page.click("#idSIButton9")   # "Next"
            log.info("Email entered; clicked Next.")
            try:
                page.wait_for_load_state("networkidle", timeout=10_000)
            except PWTimeout:
                pass
        except PWTimeout:
            raise RuntimeError(
                "Could not find the email field — the login page may not have loaded correctly."
            )

    # --- Step 2: Password ---
    try:
        page.wait_for_selector("#i0118", timeout=20_000, state="visible")
    except PWTimeout:
        raise RuntimeError(
            "Password field did not appear — account may require a different auth flow."
        )
    page.fill("#i0118", get_outlook_password() or "")
    page.click("#idSIButton9")   # "Sign in"
    log.info("Password entered; clicked Sign in.")

    # --- Step 3: "Stay signed in?" (KMSI) or MFA page ---
    # After password submission we land on one of:
    #   a) KMSI "Stay signed in?" page    → #idSIButton9 "Yes"
    #   b) MFA "Approve sign in request"  → user must act on phone
    # Give a short window for whichever page appears.
    try:
        page.wait_for_selector("#idSIButton9", timeout=8_000, state="visible")
        page.click("#idSIButton9")   # "Yes" on KMSI
        log.info("'Stay signed in' accepted.")
    except PWTimeout:
        # Check if the MFA approval page is showing
        mfa_page = page.query_selector("text=Approve sign in request")
        if mfa_page:
            # Try to extract the number shown on-screen for the user's convenience
            try:
                mfa_num_el = page.query_selector(
                    "div.displaySign, [class*='display'], strong, b"
                )
                mfa_num = mfa_num_el.inner_text().strip() if mfa_num_el else "?"
            except Exception:
                mfa_num = "?"
            log.info(
                "*** MFA REQUIRED — open Microsoft Authenticator on your phone and "
                f"approve the request (match number: {mfa_num}). "
                "You have 3 minutes."
            )
        else:
            log.info("KMSI page not shown (possibly MFA or already accepted).")

    # --- Wait for inbox (up to 3 min to allow for MFA approval) ---
    _wait_for_inbox(page, timeout=180_000)
    log.info("Login successful. Session saved to .browser_session/")


def read_aha_emails(page: Page) -> list[dict]:
    """
    Log into Outlook on the web and scrape AHA/Atlas notification emails.
    Returns a list of dicts with instructor_name, course_name, course_date.
    """
    log.info("Navigating to Outlook 365 …")
    page.goto("https://outlook.office365.com/mail/", wait_until="domcontentloaded")
    # Brief pause — the server-side redirect (when session is expired) fires
    # synchronously, so the URL is already correct after domcontentloaded.
    page.wait_for_timeout(2000)

    _OUTLOOK_DOMAINS = ("https://outlook.office365.com", "https://outlook.cloud.microsoft")
    if any(page.url.startswith(d) for d in _OUTLOOK_DOMAINS):
        # URL is on an Outlook domain — session alive; wait for inbox DOM.
        log.info("Existing session detected; waiting for inbox to render …")
        _wait_for_inbox(page, timeout=30_000)
    else:
        # Redirected to login — need to authenticate.
        _outlook_login(page)

    log.info("Outlook inbox ready.")

    # Outlook 365 search box — try several known selectors with a short
    # per-selector timeout so the list doesn't take >75 s to exhaust.
    SEARCH_SELECTORS = [
        "input[aria-label='Search']",
        "input[aria-label='Search Outlook']",
        "input[placeholder='Search']",
        "input[type='search']",
        "[role='searchbox']",
    ]
    search = None
    for sel in SEARCH_SELECTORS:
        try:
            search = page.wait_for_selector(sel, timeout=5_000)
            log.info(f"Search box found with selector: {sel}")
            break
        except PWTimeout:
            continue

    if not search:
        raise RuntimeError(
            "Could not find the Outlook search box — the inbox UI may still be loading."
        )

    # Search for AHA Atlas enrollment notification emails within the lookback window.
    # Sender is no-eccreply@heart.org; subject is "Notification from Atlas: Incoming Enrollment Request".
    lookback_date = (datetime.now() - timedelta(days=EMAIL_LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    search.click()
    search.fill(f"from:no-eccreply@heart.org after:{lookback_date}")
    search.press("Enter")
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass

    # Outlook 365 renders results as list items or rows — try multiple selectors
    EMAIL_ROW_SELECTORS = [
        "div[role='option']",
        "div[role='listitem']",
        "[data-convid]",
        ".customScrollBar div[role='option']",
    ]
    email_items = []
    for sel in EMAIL_ROW_SELECTORS:
        email_items = page.query_selector_all(sel)
        if email_items:
            log.info(f"Email rows found with selector: {sel}")
            break
    log.info(f"Found {len(email_items)} email(s) matching search.")

    email_data = []
    for i in range(len(email_items)):
        try:
            # Re-query each iteration — search results can re-render, staling refs
            fresh_items = []
            for sel in EMAIL_ROW_SELECTORS:
                fresh_items = page.query_selector_all(sel)
                if fresh_items:
                    break
            if i >= len(fresh_items):
                break
            fresh_items[i].scroll_into_view_if_needed()
            fresh_items[i].click()
            page.wait_for_timeout(1500)
            # Outlook 365 reading pane uses different roles than live.com
            body_el = (
                page.query_selector("div[role='document']")
                or page.query_selector("div[aria-label='Message body']")
                or page.query_selector(".allowTextSelection")
            )
            if not body_el:
                continue
            parsed = parse_aha_email_body(body_el.inner_text())
            if parsed.get("course_name") or parsed.get("course_date"):
                email_data.append(parsed)
                log.info(f"  Parsed AHA email: {parsed}")
        except Exception as exc:
            log.warning(f"  Skipping email: {exc}")

    return email_data


def read_acuity_emails(page: Page) -> list[AcuityRecord]:
    """
    Scrape Acuity Scheduling notification emails from Outlook and return a
    list of AcuityRecord objects.

    Reuses the existing browser session — Outlook is already open after
    read_aha_emails(), so we just run a new search without re-logging in.
    The sender address is controlled by ACUITY_SENDER_EMAIL in .env.
    """
    log.info("Searching Outlook for Acuity Scheduling emails …")

    # Always navigate to the inbox root before searching — the previous AHA search
    # may have left the reading pane open on a specific email, hiding the search bar.
    _OUTLOOK_DOMAINS = ("https://outlook.office365.com", "https://outlook.cloud.microsoft")
    page.goto("https://outlook.office365.com/mail/", wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    if not any(page.url.startswith(d) for d in _OUTLOOK_DOMAINS):
        _outlook_login(page)
    _wait_for_inbox(page, timeout=30_000)

    SEARCH_SELECTORS = [
        "input[aria-label='Search']",
        "input[aria-label='Search Outlook']",
        "input[placeholder='Search']",
        "input[type='search']",
        "[role='searchbox']",
    ]
    search = None
    for sel in SEARCH_SELECTORS:
        try:
            search = page.wait_for_selector(sel, timeout=5_000)
            break
        except PWTimeout:
            continue

    if not search:
        log.error("Could not find Outlook search box for Acuity emails. Skipping.")
        return []

    lookback_date = (datetime.now() - timedelta(days=EMAIL_LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    search.click()
    search.fill(f"from:{ACUITY_SENDER_EMAIL} after:{lookback_date}")
    search.press("Enter")
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass

    EMAIL_ROW_SELECTORS = [
        "div[role='option']",
        "div[role='listitem']",
        "[data-convid]",
        ".customScrollBar div[role='option']",
    ]
    email_items = []
    for sel in EMAIL_ROW_SELECTORS:
        email_items = page.query_selector_all(sel)
        if email_items:
            log.info(f"Acuity email rows found with selector: {sel}")
            break
    log.info(f"Found {len(email_items)} Acuity email(s).")

    records: list[AcuityRecord] = []
    for i in range(len(email_items)):
        try:
            # Re-query each iteration to avoid stale element refs
            fresh_items = []
            for sel in EMAIL_ROW_SELECTORS:
                fresh_items = page.query_selector_all(sel)
                if fresh_items:
                    break
            if i >= len(fresh_items):
                break

            # Grab sender display name and subject from the list item before opening
            sender_name = ""
            subject = ""
            try:
                sender_el = fresh_items[i].query_selector(
                    "[class*='sender' i], [class*='from' i], [aria-label*='From' i]"
                )
                if sender_el:
                    sender_name = sender_el.inner_text().strip()
            except Exception:
                pass
            try:
                subj_el = fresh_items[i].query_selector(
                    "[class*='subject' i], [class*='Subject' i]"
                )
                if subj_el:
                    subject = subj_el.inner_text().strip()
            except Exception:
                pass

            fresh_items[i].scroll_into_view_if_needed()
            fresh_items[i].click()
            page.wait_for_timeout(1500)

            body_el = (
                page.query_selector("div[role='document']")
                or page.query_selector("div[aria-label='Message body']")
                or page.query_selector(".allowTextSelection")
            )
            if not body_el:
                continue

            body = body_el.inner_text()
            parsed = parse_acuity_email_body(body, subject=subject, sender_name=sender_name)

            # Split full name into first / middle / last
            # Handles: "First Last", "First Middle Last", "First" (edge case)
            raw_name = parsed.get("student_name", "")
            name_parts = raw_name.split()
            first_name  = name_parts[0] if len(name_parts) >= 1 else ""
            middle_name = name_parts[1] if len(name_parts) == 3 else ""
            last_name   = name_parts[-1] if len(name_parts) >= 2 else ""

            # Try to find the student's email address in the body text
            email_match = re.search(r"[\w.+\-]+@[\w.\-]+\.\w+", body)
            student_email = email_match.group(0) if email_match else ""

            if not student_email and not first_name:
                log.info(f"  Acuity email {i}: could not parse name or email — skipping.")
                continue

            record = AcuityRecord(
                first_name=first_name,
                middle_name=middle_name,
                last_name=last_name,
                email=student_email,
            )
            records.append(record)
            log.info(
                f"  Parsed Acuity email: {first_name} {middle_name+' ' if middle_name else ''}{last_name}"
                f" <{student_email}>"
                + (f" | appt {parsed.get('course_date','')} {parsed.get('course_time','')}" if parsed.get("course_date") else "")
            )

        except Exception as exc:
            log.warning(f"  Skipping Acuity email {i}: {exc}")

    return records

# ---------------------------------------------------------------------------
# Part 2 – Atlas: accept pending requests & scrape student details
# ---------------------------------------------------------------------------

_ATLAS_MAX_ATTEMPTS = 15
_ATLAS_MAX_SECONDS  = 15 * 60   # 15 minutes


def _login_atlas_once(page: Page):
    """Single login attempt for atlas.heart.org. Raises on any failure."""
    page.goto("https://atlas.heart.org/dashboard", wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass

    if "atlas.heart.org/dashboard" in page.url:
        log.info("Atlas session still valid — skipping login.")
        _dismiss_cookie_banner(page)
        return

    log.info(f"Atlas session expired (URL: {page.url}). Logging in …")
    page.goto("https://atlas.heart.org", wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass

    click(page, "button:has-text('Sign In')", timeout=10_000)
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass
    log.info(f"Atlas SSO login URL: {page.url}")

    fill(page, "#Email", ATLAS_EMAIL, timeout=15_000)
    fill(page, "#Password", get_atlas_password() or "", timeout=15_000)
    click(page, "#btnSignIn", timeout=10_000)
    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass
    _dismiss_cookie_banner(page)
    log.info(f"Atlas post-login URL: {page.url}")
    log.info("Logged into Atlas.")


def login_atlas(page: Page):
    """Log into atlas.heart.org with exponential backoff retry.

    Retries up to _ATLAS_MAX_ATTEMPTS times or until _ATLAS_MAX_SECONDS
    of total wait time has elapsed, whichever comes first.
    Backoff delay: 2^attempt seconds, capped at 5 minutes per wait.
    """
    log.info("Logging into atlas.heart.org …")
    deadline = time.monotonic() + _ATLAS_MAX_SECONDS

    for attempt in range(1, _ATLAS_MAX_ATTEMPTS + 1):
        try:
            _login_atlas_once(page)
            return
        except Exception as exc:
            remaining = deadline - time.monotonic()

            if attempt >= _ATLAS_MAX_ATTEMPTS:
                raise RuntimeError(
                    f"Atlas login failed after {attempt} attempt(s): {exc}"
                ) from exc

            if remaining <= 0:
                raise RuntimeError(
                    f"Atlas login timed out (15 min elapsed): {exc}"
                ) from exc

            delay = min(2 ** attempt, 300, remaining)   # cap: 5 min per wait, 15 min total
            log.warning(
                f"Atlas login attempt {attempt} failed: {exc}. "
                f"Retrying in {delay:.0f}s (attempt {attempt+1}/{_ATLAS_MAX_ATTEMPTS}) …"
            )
            time.sleep(delay)


def set_organization(page: Page, org_name: str):
    """Switch the active organization to org_name."""
    log.info(f"Setting organization to '{org_name}' …")
    try:
        page.wait_for_selector("#organizationSelect", timeout=10_000)
        page.select_option("#organizationSelect", label=org_name)
    except PWTimeout:
        # Fallback: dropdown link pattern
        click(page, "#orgDropdown")
        click(page, f"a:has-text('{org_name}')")
    page.wait_for_load_state("networkidle")


def _dismiss_cookie_banner(page: Page):
    """Dismiss OneTrust cookie banner if present."""
    for sel in (
        "button#accept-recommended-btn-handler",
        "button:has-text('Allow All')",
        "button:has-text('Accept All')",
        "button.onetrust-accept-btn-handler",
    ):
        try:
            btn = page.query_selector(sel)
            if btn and btn.is_visible():
                btn.click()
                log.info(f"Cookie banner dismissed via: {sel}")
                page.wait_for_timeout(800)
                return
        except Exception:
            continue


def navigate_to_training_classes(page: Page):
    """Navigate Classes → Training Site Classes."""
    log.info("Navigating to Training Site Classes …")
    _dismiss_cookie_banner(page)

    # Classes is a <button id="Classes">, not an anchor
    page.wait_for_selector("button#Classes", timeout=15_000)
    page.locator("button#Classes").click(force=True)
    page.wait_for_timeout(800)

    # Training Site Classes dropdown item
    tsc = page.query_selector("button:has-text('Training Site Classes')")
    if not tsc or not tsc.is_visible():
        # Re-open dropdown and retry once
        page.locator("button#Classes").click(force=True)
        page.wait_for_timeout(500)
        tsc = page.query_selector("button:has-text('Training Site Classes')")

    if not tsc:
        raise RuntimeError("Could not find 'Training Site Classes' dropdown item.")

    try:
        with page.expect_navigation(timeout=15_000):
            tsc.click(force=True)
    except PWTimeout:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass
    log.info(f"Training Site Classes URL: {page.url}")


def _set_date_range(page: Page, start: str, end: str):
    """
    Interact with the Training Site Classes date range picker.
    Clicks "Choose a Date Range", fills start/end, then applies.
    Returns True if successful.
    """
    # Click the date range trigger button
    date_btn = None
    for sel in (
        "button:has-text('Choose a Date Range')",
        "button:has-text('Date Range')",
        "button:has-text('Date')",
    ):
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                date_btn = el
                break
        except Exception:
            continue

    if not date_btn:
        log.warning("Date range button not found.")
        return False

    date_btn.click()
    page.wait_for_timeout(800)

    # Fill start date
    for sel in (
        "input[placeholder*='Start' i]", "input[placeholder*='From' i]",
        "input[aria-label*='Start' i]", "input[aria-label*='From' i]",
        "input[id*='start' i]", "input[id*='from' i]",
    ):
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.triple_click()
                el.type(start)
                log.info(f"Start date '{start}' filled via {sel}")
                break
        except Exception:
            continue

    # Fill end date
    for sel in (
        "input[placeholder*='End' i]", "input[placeholder*='To' i]",
        "input[aria-label*='End' i]", "input[aria-label*='To' i]",
        "input[id*='end' i]", "input[id*='to' i]",
    ):
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.triple_click()
                el.type(end)
                log.info(f"End date '{end}' filled via {sel}")
                break
        except Exception:
            continue

    # Click Apply/Search/OK
    for sel in (
        "button:has-text('Apply')", "button:has-text('Search')",
        "button:has-text('OK')", "button[type='submit']",
    ):
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                log.info(f"Date range applied via '{sel}'")
                break
        except Exception:
            continue

    try:
        page.wait_for_load_state("networkidle", timeout=10_000)
    except PWTimeout:
        pass
    page.wait_for_timeout(1000)
    return True


def search_classes(page: Page, instructor_name: str, course_date: str):
    """Filter class list by date range (±1 day around course_date)."""
    log.info(f"Searching classes — instructor='{instructor_name}', date='{course_date}' …")

    try:
        date_obj = datetime.strptime(course_date, "%m/%d/%Y")
    except ValueError:
        date_obj = datetime.now()

    start = (date_obj - timedelta(days=1)).strftime("%m/%d/%Y")
    end   = (date_obj + timedelta(days=1)).strftime("%m/%d/%Y")

    _set_date_range(page, start, end)


def _parse_name(raw: str) -> tuple[str, str, str]:
    """Return (first_name, middle_initial, last_name) from a raw name string."""
    parts = raw.split()
    if len(parts) == 0:
        return "", "", ""
    if len(parts) == 1:
        return parts[0], "", ""
    if len(parts) == 2:
        return parts[0], "", parts[1]
    mid = parts[1].rstrip(".")
    if len(mid) <= 2:
        return parts[0], mid.upper(), " ".join(parts[2:])
    return parts[0], "", " ".join(parts[1:])


def _set_registration_status_filter(page: Page, status_text: str) -> bool:
    """
    Open the Registration Status dropdown on the class detail page and
    select the option matching status_text (case-insensitive substring).
    Returns True if the option was clicked.
    """
    for sel in (
        "[id='Registration Status']",
        "input[placeholder*='Status' i]",
        "[aria-label*='Registration Status' i]",
        "[class*='status' i][class*='filter' i]",
    ):
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                page.wait_for_timeout(600)
                break
        except Exception:
            continue

    # Find matching option
    for opt_sel in (
        f"li:has-text('{status_text}')",
        f"[role='option']:has-text('{status_text}')",
        f".dropdown-item:has-text('{status_text}')",
        f"[class*='item']:has-text('{status_text}')",
    ):
        try:
            opt = page.query_selector(opt_sel)
            if opt and opt.is_visible():
                opt.click()
                page.wait_for_timeout(800)
                log.info(f"  Status filter set to '{status_text}'")
                return True
        except Exception:
            continue
    return False


def _accept_modal(page: Page):
    """Confirm accept/approve modal dialog if it appears."""
    for sel in (
        ".modal-footer button.btn-primary",
        "button:has-text('Confirm')",
        "button:has-text('Yes')",
        "button:has-text('OK')",
        "[role='dialog'] button.btn-primary",
    ):
        try:
            page.wait_for_selector(sel, timeout=4_000, state="visible")
            page.click(sel)
            try:
                page.wait_for_load_state("networkidle", timeout=8_000)
            except PWTimeout:
                pass
            return
        except PWTimeout:
            continue


def scrape_students_from_class(page: Page) -> list[StudentRecord]:
    """
    On the currently open Atlas class detail page:
      1. Scroll to the student roster / pending requests section.
      2. For each student row with an Accept action, scrape email + name/phone.
      3. Click the Accept button for each, handle any confirmation modal.

    Table column order (verified from live site):
      col 1 = Email Address
      col 2 = Name / Phone Number  (phone embedded in same cell)
      col 3 = Status
      col 4 = Enrolled By
      col 5 = Action  (Accept button lives here)

    The default table view shows NO students.  Must select "Not Registered" from
    the Registration Status dropdown to reveal pending students, then wait for
    the table to reload before looking for action buttons.
    """
    students = []

    # Scroll to "Finalize Student Roster" section
    try:
        for heading_sel in (
            "h2:has-text('Finalize Student Roster')", "h3:has-text('Finalize Student Roster')",
            "h2:has-text('Pending Requests')", "h3:has-text('Pending Requests')",
            "[class*='roster']",
        ):
            el = page.query_selector(heading_sel)
            if el:
                el.scroll_into_view_if_needed()
                page.wait_for_timeout(800)
                break
    except Exception:
        pass

    # Filter to "Not Registered" — this reveals pending-approval students.
    # The dropdown trigger is the "Choose a Status" button / custom select.
    log.info("  Applying 'Not Registered' filter to reveal pending students …")
    filter_clicked = False
    for trigger_sel in (
        "[id='Registration Status']",
        "div:has-text('Choose a Status') >> nth=0",
        "button:has-text('Choose a Status')",
        "input[placeholder*='Status' i]",
        "[aria-label*='Registration Status' i]",
    ):
        try:
            el = page.query_selector(trigger_sel)
            if el and el.is_visible():
                el.click()
                page.wait_for_timeout(600)
                filter_clicked = True
                log.info(f"  Status dropdown opened via: {trigger_sel}")
                break
        except Exception:
            continue

    if filter_clicked:
        # Click "Not Registered" option
        for opt_sel in (
            "li:has-text('Not Registered')",
            "[role='option']:has-text('Not Registered')",
            "span:has-text('Not Registered')",
        ):
            try:
                opt = page.query_selector(opt_sel)
                if opt and opt.is_visible():
                    opt.click()
                    log.info("  'Not Registered' option selected.")
                    break
            except Exception:
                continue
        # Wait for table to reload after filter
        try:
            page.wait_for_load_state("networkidle", timeout=8_000)
        except PWTimeout:
            pass
        page.wait_for_timeout(1500)

    # Locate accept buttons — spec calls the action "Accept Request"
    ACCEPT_SELS = (
        "button:has-text('Accept Request')",
        "a:has-text('Accept Request')",
        "button:has-text('Accept')",
        "button:has-text('Approve')",
        "button:has-text('Approve Request')",
        "a:has-text('Accept')",
        "a:has-text('Approve')",
        "[class*='accept' i]",
        "[class*='approve' i]",
    )
    accept_buttons = []
    for sel in ACCEPT_SELS:
        accept_buttons = page.query_selector_all(sel)
        if accept_buttons:
            log.info(f"  Found {len(accept_buttons)} accept button(s) via '{sel}'")
            break

    if not accept_buttons:
        log.info("  No accept buttons found on this class page.")
        return students

    for i in range(len(accept_buttons)):
        # Re-query each iteration — DOM changes after each accept
        for sel in ACCEPT_SELS:
            btns = page.query_selector_all(sel)
            if btns:
                break
        if i >= len(btns):
            break
        btn = btns[i]

        try:
            # Scrape from the containing <tr>
            row = btn.evaluate_handle(
                "el => el.closest('tr') || el.closest('[class*=\"student\"]') || el.parentElement"
            ).as_element()

            cells = row.query_selector_all("td") if row else []

            def cell(n: int) -> str:
                """Return text of nth cell (1-based). Empty string if missing."""
                idx = n - 1
                if idx < len(cells):
                    return cells[idx].inner_text().strip()
                return ""

            raw_email     = cell(1)
            raw_name_phone = cell(2)   # "John D. Smith\n(916) 555-1234" or similar

            # Split phone from name — phone is typically last line or in parens
            phone_match = re.search(r"(\(?\d[\d\s\-\.\(\)]{6,14}\d)", raw_name_phone)
            if phone_match:
                raw_phone = phone_match.group(1).strip()
                raw_name  = raw_name_phone[:phone_match.start()].strip()
            else:
                raw_phone = ""
                raw_name  = raw_name_phone.strip()

            first_name, middle_initial, last_name = _parse_name(raw_name)

            record = StudentRecord(
                email=raw_email,
                first_name=first_name,
                middle_initial=middle_initial,
                last_name=last_name,
                phone=raw_phone,
            )
            students.append(record)
            log.info(f"  Scraped: {first_name} {middle_initial} {last_name} <{raw_email}>")

            # Click Accept
            btn.scroll_into_view_if_needed()
            btn.click()
            _accept_modal(page)
            page.wait_for_timeout(500)

        except Exception as exc:
            log.warning(f"  Error processing student row {i}: {exc}")

    return students


def _get_dashboard_request_urls(page: Page) -> list[str]:
    """
    Return hrefs / onclick targets for all 'View Requests' cards on the dashboard.
    Falls back to clicking each button and recording the resulting URL.
    """
    urls = []
    # Try to harvest href directly
    links = page.evaluate(
        """() => {
            const results = [];
            document.querySelectorAll('button, a').forEach(el => {
                if ((el.innerText||'').trim() === 'View Requests') {
                    results.push(el.href || el.getAttribute('onclick') || '__click__');
                }
            });
            return results;
        }"""
    )
    # If hrefs were found, use them
    for link in links:
        if link.startswith("http"):
            urls.append(link)

    if not urls and links:
        # Click each button and record URL
        btns = page.query_selector_all("button:has-text('View Requests'), a:has-text('View Requests')")
        log.info(f"Dashboard: {len(btns)} 'View Requests' card(s) found.")
        for btn in btns:
            try:
                with page.expect_navigation(timeout=15_000):
                    btn.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=10_000)
                except PWTimeout:
                    pass
                urls.append(page.url)
                page.go_back()
                try:
                    page.wait_for_load_state("networkidle", timeout=10_000)
                except PWTimeout:
                    pass
                page.wait_for_timeout(1000)
            except Exception as exc:
                log.warning(f"Could not record View Requests URL: {exc}")
    return urls


def _scrape_course_info(page: Page) -> tuple[str, str]:
    """
    Extract course name and date from an Atlas class detail page.
    Returns (course_name, course_date) — either may be empty string on failure.
    """
    course_name = ""
    course_date = ""

    # --- Course name ---
    # Atlas is a SPA: page.title() and h1 update to the active tab name
    # ("Student Details") not the course name.  Use the breadcrumb instead —
    # its last item is always the course name regardless of selected tab.
    _EXCLUDE_NAMES = {
        "Home", "Classes", "Student Details", "Class Details",
        "Finalize Student Roster", "Students in Progress", "View Class",
    }
    try:
        # 1. Breadcrumb last item — stable on Atlas class-detail pages.
        crumb_text = page.evaluate(
            """() => {
                const containers = [
                    document.querySelector('nav[aria-label*="breadcrumb" i]'),
                    document.querySelector('ol.breadcrumb'),
                    document.querySelector('ul.breadcrumb'),
                    document.querySelector('[class*="breadcrumb" i]'),
                ];
                for (const nav of containers) {
                    if (!nav) continue;
                    const items = [...nav.querySelectorAll('a, span, li')]
                        .map(e => e.innerText.trim())
                        .filter(t => t && t !== '/' && t !== '>' && t.length > 1);
                    if (items.length) return items[items.length - 1];
                }
                return '';
            }"""
        )
        if crumb_text and crumb_text not in _EXCLUDE_NAMES:
            course_name = crumb_text
    except Exception:
        pass

    if not course_name:
        try:
            # 2. Specific Atlas class-detail selectors (exclude tab/section headings)
            for sel in (
                "[class*='course-name' i]",
                "[class*='courseName' i]",
                "[class*='class-name' i]",
                "[class*='className' i]",
                "[class*='course-title' i]",
                "h1",
            ):
                el = page.query_selector(sel)
                if el:
                    text = el.inner_text().strip()
                    if text and text not in _EXCLUDE_NAMES:
                        course_name = text
                        break
        except Exception:
            pass

    # --- Course date ---
    # Atlas shows date as "04-15-2026 | 09:00 am" (MM-DD-YYYY with dashes).
    # Normalise to MM/DD/YYYY for the sheet.
    _DATE_PAT = re.compile(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})")

    def _normalise_date(raw: str) -> str:
        """Convert MM-DD-YYYY or MM/DD/YYYY → MM/DD/YYYY."""
        m = _DATE_PAT.search(raw)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        return ""

    try:
        page_text = page.inner_text("body")

        # 1. Look for the "Date | Time" labelled section (Atlas-specific layout)
        #    The sibling element after the label contains the value.
        for sel in (
            "[class*='classDate' i]",
            "[class*='class-date' i]",
            "[class*='dateTime' i]",
            "[data-label='Class Date']",
            "td:has-text('Class Date') + td",
            "th:has-text('Class Date') ~ td",
        ):
            try:
                el = page.query_selector(sel)
                if el:
                    course_date = _normalise_date(el.inner_text())
                    if course_date:
                        break
            except Exception:
                continue

        # 2. Scan full page text for any date near "Date | Time" label
        if not course_date:
            m = re.search(
                r"Date\s*\|\s*Time\s+(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})",
                page_text, re.IGNORECASE
            )
            if m:
                course_date = _normalise_date(m.group(1))

        # 3. General label scan
        if not course_date:
            m = re.search(
                r"(?:Class\s+(?:Date|Time|Start)[:\s|]+)(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})",
                page_text, re.IGNORECASE
            )
            if m:
                course_date = _normalise_date(m.group(1))

        # 4. Last resort: any date-like string on the page
        if not course_date:
            m = _DATE_PAT.search(page_text)
            if m:
                course_date = _normalise_date(m.group(0))

    except Exception:
        pass

    log.info(f"  Class info scraped — name='{course_name}' date='{course_date}'")
    return course_name, course_date


def process_atlas_via_dashboard(page: Page) -> list[StudentRecord]:
    """
    Dashboard-first approach: find classes with pending requests via
    'View Requests' notification cards, then accept each pending student.
    """
    all_students: list[StudentRecord] = []

    log.info("Checking Atlas dashboard for pending enrollment requests …")
    page.goto("https://atlas.heart.org/dashboard", wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
    except PWTimeout:
        pass
    _dismiss_cookie_banner(page)
    page.wait_for_timeout(2000)

    btns = page.query_selector_all("button:has-text('View Requests'), a:has-text('View Requests')")
    log.info(f"Dashboard: {len(btns)} pending class(es) with enrollment requests.")

    if not btns:
        log.info("No pending enrollment requests on dashboard.")
        return all_students

    # Navigate to each class by index (re-query each time to avoid stale refs)
    for i in range(len(btns)):
        try:
            page.goto("https://atlas.heart.org/dashboard", wait_until="domcontentloaded")
            try:
                page.wait_for_load_state("networkidle", timeout=10_000)
            except PWTimeout:
                pass
            page.wait_for_timeout(1500)
            _dismiss_cookie_banner(page)

            view_btns = page.query_selector_all("button:has-text('View Requests'), a:has-text('View Requests')")
            if i >= len(view_btns):
                break

            try:
                with page.expect_navigation(timeout=15_000):
                    view_btns[i].click()
            except PWTimeout:
                pass
            try:
                page.wait_for_load_state("networkidle", timeout=15_000)
            except PWTimeout:
                pass
            page.wait_for_timeout(2000)
            log.info(f"Class detail URL: {page.url}")

            # Scrape course name + date from the class detail page heading/metadata
            course_name, course_date = _scrape_course_info(page)

            students = scrape_students_from_class(page)
            for s in students:
                if not s.course_name:
                    s.course_name = course_name
                if not s.course_date:
                    s.course_date = course_date
            all_students.extend(students)

        except Exception as exc:
            log.warning(f"Error processing dashboard card {i}: {exc}")

    return all_students


def process_atlas_classes(page: Page, email_data: list[dict]) -> list[StudentRecord]:
    """
    Accept pending Atlas requests and return collected student records.

    Strategy:
      1. Dashboard-driven: check 'View Requests' notification cards first.
      2. Email-driven fallback: for each email notification, navigate to
         Training Site Classes, apply date filter, and open each class.
    """
    all_students: list[StudentRecord] = []
    login_atlas(page)

    # --- Strategy 1: dashboard ---
    dashboard_students = process_atlas_via_dashboard(page)
    all_students.extend(dashboard_students)
    log.info(f"Dashboard pass: {len(dashboard_students)} student(s) collected.")

    # --- Strategy 2: email-driven (only if we have email data) ---
    if email_data:
        navigate_to_training_classes(page)

        # Use a wide date range covering all notified courses
        dates = [n.get("course_date", "") for n in email_data if n.get("course_date")]
        if dates:
            try:
                date_objs = [datetime.strptime(d, "%m/%d/%Y") for d in dates]
                start = (min(date_objs) - timedelta(days=1)).strftime("%m/%d/%Y")
                end   = (max(date_objs) + timedelta(days=1)).strftime("%m/%d/%Y")
            except ValueError:
                today = datetime.now()
                start = (today - timedelta(days=7)).strftime("%m/%d/%Y")
                end   = (today + timedelta(days=30)).strftime("%m/%d/%Y")
            _set_date_range(page, start, end)

        # Find Action links in the listing
        action_links = page.query_selector_all(
            "a:has-text('Action'), a:has-text('View'), button:has-text('Action')"
        )
        log.info(f"Email-driven: {len(action_links)} class(es) found in listing.")

        for i in range(len(action_links)):
            action_links = page.query_selector_all(
                "a:has-text('Action'), a:has-text('View'), button:has-text('Action')"
            )
            if i >= len(action_links):
                break
            try:
                try:
                    with page.expect_navigation(timeout=15_000):
                        action_links[i].click()
                except PWTimeout:
                    pass
                try:
                    page.wait_for_load_state("networkidle", timeout=15_000)
                except PWTimeout:
                    pass
                page.wait_for_timeout(1500)

                # Find matching notification by URL
                course_name = ""
                course_date = ""
                m = re.search(r"id=(\d+)", page.url)
                _ = m.group(1) if m else ""

                for notif in email_data:
                    if notif.get("email_type") == "rescheduled":
                        # Rescheduled appointments update an existing record —
                        # skip here and handle as a future update pass.
                        continue
                    if notif.get("course_date"):
                        course_date = notif["course_date"]
                        course_name = notif.get("course_name", "")
                        break

                students = scrape_students_from_class(page)
                for s in students:
                    if not s.course_name:
                        s.course_name = course_name
                    if not s.course_date:
                        s.course_date = course_date
                all_students.extend(students)

                page.go_back()
                try:
                    page.wait_for_load_state("networkidle", timeout=10_000)
                except PWTimeout:
                    pass
                page.wait_for_timeout(1000)
            except Exception as exc:
                log.warning(f"Error processing class listing item {i}: {exc}")

    return all_students

# ---------------------------------------------------------------------------
# Part 3 – Google Sheets: write student records
# ---------------------------------------------------------------------------

# Column order matches AHARegistration.xlsx exactly:
# A=EMAIL, B=First Name, C=M, D=Last Name, E=Phone,
# F=Course, G=Date, H=Acuity Regist., I=AHA Regist., J=Reminder email sent
# K=RQI Uploaded  (internal tracking — not sent to RQI)
SHEET_HEADERS = [
    "EMAIL", "First Name", "M", "Last Name", "Phone",
    "Course", "Date", "Acuity Regist.", "AHA Regist.", "Reminder email sent",
    "RQI Uploaded",   # format: "MM/DD/YYYY HH:MM|<hash>" — empty = never sent
]

# Delta sheet columns match the RQI Registration Sheet exactly (17 RQI columns).
# The AHA Registration Sheet (SHEET_HEADERS) is record-keeping only and is NOT
# the source for SFTP uploads — the RQI Registration Sheet (ACUITY_SHEET_HEADERS)
# is populated from Acuity emails and is the sole SFTP source.
RQI_DELTA_HEADERS = ACUITY_SHEET_HEADERS  # 17 RQI columns (LocationID → Group)

# Hash all 17 RQI data columns for change detection.
# Excludes the internal "RQI Uploaded" column (col 18) which is not data.
_RQI_HASH_COLS = 17


def get_sheet(spreadsheet_name: str) -> gspread.Worksheet:
    """Return the first worksheet of the named Google Sheet.

    Uses gspread's built-in OAuth flow:
      - First run: opens a browser window for you to log in and grant access.
      - Subsequent runs: loads the cached token from 'token.json' silently.
    No service account or manual credential handling required.
    Raises OAuthExpiredError if the token is invalid or revoked.
    """
    gc = _gspread_client()
    return gc.open(spreadsheet_name).sheet1


def ensure_headers(ws: gspread.Worksheet):
    """
    Write header row if the sheet is empty.
    If the sheet already has a header row that is missing 'RQI Uploaded',
    append that column so existing sheets are upgraded automatically.
    """
    existing = _gs_call(ws.row_values, 1)
    if not existing:
        _gs_call(ws.insert_row, SHEET_HEADERS, index=1)
        log.info("Header row written to '%s'.", GOOGLE_SHEET_NAME)
    elif "RQI Uploaded" not in existing:
        # Append the missing tracking column to an existing sheet
        next_col = len(existing) + 1
        _gs_call(ws.update_cell, 1, next_col, "RQI Uploaded")
        log.info("'RQI Uploaded' column added to '%s' (col %d).", GOOGLE_SHEET_NAME, next_col)


def ensure_rqi_headers(ws: gspread.Worksheet):
    """
    Write the RQI Registration Sheet header row if the sheet is empty, and
    ensure the internal 'RQI Uploaded' tracking column (col 18) exists.

    The 17 public data columns (LocationID → Group) match the RQI sheet format
    exactly and are what gets exported to the delta sheet and SFTP.  Column 18
    ('RQI Uploaded') is internal only — never sent to RQI.
    """
    existing = _gs_call(ws.row_values, 1)
    if not any(existing):
        _gs_call(ws.insert_row, ACUITY_SHEET_HEADERS + ["RQI Uploaded"], index=1)
        log.info("Header row written to '%s'.", ACUITY_GOOGLE_SHEET_NAME)
    elif "RQI Uploaded" not in existing:
        next_col = len([h for h in existing if h]) + 1
        _gs_call(ws.update_cell, 1, next_col, "RQI Uploaded")
        log.info(
            "'RQI Uploaded' column added to '%s' (col %d).",
            ACUITY_GOOGLE_SHEET_NAME, next_col,
        )


def append_students_to_sheet(students: list[StudentRecord]):
    """Append each student as a new row in the AHA Registration Sheet (record-keeping).

    Writes students sourced from AHA Atlas emails.  Duplicate = same email
    address already exists in column A.
    """
    if not students:
        log.info("No students to write to '%s'.", GOOGLE_SHEET_NAME)
        return

    ws = get_sheet(GOOGLE_SHEET_NAME)
    ensure_headers(ws)

    # Collect existing emails (col A) to avoid duplicate rows
    existing_emails = {v.strip().lower() for v in _gs_call(ws.col_values, 1) if v.strip()}

    new_students = []
    for s in students:
        if s.email.strip().lower() in existing_emails:
            log.info(f"  Skipping duplicate: {s.email}")
        else:
            new_students.append(s)
            existing_emails.add(s.email.strip().lower())

    if not new_students:
        log.info("All students already in '%s' — nothing to write.", GOOGLE_SHEET_NAME)
        return

    rows = [
        [
            s.email, s.first_name, s.middle_initial, s.last_name, s.phone,
            s.course_name, s.course_date, s.acuity_registered,
            s.aha_registered, s.reminder_email_sent,
        ]
        for s in new_students
    ]
    _gs_call(ws.append_rows, rows, value_input_option="USER_ENTERED")
    log.info("Wrote %d student row(s) to '%s'.", len(rows), GOOGLE_SHEET_NAME)


# ---------------------------------------------------------------------------
# Part 3b – Google Sheets: Acuity appointment sheet
# ---------------------------------------------------------------------------

def get_acuity_sheet() -> gspread.Worksheet:
    """Return the first worksheet of the Acuity Google Sheet.
    Raises OAuthExpiredError if the token is invalid or revoked.
    """
    gc = _gspread_client()
    return gc.open(ACUITY_GOOGLE_SHEET_NAME).sheet1


def append_acuity_to_sheet(records: list[AcuityRecord]):
    """
    Append Acuity appointment records to the 'RQI Registration Sheet'.

    Row layout matches the sheet's existing 17-column schema:
        LocationID, LocationName, UserID, FirstName, MiddleName, LastName,
        Email, JobCode, JobName, HireDate, Status, DateOfBirth, Gender,
        YearsofExperiences, ActiveDate, InactiveDate, Group

    Duplicate detection: a record whose Email is already present in the
    sheet is skipped.  Only FirstName, MiddleName, LastName, and Email are
    populated automatically; all other columns are left blank for manual
    or downstream completion.
    """
    if not records:
        log.info("No Acuity records to write.")
        return
    if not ACUITY_GOOGLE_SHEET_NAME:
        log.warning("ACUITY_GOOGLE_SHEET_NAME not set — skipping Acuity sheet write.")
        return

    ws = get_acuity_sheet()
    if not _gs_call(ws.row_values, 1):
        _gs_call(ws.insert_row, ACUITY_SHEET_HEADERS, index=1)
        log.info("Header row written to '%s'.", ACUITY_GOOGLE_SHEET_NAME)

    # Build a set of emails already present.
    # Email is column 7 (0-based index 6) per the sheet schema.
    all_rows = _gs_call(ws.get_all_values)
    headers = all_rows[0] if all_rows else []
    email_col_idx = next(
        (i for i, h in enumerate(headers) if h.strip() == "Email"), 6
    )
    existing_emails: set[str] = set()
    for row in all_rows[1:]:
        if len(row) > email_col_idx:
            existing_emails.add(row[email_col_idx].strip().lower())

    new_rows = []
    for r in records:
        if r.email.strip().lower() in existing_emails:
            log.info("  Skipping duplicate Acuity record: %s", r.email)
        else:
            new_rows.append([
                r.location_id, r.location_name, r.user_id,
                r.first_name, r.middle_name, r.last_name, r.email,
                r.job_code, r.job_name, r.hire_date, r.status,
                r.date_of_birth, r.gender, r.years_of_experience,
                r.active_date, r.inactive_date, r.group,
            ])
            existing_emails.add(r.email.strip().lower())

    if new_rows:
        _gs_call(ws.append_rows, new_rows, value_input_option="USER_ENTERED")
        log.info("Wrote %d record(s) to '%s'.", len(new_rows), ACUITY_GOOGLE_SHEET_NAME)
    else:
        log.info("All Acuity records already in '%s' — nothing to write.", ACUITY_GOOGLE_SHEET_NAME)

# ---------------------------------------------------------------------------
# Part 4 – Outlook: send confirmation email to each student
# ---------------------------------------------------------------------------

AHA_SIGNATURE_HTML = """
<br><br>
<div style="font-family:Calibri,sans-serif;font-size:11pt;">
  <strong>AHA Registration Team</strong><br>
  American Heart Association Training Site<br>
  <em>AHA Registration</em>
</div>
"""

EMAIL_SUBJECT = "Complete your AHA Registration"

EMAIL_BODY_TEMPLATE = """\
<p>Dear {first_name},</p>
<p>Thank you for registering for <strong>{course_name}</strong> on <strong>{course_date}</strong>.</p>
<p>Please complete your AHA registration by following the instructions you will receive shortly.</p>
<p>If you have any questions, do not hesitate to reach out.</p>
{signature}
"""


def send_confirmation_emails(page: Page, students: list[StudentRecord]):
    """
    Open Outlook Web, compose and send a confirmation email to each student.
    Stamps reminder_email_sent on the record after a successful send.
    """
    if not students:
        log.info("No students to email.")
        return

    log.info("Navigating to Outlook to send confirmation emails …")
    page.goto("https://outlook.office365.com/mail/")
    page.wait_for_load_state("networkidle")

    for student in students:
        if not student.email:
            log.warning(f"  Skipping student with no email: {student.first_name} {student.last_name}")
            continue

        log.info(f"  Composing email for {student.email} …")
        try:
            # New message
            click(page, "button[aria-label='New mail'], [data-icon-name='ComposeNewFilled']")
            page.wait_for_load_state("domcontentloaded")

            # To field
            to_field = page.wait_for_selector(
                "input[aria-label='To'], div[aria-label='To']", timeout=20_000
            )
            to_field.fill(student.email)
            to_field.press("Tab")

            # Subject
            page.fill("input[aria-label='Subject']", EMAIL_SUBJECT)

            # Insert named signature via Outlook's signature menu if available
            try:
                page.wait_for_selector("[aria-label='Signature']", timeout=5_000)
                page.click("[aria-label='Signature']")
                page.click("button:has-text('AHA Registration')", timeout=5_000)
            except PWTimeout:
                log.warning("  'AHA Registration' signature not found; inserting inline.")

            # Compose body
            body_area = page.wait_for_selector(
                "div[aria-label='Message body'], div[contenteditable='true']",
                timeout=20_000,
            )
            body_html = EMAIL_BODY_TEMPLATE.format(
                first_name=student.first_name or "Student",
                course_name=student.course_name,
                course_date=student.course_date,
                signature=AHA_SIGNATURE_HTML,
            )
            page.evaluate(
                "(el, html) => { el.innerHTML = html; }",
                body_area,
                body_html,
            )

            # Send
            click(page, "button[aria-label='Send'], [data-icon-name='Send']")
            page.wait_for_load_state("networkidle")

            student.reminder_email_sent = datetime.now().strftime("%m/%d/%Y")
            log.info(f"  Email sent to {student.email}.")

        except Exception as exc:
            log.error(f"  Failed to send email to {student.email}: {exc}")

# ---------------------------------------------------------------------------
# Part 5 – SFTP: export AHA sheet to CSV and upload
# ---------------------------------------------------------------------------

def _sftp_verify_remote_dir(sftp: "paramiko.SFTPClient", remote_dir: str) -> None:
    """Verify the remote directory exists. Raises RuntimeError if not."""
    if not remote_dir or remote_dir in (".", "/"):
        return
    remote_dir = remote_dir.replace("\\", "/").rstrip("/")
    try:
        sftp.stat(remote_dir)
    except IOError as e:
        err = getattr(e, "errno", None)
        if err == errno.ENOENT or "No such file" in str(e):
            raise RuntimeError(
                f"Remote directory '{remote_dir}' does not exist. Upload aborted."
            ) from e
        if err == errno.EACCES or "Permission denied" in str(e):
            raise RuntimeError(
                f"No permission to access remote directory '{remote_dir}'. Upload aborted."
            ) from e
        raise


def _sftp_sha256(path: str, chunk_size: int = 65_536) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()


def _sftp_upload(
    host: str,
    port: int,
    username: str,
    password: str,
    local_dir: str,
    filename: str,
    remote_dir: str,
    verify_sha256: bool = False,
    verify_size: bool = True,
    auto_add_host_key: bool = False,
    timeout: int = 30,
) -> bool:
    """
    Upload a single file via SFTP using password authentication.
    Returns True on success, False on any failure.

    Timeout behaviour
    -----------------
    ``timeout`` is applied to every blocking phase of the connection:
    * TCP socket connect  (``timeout``)
    * SSH banner exchange (``banner_timeout``)
    * Authentication      (``auth_timeout``)  ← was missing; caused hang
    * Channel open        (``channel_timeout``)
    * Individual SFTP operations (``sftp.get_channel().settimeout()``)

    Without ``auth_timeout`` paramiko waits indefinitely during the SSH
    authentication handshake, causing the bot to freeze if the server
    accepts the TCP connection but is slow to complete auth.
    """
    import socket as _socket

    if not _SFTP_AVAILABLE:
        log.warning("SFTP skipped — paramiko/keyring not installed. Run: pip install paramiko keyring")
        return False

    local_path = os.path.join(local_dir, filename)
    if not os.path.isfile(local_path):
        log.error("SFTP: local file not found: %s", local_path)
        return False

    remote_dir  = (remote_dir or "").replace("\\", "/").rstrip("/") or "/"
    remote_path = (remote_dir.rstrip("/") + "/" + filename) if remote_dir != "/" else "/" + filename
    if host.startswith("sftp://"):
        host = host[len("sftp://"):]

    client = paramiko.SSHClient()
    try:
        client.load_system_host_keys()
    except Exception:
        pass

    if auto_add_host_key:
        log.warning("SFTP: auto-adding unknown host keys (MITM risk).")
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    else:
        # Warn instead of silently reject so misconfiguration is obvious in logs.
        # If the host key is not in known_hosts this will raise SSHException.
        # Set SFTP_AUTO_ADD_HOST_KEY=true in .env to trust the host on first connect.
        client.set_missing_host_key_policy(paramiko.WarningPolicy())

    sftp = None
    try:
        log.info("SFTP: connecting to %s:%d as %s …", host, port, username)
        client.connect(
            hostname=host,
            port=port,
            username=username,
            password=password,
            timeout=timeout,           # TCP socket connect timeout
            banner_timeout=timeout,    # SSH banner wait timeout
            auth_timeout=timeout,      # SSH authentication timeout (was None → infinite hang)
            channel_timeout=timeout,   # SSH channel open timeout
            look_for_keys=False,
            allow_agent=False,
        )
        log.info("SFTP: authentication successful.")
        sftp = client.open_sftp()

        # Apply timeout to all subsequent SFTP operations (stat, put, etc.)
        # Without this, individual SFTP commands can hang indefinitely even
        # though the connection itself has a timeout.
        try:
            sftp.get_channel().settimeout(timeout)
        except Exception:
            pass

        _sftp_verify_remote_dir(sftp, remote_dir)

        log.info("SFTP: uploading %s → %s", local_path, remote_path)
        sftp.put(local_path, remote_path)

        if verify_size:
            local_size  = os.path.getsize(local_path)
            remote_size = sftp.stat(remote_path).st_size
            if local_size != remote_size:
                log.error("SFTP: size mismatch (local=%d, remote=%d).", local_size, remote_size)
                return False
            log.info("SFTP: size verified (%d bytes).", local_size)

        if verify_sha256:
            local_hash = _sftp_sha256(local_path)
            h = hashlib.sha256()
            with sftp.open(remote_path, "rb") as rf:
                for chunk in iter(lambda: rf.read(65_536), b""):
                    h.update(chunk)
            if local_hash != h.hexdigest():
                log.error("SFTP: SHA-256 mismatch after upload.")
                return False
            log.info("SFTP: SHA-256 verified.")

        log.info("SFTP: upload complete.")
        return True

    except paramiko.AuthenticationException:
        log.error("SFTP: authentication failed — check SFTP username and keychain password.")
        return False
    except paramiko.SSHException as e:
        log.error(
            "SFTP: SSH error: %s. "
            "If 'not found in known_hosts', set SFTP_AUTO_ADD_HOST_KEY=true in .env "
            "for first-connect trust, then disable it.", e
        )
        return False
    except _socket.timeout:
        log.error(
            "SFTP: connection timed out after %ds — server unreachable or "
            "firewall is dropping packets on %s:%d.", timeout, host, port
        )
        return False
    except OSError as e:
        log.error("SFTP: network error: %s", e)
        return False
    except RuntimeError as e:
        log.error("SFTP: %s", e)
        return False
    finally:
        if sftp:
            try:
                sftp.close()
            except Exception:
                pass
        client.close()


def _export_sheet_to_csv(sheet_name: str, local_dir: str, filename: str) -> str:
    """Download a Google Sheet as CSV and save it locally. Returns the file path."""
    ws = _gspread_client().open(sheet_name).sheet1
    rows = _gs_call(ws.get_all_values)
    os.makedirs(local_dir, exist_ok=True)
    path = os.path.join(local_dir, filename)
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    log.info("Exported %d row(s) to %s", len(rows), path)
    return path


# ---------------------------------------------------------------------------
# RQI delta helpers
# ---------------------------------------------------------------------------

def _rqi_row_hash(row: list[str]) -> str:
    """
    Return a short MD5 hash of the RQI Registration Sheet data columns
    (LocationID → Group, i.e. the first 17 columns).

    Only the first ``_RQI_HASH_COLS`` columns are hashed so that changes to
    the internal tracking column (RQI Uploaded, col 18) do not trigger a
    spurious re-upload.
    """
    data = "|".join(row[:_RQI_HASH_COLS])
    return hashlib.md5(data.encode()).hexdigest()[:12]


def _rqi_col_index(ws: gspread.Worksheet) -> int:
    """
    Return the 1-based column index of 'RQI Uploaded' in the RQI Registration
    Sheet, appending the column header if it is not already present.

    The RQI Registration Sheet has 17 data columns (LocationID → Group).
    'RQI Uploaded' is added as column 18 as an internal tracking column;
    it is never included in the delta or SFTP export.
    """
    headers = _gs_call(ws.row_values, 1)
    if "RQI Uploaded" in headers:
        return headers.index("RQI Uploaded") + 1
    # Not present — append it after the last non-empty header
    next_col = len([h for h in headers if h]) + 1
    _gs_call(ws.update_cell, 1, next_col, "RQI Uploaded")
    log.info(
        "'RQI Uploaded' column added to RQI Registration Sheet (col %d).", next_col
    )
    return next_col


def build_rqi_delta() -> list[tuple[int, str]]:
    """
    Compare every row in the RQI Registration Sheet against its last-uploaded
    snapshot and populate the RQI delta sheet with rows that are new or changed.

    Source sheet : ``ACUITY_GOOGLE_SHEET_NAME`` (the RQI Registration Sheet,
                   populated from Acuity emails, columns LocationID → Group).
    Delta sheet  : ``RQI_DELTA_SHEET_NAME`` (same 17-column format, cleared and
                   rebuilt each cycle with only new / changed rows).

    A row is included in the delta when:
    * ``RQI Uploaded`` (col 18, internal) is empty → never sent to RQI, or
    * the hash in ``RQI Uploaded`` differs from the current row hash →
      one or more data fields changed since the last upload.

    The ``RQI Uploaded`` cell stores ``"MM/DD/YYYY HH:MM|<hash>"`` so the
    upload state survives bot restarts.

    Returns
    -------
    list of (rqi_row_index, new_rqi_stamp)
        Each tuple is the 1-based row index in the RQI Registration Sheet and
        the stamp value to write back after a successful SFTP upload.  Callers
        must call ``stamp_rqi_uploaded()`` only after the upload succeeds.
    """
    if not RQI_DELTA_SHEET_NAME:
        log.warning("RQI_DELTA_SHEET_NAME not set — skipping delta build.")
        return []
    if not ACUITY_GOOGLE_SHEET_NAME:
        log.warning("ACUITY_GOOGLE_SHEET_NAME not set — cannot build RQI delta.")
        return []

    gc = _gspread_client()
    master_ws = gc.open(ACUITY_GOOGLE_SHEET_NAME).sheet1
    ensure_rqi_headers(master_ws)
    rqi_col   = _rqi_col_index(master_ws)

    all_rows = _gs_call(master_ws.get_all_values)
    if len(all_rows) < 2:
        log.info("'%s' has no student rows — delta is empty.", ACUITY_GOOGLE_SHEET_NAME)
        return []

    delta_data:  list[list[str]]        = []   # rows for the delta sheet
    stamp_list:  list[tuple[int, str]]  = []   # (master_row_idx, new_stamp)
    now_str = datetime.now().strftime("%m/%d/%Y %H:%M")

    for row_idx, row in enumerate(all_rows[1:], start=2):   # 1-based; row 1 = header
        # Skip completely blank rows or rows with no email (col 7, index 6)
        if not any(cell.strip() for cell in row):
            continue
        email_val = row[6].strip() if len(row) > 6 else ""
        if not email_val:
            continue
        current_hash  = _rqi_row_hash(row)
        rqi_cell      = row[rqi_col - 1].strip() if rqi_col <= len(row) else ""
        new_stamp     = f"{now_str}|{current_hash}"

        if not rqi_cell:
            reason = "new record"
        else:
            stored_hash = rqi_cell.split("|", 1)[-1]
            if stored_hash == current_hash:
                continue   # unchanged — omit from delta
            reason = "data changed"

        # Include only the 17 data columns; strip the internal RQI Uploaded (col 18)
        delta_row = (row + [""] * max(0, len(RQI_DELTA_HEADERS) - len(row)))
        delta_data.append(delta_row[:len(RQI_DELTA_HEADERS)])
        stamp_list.append((row_idx, new_stamp))
        log.info("  Delta: row %d (email: %s) — %s", row_idx, email_val, reason)

    log.info("Delta: %d record(s) to upload (%d unchanged, omitted).",
             len(delta_data), len(all_rows) - 1 - len(delta_data))

    # Rebuild the delta sheet (clear → header → data)
    delta_ws = gc.open(RQI_DELTA_SHEET_NAME).sheet1
    _gs_call(delta_ws.clear)
    _gs_call(delta_ws.insert_row, RQI_DELTA_HEADERS, index=1)
    if delta_data:
        _gs_call(delta_ws.append_rows, delta_data, value_input_option="USER_ENTERED")
    log.info("Delta sheet '%s' refreshed with %d row(s).", RQI_DELTA_SHEET_NAME, len(delta_data))

    return stamp_list


def stamp_rqi_uploaded(stamp_list: list[tuple[int, str]]):
    """
    Write the ``RQI Uploaded`` stamp back to the RQI Registration Sheet after
    a successful SFTP upload.

    Parameters
    ----------
    stamp_list : list of (row_index, stamp_value)
        As returned by ``build_rqi_delta()``.  Only called after a successful
        upload so that failures are automatically retried next cycle.
    """
    if not stamp_list:
        return
    master_ws = _gspread_client().open(ACUITY_GOOGLE_SHEET_NAME).sheet1
    rqi_col = _rqi_col_index(master_ws)

    updates = [
        {"range": gspread.utils.rowcol_to_a1(row_idx, rqi_col), "values": [[stamp]]}
        for row_idx, stamp in stamp_list
    ]
    _gs_call(master_ws.batch_update, updates)
    log.info(
        "Stamped 'RQI Uploaded' on %d row(s) in '%s'.",
        len(stamp_list), ACUITY_GOOGLE_SHEET_NAME,
    )


def sftp_keyring_configured() -> bool:
    """
    Return True if the SFTP password is present in the OS keychain.

    Used by the GUI to detect whether first-time keyring setup is needed
    and to gate the SFTP upload button.  Always returns False when paramiko
    or keyring are not installed.
    """
    if not _SFTP_AVAILABLE:
        return False
    try:
        return bool(keyring.get_password(SFTP_KEYRING_SERVICE, SFTP_USERNAME))
    except Exception:
        return False


def set_sftp_password(password: str) -> bool:
    """
    Store *password* in the OS keychain under the configured service/username.

    Intended to be called from the GUI's first-run setup dialog so the user
    never has to open a terminal.  Returns True on success, False on any error
    (check the log for details).
    """
    if not _SFTP_AVAILABLE:
        log.warning("keyring not installed — cannot store SFTP password.")
        return False
    if not SFTP_USERNAME:
        log.warning("SFTP_USERNAME is not set in .env — cannot store password.")
        return False
    try:
        keyring.set_password(SFTP_KEYRING_SERVICE, SFTP_USERNAME, password)
        log.info(
            "SFTP password stored in OS keychain (service='%s', username='%s').",
            SFTP_KEYRING_SERVICE, SFTP_USERNAME,
        )
        return True
    except Exception as exc:
        log.error("Failed to store SFTP password in keychain: %s", exc)
        return False


# ── Outlook credential helpers ────────────────────────────────────────────────

def outlook_keyring_configured() -> bool:
    """Return True if the Outlook password is present in the OS keychain."""
    if not _KEYRING_AVAILABLE:
        return False
    try:
        return bool(keyring.get_password(OUTLOOK_KEYRING_SERVICE, OUTLOOK_EMAIL or ""))
    except Exception:
        return False


def set_outlook_password(password: str) -> bool:
    """
    Store the Outlook password in the OS keychain.

    Returns True on success, False on any error (check the log for details).
    """
    if not _KEYRING_AVAILABLE:
        log.warning("keyring not installed — cannot store Outlook password.")
        return False
    if not OUTLOOK_EMAIL:
        log.warning("OUTLOOK_EMAIL is not set in .env — cannot store Outlook password.")
        return False
    try:
        keyring.set_password(OUTLOOK_KEYRING_SERVICE, OUTLOOK_EMAIL, password)
        log.info(
            "Outlook password stored in OS keychain (service='%s', username='%s').",
            OUTLOOK_KEYRING_SERVICE, OUTLOOK_EMAIL,
        )
        return True
    except Exception as exc:
        log.error("Failed to store Outlook password in keychain: %s", exc)
        return False


def get_outlook_password() -> "str | None":
    """
    Return the Outlook password.

    Priority:
      1. OS keychain (service ``aha-outlook``, username = OUTLOOK_EMAIL)
      2. ``OUTLOOK_PASSWORD`` env var — backward-compatibility fallback so
         existing ``.env`` files keep working while users migrate to keyring.
    """
    if _KEYRING_AVAILABLE and OUTLOOK_EMAIL:
        try:
            pw = keyring.get_password(OUTLOOK_KEYRING_SERVICE, OUTLOOK_EMAIL)
            if pw:
                return pw
        except Exception:
            pass
    return OUTLOOK_PASSWORD   # may be None if .env not populated


# ── Atlas credential helpers ──────────────────────────────────────────────────

def atlas_keyring_configured() -> bool:
    """Return True if the Atlas password is present in the OS keychain."""
    if not _KEYRING_AVAILABLE:
        return False
    try:
        return bool(keyring.get_password(ATLAS_KEYRING_SERVICE, ATLAS_EMAIL or ""))
    except Exception:
        return False


def set_atlas_password(password: str) -> bool:
    """
    Store the Atlas password in the OS keychain.

    Returns True on success, False on any error (check the log for details).
    """
    if not _KEYRING_AVAILABLE:
        log.warning("keyring not installed — cannot store Atlas password.")
        return False
    if not ATLAS_EMAIL:
        log.warning("ATLAS_EMAIL is not set in .env — cannot store Atlas password.")
        return False
    try:
        keyring.set_password(ATLAS_KEYRING_SERVICE, ATLAS_EMAIL, password)
        log.info(
            "Atlas password stored in OS keychain (service='%s', username='%s').",
            ATLAS_KEYRING_SERVICE, ATLAS_EMAIL,
        )
        return True
    except Exception as exc:
        log.error("Failed to store Atlas password in keychain: %s", exc)
        return False


def get_atlas_password() -> "str | None":
    """
    Return the Atlas password.

    Priority:
      1. OS keychain (service ``aha-atlas``, username = ATLAS_EMAIL)
      2. ``ATLAS_PASSWORD`` env var — backward-compatibility fallback so
         existing ``.env`` files keep working while users migrate to keyring.
    """
    if _KEYRING_AVAILABLE and ATLAS_EMAIL:
        try:
            pw = keyring.get_password(ATLAS_KEYRING_SERVICE, ATLAS_EMAIL)
            if pw:
                return pw
        except Exception:
            pass
    return ATLAS_PASSWORD   # may be None if .env not populated


# ── Test-connection helpers ───────────────────────────────────────────────────

def test_outlook_connection() -> "tuple[bool, str]":
    """
    Verify Outlook credentials are configured and the server is reachable.

    Checks:
      1. OUTLOOK_EMAIL is set.
      2. A password exists (keyring or env fallback).
      3. outlook.cloud.microsoft is reachable on port 443.

    Returns ``(True, message)`` on success, ``(False, reason)`` on failure.
    No browser is opened; a full login test runs automatically with "Run Once".
    """
    import socket
    email = OUTLOOK_EMAIL or ""
    if not email:
        return False, "OUTLOOK_EMAIL not set — add it in Settings"
    if not get_outlook_password():
        return False, "Outlook password not configured — use 'Set Password' on the Settings page"
    try:
        conn = socket.create_connection(("outlook.cloud.microsoft", 443), timeout=6)
        conn.close()
    except OSError as exc:
        return False, f"Cannot reach Outlook server: {exc}"
    src = "keychain" if outlook_keyring_configured() else ".env fallback"
    return True, f"Credentials set for {email} ({src}) — server reachable ✓"


def test_sheets_connection() -> "tuple[bool, str]":
    """
    Test Google Sheets authentication and sheet access.

    Opens the configured spreadsheet and returns its title and row count.
    Returns ``(True, message)`` on success, ``(False, reason)`` on failure.
    """
    if not GOOGLE_SHEET_NAME:
        return False, "GOOGLE_SHEET_NAME not set — add it in Settings"
    try:
        gc = _gspread_client()
        sh = gc.open(GOOGLE_SHEET_NAME)
        rows = sh.sheet1.row_count
        return True, f"Connected to '{sh.title}' — {rows} row(s) in Sheet 1 ✓"
    except OAuthExpiredError:
        return False, "Google OAuth token expired — click Re-authenticate in the banner"
    except gspread.exceptions.SpreadsheetNotFound:
        return False, f"Sheet '{GOOGLE_SHEET_NAME}' not found — check the name in Settings"
    except Exception as exc:
        return False, str(exc)


def test_sftp_connection() -> "tuple[bool, str]":
    """
    Test SFTP connectivity — connects, lists the remote directory, disconnects.

    No file is transferred.  Returns ``(True, message)`` or ``(False, reason)``.
    """
    if not _SFTP_AVAILABLE:
        return False, "paramiko / keyring not installed — SFTP unavailable"
    if not SFTP_HOST or not SFTP_USERNAME:
        return False, "SFTP_HOST or SFTP_USERNAME not configured — add them in Settings"
    password = keyring.get_password(SFTP_KEYRING_SERVICE, SFTP_USERNAME)
    if not password:
        return False, "No SFTP password in OS keychain — run 'Set Up Password' first"
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(
            hostname=SFTP_HOST,
            port=SFTP_PORT,
            username=SFTP_USERNAME,
            password=password,
            timeout=10,
        )
        sftp = client.open_sftp()
        entries = sftp.listdir(SFTP_REMOTE_DIR or ".")
        sftp.close()
        client.close()
        return True, f"Connected to {SFTP_HOST}:{SFTP_PORT} — {len(entries)} item(s) in remote dir ✓"
    except Exception as exc:
        return False, str(exc)


def sftp_upload_sheet():
    """
    Build the RQI delta sheet (new + changed records only), export it to CSV,
    and push it to the RQI SFTP server.

    Two Google Sheets are involved:
    * **Master sheet** (``GOOGLE_SHEET_NAME``) — full record-keeping history.
      Each row's ``RQI Uploaded`` column stores the upload timestamp and a hash
      of the student's data.
    * **Delta sheet** (``RQI_DELTA_SHEET_NAME``) — cleared and repopulated each
      cycle with only the rows that are new or whose data has changed.  This is
      the file exported to CSV and sent to RQI.

    The master sheet is stamped *only after* a successful upload, so any upload
    failure automatically causes those records to be retried next cycle.

    SFTP password is retrieved from the OS keychain via ``keyring``.
    One-time setup:
        python -c "import keyring; keyring.set_password('rqi-sftp', '<username>', input('Password: '))"
    """
    if not _SFTP_AVAILABLE:
        log.warning("SFTP skipped — paramiko/keyring not installed.")
        return

    required = {
        "SFTP_HOST": SFTP_HOST, "SFTP_USERNAME": SFTP_USERNAME,
        "SFTP_REMOTE_DIR": SFTP_REMOTE_DIR, "SFTP_LOCAL_DIR": SFTP_LOCAL_DIR,
        "SFTP_FILENAME": SFTP_FILENAME, "RQI_DELTA_SHEET_NAME": RQI_DELTA_SHEET_NAME,
    }
    missing = [k for k, v in required.items() if not v]
    if missing:
        log.warning("SFTP skipped — missing env vars: %s", missing)
        return

    # ── Step 1: build delta sheet ─────────────────────────────────────────────
    log.info("Building RQI delta (new / changed records) …")
    try:
        stamp_list = build_rqi_delta()
    except Exception as exc:
        log.error("Failed to build '%s': %s", RQI_DELTA_SHEET_NAME, exc)
        return

    if not stamp_list:
        log.info("Delta is empty — no new or changed records to upload.")
        return

    # ── Step 2: export delta sheet → CSV ──────────────────────────────────────
    log.info("Exporting '%s' to CSV …", RQI_DELTA_SHEET_NAME)
    try:
        _export_sheet_to_csv(RQI_DELTA_SHEET_NAME, SFTP_LOCAL_DIR, SFTP_FILENAME)
    except Exception as exc:
        log.error("Failed to export '%s' to CSV: %s", RQI_DELTA_SHEET_NAME, exc)
        return

    # ── Step 3: retrieve SFTP password ────────────────────────────────────────
    try:
        password = keyring.get_password(SFTP_KEYRING_SERVICE, SFTP_USERNAME)
        if not password:
            log.error(
                "No SFTP password in keyring. "
                "Run: python -c \"import keyring; keyring.set_password('%s', '%s', input('Password: '))\"",
                SFTP_KEYRING_SERVICE, SFTP_USERNAME,
            )
            return
    except Exception as exc:
        log.error("Keyring error: %s", exc)
        return

    # ── Step 4: upload ────────────────────────────────────────────────────────
    ok = _sftp_upload(
        host=SFTP_HOST, port=SFTP_PORT, username=SFTP_USERNAME, password=password,
        local_dir=SFTP_LOCAL_DIR, filename=SFTP_FILENAME, remote_dir=SFTP_REMOTE_DIR,
        verify_sha256=SFTP_VERIFY_SHA256, verify_size=SFTP_VERIFY_SIZE,
        auto_add_host_key=SFTP_AUTO_ADD_HOST_KEY,
    )

    # ── Step 5: stamp master sheet only on success ────────────────────────────
    if ok:
        _runtime["last_sftp_upload_time"] = datetime.now()
        _runtime["last_delta_count"]      = len(stamp_list)
        try:
            stamp_rqi_uploaded(stamp_list)
        except Exception as exc:
            log.error(
                "Upload succeeded but failed to stamp master sheet: %s. "
                "Records will be re-included in the next delta — this is safe.", exc
            )
    else:
        log.warning(
            "SFTP upload failed — master sheet NOT stamped. "
            "%d record(s) will be retried next cycle.", len(stamp_list)
        )

# ---------------------------------------------------------------------------
# Part 6 – Appointment reminder emails (3 days / 1 day / 2 hours before)
# ---------------------------------------------------------------------------

REMINDER_SUBJECTS = {
    "3d":  "Reminder: Your AHA Course is in 3 Days",
    "1d":  "Reminder: Your AHA Course is Tomorrow",
    "2hr": "Reminder: Your AHA Course Starts in 2 Hours",
}

REMINDER_BODY_TEMPLATE = """\
<p>Dear {first_name},</p>
<p>This is a friendly reminder that you are registered for
<strong>{course_name}</strong> on <strong>{course_date}</strong>
at <strong>{course_time}</strong>.</p>
{location_block}
<p>If you have any questions, please don't hesitate to reach out.</p>
{signature}
"""


def _parse_appointment_dt(date_str: str, time_str: str) -> datetime | None:
    """Combine a date string (MM/DD/YYYY) and time string (e.g. '9:00am') into a datetime."""
    if not date_str:
        return None
    t = time_str.strip().upper() if time_str else "12:00AM"
    # Normalise: "9AM" → "9:00AM"
    if ":" not in t:
        t = t.replace("AM", ":00AM").replace("PM", ":00PM")
    for fmt in ("%m/%d/%Y %I:%M%p", "%m/%d/%Y"):
        try:
            return datetime.strptime(f"{date_str} {t}".strip(), fmt)
        except ValueError:
            continue
    return None


def _in_reminder_window(appt_dt: datetime, hours_before: float, window_hours: float = 1.5) -> bool:
    """
    Return True if the appointment falls within the reminder window centred on
    `hours_before` hours from now (±window_hours/2).

    With a default window of 1.5 h, any hourly bot run will capture the trigger
    exactly once.  The "already sent" column prevents double-sending on overlap.
    """
    time_until_hours = (appt_dt - datetime.now()).total_seconds() / 3600
    half = window_hours / 2
    return (hours_before - half) <= time_until_hours <= (hours_before + half)


def _compose_and_send(page: Page, to_email: str, subject: str, body_html: str) -> bool:
    """Open a new Outlook compose window, fill it, and send. Returns True on success."""
    try:
        click(page, "button[aria-label='New mail'], [data-icon-name='ComposeNewFilled']")
        page.wait_for_load_state("domcontentloaded")

        to_field = page.wait_for_selector(
            "input[aria-label='To'], div[aria-label='To']", timeout=20_000
        )
        to_field.fill(to_email)
        to_field.press("Tab")

        page.fill("input[aria-label='Subject']", subject)

        body_area = page.wait_for_selector(
            "div[aria-label='Message body'], div[contenteditable='true']",
            timeout=20_000,
        )
        page.evaluate("(el, html) => { el.innerHTML = html; }", body_area, body_html)

        click(page, "button[aria-label='Send'], [data-icon-name='Send']")
        page.wait_for_load_state("networkidle")
        return True
    except Exception as exc:
        log.error("  Failed to send email to %s: %s", to_email, exc)
        return False


def send_reminder_emails(page: Page):
    """
    Appointment reminder emails require Date, Time, and reminder-timestamp
    columns which are not present in the current 'RQI Registration Sheet'
    schema.  This step is a no-op until those columns are added.
    """
    log.info(
        "'%s' does not include appointment Date/Time columns — "
        "reminder email step skipped.", ACUITY_GOOGLE_SHEET_NAME
    )
    return 0

# ---------------------------------------------------------------------------
# Part 7 – Analytics
# ---------------------------------------------------------------------------

def get_analytics() -> dict:
    """
    Return a snapshot of all bot analytics as a flat dictionary.

    Reads the master AHA sheet and the Acuity sheet, then merges the results
    with the in-memory ``_runtime`` counters updated by each scan / upload.

    All keys are always present; any value that cannot be computed defaults to
    ``0``, ``""`` or ``{}`` so callers never need to guard for missing keys.

    Keys
    ----
    Core counters
        total_unique_students       – distinct email addresses in master sheet
        students_uploaded_rqi       – rows with a non-empty RQI Uploaded stamp
        students_pending_new        – rows never uploaded (RQI Uploaded empty)
        students_pending_changed    – rows whose data changed since last upload
        students_pending_total      – new + changed combined
        total_acuity_appointments   – total rows in Acuity sheet

    Reminder stats
        reminders_3d_sent           – Acuity rows with a 3d stamp
        reminders_1d_sent           – Acuity rows with a 1d stamp
        reminders_2hr_sent          – Acuity rows with a 2hr stamp
        upcoming_appointments_7d    – Acuity appointments in the next 7 days

    Registration breakdown
        students_per_course         – {course_name: student_count}
        top_course                  – course with the most students
        most_recent_registration    – latest Date value in master sheet
        cross_registered            – students present in both sheets (by email)
        last_sftp_upload            – timestamp string from most recent RQI stamp

    System health
        last_scan_time              – wall-clock time of last completed scan
        last_sftp_upload_time       – wall-clock time of last successful upload
        consecutive_errors          – scans that failed without a success between them
        total_scans                 – scans run this session
        successful_scans            – scans that returned True this session
        last_delta_count            – records in the most recent SFTP upload
        total_students_found        – cumulative new Atlas students this session
        total_reminders_sent        – cumulative reminder emails this session
    """
    def _fmt(dt) -> str:
        return dt.strftime("%m/%d/%Y %H:%M:%S") if dt else "Never"

    out = {
        # core
        "total_unique_students":     0,
        "students_uploaded_rqi":     0,
        "students_pending_new":      0,
        "students_pending_changed":  0,
        "students_pending_total":    0,
        "total_acuity_appointments": 0,
        # reminders
        "reminders_3d_sent":         0,
        "reminders_1d_sent":         0,
        "reminders_2hr_sent":        0,
        "upcoming_appointments_7d":  0,
        # breakdown
        "students_per_course":       {},
        "top_course":                "",
        "most_recent_registration":  "",
        "cross_registered":          0,
        "last_sftp_upload":          "",
        # health
        "last_scan_time":            _fmt(_runtime["last_scan_time"]),
        "last_sftp_upload_time":     _fmt(_runtime["last_sftp_upload_time"]),
        "consecutive_errors":        _runtime["consecutive_errors"],
        "total_scans":               _runtime["total_scans"],
        "successful_scans":          _runtime["successful_scans"],
        "last_delta_count":          _runtime["last_delta_count"],
        "total_students_found":      _runtime["total_students_found"],
        "total_reminders_sent":      _runtime["total_reminders_sent"],
    }

    master_emails: set[str] = set()

    # ── Master AHA sheet ──────────────────────────────────────────────────────
    try:
        gc = _gspread_client()
        try:
            master_ws = gc.open(GOOGLE_SHEET_NAME).sheet1
        except gspread.SpreadsheetNotFound:
            log.warning(
                "Analytics: master sheet '%s' not found — create it in Google Drive "
                "and share it with your Google account.", GOOGLE_SHEET_NAME
            )
            master_ws = None
        if master_ws is None:
            raise StopIteration   # skip to except block cleanly
        all_rows  = _gs_call(master_ws.get_all_values)

        if len(all_rows) >= 2:
            headers = all_rows[0]
            col     = {h.strip(): i for i, h in enumerate(headers)}

            # AHA Registration Sheet columns
            email_idx  = col.get("EMAIL", 0)
            course_idx = col.get("Course")
            date_idx   = col.get("Date")

            courses:      dict[str, int] = {}
            date_strings: list[str]      = []

            for row in all_rows[1:]:
                def _get(idx) -> str:
                    return row[idx].strip() if idx is not None and idx < len(row) else ""

                email = _get(email_idx).lower()
                if not email:
                    continue   # skip blank rows and rows with no email address

                master_emails.add(email)

                course = _get(course_idx)
                if course:
                    courses[course] = courses.get(course, 0) + 1

                date_str = _get(date_idx)
                if date_str:
                    date_strings.append(date_str)

            # Most recent registration date
            most_recent = ""
            if date_strings:
                try:
                    parsed = [datetime.strptime(d, "%m/%d/%Y") for d in date_strings if d]
                    most_recent = max(parsed).strftime("%m/%d/%Y") if parsed else ""
                except ValueError:
                    most_recent = date_strings[-1]

            top_course = max(courses, key=courses.get) if courses else ""

            out.update({
                "total_unique_students":    len(master_emails),
                "students_per_course":      courses,
                "top_course":               top_course,
                "most_recent_registration": most_recent,
            })

    except (StopIteration, OAuthExpiredError, Exception) as exc:
        if isinstance(exc, OAuthExpiredError):
            raise   # let the caller surface it to the GUI
        if not isinstance(exc, StopIteration):
            log.warning("Analytics: could not read '%s': %s", GOOGLE_SHEET_NAME, exc)

    # ── Acuity sheet ──────────────────────────────────────────────────────────
    if ACUITY_GOOGLE_SHEET_NAME:
        try:
            gc        = _gspread_client()
            try:
                acuity_ws = gc.open(ACUITY_GOOGLE_SHEET_NAME).sheet1
            except gspread.SpreadsheetNotFound:
                log.warning(
                    "Analytics: Acuity sheet '%s' not found — create it in Google Drive "
                    "and share it with your Google account.", ACUITY_GOOGLE_SHEET_NAME
                )
                acuity_ws = None
            if acuity_ws is None:
                raise StopIteration   # skip to except block cleanly
            a_rows    = _gs_call(acuity_ws.get_all_values)

            if len(a_rows) >= 2:
                headers  = a_rows[0]
                acol     = {h.strip(): i for i, h in enumerate(headers)}

                # 'Email' is column 7 (0-based index 6) in the RQI Registration Sheet.
                # Fall back to scanning headers in case the sheet order shifts.
                email_idx = acol.get("Email", acol.get("EMAIL", 6))
                rqi_idx   = acol.get("RQI Uploaded")  # col 18, internal tracking

                acuity_emails: set[str] = set()
                uploaded = pending_new = pending_changed = 0
                last_upload_ts = ""

                for row in a_rows[1:]:
                    def _aget(idx) -> str:
                        return row[idx].strip() if idx is not None and idx < len(row) else ""

                    email_val = _aget(email_idx).lower()
                    if not email_val:
                        continue   # skip blank rows
                    acuity_emails.add(email_val)

                    # RQI upload tracking (from col 18 of the RQI Registration Sheet)
                    if rqi_idx is not None:
                        rqi_val = _aget(rqi_idx)
                        if not rqi_val:
                            pending_new += 1
                        else:
                            stored_hash  = rqi_val.split("|", 1)[-1]
                            current_hash = _rqi_row_hash(row)
                            if stored_hash != current_hash:
                                pending_changed += 1
                            else:
                                uploaded += 1
                            ts_part = rqi_val.split("|", 1)[0]
                            if not last_upload_ts or ts_part > last_upload_ts:
                                last_upload_ts = ts_part
                    else:
                        # RQI Uploaded column not yet present — all rows are pending
                        pending_new += 1

                out.update({
                    "total_acuity_appointments": len(acuity_emails),
                    "students_uploaded_rqi":     uploaded,
                    "students_pending_new":      pending_new,
                    "students_pending_changed":  pending_changed,
                    "students_pending_total":    pending_new + pending_changed,
                    "last_sftp_upload":          last_upload_ts,
                    "reminders_3d_sent":         0,
                    "reminders_1d_sent":         0,
                    "reminders_2hr_sent":        0,
                    "upcoming_appointments_7d":  0,
                    "cross_registered":          len(master_emails & acuity_emails),
                })

        except (StopIteration, OAuthExpiredError, Exception) as exc:
            if isinstance(exc, OAuthExpiredError):
                raise   # let the caller surface it to the GUI
            if not isinstance(exc, StopIteration):
                log.warning("Analytics: could not read '%s': %s", ACUITY_GOOGLE_SHEET_NAME, exc)

    return out

# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def _check_env() -> list[str]:
    """Return a list of missing required configuration items."""
    missing = [v for v in ["OUTLOOK_EMAIL", "ATLAS_EMAIL", "GOOGLE_SHEET_NAME"]
               if not os.getenv(v)]
    # Passwords may live in keyring or env — flag only if neither source has them.
    if not get_outlook_password():
        missing.append("OUTLOOK_PASSWORD (not in keyring or .env)")
    if not get_atlas_password():
        missing.append("ATLAS_PASSWORD (not in keyring or .env)")
    return missing


def _sleep_interruptible(seconds: int, stop_event: "threading.Event | None" = None):
    """Sleep for ``seconds`` in 1-second ticks, waking early if stop_event is set."""
    for _ in range(max(0, seconds)):
        if stop_event is not None and stop_event.is_set():
            return
        time.sleep(1)


def _sftp_quarter_id(dt: datetime) -> int:
    """
    Return a unique integer that identifies the current 15-minute quarter-hour.

    Each hour has 4 quarters (minute // 15 ∈ {0, 1, 2, 3}), giving 96 unique
    IDs per day.  Comparing this value between loop iterations lets auto_mode()
    detect when a new quarter has started without relying on wall-clock alignment.
    """
    return dt.hour * 4 + dt.minute // 15


def _sftp_due(dt: datetime) -> bool:
    """
    Return True when ``dt`` falls inside the SFTP upload window for its quarter.

    The RQI server processes uploads at :00, :15, :30, :45.  We upload at :12
    of each quarter (minute % 15 == 12) to arrive 3 minutes before processing.
    A ±2-minute tolerance (minutes 10-14 of each quarter) ensures the 2-minute
    scan loop never misses the window even if a scan runs long.
    """
    return 10 <= (dt.minute % 15) <= 14


def run_scan(stop_event: "threading.Event | None" = None) -> bool:
    """
    Execute the fast-cycle portion of the bot (runs every ~2 minutes):
      1. Read AHA Atlas notification emails
      2. Read Acuity emails → 2nd Google Sheet
      3. Accept Atlas enrollment requests
      4. Send AHA confirmation emails & write to AHA sheet
      5. Send appointment reminders (3d / 1d / 2hr)

    SFTP upload is intentionally excluded — it runs on its own schedule
    inside ``auto_mode()``.  For a full one-shot run use ``run_once()``.

    Parameters
    ----------
    stop_event : threading.Event, optional
        Set by the GUI or auto-mode loop to request a clean early exit.

    Returns
    -------
    bool
        ``True`` on success, ``False`` if stopped early or an error occurred.
    """
    def _stopped() -> bool:
        return stop_event is not None and stop_event.is_set()

    _runtime["total_scans"] += 1
    log.info("--- Scan: starting (run #%d) ---", _runtime["total_scans"])

    missing = _check_env()
    if missing:
        log.error(f"Missing required environment variables: {', '.join(missing)}")
        _runtime["consecutive_errors"] += 1
        return False

    _scan_t0      = datetime.now()
    _n_aha_emails = 0
    _n_students   = 0
    _n_acuity     = 0

    global _scan_step
    with sync_playwright() as playwright:
        context, page = build_page(playwright)
        try:
            # Step 1 – AHA Atlas notification emails
            if _stopped(): return False
            _scan_step = "Step 1 / 5 — Reading AHA emails…"
            email_notifications = read_aha_emails(page)
            _n_aha_emails = len(email_notifications)
            log.info(f"Retrieved {_n_aha_emails} AHA notification(s).")

            # Step 2 – Acuity emails → 2nd Google Sheet
            if _stopped(): return False
            _scan_step = "Step 2 / 5 — Reading Acuity emails…"
            acuity_records = read_acuity_emails(page)
            _n_acuity = len(acuity_records)
            log.info(f"Retrieved {_n_acuity} Acuity appointment(s).")
            append_acuity_to_sheet(acuity_records)

            # Step 3 – Accept Atlas requests
            if _stopped(): return False
            _scan_step = "Step 3 / 5 — Checking Atlas classes…"
            students = process_atlas_classes(page, email_notifications)
            log.info(f"Collected {len(students)} student record(s).")

            # Step 4 – Send confirmations + write to sheet
            if _stopped(): return False
            _scan_step = "Step 4 / 5 — Registering students…"
            if students:
                send_confirmation_emails(page, students)   # Step 4a
                append_students_to_sheet(students)         # Step 4b
                _n_students = len(students)
                _runtime["total_students_found"] += _n_students
            else:
                log.info("No new Atlas students this scan.")

            # Step 5 – Appointment reminders
            if _stopped(): return False
            _scan_step = "Step 5 / 5 — Sending reminders…"
            send_reminder_emails(page)

            # Update runtime stats on success
            _runtime["last_scan_time"]     = datetime.now()
            _runtime["successful_scans"]  += 1
            _runtime["consecutive_errors"] = 0

            _last_scan_result.clear()
            _last_scan_result.update({
                "ok":         True,
                "ts":         datetime.now(),
                "aha_emails": _n_aha_emails,
                "students":   _n_students,
                "acuity":     _n_acuity,
                "duration_s": (datetime.now() - _scan_t0).total_seconds(),
            })

            _scan_step = ""
            log.info("--- Scan: complete ---")
            return True

        except Exception as exc:
            log.exception(f"Unhandled error during scan: {exc}")
            _runtime["consecutive_errors"] += 1
            _last_scan_result.clear()
            _last_scan_result.update({
                "ok":         False,
                "ts":         datetime.now(),
                "error":      str(exc),
                "aha_emails": _n_aha_emails,
                "students":   _n_students,
                "acuity":     _n_acuity,
                "duration_s": (datetime.now() - _scan_t0).total_seconds(),
            })
            _scan_step = ""
            return False
        finally:
            context.close()


def run_once(stop_event: "threading.Event | None" = None) -> bool:
    """
    Execute one complete bot cycle (scan + SFTP upload).

    Use this for on-demand / manual runs.  In auto mode the two halves run
    on separate schedules via ``auto_mode()``.

    Parameters
    ----------
    stop_event : threading.Event, optional
        Set to request a clean early exit between steps.

    Returns
    -------
    bool
        ``True`` on success, ``False`` if stopped early or an error occurred.
    """
    log.info("=== AHA Registration Bot: starting run ===")
    ok = run_scan(stop_event=stop_event)
    if stop_event is not None and stop_event.is_set():
        return False
    sftp_upload_sheet()
    log.info("=== AHA Registration Bot: run complete ===")
    return ok


def auto_mode(
    scan_interval_seconds: int = 120,
    stop_event: "threading.Event | None" = None,
):
    """
    Two-speed continuous loop — runs until Ctrl+C or ``stop_event`` is set.

    Fast cycle  — every ``scan_interval_seconds`` (default 120 s / 2 min):
        Outlook scan, Atlas acceptance, Google Sheets update, reminder emails.

    Slow cycle  — once per 15-minute quarter, during the :10-:14 window:
        SFTP export + upload.  Aligned to x:12, x:27, x:42, x:57 so the file
        arrives at the RQI server ~3 minutes before the :00/:15/:30/:45
        processing batch.

    Cancellation
    ------------
    * **Ctrl+C** from the terminal raises ``KeyboardInterrupt`` and exits cleanly.
    * Setting ``stop_event`` from another thread (GUI "Stop" button) stops the
      loop within one second without interrupting an in-flight scan.

    Parameters
    ----------
    scan_interval_seconds : int
        Seconds between Outlook / Sheets scans (default: 120).
    stop_event : threading.Event, optional
        External cancellation handle (GUI integration).
    """
    def _stopped() -> bool:
        return stop_event is not None and stop_event.is_set()

    log.info(
        f"=== Auto mode ON | scan every {scan_interval_seconds}s | "
        f"SFTP at :12/:27/:42/:57 | Ctrl+C to stop ==="
    )

    last_sftp_quarter = -1   # tracks which quarter we last uploaded in
    scan_number = 0

    try:
        while not _stopped():
            scan_number += 1
            log.info(f"--- Auto mode: scan #{scan_number} ---")
            run_scan(stop_event=stop_event)

            if _stopped():
                break

            # ── SFTP check ────────────────────────────────────────────────
            now = datetime.now()
            current_quarter = _sftp_quarter_id(now)

            if _sftp_due(now) and current_quarter != last_sftp_quarter:
                log.info(
                    f"SFTP upload window: {now.strftime('%H:%M')} "
                    f"(minute {now.minute} of quarter — uploading now) …"
                )
                sftp_upload_sheet()
                last_sftp_quarter = current_quarter
            else:
                # Log next expected upload time for visibility
                mins_into_quarter = now.minute % 15
                mins_until_window = (10 - mins_into_quarter) % 15
                if mins_until_window == 0:
                    mins_until_window = 15
                next_upload = (now + timedelta(minutes=mins_until_window)).strftime("%H:%M")
                log.info(f"SFTP: next upload window ~{next_upload}.")

            # ── Wait before next scan ─────────────────────────────────────
            if not _stopped():
                global _next_scan_time
                _next_scan_time = datetime.now() + timedelta(seconds=scan_interval_seconds)
                log.info(
                    f"Next scan at ~{_next_scan_time.strftime('%H:%M:%S')} "
                    f"(sleeping {scan_interval_seconds}s) …"
                )
                _sleep_interruptible(scan_interval_seconds, stop_event)
                _next_scan_time = None   # clear once sleep is done

    except KeyboardInterrupt:
        log.info("Auto mode: Ctrl+C received — shutting down.")

    log.info("=== Auto mode stopped ===")


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="AHA Registration Bot",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python aha_registration_bot.py                        # run once (scan + SFTP)\n"
            "  python aha_registration_bot.py --auto                 # auto mode, scan every 2 min\n"
            "  python aha_registration_bot.py --auto --scan-interval 60  # scan every 60 s\n"
        ),
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Enable auto mode: scan continuously and upload via SFTP on the quarter-hour schedule.",
    )
    parser.add_argument(
        "--scan-interval",
        type=int,
        default=None,
        metavar="SECONDS",
        dest="scan_interval",
        help=(
            "Seconds between Outlook / Sheets scans in auto mode "
            "(default: SCAN_INTERVAL_SECONDS env var or 120)."
        ),
    )
    args = parser.parse_args()

    scan_interval = args.scan_interval or int(os.getenv("SCAN_INTERVAL_SECONDS", "120"))

    if args.auto:
        auto_mode(scan_interval_seconds=scan_interval)
    else:
        success = run_once()
        if not success:
            raise SystemExit(1)


if __name__ == "__main__":
    main()