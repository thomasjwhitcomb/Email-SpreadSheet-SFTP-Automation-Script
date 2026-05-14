"""
Graph API reminder email system for AHA Registration Bot.

Authentication: Azure AD delegated permissions (device code flow, no admin consent needed).
Scheduling:     APScheduler BackgroundScheduler with date triggers.

Required environment variables (see .env.example):
    AZURE_CLIENT_ID, AZURE_TENANT_ID, OUTLOOK_EMAIL
    (AZURE_CLIENT_SECRET is not needed for delegated/device-code flow)

Azure setup checklist (one-time):
    1. Register an app in Azure AD -> App registrations -> New registration
    2. Authentication -> Add a platform -> Mobile and desktop applications
       Check the box: https://login.microsoftonline.com/common/oauth2/nativeclient
       Set "Allow public client flows" to YES
    3. API permissions -> Add -> Microsoft Graph -> Delegated permissions:
           Mail.Send
       (No admin consent needed — the user consents on first login)
    4. Copy Application (client) ID  -> AZURE_CLIENT_ID
       Copy Directory (tenant) ID    -> AZURE_TENANT_ID
       Copy the sending mailbox UPN  -> OUTLOOK_EMAIL

First run: prints a device code + URL. Open the URL, enter the code, sign in with
           OUTLOOK_EMAIL. Token is cached to graph_token.json beside this file.
           All future runs are silent.
"""

from __future__ import annotations

import logging
import os
import re
from datetime import datetime, timedelta, timezone
from typing import Callable

from dotenv import load_dotenv

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Reminder window definitions  (code, delta before class)
# ---------------------------------------------------------------------------
REMINDER_WINDOWS: list[tuple[str, timedelta]] = [
    ("3d", timedelta(days=3)),
    ("1d", timedelta(days=1)),
    ("1h", timedelta(hours=1)),
]

# Grace period: if APScheduler missed a fire time by up to this many seconds,
# it will still run the job (handles brief app restarts).
_MISFIRE_GRACE = 3 * 3600   # 3 hours

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _validate_email(addr: str) -> None:
    if not _EMAIL_RE.match(addr.strip()):
        raise ValueError(f"Invalid email address: {addr!r}")


# ---------------------------------------------------------------------------
# GraphMailClient
# ---------------------------------------------------------------------------

import json
import pathlib

_TOKEN_CACHE_FILE = pathlib.Path(__file__).with_name("graph_token.json")


class GraphMailClient:
    """
    Sends email via Microsoft Graph API using OAuth2 delegated permissions
    (device code flow).  No admin consent required — only Mail.Send delegated.

    First run: prints a one-time device code + URL to stdout/log. Sign in with
    OUTLOOK_EMAIL at that URL.  The token (including refresh token) is persisted
    to graph_token.json beside this file, so every subsequent run is silent.
    """

    # Delegated scope — acts on behalf of the signed-in user
    _SCOPE    = ["https://graph.microsoft.com/Mail.Send"]
    _SEND_URL = "https://graph.microsoft.com/v1.0/me/sendMail"
    _SEND_TIMEOUT = 20   # seconds

    def __init__(
        self,
        client_id: str | None = None,
        tenant_id: str | None = None,
        sender_email: str | None = None,
    ) -> None:
        load_dotenv()
        self._client_id  = client_id     or os.environ.get("AZURE_CLIENT_ID",  "")
        self._tenant_id  = tenant_id     or os.environ.get("AZURE_TENANT_ID",  "")
        self._sender     = sender_email  or os.environ.get("OUTLOOK_EMAIL",    "")

        missing = [
            k for k, v in {
                "AZURE_CLIENT_ID":  self._client_id,
                "AZURE_TENANT_ID":  self._tenant_id,
                "OUTLOOK_EMAIL":    self._sender,
            }.items()
            if not v
        ]
        if missing:
            raise EnvironmentError(
                f"Missing required environment variables: {', '.join(missing)}"
            )

        self._msal_app = None

    def _get_app(self):
        if self._msal_app is None:
            try:
                import msal
            except ImportError as exc:
                raise ImportError(
                    "msal is required. Install it with: pip install msal"
                ) from exc

            cache = msal.SerializableTokenCache()
            if _TOKEN_CACHE_FILE.exists():
                cache.deserialize(_TOKEN_CACHE_FILE.read_text(encoding="utf-8"))

            self._msal_app = msal.PublicClientApplication(
                self._client_id,
                authority=f"https://login.microsoftonline.com/{self._tenant_id}",
                token_cache=cache,
            )
            self._token_cache = cache
        return self._msal_app

    def _save_cache(self) -> None:
        if self._msal_app and self._token_cache.has_state_changed:
            _TOKEN_CACHE_FILE.write_text(
                self._token_cache.serialize(), encoding="utf-8"
            )

    def _acquire_token(self) -> str:
        app = self._get_app()

        # Try cached token first (silent)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(self._SCOPE, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]

        # No cached token — start device code flow
        flow = app.initiate_device_flow(scopes=self._SCOPE)
        if "user_code" not in flow:
            raise PermissionError(f"Device flow failed to start: {flow}")

        # Print prominently so the user sees it in console or logs
        msg = (
            "\n"
            "-------------------------------------------------\n"
            "  ACTION REQUIRED -- Microsoft Graph API sign-in\n"
            "-------------------------------------------------\n"
            f"  1. Open:     {flow['verification_uri']}\n"
            f"  2. Enter:    {flow['user_code']}\n"
            f"  3. Sign in as: {self._sender}\n"
            "  Waiting for sign-in (expires in ~15 min) ...\n"
            "-------------------------------------------------\n"
        )
        print(msg, flush=True)
        log.info("Graph API device code flow initiated. Follow the prompt above.")

        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            err = result.get("error_description") or result.get("error") or str(result)
            raise PermissionError(f"Graph API authentication failed: {err}")

        self._save_cache()
        log.info("Graph API sign-in successful. Token cached to %s.", _TOKEN_CACHE_FILE)
        return result["access_token"]

    def send(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        is_html: bool = True,
    ) -> None:
        """
        Send an email immediately via Microsoft Graph API.

        Args:
            to:       One recipient address or a list of addresses.
            subject:  Email subject line.
            body:     Email body (HTML by default; set is_html=False for plain text).
            is_html:  True → Content-Type HTML, False → plain text.

        Raises:
            ValueError:       Invalid recipient address.
            EnvironmentError: Missing Azure credentials in environment.
            PermissionError:  Auth failure (bad credentials or missing Mail.Send consent).
            ConnectionError:  Network timeout or non-2xx Graph API response.
        """
        try:
            import requests
        except ImportError as exc:
            raise ImportError(
                "requests is required. Install it with: pip install requests"
            ) from exc

        recipients = [to] if isinstance(to, str) else list(to)
        for addr in recipients:
            _validate_email(addr)

        token = self._acquire_token()

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML" if is_html else "Text",
                    "content": body,
                },
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in recipients
                ],
            },
            "saveToSentItems": True,
        }

        url = self._SEND_URL.format(sender=self._sender)
        try:
            resp = requests.post(
                url,
                json=payload,
                headers={"Authorization": f"Bearer {token}"},
                timeout=self._SEND_TIMEOUT,
            )
        except requests.Timeout as exc:
            raise ConnectionError(
                f"Graph API request timed out after {self._SEND_TIMEOUT}s"
            ) from exc
        except requests.ConnectionError as exc:
            raise ConnectionError(f"Network error reaching Graph API: {exc}") from exc

        if resp.status_code == 202:
            log.debug("Graph API accepted email to %s (subject: %s)", recipients, subject)
            return

        if resp.status_code in (401, 403):
            raise PermissionError(
                f"Graph API auth/permission error {resp.status_code}: {resp.text[:300]}"
            )

        raise ConnectionError(
            f"Graph API returned unexpected status {resp.status_code}: {resp.text[:300]}"
        )


# ---------------------------------------------------------------------------
# ReminderEmail
# ---------------------------------------------------------------------------

class ReminderEmail:
    """
    High-level email interface with optional APScheduler-based scheduling.

    Usage::

        re = ReminderEmail()

        # Send immediately
        re.send_reminder("student@example.com", "Reminder", "<p>Hi!</p>")

        # Schedule for a specific time
        from datetime import datetime, timezone
        re.send_reminder(
            "student@example.com",
            "Reminder",
            "<p>Hi!</p>",
            send_at=datetime(2026, 5, 20, 9, 0, tzinfo=timezone.utc),
        )

        # Shut down the scheduler when the app exits
        re.shutdown()
    """

    def __init__(self, client: GraphMailClient | None = None) -> None:
        self._client = client or GraphMailClient()
        self._scheduler = None

    def _get_scheduler(self):
        if self._scheduler is None:
            try:
                from apscheduler.schedulers.background import BackgroundScheduler
            except ImportError as exc:
                raise ImportError(
                    "apscheduler is required for scheduling. "
                    "Install it with: pip install apscheduler"
                ) from exc
            self._scheduler = BackgroundScheduler(
                timezone="UTC",
                job_defaults={"misfire_grace_time": _MISFIRE_GRACE},
            )
            self._scheduler.start()
            log.debug("APScheduler BackgroundScheduler started.")
        return self._scheduler

    def send_reminder(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        send_at: datetime | None = None,
        is_html: bool = True,
    ) -> str | None:
        """
        Send an email immediately, or schedule it for later.

        Args:
            to:      Recipient address(es).
            subject: Subject line.
            body:    HTML (or plain-text) body.
            send_at: If None, sends right now. Otherwise, schedules for that moment.
                     Naive datetimes are treated as UTC.
            is_html: Whether body is HTML (default True).

        Returns:
            APScheduler job ID when scheduled, None when sent immediately.
        """
        if send_at is None:
            self._client.send(to, subject, body, is_html=is_html)
            return None
        return self.schedule_reminder(to, subject, body, send_at, is_html=is_html)

    def schedule_reminder(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        send_at: datetime,
        is_html: bool = True,
    ) -> str:
        """
        Schedule a one-off email. Returns the APScheduler job ID.

        send_at must be in the future; naive datetimes are treated as UTC.
        """
        if send_at.tzinfo is None:
            send_at = send_at.replace(tzinfo=timezone.utc)

        sched = self._get_scheduler()
        job = sched.add_job(
            self._client.send,
            trigger="date",
            run_date=send_at,
            args=[to, subject, body],
            kwargs={"is_html": is_html},
            misfire_grace_time=_MISFIRE_GRACE,
        )
        log.info(
            "Reminder email scheduled: to=%s  send_at=%s  job_id=%s",
            to, send_at.isoformat(), job.id,
        )
        return job.id

    def shutdown(self, wait: bool = True) -> None:
        """Shut down the background scheduler gracefully."""
        if self._scheduler and self._scheduler.running:
            self._scheduler.shutdown(wait=wait)
            log.debug("APScheduler shut down.")


# ---------------------------------------------------------------------------
# ReminderScheduler  (class-level 3d/1d/1h logic)
# ---------------------------------------------------------------------------

class ReminderScheduler:
    """
    Schedules 3-day, 1-day, and 1-hour pre-class reminder emails for a student.

    At-least-one guarantee
    ----------------------
    Windows whose send time has already passed are silently skipped, but the
    remaining future windows are still scheduled.  If *all* windows have passed
    (the class is in the past or under an hour away), a warning is logged.

    Integration example (call this when a student row is written to the sheet)::

        scheduler = ReminderScheduler()

        scheduler.schedule_class_reminders(
            student_email="student@example.com",
            class_datetime=datetime(2026, 5, 25, 9, 0, tzinfo=timezone.utc),
            location="Bartlett Hall",
            course_name="Heartsaver CPR/AED",
            already_sent="3d",   # pulled from the Google Sheet tracking column
        )

        # On app exit:
        scheduler.shutdown()
    """

    def __init__(
        self,
        reminder_email: ReminderEmail | None = None,
        subject_template: str = "Reminder: Your Upcoming CPR Lifeline Course on {course_date}",
        build_body: Callable[[str, str, str, str], str] | None = None,
        dry_run: bool = False,
        test_mode: bool = False,
    ) -> None:
        """
        Args:
            reminder_email:    ReminderEmail instance (created automatically if None).
            subject_template:  Format string with ``{course_date}`` placeholder.
            build_body:        Callable(student_email, course_date_str, location, code) -> HTML str.
                               Falls back to a minimal default if None.
            dry_run:           When True, logs intent but sends nothing.
            test_mode:         Same as dry_run.
        """
        self._re          = reminder_email or ReminderEmail()
        self._subject_tpl = subject_template
        self._build_body  = build_body or self._default_body
        self._dry_run     = dry_run
        self._test_mode   = test_mode

    @staticmethod
    def _default_body(
        student_email: str,   # noqa: ARG004
        course_date: str,
        location: str,
        code: str,
    ) -> str:
        window_label = {"3d": "3 days", "1d": "tomorrow", "1h": "1 hour"}.get(code, "soon")
        location_line = f" at <strong>{location}</strong>" if location else ""
        return (
            f"<p>This is a friendly reminder that your CPR Lifeline course is "
            f"<strong>{window_label}</strong> away.</p>"
            f"<p>Your class is scheduled for <strong>{course_date}</strong>"
            f"{location_line}.</p>"
            f"<p>Please reply to this email if you have any questions.</p>"
        )

    def schedule_class_reminders(
        self,
        student_email: str,
        class_datetime: datetime,
        location: str = "",
        course_name: str = "",
        already_sent: str = "",
    ) -> list[str]:
        """
        Schedule future 3d/1d/1h reminders for one student.

        Args:
            student_email:  Recipient address.
            class_datetime: When the class starts (naive → treated as UTC).
            location:       Location name used in the email body.
            course_name:    Course name used in log messages.
            already_sent:   Pipe-separated reminder codes already sent,
                            e.g. ``"3d"`` or ``"3d|1d"``.  These are skipped.

        Returns:
            List of APScheduler job IDs for the scheduled jobs (empty if
            dry_run / test_mode or if all windows have already passed).
        """
        if self._dry_run or self._test_mode:
            flag = "DRY_RUN" if self._dry_run else "TEST_MODE"
            log.info(
                "%s: skipping reminder scheduling for %s (class %s).",
                flag, student_email, class_datetime.isoformat(),
            )
            return []

        _validate_email(student_email)

        if class_datetime.tzinfo is None:
            class_datetime = class_datetime.replace(tzinfo=timezone.utc)

        now        = datetime.now(tz=timezone.utc)
        sent_codes = set(filter(None, already_sent.split("|")))
        job_ids: list[str] = []

        for code, delta in REMINDER_WINDOWS:
            if code in sent_codes:
                log.debug("  %s reminder already sent for %s — skipping.", code, student_email)
                continue

            send_at = class_datetime - delta

            if send_at <= now:
                log.debug(
                    "  %s reminder window already passed for %s (would have sent at %s).",
                    code, student_email, send_at.isoformat(),
                )
                continue

            course_date_str = class_datetime.strftime("%m/%d/%Y")
            subject = self._subject_tpl.format(course_date=course_date_str)
            body    = self._build_body(student_email, course_date_str, location, code)

            try:
                job_id = self._re.schedule_reminder(student_email, subject, body, send_at)
                job_ids.append(job_id)
                log.info(
                    "  Scheduled %s reminder for %s (%s @ %s) -> fires %s",
                    code,
                    student_email,
                    course_name or "class",
                    location or "unknown",
                    send_at.isoformat(),
                )
            except Exception as exc:
                log.error(
                    "  Failed to schedule %s reminder for %s: %s",
                    code, student_email, exc,
                )

        # At-least-one warning
        all_codes = {c for c, _ in REMINDER_WINDOWS}
        if not job_ids and not sent_codes.issuperset(all_codes):
            log.warning(
                "No future reminder windows remain for %s (class at %s). "
                "All windows have passed or the class is within 1 hour.",
                student_email,
                class_datetime.isoformat(),
            )

        return job_ids

    def shutdown(self, wait: bool = True) -> None:
        """Propagate shutdown to the underlying ReminderEmail scheduler."""
        self._re.shutdown(wait=wait)


# ---------------------------------------------------------------------------
# send_due_reminders  — scan-based alternative (drop-in for send_reminder_emails)
# ---------------------------------------------------------------------------

def send_due_reminders(
    ws,
    *,
    client: GraphMailClient | None = None,
    build_body: Callable[[str, str, str, str], str] | None = None,
    subject_template: str = "Reminder: Your Upcoming CPR Lifeline Course on {course_date}",
    dry_run: bool = False,
    test_mode: bool = False,
) -> int:
    """
    Scan-based reminder sender — a drop-in replacement for the Playwright-based
    ``send_reminder_emails(page)`` function in aha_registration_bot.py.

    Reads the RQI/Acuity Google Sheet worksheet ``ws``, finds students whose
    class is exactly 3 days, 1 day, or (within 90 minutes of) 1 hour away, and
    sends reminders via the Graph API.  Stamps each sent code in the sheet's
    ``Reminder email sent`` column to prevent duplicates.

    Column expectations (same as the existing sheet layout):
        Email          → index 6
        LocationName   → index 1
        JobName        → index 8
        HireDate       → index 9  (format MM/DD/YYYY)
        Reminder email sent → last appended column (found by header name)

    Args:
        ws:               gspread Worksheet for the RQI Registration Sheet.
        client:           GraphMailClient (created from env vars if None).
        build_body:       Callable(student_email, course_date, location, code) -> HTML.
        subject_template: Subject format string with ``{course_date}``.
        dry_run:          Log without sending or writing to sheet.
        test_mode:        Same as dry_run.

    Returns:
        Number of reminder emails sent this call.
    """
    if dry_run or test_mode:
        flag = "DRY_RUN" if dry_run else "TEST_MODE"
        log.info("%s: send_due_reminders called — no emails will be sent.", flag)
        return 0

    mail_client = client or GraphMailClient()
    _body_fn    = build_body or ReminderScheduler._default_body

    all_rows = ws.get_all_values()
    if len(all_rows) < 2:
        log.info("RQI sheet has no student rows — reminder step skipped.")
        return 0

    headers = all_rows[0]
    col = {h.strip(): i for i, h in enumerate(headers)}

    email_idx   = col.get("Email",              6)
    loc_idx     = col.get("LocationName",       1)
    job_idx     = col.get("JobName",            8)
    date_idx    = col.get("HireDate",           9)
    reminder_idx = col.get("Reminder email sent")

    if reminder_idx is None:
        log.error("'Reminder email sent' column not found — aborting reminder step.")
        return 0

    now        = datetime.now(tz=timezone.utc)
    total_sent = 0

    # Window definitions for scan mode: (code, min_hours_before, max_hours_before)
    # "due" means the class is between min and max hours away right now.
    scan_windows = [
        ("3d", 71.0, 73.0),   # ±1 h window centred on 72 h
        ("1d", 23.0, 25.0),   # ±1 h window centred on 24 h
        ("1h",  0.0,  1.5),   # within the next 90 minutes
    ]

    def _row_get(row: list, idx: int) -> str:
        return row[idx].strip() if idx < len(row) else ""

    def _reminder_sent(cell: str, code: str) -> bool:
        return f"|{code}|" in f"|{cell}|"

    def _append_code(cell: str, code: str) -> str:
        return f"{cell}|{code}" if cell else code

    for row_num, row in enumerate(all_rows[1:], start=2):
        email    = _row_get(row, email_idx)
        date_str = _row_get(row, date_idx)
        if not email or not date_str:
            continue

        try:
            _validate_email(email)
        except ValueError:
            log.warning("  Skipping invalid email address %r in row %d.", email, row_num)
            continue

        try:
            class_date = datetime.strptime(date_str, "%m/%d/%Y")
            # Treat class as starting at 08:00 local time (UTC assumed here).
            class_dt = class_date.replace(hour=8, tzinfo=timezone.utc)
        except ValueError:
            log.warning("  Could not parse date %r for %s — skipping.", date_str, email)
            continue

        hours_until = (class_dt - now).total_seconds() / 3600
        sent_cell   = _row_get(row, reminder_idx)
        location    = _row_get(row, loc_idx)
        course      = _row_get(row, job_idx) or "class"

        for code, min_h, max_h in scan_windows:
            if not (min_h <= hours_until < max_h):
                continue
            if _reminder_sent(sent_cell, code):
                continue

            subject = subject_template.format(course_date=date_str)
            body    = _body_fn(email, date_str, location, code)

            try:
                mail_client.send(email, subject, body, is_html=True)
                sent_cell = _append_code(sent_cell, code)
                ws.update_cell(row_num, reminder_idx + 1, sent_cell)
                total_sent += 1
                log.info(
                    "  Reminder (%s) sent via Graph API -> %s (%s on %s).",
                    code, email, course, date_str,
                )
            except Exception as exc:
                log.error("  Failed to send %s reminder to %s: %s", code, email, exc)

    log.info("send_due_reminders complete — %d email(s) sent.", total_sent)
    return total_sent


# ---------------------------------------------------------------------------
# Quick usage example (run this file directly to test credentials)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    if len(sys.argv) < 2:
        print(
            "Usage: python reminder_email.py <recipient@example.com>\n"
            "Sends a test email immediately to verify your Azure credentials."
        )
        sys.exit(1)

    recipient = sys.argv[1]

    print(f"Sending test reminder email to {recipient} …")
    print("(If prompted, open the URL and enter the device code to sign in.)\n")
    try:
        client = GraphMailClient()
        client.send(
            to=recipient,
            subject="[TEST] AHA Bot reminder email check",
            body=(
                "<h2>Credentials OK</h2>"
                "<p>If you received this, your Azure Graph API setup is working correctly.</p>"
            ),
        )
        print("Done — email accepted by Graph API (HTTP 202).")
    except (EnvironmentError, PermissionError, ConnectionError, ValueError) as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    # --- Scheduling smoke-test ---
    print("\nScheduling a reminder 5 seconds from now …")
    from datetime import timezone as _tz
    re = ReminderEmail(client=client)
    fire_at = datetime.now(tz=_tz.utc) + timedelta(seconds=5)
    job_id = re.schedule_reminder(recipient, "[TEST] Scheduled reminder", "<p>Scheduled OK</p>", fire_at)
    print(f"  Job {job_id} scheduled for {fire_at.isoformat()}")

    import time
    print("  Waiting 10 s for the job to fire …")
    time.sleep(10)
    re.shutdown()
    print("Scheduler shut down. Check your inbox.")
