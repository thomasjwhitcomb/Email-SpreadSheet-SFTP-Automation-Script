"""
test_acuity_parsing.py
----------------------
Reads every .msg file in the Acuity Emails folder, runs each one through the
bot's real parser, and writes the results to the RQI Registration Google Sheet.

This is functionally identical to what the bot does when it reads Acuity emails
from Outlook - it uses the exact same parse_acuity_email_body() function and
the same append_acuity_to_sheet() writer - but bypasses Outlook entirely.

Usage (with venv active):
    python test_acuity_parsing.py

Output:
    Prints a summary table and writes records to the RQI Registration Sheet.
"""

import os
import re
import sys
from pathlib import Path

# --- Locate project root and load .env ---
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
load_dotenv(ROOT / ".env")

from aha_registration_bot import (
    parse_acuity_email_body,
    append_acuity_to_sheet,
    AcuityRecord,
    _course_to_group,
    _location_to_id,
)

MSG_FOLDER = Path(r"C:\Users\thoma\Downloads\Acuity Emails")

# --- Extract readable text from a .msg binary file ---
def extract_msg_body(path: Path) -> tuple[str, str, str]:
    """
    Return (sender_display_name, subject, body_text) extracted from a .msg file.

    Uses UTF-16-LE decoding (standard MSG encoding) and falls back to ASCII
    string extraction.  Good enough for Acuity's plain-text email bodies.
    """
    data = path.read_bytes()

    # UTF-16-LE decode - this is how Outlook stores the body in .msg files
    text = data.decode("utf-16-le", errors="replace")

    # --- Subject ---
    # Subject appears as a long readable line in the decoded text
    subject = ""
    m = re.search(r"New Appointment[^\r\n]{10,}", text)
    if not m:
        m = re.search(r"Appointment Rescheduled[^\r\n]{10,}", text)
    if m:
        subject = m.group(0).strip()

    # --- Sender display name ---
    # Format in MSG: "Aanyah Sowell SMTP no-reply@acuityscheduling.com"
    sender_name = ""
    m = re.search(r"for\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s*\r?\n", text)
    if m:
        sender_name = m.group(1).strip()

    # --- Body: find the Acuity structured block ---
    # The key block starts with "Appointment Scheduled" or "Scheduled by a client"
    body = ""
    m = re.search(
        r"((?:Scheduled by a client|Appointment Scheduled).+?)"
        r"(?:Add to iCal|Add to Google|Name:\s)",
        text, re.DOTALL
    )
    if m:
        body = m.group(1)
        # Also append the Name/Phone block that follows
        name_block = re.search(r"Name:\s.+?(?=\s{10,}|\Z)", text, re.DOTALL)
        if name_block:
            body += "\n" + name_block.group(0)
    else:
        # Fallback: grab any block containing "What" and "When"
        m = re.search(r"(for\s+\S.+?Where\s+.+?)(?:\r?\n\r?\n)", text, re.DOTALL)
        if m:
            body = m.group(1)

    # Clean up replacement characters from decode errors
    body = body.replace("ï¿½", "")

    return sender_name, subject, body


# --- Main ---
def main():
    msg_files = sorted(MSG_FOLDER.glob("*.msg"))
    if not msg_files:
        print(f"No .msg files found in {MSG_FOLDER}")
        sys.exit(1)

    print(f"Found {len(msg_files)} .msg file(s). Parsing...\n")
    print(f"{'#':<3}  {'Name':<25}  {'Course':<35}  {'Date':<12}  {'Location':<30}  {'Email'}")
    print("-" * 130)

    records: list[AcuityRecord] = []

    for i, path in enumerate(msg_files, 1):
        sender_name, subject, body = extract_msg_body(path)

        parsed = parse_acuity_email_body(body, subject=subject, sender_name=sender_name)

        raw_name   = parsed.get("student_name", "")
        name_parts = raw_name.split()
        first_name  = name_parts[0] if len(name_parts) >= 1 else ""
        middle_name = name_parts[1] if len(name_parts) == 3 else ""
        last_name   = name_parts[-1] if len(name_parts) >= 2 else ""

        # Extract student email - exclude known non-student domains
        _EXCLUDE = {"acuityscheduling.com", "cprlifeline.net", "rqi1stop.com",
                    "squarespace.com", "outlook.com", "microsoft.com"}
        student_email = ""
        for m in re.finditer(r"[\w.+\-]+@[\w.\-]+\.\w+", body):
            addr   = m.group(0).lower()
            domain = addr.split("@", 1)[-1]
            if not any(domain == ex or domain.endswith("." + ex) for ex in _EXCLUDE):
                student_email = m.group(0)
                break

        course_date = parsed.get("course_date", "")
        course_name = parsed.get("course_name", "")
        location    = parsed.get("course_location", "")

        record = AcuityRecord(
            first_name=first_name,
            middle_name=middle_name,
            last_name=last_name,
            email=student_email,
            user_id=student_email,
            location_name=_location_to_id(location),
            job_name="",
            hire_date="",
            status="Active",
            group=_course_to_group(course_name),
        )
        records.append(record)

        full_name = f"{first_name} {last_name}".strip() or "(unparsed)"
        print(
            f"{i:<3}  {full_name:<25}  {course_name:<35}  "
            f"{course_date:<12}  {location:<30}  {student_email or '(none)'}"
        )

    print(f"\nParsed {len(records)} record(s).")

    # --- Write to Google Sheet ---
    answer = input("\nWrite these records to the RQI Registration Sheet? (yes/no): ").strip().lower()
    if answer not in ("yes", "y"):
        print("Aborted - nothing written.")
        return

    print("Writing to Google Sheet...")
    try:
        append_acuity_to_sheet(records)
        print(f"Done. {len(records)} record(s) written to the RQI Registration Sheet.")
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()

