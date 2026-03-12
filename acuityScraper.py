import re
from dataclasses import dataclass
from typing import Optional
from dateutil import parser as dateparser

ADMIN_MARKER = "The info below is just sent to you as the admin"

NAME_RE  = re.compile(r"^Name\s*:\s*(.+)$",  re.IGNORECASE | re.MULTILINE)
PHONE_RE = re.compile(r"^Phone\s*:\s*(.+)$", re.IGNORECASE | re.MULTILINE)
EMAIL_RE = re.compile(r"^Email\s*:\s*(.+)$", re.IGNORECASE | re.MULTILINE)
WHEN_RE  = re.compile(r"^When\s+(.+)$",      re.IGNORECASE | re.MULTILINE)

def pick_course(text: str) -> str:
    t = (text or "").lower()
    if "pals" in t: return "PALS"
    if "bls" in t:  return "BLS"
    if "acls" in t: return "ACLS"
    return ""

def split_name(full_name: str):
    parts = (full_name or "").strip().split()
    if not parts:
        return ("", "", "")
    if len(parts) == 1:
        return (parts[0], "", "")
    first = parts[0]
    last = parts[-1]
    middle = " ".join(parts[1:-1]) if len(parts) > 2 else ""
    return (first, middle[:1] if middle else "", last)

def extract_field(regex: re.Pattern, text: str) -> str:
    m = regex.search(text or "")
    return m.group(1).strip() if m else ""

def extract_admin_section(body: str) -> str:
    if not body:
        return ""
    idx = body.lower().find(ADMIN_MARKER.lower())
    return body[idx:] if idx != -1 else ""

def extract_when_date(body: str) -> str:
    m = WHEN_RE.search(body or "")
    if not m:
        return ""
    when_line = m.group(1).strip()
    when_line = re.sub(r"\(.*?\)\s*$", "", when_line).strip()  # remove "(1 hour)" etc.
    try:
        dt = dateparser.parse(when_line, fuzzy=True)
        return f"{dt.month}/{dt.day}/{str(dt.year)[-2:]}"  # e.g. 2/26/26
    except Exception:
        return when_line  # fallback raw

def scrape_acuity_admin_fields(email_body: str) -> dict:
    """
    Input: email body text (as a string).
    Output: dict matching the columns
    """
    admin_section = extract_admin_section(email_body)

    full_name = extract_field(NAME_RE, admin_section)
    phone     = extract_field(PHONE_RE, admin_section)
    email     = extract_field(EMAIL_RE, admin_section)

    course = pick_course(email_body)
    date_str = extract_when_date(email_body)

    first, m_init, last = split_name(full_name)

    return {
        "EMAIL": email,
        "First Name": first,
        "M": m_init,
        "Last Name": last,
        "Phone": phone,
        "Course": course,
        "Date": date_str,
        "Acuity Regist.": "YES",
    }

# ---- Example usage with a fake sample email body ----
if __name__ == "__main__":
    sample_body = """
New Appointment

When Thursday, February 26, 2026 5:00pm CST (1 hour)
Service ACLS Skills Check Only

The info below is just sent to you as the admin
Name: Jayden Kearney
Phone: (877) 422-7755
Email: jayden@example.com
"""
    print(scrape_acuity_admin_fields(sample_body))