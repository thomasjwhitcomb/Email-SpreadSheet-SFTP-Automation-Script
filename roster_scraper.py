cat > /Users/sachintomy/Desktop/aha_project/roster_scraper.py << 'EOF'
import os
import re
import csv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

ATLAS_URL          = "https://atlas.heart.org/"
TSC_URL            = "https://atlas.heart.org/organisation/class-listing?applyTsFilter=true"
TRAINING_SITE_NAME = os.getenv("ATLAS_SITE", "Sac State")
MAX_PAGES          = 25

def login(page, user: str, password: str) -> None:
    """Log into the AHA Atlas website."""
    print("Logging in...")
    page.goto(ATLAS_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    page.locator("div").filter(
        has_text=re.compile(r"^Sign In \| Sign Up$")
    ).first.click()
    page.wait_for_timeout(2000)
    page.get_by_role("textbox", name="Username / Email").wait_for(state="visible", timeout=15000)
    page.get_by_role("textbox", name="Username / Email").fill(user)
    page.get_by_role("textbox", name="Password").fill(password)
    page.get_by_role("button", name="Sign In").click()
    page.wait_for_timeout(8000)
    print(f"✅ Logged in — {page.url}")

def go_to_training_site_classes(page) -> None:
    """Navigate directly to the Training Site Classes page."""
    page.goto(TSC_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)

def select_site(page) -> None:
    """Select the training site from the dropdown."""
    print(f"Selecting site: {TRAINING_SITE_NAME}...")
    page.locator(".css-19bb58m").first.click()
    page.wait_for_timeout(2000)
    try:
        page.get_by_role("option", name=TRAINING_SITE_NAME).click(timeout=8000)
    except Exception:
        print(f"⚠️  Could not find site '{TRAINING_SITE_NAME}' — check ATLAS_SITE env var")
        raise
    page.wait_for_timeout(3000)
    print(f"✅ Site selected: {TRAINING_SITE_NAME}")

def get_class_count(page) -> int:
    """Count how many classes are listed on the current page."""
    return page.locator("div").filter(
        has_text=re.compile(r"^ViewEditDuplicateCancel$")
    ).count()

def open_class_by_index(page, index: int) -> str:
    """Open a class by its position in the list. Returns the class name."""
    rows = page.locator("div").filter(
        has_text=re.compile(r"^ViewEditDuplicateCancel$")
    )
    row = rows.nth(index)
    try:
        class_name = " ".join(
            row.locator("..").inner_text().split()
        ).split("View")[0].strip()
    except Exception:
        class_name = f"Class {index + 1}"
    try:
        row.scroll_into_view_if_needed()
        page.wait_for_timeout(300)
        row.hover()
        page.wait_for_timeout(800)
    except Exception:
        pass
    page.get_by_test_id(f"action-menus-0-{index}").evaluate("el => el.click()")
    page.wait_for_timeout(3000)
    print(f"  Opened: {class_name}")
    return class_name

def clean_text(value: str) -> str:
    return " ".join(value.split()).strip()

def extract_email(text: str) -> str:
    m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    return m.group(0) if m else ""

def extract_phone(text: str) -> str:
    m = re.search(r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", text)
    return m.group(0) if m else ""

def extract_name(text: str, email: str, phone: str) -> str:
    """Extract student name by removing email, phone, and status words."""
    result = text
    if email:
        result = result.replace(email, "")
    if phone:
        result = result.replace(phone, "")
    for status in ["Not Registered", "Registered", "Invited", "Completed", "Pending"]:
        result = result.replace(status, "")
    return clean_text(result)

def get_course_type(page) -> str:
    """Extract the course type from the page heading."""
    try:
        return clean_text(page.locator("h1").first.inner_text())
    except Exception:
        return ""

def get_class_date(page) -> str:
    """Extract the class date from the class detail page."""
    try:
        text = page.locator("text=/Date | Time/").locator("..").inner_text()
        m = re.search(r"\d{2}-\d{2}-\d{4}", text)
        return m.group(0) if m else ""
    except Exception:
        return ""

def get_rows(page):
    """Find student roster rows on the current page."""
    for locator in [page.locator("tr"), page.locator("[role='row']")]:
        if locator.count() > 0:
            return locator
    return None

def scrape_students_from_class(page, class_name: str) -> list[dict]:
    """Scrape all enrolled students from the currently open class."""
    students = []
    course_type = get_course_type(page)
    class_date  = get_class_date(page)

    for page_num in range(1, MAX_PAGES + 1):
        page.wait_for_timeout(1500)
        rows = get_rows(page)
        if rows is None:
            break
        count = rows.count()
        print(f"    Page {page_num}: {count} row(s)")
        for i in range(count):
            try:
                text = clean_text(rows.nth(i).inner_text())
            except Exception:
                continue
            if not text or "no data found" in text.lower():
                continue
            email = extract_email(text)
            phone = extract_phone(text)
            if not email and not phone:
                continue
            name = extract_name(text, email, phone)
            students.append({
                "name":        name,
                "email":       email,
                "phone":       phone,
                "course_type": course_type,
                "class_date":  class_date,
                "class":       class_name,
            })
            print(f"    Found: {name} | {email} | {phone}")
        next_btn = page.get_by_test_id(f"pagination-link-{page_num + 1}")
        try:
            if next_btn.count() > 0 and next_btn.is_visible():
                next_btn.click()
                page.wait_for_timeout(2500)
            else:
                break
        except Exception:
            break
    return students

def collect_all_students(page) -> list[dict]:
    """Loop through every class and collect all enrolled students."""
    all_students = []
    go_to_training_site_classes(page)
    select_site(page)
    count = get_class_count(page)
    print(f"\nFound {count} class(es) to search.\n")
    for i in range(count):
        print(f"Opening class {i + 1} of {count}...")
        try:
            go_to_training_site_classes(page)
            select_site(page)
            page.wait_for_timeout(2000)
            class_name = open_class_by_index(page, i)
            found = scrape_students_from_class(page, class_name)
            print(f"  → {len(found)} student(s) found\n")
            all_students.extend(found)
        except Exception as e:
            print(f"  ⚠️  Error on class {i + 1}: {e}\n")
            continue
    return all_students

def save_to_csv(students: list[dict], filename: str = "all_students.csv") -> None:
    """Save collected student data to a CSV file."""
    if not students:
        print("No students found.")
        return
    fieldnames = ["name", "email", "phone", "course_type", "class_date", "class"]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(students)
    print(f"✅ Saved {len(students)} student(s) to {filename}")

def run() -> None:
    user     = os.getenv("ATLAS_USER")
    password = os.getenv("ATLAS_PASS")
    if not user or not password:
        print("❌ Missing credentials. Set environment variables:")
        print('   export ATLAS_USER="your_email"')
        print('   export ATLAS_PASS="your_password"')
        print('   export ATLAS_SITE="Sac State"   # optional')
        return
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=300)
        context = browser.new_context()
        page    = context.new_page()
        try:
            login(page, user, password)
            students = collect_all_students(page)
            save_to_csv(students)
            input("\nPress Enter to close the browser...")
        except PlaywrightTimeoutError as e:
            print(f"⏱  Timeout error: {e}")
            input("Press Enter to close...")
        except Exception as e:
            print(f"💥 Unexpected error: {e}")
            input("Press Enter to close...")
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    run()
EOF
