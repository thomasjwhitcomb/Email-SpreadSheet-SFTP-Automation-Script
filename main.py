from src.email_handler import parse_email
from src.sheets_handler import get_sheet


def main():

    print("CPR Automation System Initialized")

    sample_email = """
    Subject: AHA Registration Confirmation

    Student Name: John Smith
    Email: johnsmith@email.com
    Phone: 916-555-1234
    Course: BLS Provider
    Date: March 10, 2026
    """

    email_data = parse_email(sample_email)

    print("Parsed student data:", email_data)

    # Replace with your actual RQI sheet URL
    RQI_SHEET_URL = "https://docs.google.com/spreadsheets/d/1mw7St3F-iITUtkQCjKFU0PmRbjNYiuuctEMgQtMDOHE/edit?gid=0#gid=0"

    worksheet = get_sheet(RQI_SHEET_URL)

    worksheet.append_row([
        email_data["name"],
        email_data["email"],
        email_data["phone"],
        email_data["course"],
        email_data["date"],
    ])

    print("Data pushed to Google Sheet successfully.")


if __name__ == "__main__":
    main()