
import gspread
from google.oauth2.service_account import Credentials

# Define Google API scope

SCOPE = [

    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"

]

def get_sheet(sheet_url):

    """
    Authenticates and return worksheet object.
    """

    creds = Credentials.from_service_account_file(
        "config/google_credentials.json",
        scopes=SCOPE

    )

    client = gspread.authorize(creds)
    sheet = client.open_by_url(sheet_url)
    return sheet.sheet1