from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import os

load_dotenv()

EMAIL = os.getenv("OUTLOOK_EMAIL")
PASSWORD = os.getenv("OUTLOOK_PASSWORD")

def login_outlook():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        print("Opening Outlook...")
        page.goto("https://outlook.office.com")

        page.fill("input[type='email']", EMAIL)
        page.click("input[type='submit']")

        page.wait_for_timeout(2000)

        page.fill("input[type='password']", PASSWORD)
        page.click("input[type='submit']")

        page.wait_for_timeout(5000)

        print("Login attempt finished.")

        browser.close()

if __name__ == "__main__":
    login_outlook()