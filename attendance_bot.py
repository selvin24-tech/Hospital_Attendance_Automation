from datetime import datetime
from openpyxl import Workbook
import os

from dotenv import load_dotenv
from playwright.sync_api import Playwright, sync_playwright

load_dotenv(override=True)

USERNAME = os.getenv("HR_USERNAME")
PASSWORD = os.getenv("HR_PASSWORD")


def run(playwright: Playwright):

    browser = playwright.chromium.launch(headless=False)

    context = browser.new_context()

    page = context.new_page()

    page.goto("https://gleneagles.myadrenalin.com/AdrenalinMax/#/")

    # Login
    page.get_by_role("textbox", name="User ID").fill(USERNAME)
    page.get_by_role("textbox", name="Password").fill(PASSWORD)
    page.get_by_role("button", name="Login").click()

    page.wait_for_load_state("networkidle")

    # Attendance Menu
    page.get_by_label("Attendance, claims & requests").click()
    page.locator("a").filter(has_text="Attendance & leave").click()

    # Attendance iframe
    frame = page.locator("#commonFormRender").content_frame

    # Wait for Show Time checkbox
    frame.get_by_role("checkbox", name="Show time").wait_for(timeout=15000)

    # Tick Show Time
    frame.get_by_role("checkbox", name="Show time").check()

    # Wait for Show Time checkbox
    frame.get_by_role("checkbox", name="Show time").wait_for(timeout=15000)

# Tick Show Time
    frame.get_by_role("checkbox", name="Show time").check()

# Wait a little longer after checking
    page.wait_for_timeout(8000)

    # Date & Time
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    current_datetime = now.strftime("%Y-%m-%d %H:%M")

    # Screenshot
    page.screenshot(
        path=f"screenshots/attendance_{today}.png",
        full_page=True
    )

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"

    ws.append(["Date & Time", "Information", "Comment"])
    ws.append([
        current_datetime,
        "Attendance & Leave Page",
        "Show Time Enabled"
    ])

    wb.save(f"reports/daily_report_{today}.xlsx")

    print("✅ Screenshot Saved")
    print("✅ Excel Saved")

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)