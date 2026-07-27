from datetime import datetime
import os
import time

import pyautogui
import pyperclip
from openpyxl import Workbook

from dotenv import load_dotenv
from playwright.sync_api import Playwright, sync_playwright

load_dotenv(override=True)

USERNAME = os.getenv("HR_USERNAME")
PASSWORD = os.getenv("HR_PASSWORD")


def run(playwright: Playwright):

    browser = playwright.chromium.launch(
    headless=False,
    args=["--start-maximized"]
    )

    context = browser.new_context(
    no_viewport=True
    )

    page = context.new_page()

    page.goto("https://gleneagles.myadrenalin.com/AdrenalinMax/#/")

    # Login
    page.get_by_role("textbox", name="User ID").fill(USERNAME)
    page.get_by_role("textbox", name="Password").fill(PASSWORD)
    page.get_by_role("button", name="Login").click()

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)

    # Attendance Menu
    page.get_by_label("Attendance, claims & requests").click()
    page.locator("a").filter(has_text="Attendance & leave").click()

    # Attendance iframe
    frame = page.locator("#commonFormRender").content_frame

    # Wait for Show Time checkbox
    frame.get_by_role("checkbox", name="Show time").wait_for(timeout=15000)

    # Tick Show Time
    frame.get_by_role("checkbox", name="Show time").check()

    # Click Refresh
    frame.get_by_text("Refresh").click()

    # Wait for the calendar to reload
    page.wait_for_timeout(5000)

    # Create screenshots folder if it doesn't exist
    os.makedirs("screenshots", exist_ok=True)

    today = datetime.now().strftime("%Y-%m-%d")

    frame.locator("body").screenshot(
    path=f"screenshots/attendance_{today}.png"
    )

    print("✅ Screenshot saved.")

    # Click inside attendance page
    frame.locator("body").click()
    time.sleep(1)

    # Copy everything
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)

    pyautogui.hotkey("ctrl", "c")
    time.sleep(2)

    # Read clipboard
    clipboard_data = pyperclip.paste()

    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"

    lines = [line.strip() for line in clipboard_data.splitlines() if line.strip()]

    ws.append(["Date", "In Time", "Out Time", "Working Hours"])

    i = 0

    while i < len(lines):

        # Look for a day number like 01, 02, 15...
        if lines[i].isdigit() and len(lines[i]) <= 2:

            # Make sure there are enough lines left
            if i + 3 < len(lines):

                # The 4th line must contain "Hrs"
                if "Hrs" in lines[i + 3]:

                    date = lines[i]
                    in_time = lines[i + 1]
                    out_time = lines[i + 2]
                    hours = lines[i + 3]

                    ws.append([date, in_time, out_time, hours])

                    i += 4
                    continue

        i += 1

        os.makedirs("reports", exist_ok=True)

    today = datetime.now().strftime("%Y-%m-%d")
    save_path = os.path.join("reports", f"Attendance_{today}.xlsx")

    wb.save(save_path)

    print("✅ Attendance exported successfully!")
    print(save_path)

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)