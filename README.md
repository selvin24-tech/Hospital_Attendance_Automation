# Hospital Attendance Automation

This repository includes a Python automation template for logging into the hospital attendance portal, capturing attendance details, and sending a report by email or WhatsApp.

## What you need to provide

1. `attendance_url`
   - The exact login URL, e.g. `https://gleneagles.myadrenalin.com/AdrenalinMax/#/`

2. Login field details
   - If using Playwright: the CSS or text selectors for username, password, and login button.
   - If using PyAutoGUI: screenshots of the username field, password field, login button, menu items, and any steps needed after login.

3. Navigation flow after login
   - The exact menu item names and the order to click them (for example: Attendance → Attendance & Leave).
   - Screenshots of the attendance page and the calendar popup if you need a specific shift view.

4. Report contents
   - Whether you want a screenshot only or text extracted from the portal.
   - Which fields matter: shift details, punch in/out times, present/absent days, total hours.

5. Delivery method
   - For WhatsApp: the phone number and whether you want the site screenshot attached or a text summary.
   - For email: SMTP settings and sender/recipient addresses.

6. Scheduling details
   - When and how often this should run (daily, weekdays only, time of day).

## Setup

1. Install Python 3.10+.

2. Install packages:

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
```

3. Open this folder in Visual Studio Code and edit `config.json`.
   - Set `username`, `password`.
   - Set `use_playwright` to `true` for browser automation.
   - Update selectors if needed.
   - Add the WhatsApp phone number or enable email delivery.

4. If you use PyAutoGUI instead of Playwright, capture reference images and place them under `images/`.
   - Update the `images` section in `config.json`.

## Running

```powershell
python attendance_automation.py
```

## Notes

- Playwright is more reliable than PyAutoGUI for browser automation because it uses the page DOM instead of screen coordinates.
- If the website has an unusual login flow or if selectors do not work, send screenshots of the login page and the attendance page so the script can be updated.
- Do not store your real password in a shared file. You can replace it later with environment variables or a secure vault.
