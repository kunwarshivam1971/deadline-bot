import openpyxl
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText

# Load credentials from .env file
load_dotenv()
SENDER = os.getenv("OUTLOOK_EMAIL")
PASSWORD = os.getenv("OUTLOOK_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
RECIPIENT = SENDER  # send to yourself for testing

print("📧 Using sender:", SENDER)

# Load Excel file
wb = openpyxl.load_workbook("deadlines.xlsx")
sheet = wb.active

today = datetime.today().date()
print("📅 Today:", today)

# Iterate through tasks
for row in sheet.iter_rows(min_row=2, values_only=True):
    task, deadline = row
    print("➡️ Checking task:", task, "Deadline:", deadline)

    # Handle Excel date formats safely
    if isinstance(deadline, datetime):
        deadline_date = deadline.date()
    else:
        deadline_date = datetime.strptime(str(deadline), "%Y-%m-%d").date()

    # If deadline is within 2 days
    if deadline_date - today <= timedelta(days=2):
        print("⚠️ Sending reminder for:", task)
        msg = MIMEText(f"Reminder: {task} is due on {deadline_date}")
        msg["Subject"] = f"Deadline Reminder: {task}"
        msg["From"] = SENDER
        msg["To"] = RECIPIENT

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER, PASSWORD)
                server.send_message(msg)
            print("✅ Email sent for:", task)
        except Exception as e:
            print("❌ Error sending email:", e)

# Log run
with open("log.txt", "a") as log:
    log.write(f"{datetime.now()} - Checked deadlines\n")