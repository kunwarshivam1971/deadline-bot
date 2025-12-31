import openpyxl
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText

# Load credentials from .env file
load_dotenv()
SENDER = os.getenv("OUTLOOK_EMAIL")   # Your Gmail or Outlook address
PASSWORD = os.getenv("OUTLOOK_PASSWORD")  # App password
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")  # Use smtp.office365.com for Outlook
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))

print("📧 Using sender:", SENDER)

# Load Excel file
wb = openpyxl.load_workbook("deadlines.xlsx")
sheet = wb.active

today = datetime.today().date()
print("📅 Today:", today)

# Iterate through tasks
for row in sheet.iter_rows(min_row=2, values_only=True):
    task, deadline, recipient_emails = row
    print("➡️ Checking task:", task, "Deadline:", deadline)

    # Convert deadline to date
    if isinstance(deadline, datetime):
        deadline_date = deadline.date()
    else:
        deadline_date = datetime.strptime(str(deadline), "%Y-%m-%d").date()

    # If deadline is within 2 days
    if deadline_date - today <= timedelta(days=2):
        print("⚠️ Sending reminder for:", task)

        # Split and clean recipient emails
        recipients = [email.strip() for email in recipient_emails.split(",")]

        # Message body
        body = f"""
        Hello,

        This is a reminder that the task **{task}** is due on {deadline_date}.

        Please make sure to complete and submit it before the deadline.

        Regards,
        Deadline Bot
        """

        msg = MIMEText(body, "plain")
        msg["Subject"] = f"Deadline Reminder: {task}"
        msg["From"] = SENDER
        msg["To"] = ", ".join(recipients)

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER, PASSWORD)
                server.sendmail(SENDER, recipients, msg.as_string())
            print("✅ Email sent to:", recipients)
        except Exception as e:
            print("❌ Error sending email:", e)

# Log run
with open("log.txt", "a") as log:
    log.write(f"{datetime.now()} - Checked deadlines\n")
