# 📅 Deadline Reminder Bot

A Python script that reads tasks and deadlines from an Excel file and automatically sends email reminders.

---

## 🚀 Features
- Reads tasks and deadlines from `deadlines.xlsx`
- Sends reminders if a deadline is within 2 days
- Supports multiple recipients (comma‑separated emails in Excel)
- Secure credential handling via `.env` file
- Logs each run in `log.txt`

---

## 📊 Excel Format
Your `deadlines.xlsx` should have 3 columns:

| Task           | Deadline   | Recipient Email                                      |
|----------------|------------|------------------------------------------------------|
| Assignment     | 2026-01-01 | user1@example.com, user2@example.com                 |
| Project Report | 2026-02-01 | user3@example.com, user4@example.com                 |

---

## 🔑 Environment Variables
Create a `.env` file in the project root:

OUTLOOK_EMAIL=your_email@gmail.com
OUTLOOK_PASSWORD=your_app_password_here
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

⚠️ Use an **App Password** (not your normal password).
For Outlook: smtp.office365.com
For Gmail: smtp.gmail.com

---

## ▶️ Usage
1. Install dependencies:
   pip install openpyxl python-dotenv
2. Run the bot:
   python main.py
3. Check your inbox for reminders!

---

## 🛡️ Security Notes
- Never commit your real `.env` file to GitHub.
- Add `.env` to `.gitignore`.
- Use `example.env` with placeholder values for sharing.
