# Automated Excel Email Reporter

This Python application reads `report.xlsx`, builds a plain-text summary, sends an email with the Excel file attached, and records sent emails in `log.txt`.

## Setup

Requires Python 3.10 or newer.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Place `report.xlsx` in this folder, or set `REPORT_FILE` to another path.

## Configuration

Set these environment variables before running the app:

```powershell
$env:SMTP_HOST = "smtp.example.com"
$env:SMTP_PORT = "587"
$env:SMTP_SECURITY = "starttls"  # starttls, ssl, or none
$env:SMTP_USERNAME = "your_username"
$env:SMTP_PASSWORD = "your_password"
$env:EMAIL_FROM = "sender@example.com"
$env:EMAIL_TO = "recipient@example.com, manager@example.com"
$env:EMAIL_SUBJECT = "Automated Excel Report"
$env:EMAIL_BODY = "Hello,\n\nPlease find attached the latest Excel report."
$env:REPORT_FILE = "report.xlsx"
$env:LOG_FILE = "log.txt"
```

`SMTP_USERNAME` and `SMTP_PASSWORD` may both be left empty only if your SMTP server does not require authentication.

## Run

```powershell
python app.py
```

Successful sends are appended to `log.txt` with the date, time, recipients, subject, and attachment path. Expected errors are also logged so failed runs are easier to diagnose.
