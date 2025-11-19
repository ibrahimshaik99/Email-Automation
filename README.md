# HR Email Automation System

Complete Python automation system for sending greeting emails to employees for birthdays, work anniversaries, and marriage anniversaries using **India Standard Time (IST, GMT+5:30)**.

## 📋 Table of Contents
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Excel File Setup](#excel-file-setup)
- [Running the Script](#running-the-script)
- [Scheduling (IST Timezone)](#scheduling-ist-timezone)
- [Email Provider Setup](#email-provider-setup)
- [Logs](#logs)
- [Troubleshooting](#troubleshooting)

---

## ✨ Features

✅ **IST Timezone-Aware**: All date/time operations use India Standard Time (GMT+5:30)  
✅ **Automated Email Triggers**:
   - 🎂 Birthday greetings
   - 🎊 Work anniversary (years completed from joining date)
   - 💕 Marriage anniversary

✅ **Smart Features**:
   - Duplicate email prevention (per day)
   - Retry logic (3 attempts per email)
   - Comprehensive error handling
   - Detailed logging with IST timestamps
   - Support for multiple date formats (DD-MM-YYYY, DD/MM/YYYY, etc.)

✅ **Email Provider Support**: Gmail, Outlook, Custom SMTP

---

## 🔧 Requirements

### Python Version
- Python 3.8 or higher

### Dependencies
Install required packages:

```bash
pip install pandas openpyxl pytz
```

**Package Details:**
- `pandas` - Excel file reading
- `openpyxl` - Excel file handling (.xlsx format)
- `pytz` - Timezone handling (IST)
- `smtplib` - Built-in (email sending)

---

## 📦 Installation

### Step 1: Download/Clone the Project

```bash
# Create project directory
mkdir hr-email-automation
cd hr-email-automation
```

### Step 2: Install Dependencies

```bash
pip install pandas openpyxl pytz
```

### Step 3: Verify Folder Structure

```
hr-email-automation/
├── main.py
├── config.json
├── templates/
│   ├── birthday.html
│   ├── work_anniversary.html
│   └── marriage_anniversary.html
├── data/
│   └── employees.xlsx
├── logs/                    # Auto-created
└── README.md
```

---

## ⚙️ Configuration

### Edit `config.json`

```json
{
  "email": {
    "provider": "gmail",
    "sender_email": "your-email@gmail.com",
    "password": "your-app-password-here",
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587
  },
  "excel_file_path": "data/employees.xlsx",
  "template_directory": "templates",
  "log_directory": "logs"
}
```

### Configuration Fields

| Field | Description | Example |
|-------|-------------|---------|
| `provider` | Email service | `gmail`, `outlook`, or `custom` |
| `sender_email` | Your email address | `hr@company.com` |
| `password` | App password (NOT regular password) | See [Email Provider Setup](#email-provider-setup) |
| `smtp_host` | SMTP server | `smtp.gmail.com` |
| `smtp_port` | SMTP port | `587` (TLS) or `465` (SSL) |

---

## 📊 Excel File Setup

### Required Columns

Your Excel file (`data/employees.xlsx`) must have these columns:

| Column Name | Format | Required | Example |
|-------------|--------|----------|---------|
| Employee Name | Text | Yes | Rajesh Kumar |
| Email | Email address | Yes | rajesh@company.com |
| Date of Birth | DD-MM-YYYY | Yes | 15-03-1990 |
| Date of Joining | DD-MM-YYYY | Yes | 10-05-2020 |
| Marriage Anniversary | DD-MM-YYYY | No | 12-12-2015 |
| Department | Text | No | IT |

### Supported Date Formats
- `DD-MM-YYYY` (Preferred for India)
- `DD/MM/YYYY`
- `DD.MM.YYYY`
- `YYYY-MM-DD`

### Sample Excel Data

```
Employee Name       Email                      Date of Birth  Date of Joining  Marriage Anniversary  Department
Rajesh Kumar        rajesh@company.com         15-03-1990     10-05-2020       12-12-2015           IT
Priya Sharma        priya@company.com          22-07-1988     15-03-2018       25-11-2012           HR
Amit Patel          amit@company.com           08-11-1992     01-01-2019       14-02-2017           Finance
```

**Note**: Leave Marriage Anniversary blank if not applicable.

---

## 🚀 Running the Script

### Manual Run (One-Time)

```bash
python main.py
```

### Test Run (Check without sending)
You can modify the script to add a `--test` flag or manually comment out the `send_email()` calls for testing.

---

## ⏰ Scheduling (IST Timezone)

Run the script automatically every day at **8:00 AM IST**.

### Windows - Task Scheduler

1. **Open Task Scheduler**
   - Press `Win + R`, type `taskschd.msc`, press Enter

2. **Create Basic Task**
   - Click "Create Basic Task"
   - Name: `HR Email Automation`
   - Description: `Send birthday/anniversary emails (IST)`

3. **Trigger**
   - Select "Daily"
   - Start time: `08:00:00 AM`
   - **Important**: Set your system timezone to IST (Asia/Kolkata)

4. **Action**
   - Action: "Start a program"
   - Program/script: `C:\Python39\python.exe` (your Python path)
   - Arguments: `main.py`
   - Start in: `C:\hr-email-automation` (your project path)

5. **Finish & Test**
   - Right-click task → "Run" to test

### Linux/Ubuntu - Cron Job (IST)

1. **Edit Crontab**
```bash
crontab -e
```

2. **Add Cron Job with IST Timezone**
```bash
# Run daily at 8:00 AM IST
0 8 * * * TZ='Asia/Kolkata' /usr/bin/python3 /home/user/hr-email-automation/main.py >> /home/user/hr-email-automation/logs/cron.log 2>&1
```

**Breakdown:**
- `0 8 * * *` - Every day at 8:00 AM
- `TZ='Asia/Kolkata'` - Force IST timezone
- `/usr/bin/python3` - Python path (check with `which python3`)
- `/home/user/hr-email-automation/main.py` - Full script path
- `>> logs/cron.log 2>&1` - Log output

3. **Verify Cron Job**
```bash
crontab -l
```

4. **Check Cron Logs**
```bash
cat /home/user/hr-email-automation/logs/cron.log
```

### macOS - Launchd (IST)

1. **Create plist file**: `~/Library/LaunchAgents/com.hrmail.automation.plist`

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.hrmail.automation</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>TZ</key>
        <string>Asia/Kolkata</string>
    </dict>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/local/bin/python3</string>
        <string>/Users/username/hr-email-automation/main.py</string>
    </array>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>8</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>/Users/username/hr-email-automation/logs/launchd.log</string>
    <key>StandardErrorPath</key>
    <string>/Users/username/hr-email-automation/logs/launchd-error.log</string>
</dict>
</plist>
```

2. **Load the job**
```bash
launchctl load ~/Library/LaunchAgents/com.hrmail.automation.plist
```

---

## 📧 Email Provider Setup

### Gmail Setup

1. **Enable 2-Step Verification**
   - Go to: https://myaccount.google.com/security
   - Enable "2-Step Verification"

2. **Generate App Password**
   - Go to: https://myaccount.google.com/apppasswords
   - Select "Mail" and "Other (Custom name)"
   - Name it: `HR Email Automation`
   - Copy the 16-character password

3. **Update config.json**
```json
{
  "email": {
    "provider": "gmail",
    "sender_email": "your-email@gmail.com",
    "password": "abcd efgh ijkl mnop",
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587
  }
}
```

### Outlook/Office 365 Setup

1. **Update config.json**
```json
{
  "email": {
    "provider": "outlook",
    "sender_email": "your-email@outlook.com",
    "password": "your-outlook-password",
    "smtp_host": "smtp.office365.com",
    "smtp_port": 587
  }
}
```

### Custom SMTP Setup

```json
{
  "email": {
    "provider": "custom",
    "sender_email": "hr@company.com",
    "password": "your-password",
    "smtp_host": "mail.company.com",
    "smtp_port": 587
  }
}
```

---

## 📝 Logs

Logs are saved in the `logs/` directory with IST timestamps.

### Log File Format
```
2025-01-18_IST.log
```

### Sample Log Output
```
2025-01-18 08:00:05 [IST] - INFO - === HR Email Automation Started (IST: 2025-01-18 08:00:05) ===
2025-01-18 08:00:06 [IST] - INFO - Loaded 50 employees from data/employees.xlsx
2025-01-18 08:00:07 [IST] - INFO - 🎂 Birthday detected: Rajesh Kumar (Age: 35)
2025-01-18 08:00:09 [IST] - INFO - ✓ Email sent successfully to rajesh@company.com - 🎉 Happy Birthday, Rajesh Kumar!
2025-01-18 08:00:10 [IST] - INFO - 🎊 Work Anniversary detected: Priya Sharma (7 years)
2025-01-18 08:00:12 [IST] - INFO - ✓ Email sent successfully to priya@company.com - 🌟 Happy 7 Year Work Anniversary, Priya Sharma!
2025-01-18 08:00:13 [IST] - INFO - === Summary ===
2025-01-18 08:00:13 [IST] - INFO - Birthday emails sent: 1
2025-01-18 08:00:13 [IST] - INFO - Work anniversary emails sent: 1
2025-01-18 08:00:13 [IST] - INFO - Marriage anniversary emails sent: 0
2025-01-18 08:00:13 [IST] - INFO - Total emails sent: 2
```

---

## 🔧 Troubleshooting

### Common Issues

#### 1. **ModuleNotFoundError: No module named 'pandas'**
**Solution:**
```bash
pip install pandas openpyxl pytz
```

#### 2. **SMTPAuthenticationError: Username and Password not accepted**
**Solution:**
- Gmail: Use App Password (not regular password)
- Outlook: Enable "Allow less secure apps" or use App Password
- Verify credentials in `config.json`

#### 3. **FileNotFoundError: data/employees.xlsx**
**Solution:**
- Ensure Excel file exists at the specified path
- Check `excel_file_path` in `config.json`

#### 4. **Wrong timezone (emails sent at wrong time)**
**Solution:**
- Verify system timezone: `date` (Linux/Mac) or `tzutil /g` (Windows)
- For cron jobs, use `TZ='Asia/Kolkata'` in the command
- The script uses `pytz` for IST, so internal logic should be correct

#### 5. **No emails sent (but no errors)**
**Solution:**
- Check if today's date matches any employee dates
- Verify date formats in Excel (DD-MM-YYYY)
- Check logs for "detected" messages

#### 6. **Duplicate emails sent**
**Solution:**
- The script prevents duplicates per run
- If running multiple times per day, duplicates may occur
- Check logs for "Duplicate email prevented" messages

### Debug Mode

Add print statements to `main.py` for debugging:

```python
# After loading data
print(f"Today (IST): {self.today}")
print(f"Total employees: {len(df)}")
```

---

## 📚 Additional Resources

- **Python pytz documentation**: https://pypi.org/project/pytz/
- **Pandas documentation**: https://pandas.pydata.org/docs/
- **Gmail App Passwords**: https://support.google.com/accounts/answer/185833
- **Cron Job Guide**: https://crontab.guru/

---

## 🆘 Support

For issues or questions:
1. Check the `logs/` directory for error messages
2. Verify all configuration settings
3. Test email credentials manually
4. Ensure Excel file format is correct

---

## 📜 License

This project is provided as-is for internal HR automation purposes.

---

**Last Updated**: 2025-01-18  
**IST Timezone**: Asia/Kolkata (GMT+5:30)
