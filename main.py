"""
HR Email Automation System
Sends automated greeting emails for birthdays, work anniversaries, and marriage anniversaries
Uses India Standard Time (IST, GMT+5:30)
"""

import pandas as pd
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import pytz
import os
import logging
from pathlib import Path
import time

# Set up IST timezone
IST = pytz.timezone("Asia/Kolkata")

class HREmailAutomation:
    def __init__(self, config_path="config.json"):
        """Initialize the HR Email Automation system"""       
        self.config = self.load_config(config_path)
        self.now = datetime.now(IST)
        self.today = self.now.date()
        self.setup_logging()
        self.email_sent_today = set()  # Track (email, event_type) to allow multiple events per day
        self.target_hour = 16
        self.target_minute = 49
        
    def load_config(self, config_path):
        """Load configuration from JSON file"""
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Error: {config_path} not found. Please create the configuration file.")
            raise
        except json.JSONDecodeError:
            print(f"Error: Invalid JSON in {config_path}")
            raise
    
    def setup_logging(self):
        """Set up logging with IST timezone"""
        log_dir = Path(self.config.get('log_directory', 'logs'))
        log_dir.mkdir(exist_ok=True)
        
        # Format: YYYY-MM-DD_IST.log now change to DD-MM-YYYY
        log_filename = f"{self.today.strftime('%Y-%m-%d')}_IST.log"
        log_path = log_dir / log_filename
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [IST] - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        # Set IST for logging timestamps
        logging.Formatter.converter = lambda *args: datetime.now(IST).timetuple()
        
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"=== HR Email Automation Started (IST: {datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S')}) ===")
    
    def load_employee_data(self):
        """Load employee data from Excel file"""
        excel_path = self.config.get('excel_file_path', 'data/employees.xlsx')
        
        try:
            df = pd.read_excel(excel_path)
            self.logger.info(f"Loaded {len(df)} employees from {excel_path}")
            return df
        except FileNotFoundError:
            self.logger.error(f"Excel file not found: {excel_path}")
            raise
        except Exception as e:
            self.logger.error(f"Error reading Excel file: {str(e)}")
            raise
    
    def parse_date(self, date_value):
        """Parse date from Excel (handles various formats)"""
        if pd.isna(date_value):
            return None
        try:
            # If already a datetime object
            if isinstance(date_value, datetime):
                return date_value.date()
            # Try parsing string dates (DD-MM-YYYY, DD/MM/YYYY, etc.)
            date_formats = ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%d.%m.%Y']
            for fmt in date_formats:
                try:
                    return datetime.strptime(str(date_value), fmt).date()
                except ValueError:
                    continue
            
            # If all format fail then 
            return None
        except Exception as e:
            self.logger.warning(f"Could not parse date: {date_value} - {str(e)}")
            return None
    
    def load_email_template(self, template_name):
        """Load HTML email template"""
        template_path = Path(self.config.get('template_directory', 'templates')) / template_name
        
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                return f.read()
        except FileNotFoundError:
            self.logger.error(f"Template not found: {template_path}")
            raise
    
    def send_email(self, recipient_email, subject, html_content, event_type, max_retries=3):
        """Send email with retry logic"""
        email_event_key = (recipient_email, event_type)
        if email_event_key in self.email_sent_today:
            self.logger.warning(f"Duplicate email prevented for {recipient_email} - {event_type}")
            return False
        
        for attempt in range(1, max_retries + 1):
            try:
                msg = MIMEMultipart('alternative')
                msg['From'] = self.config['email']['sender_email']
                msg['To'] = recipient_email
                msg['Subject'] = subject
                
                html_part = MIMEText(html_content, 'html')
                msg.attach(html_part)
                
                # this feature use to connect smtp google gmail server and send email 
                if self.config['email']['provider'] == 'gmail':
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                elif self.config['email']['provider'] == 'outlook':
                    server = smtplib.SMTP('smtp.office365.com', 587)
                else:
                    # Custom SMTP
                    server = smtplib.SMTP(
                        self.config['email']['smtp_host'],
                        self.config['email']['smtp_port']
                    )
                
                server.starttls()
                server.login(
                    self.config['email']['sender_email'],
                    self.config['email']['password']
                 
                
                server.send_message(msg)
                server.quit()
                
                self.email_sent_today.add(email_event_key)
                self.logger.info(f"✓ Email sent successfully to {recipient_email} - {subject}")
                return True
                
            except Exception as e:
                self.logger.warning(f"Attempt {attempt}/{max_retries} failed for {recipient_email}: {str(e)}")
                if attempt < max_retries:
                    time.sleep(2)  # Wait before retry
                else:
                    self.logger.error(f"✗ Failed to send email to {recipient_email} after {max_retries} attempts")
                    return False
    
    def check_birthday(self, row):
        """Check if today is employee's birthday"""
        dob = self.parse_date(row.get('Date of Birth'))
        
        if not dob:
            return False
        
        if self.today.day == dob.day and self.today.month == dob.month:
            name = row.get('Employee Name', 'Employee')
            email = row.get('Email')
            
            if not email or pd.isna(email):
                self.logger.warning(f"Skipped birthday for {name}: No email address")
                return False
            
            age = self.today.year - dob.year
            self.logger.info(f"🎂 Birthday detected: {name} (Age: {age})")
            
            # Load and personalize template 
            template = self.load_email_template('birthday.html')
            html_content = template.replace('{{name}}', name).replace('{{age}}', str(age))
            
            subject = f"🎉 Happy Birthday, {name}!"
            return self.send_email(email, subject, html_content, 'birthday')
        
        return False
    
    def check_work_anniversary(self, row):
        """Check if today is employee's work anniversary"""
        doj = self.parse_date(row.get('Date of Joining'))

        if not doj:
            return False

        if self.today.month == doj.month and self.today.day == doj.day:
            years_completed = self.today.year - doj.year

            name = row.get('Employee Name', 'Employee')
            email = row.get('Email')

            if not email or pd.isna(email):
                self.logger.warning(f"Skipped work anniversary for {name}: No email address")
                return False

            self.logger.info(f"🎊 Work Anniversary detected: {name} ({years_completed} years)")

            # Load and personalize template
            template = self.load_email_template('work_anniversary.html')
            html_content = (template
                          .replace('{{name}}', name)
                          .replace('{{years}}', str(years_completed)))

            subject = f"🌟 Happy {years_completed} Year Work Anniversary, {name}!"
            return self.send_email(email, subject, html_content, 'work_anniversary')

        return False
    
    def check_marriage_anniversary(self, row):
        """Check if today is employee's marriage anniversary"""
        marriage_date = self.parse_date(row.get('Marriage Anniversary'))

        if not marriage_date:
            return False

        if self.today.day == marriage_date.day and self.today.month == marriage_date.month:
            years_completed = self.today.year - marriage_date.year

            name = row.get('Employee Name', 'Employee')
            email = row.get('Email')

            if not email or pd.isna(email):
                self.logger.warning(f"Skipped marriage anniversary for {name}: No email address")
                return False

            self.logger.info(f"💑 Marriage Anniversary detected: {name} ({years_completed} years)")

            # Load and personalize template
            template = self.load_email_template('marriage_anniversary.html')
            html_content = (template
                          .replace('{{name}}', name)
                          .replace('{{years}}', str(years_completed)))

            subject = f"💕 Happy {years_completed} Year Marriage Anniversary, {name}!"
            return self.send_email(email, subject, html_content, 'marriage_anniversary')

        return False
    
    def run(self):
        """Main execution method"""
        try:
            # Check if current time is 16:48AM IST (with 1-minute tolerance)s
            current_time = datetime.now(IST)
            if current_time.hour != self.target_hour or current_time.minute != self.target_minute:
                self.logger.info(f"Not scheduled time. Current time: {current_time.strftime('%H:%M')} IST. Target time: {self.target_hour:02d}:{self.target_minute:02d} IST")
                self.logger.info("Emails will only be sent at 16:49PM IST")
                return
            
            self.logger.info(f"Scheduled time reached: {current_time.strftime('%H:%M')} IST - Processing emails...")
            
            df = self.load_employee_data()
            birthday_count = 0
            work_anniversary_count = 0
            marriage_anniversary_count = 0           
            for index, row in df.iterrows():
                try:
                    # Check each type of greeting
                    if self.check_birthday(row):
                        birthday_count += 1
                    
                    if self.check_work_anniversary(row):
                        work_anniversary_count += 1
                    
                    if self.check_marriage_anniversary(row):
                        marriage_anniversary_count += 1
                        
                except Exception as e:
                    name = row.get('Employee Name', f'Row {index}')
                    self.logger.error(f"Error processing employee {name}: {str(e)}")
            
            # Summary
            self.logger.info("=== Summary ===")
            self.logger.info(f"Birthday emails sent: {birthday_count}")
            self.logger.info(f"Work anniversary emails sent: {work_anniversary_count}")
            self.logger.info(f"Marriage anniversary emails sent: {marriage_anniversary_count}")
            self.logger.info(f"Total emails sent: {birthday_count + work_anniversary_count + marriage_anniversary_count}")
            self.logger.info(f"=== HR Email Automation Completed (IST: {datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S')}) ===")
            
        except Exception as e:
            self.logger.error(f"Critical error in main execution: {str(e)}")
            raise


if __name__ == "__main__":
    try:
        automation = HREmailAutomation()
        automation.run()
    except Exception as e:
        print(f"Failed to run HR Email Automation: {str(e)}")
        exit(1)
