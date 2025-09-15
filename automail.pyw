

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
import glob
import logging
import time

# ---------------- Logging Configuration ----------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ---------------- Email Configuration ----------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "incraftiq@gmail.com"
SENDER_PASSWORD = "ehqd hjtm kyvx msyc"  # Gmail App Password

# ✅ Recipients list
RECIPIENT_EMAILS = [
    "Avantika.Trivedi@unilever.com",
    "Aneesh.Bhagwat@unilever.com",
    "Sumit.Kumar2@unilever.com",
    "Rahul.M@unilever.com",
    "Diksha.Mishra@unilever.com",
    "VIKAS.GHURIANI@unilever.com"
]

# ---------------- File Paths ----------------
REPORTS_DIR = r"D:\Summerpur_Reports\Summerpur_Reports\Reports"

# ---------------- Email Sending Function ----------------
def send_shift_reports():
    """Send email with today's shift reports as attachments"""
    try:
        today = datetime.now().strftime("%d-%m-%Y")  # e.g., 09-09-2025
        logging.info(f"Looking for shift reports for {today}")

        # Shift names as per report generation
        shift_files = []
        for shift in ["Shift-A", "Shift-B", "Shift-C", "Full-Day"]:
            pattern = os.path.join(REPORTS_DIR, f"*{today}*{shift}*.xlsx")
            found_files = glob.glob(pattern)
            if found_files:
                latest_file = max(found_files, key=os.path.getctime)
                shift_files.append(latest_file)
                logging.info(f"Found {shift} file: {latest_file}")
            else:
                logging.warning(f"No file found for {shift} on {today}")

        if not shift_files:
            logging.error("No shift reports found, email not sent.")
            return

        # Create email
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = ", ".join(RECIPIENT_EMAILS)
        msg["Subject"] = f"Daily Shift Reports - {today}"

        # Email body
        body = f"""
Dear Team,

Please find attached the daily shift reports for {today}.

Shifts Covered: {', '.join([os.path.basename(f).split('_')[-1].replace('.xlsx','') for f in shift_files])}

Regards,
InCraftIQ-Reports powered by Innovexer TechCraft Pvt Ltd
"""
        msg.attach(MIMEText(body, "plain"))

        # Attach files
        for file in shift_files:
            try:
                with open(file, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(file)}"
                )
                msg.attach(part)
                logging.info(f"Attached file: {file}")
            except Exception as e:
                logging.error(f"Error attaching {file}: {e}")

        # Send email
        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECIPIENT_EMAILS, msg.as_string())
            server.quit()
            logging.info("Shift reports email sent successfully!")
        except Exception as e:
            logging.error(f"Email sending failed: {e}")
    except Exception as e:
        logging.error(f"Error in send_shift_reports: {e}")

# ---------------- Scheduler ----------------
scheduler = BackgroundScheduler()
scheduler.add_job(send_shift_reports, "cron", hour=7, minute=0)  # ✅ Daily 7:00 AM
scheduler.start()

if __name__ == "__main__":
    logging.info("Email scheduler started. Waiting for 7:00 AM daily to send the shift reports...")
    try:
        while True:
            time.sleep(60)  # ✅ Keep script alive
    except (KeyboardInterrupt, SystemExit):
        scheduler.shutdown()
        logging.info("Scheduler stopped.")
