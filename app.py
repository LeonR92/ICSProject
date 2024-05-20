import imaplib
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List
import os
import dotenv
from dotenv import load_dotenv
from datetime import datetime, timedelta


load_dotenv()

logging.basicConfig(level=logging.INFO)
user_email = os.environ.get("EMAIL")
user_password = os.environ.get("PASSWORD")


def connect_to_outlook_imap(user_email: str, user_password: str) -> imaplib.IMAP4_SSL:
    """Connect to the Outlook IMAP server using provided email and password.
    
    Args:
        user_email (str): The user's email address.
        user_password (str): The user's password.

    Returns:
        imaplib.IMAP4_SSL: The IMAP connection object.
    
    Raises:
        imaplib.IMAP4.error: If the connection or login fails.
    """
    imap_server = 'outlook.office365.com'
    imap_port = 993
    
    imap_conn = imaplib.IMAP4_SSL(imap_server, imap_port)
    
    try:
        imap_conn.login(user_email, user_password)
        logging.info("IMAP connection established.")
    except imaplib.IMAP4.error as e:
        logging.error(f"IMAP connection failed: {e}")
        raise

    return imap_conn

def fetch_inbox_emails(imap_conn: imaplib.IMAP4_SSL, num_emails: int = 10) -> List[str]:
    """Fetch the latest emails from the inbox.
    
    Args:
        imap_conn (imaplib.IMAP4_SSL): The IMAP connection object.
        num_emails (int): Number of emails to fetch.
    
    Returns:
        List[str]: List of email subjects.
    """
    imap_conn.select('inbox')
    result, data = imap_conn.search(None, 'ALL')
    if result != 'OK':
        raise Exception("Failed to search inbox.")
    
    email_ids = data[0].split()
    latest_email_ids = email_ids[-num_emails:]
    
    emails = []
    for email_id in latest_email_ids:
        result, msg_data = imap_conn.fetch(email_id, '(RFC822)')
        if result != 'OK':
            raise Exception("Failed to fetch email.")
        emails.append(msg_data[0][1].decode('utf-8'))
    
    return emails

def connect_to_outlook_smtp(user_email: str, user_password: str) -> smtplib.SMTP:
    """Connect to the Outlook SMTP server using provided email and password.
    
    Args:
        user_email (str): The user's email address.
        user_password (str): The user's password.

    Returns:
        smtplib.SMTP: The SMTP connection object.
    
    Raises:
        smtplib.SMTPException: If the connection or login fails.
    """
    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    
    smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
    smtp_conn.starttls()
    
    try:
        smtp_conn.login(user_email, user_password)
        logging.info("SMTP connection established.")
    except smtplib.SMTPException as e:
        logging.error(f"SMTP connection failed: {e}")
        raise
    
    return smtp_conn

def create_ics_file(event_name: str, start_time: datetime, end_time: datetime, description: str, location: str) -> str:
    """Create an ICS file for a calendar event."""
    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Your Organization//Your Product//EN
BEGIN:VEVENT
UID:{datetime.now().strftime('%Y%m%dT%H%M%S')}@yourdomain.com
DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}
DTSTART:{start_time.strftime('%Y%m%dT%H%M%S')}
DTEND:{end_time.strftime('%Y%m%dT%H%M%S')}
SUMMARY:{event_name}
DESCRIPTION:{description}
LOCATION:{location}
END:VEVENT
END:VCALENDAR"""
    
    ics_filename = f"{event_name.replace(' ', '_')}.ics"
    with open(ics_filename, 'w') as ics_file:
        ics_file.write(ics_content)
    
    return ics_filename


def send_email(smtp_conn: smtplib.SMTP, user_email: str, to_email: str, subject: str, body: str, attachment_path: str = None):
    """Send an email using the provided SMTP connection.
    
    Args:
        smtp_conn (smtplib.SMTP): The SMTP connection object.
        user_email (str): The sender's email address.
        to_email (str): The recipient's email address.
        subject (str): The email subject.
        body (str): The email body.
    """
    msg = MIMEMultipart()
    msg['From'] = user_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    if attachment_path:
        attachment = MIMEBase('application', 'octet-stream')
        with open(attachment_path, 'rb') as attachment_file:
            attachment.set_payload(attachment_file.read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
        msg.attach(attachment)

    try:
        smtp_conn.sendmail(user_email, to_email, msg.as_string())
        logging.info("Email sent successfully.")
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        raise

def main():
    user_email = os.environ.get("EMAIL")
    user_password = os.environ.get("PASSWORD")

    if not user_email or not user_password:
        logging.error("Email or password environment variables not set.")
        return

    try:
        # Connect to Outlook IMAP
        imap_conn = connect_to_outlook_imap(user_email, user_password)
        
        # Fetch and print the latest 10 emails
        emails = fetch_inbox_emails(imap_conn)
        print("Latest 10 emails:")
        for email in emails:
            print(email)
        
        # Logout from the IMAP server
        imap_conn.logout()

        # Connect to Outlook SMTP
        smtp_conn = connect_to_outlook_smtp(user_email, user_password)
        
        # Send an email
        to_email = input("Enter recipient's email: ")
        subject = input("Enter subject: ")
        body = input("Enter email body: ")

        # Create an ICS file for an event
        event_name = "Meeting with Team"
        start_time = datetime.now() + timedelta(days=1)
        end_time = start_time + timedelta(hours=1)
        description = "Discussing project updates."
        location = "Office"

        ics_filename = create_ics_file(event_name, start_time, end_time, description, location)

        send_email(smtp_conn, user_email, to_email, subject, body, ics_filename)
        
        # Quit the SMTP server
        smtp_conn.quit()
        
    except Exception as e:
        logging.error(f"Error: {e}")

if __name__ == '__main__':
    main()
