import imaplib
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email import message_from_bytes
from typing import List
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta


class MailHandler:
    def __init__(self):
        """Initialize the MailHandler with email and password from environment variables."""
        self.load_env_vars()
        self.setup_logging()

        if not self.user_email or not self.user_password:
            raise ValueError("Email or password environment variables not set.")

    def load_env_vars(self):
        """Load environment variables."""
        load_dotenv()
        self.user_email = os.environ.get("EMAIL")
        self.user_password = os.environ.get("PASSWORD")

    def setup_logging(self):
        """Setup logging configuration."""
        logging.basicConfig(level=logging.INFO)

    def connect_to_outlook_imap(self) -> imaplib.IMAP4_SSL:
        """Connect to the Outlook IMAP server using provided email and password.

        Returns:
            imaplib.IMAP4_SSL: The IMAP connection object.

        Raises:
            imaplib.IMAP4.error: If the connection or login fails.
        """
        imap_server = "outlook.office365.com"
        imap_port = 993
        imap_conn = imaplib.IMAP4_SSL(imap_server, imap_port)

        try:
            imap_conn.login(self.user_email, self.user_password)
            logging.info("IMAP connection established.")
        except imaplib.IMAP4.error as e:
            logging.error(f"IMAP connection failed: {e}")
            raise

        return imap_conn

    def fetch_calendar_responses(self, imap_conn: imaplib.IMAP4_SSL) -> List[str]:
        """Fetch the latest calendar responses from the inbox.

        Args:
            imap_conn (imaplib.IMAP4_SSL): The IMAP connection object.
            num_emails (int): Number of emails to fetch.

        Returns:
            List[str]: List of calendar response email subjects.

        Raises:
            Exception: If fetching or searching emails fails.
        """
        imap_conn.select("inbox")
        result, data = imap_conn.search(None, "ALL")
        if result != "OK":
            raise Exception("Failed to search inbox.")

        email_ids = data[0].split()

        calendar_responses = []
        for email_id in email_ids:
            result, msg_data = imap_conn.fetch(email_id, "(RFC822)")
            if result != "OK":
                raise Exception("Failed to fetch email.")

            email_msg = message_from_bytes(msg_data[0][1])

            if email_msg.get_content_type() == "text/calendar":
                calendar_responses.append(email_msg["Subject"])
            else:
                for part in email_msg.walk():
                    if part.get_content_type() == "text/calendar":
                        calendar_responses.append(email_msg["Subject"])
                        break

        return calendar_responses

    def connect_to_outlook_smtp(self) -> smtplib.SMTP:
        """Connect to the Outlook SMTP server using provided email and password.

        Returns:
            smtplib.SMTP: The SMTP connection object.

        Raises:
            smtplib.SMTPException: If the connection or login fails.
        """
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
        smtp_conn.starttls()

        try:
            smtp_conn.login(self.user_email, self.user_password)
            logging.info("SMTP connection established.")
        except smtplib.SMTPException as e:
            logging.error(f"SMTP connection failed: {e}")
            raise

        return smtp_conn

    def create_ics_file(
        self,
        event_name: str,
        start_time: datetime,
        end_time: datetime,
        description: str,
        location: str,
        attendee_email: str,
    ) -> str:
        """Create an ICS file for a calendar event.

        Args:
            event_name (str): The name of the event.
            start_time (datetime): The start time of the event.
            end_time (datetime): The end time of the event.
            description (str): The description of the event.
            location (str): The location of the event.
            attendee_email (str): The email of the attendee.

        Returns:
            str: The filename of the created ICS file.
        """
        ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Your Organization//Your Product//EN
METHOD:REQUEST
BEGIN:VEVENT
UID:{datetime.now().strftime('%Y%m%dT%H%M%S')}@yourdomain.com
DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}
DTSTART:{start_time.strftime('%Y%m%dT%H%M%S')}
DTEND:{end_time.strftime('%Y%m%dT%H%M%S')}
SUMMARY:{event_name}
DESCRIPTION:{description}
LOCATION:{location}
ORGANIZER;CN=Organizer:MAILTO:{self.user_email}
ATTENDEE;RSVP=TRUE;CN=Attendee;PARTSTAT=NEEDS-ACTION:MAILTO:{attendee_email}
END:VEVENT
END:VCALENDAR"""

        ics_filename = f"{event_name.replace(' ', '_')}.ics"
        with open(ics_filename, "w") as ics_file:
            ics_file.write(ics_content)

        return ics_filename

    def send_email(
        self,
        smtp_conn: smtplib.SMTP,
        to_email: str,
        subject: str,
        body: str,
        attachment_path: str = None,
    ):
        """Send an email using the provided SMTP connection.

        Args:
            smtp_conn (smtplib.SMTP): The SMTP connection object.
            to_email (str): The recipient's email address.
            subject (str): The email subject.
            body (str): The email body.
            attachment_path (str, optional): The path to the attachment file. Defaults to None.
        """
        msg = MIMEMultipart("mixed")
        msg["From"] = self.user_email
        msg["To"] = to_email
        msg["Subject"] = subject

        msg_alt = MIMEMultipart("alternative")
        msg.attach(msg_alt)

        msg_text = MIMEText(body, "plain")
        msg_alt.attach(msg_text)

        if attachment_path:
            attachment = MIMEBase(
                "text",
                "calendar",
                method="REQUEST",
                name=os.path.basename(attachment_path),
            )
            with open(attachment_path, "rb") as attachment_file:
                attachment.set_payload(attachment_file.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(attachment_path)}",
            )
            attachment.add_header(
                "Content-class", "urn:content-classes:calendarmessage"
            )
            msg.attach(attachment)

        try:
            smtp_conn.sendmail(self.user_email, to_email, msg.as_string())
            logging.info("Email sent successfully.")
        except Exception as e:
            logging.error(f"Failed to send email: {e}")
            raise
