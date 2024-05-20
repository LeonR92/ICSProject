from MailHandler import MailHandler
from datetime import datetime, timedelta
import logging


def main():
    try:
        mail_handler = MailHandler()

        # Connect to Outlook IMAP
        imap_conn = mail_handler.connect_to_outlook_imap()

        # Fetch and print the latest calendar responses
        calendar_responses = mail_handler.fetch_calendar_responses(imap_conn)
        print("Latest calendar responses:")
        for response in calendar_responses:
            print(response)

        # Logout from the IMAP server
        imap_conn.logout()

        # Connect to Outlook SMTP
        smtp_conn = mail_handler.connect_to_outlook_smtp()

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

        ics_filename = mail_handler.create_ics_file(
            event_name, start_time, end_time, description, location, to_email
        )

        mail_handler.send_email(smtp_conn, to_email, subject, body, ics_filename)

        # Quit the SMTP server
        smtp_conn.quit()

    except Exception as e:
        logging.error(f"Error: {e}")


if __name__ == "__main__":
    main()
