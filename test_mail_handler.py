import pytest
import imaplib
import smtplib
from datetime import datetime, timedelta
from MailHandler import MailHandler  # Import the class from the correct module

@pytest.fixture
def mail_handler():
    return MailHandler()

def test_load_env_vars(mocker):
    mocker.patch.dict('os.environ', {'EMAIL': 'test@example.com', 'PASSWORD': 'testpassword'})
    handler = MailHandler()
    assert handler.user_email == 'test@example.com'
    assert handler.user_password == 'testpassword'

def test_connect_to_outlook_imap(mocker, mail_handler):
    mock_imap = mocker.patch('imaplib.IMAP4_SSL')
    instance = mock_imap.return_value
    instance.login.return_value = 'OK'
    
    imap_conn = mail_handler.connect_to_outlook_imap()
    
    instance.login.assert_called_with(mail_handler.user_email, mail_handler.user_password)
    assert imap_conn == instance

def test_fetch_calendar_responses(mocker, mail_handler):
    mock_imap = mocker.patch('imaplib.IMAP4_SSL')
    instance = mock_imap.return_value
    instance.search.return_value = ('OK', [b'1 2'])
    instance.fetch.return_value = ('OK', [(b'1 (RFC822 {email data})', b'raw email data')])

    # Mocking email.message_from_bytes
    mock_message = mocker.patch('email.message_from_bytes')
    email_message = mock_message.return_value
    email_message.get_content_type.return_value = 'text/calendar'
    email_message['Subject'] = 'Test Subject'

    # Mock the walk method to return parts
    part = mocker.Mock()
    part.get_content_type.return_value = 'text/calendar'
    email_message.walk.return_value = [part]

    responses = mail_handler.fetch_calendar_responses(instance)

    # TODO !!!!!!!!!! Testfile
    assert len(responses) == 0
    assert responses[0] == 'Test Subject'

def test_connect_to_outlook_smtp(mocker, mail_handler):
    mock_smtp = mocker.patch('smtplib.SMTP')
    instance = mock_smtp.return_value
    instance.login.return_value = 'OK'
    
    smtp_conn = mail_handler.connect_to_outlook_smtp()
    
    instance.login.assert_called_with(mail_handler.user_email, mail_handler.user_password)
    assert smtp_conn == instance

def test_create_ics_file(mocker, mail_handler):
    mocker.patch('builtins.open', mocker.mock_open())
    mock_datetime = mocker.patch('MailHandler.datetime')
    mock_datetime.now.return_value = datetime(2021, 1, 1, 12, 0, 0)
    mock_datetime.strftime.return_value = '20210101T120000'
    
    start_time = datetime(2021, 1, 2, 12, 0, 0)
    end_time = start_time + timedelta(hours=1)
    ics_filename = mail_handler.create_ics_file('Test Event', start_time, end_time, 'Description', 'Location', 'attendee@example.com')
    
    assert ics_filename == 'Test_Event.ics'

def test_send_email(mocker, mail_handler):
    mock_smtp = mocker.patch('smtplib.SMTP')
    instance = mock_smtp.return_value
    mock_open = mocker.patch('builtins.open', mocker.mock_open(read_data='ICS DATA'))
    
    smtp_conn = instance
    to_email = 'recipient@example.com'
    subject = 'Test Subject'
    body = 'Test Body'
    attachment_path = 'test.ics'
    
    mail_handler.send_email(smtp_conn, to_email, subject, body, attachment_path)
    
    instance.sendmail.assert_called_once()
