import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import config     # stores the email
from datetime import datetime, timedelta
import re

import os.path
# libraries for google API
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.metadata.readonly']

SPREADSHEET_ID = '1zM-9tdsbCMwqEdGILtiHE6WPUMCkpEk5kdKcYBAICA4'
RANGE_NAME = 'N:N'


def get_values(spreadsheet_id: 'str', range_name: 'str', creds: 'Credentials') -> 'list[list[str]]':
    """
    Get the value from a google sheets by specific sheets id, range and credentials

    Parameters
    ----------
    spreadsheet_id : str
        The id of the spreadsheet.

    range_name : str
        Sheets name, colume and row name of the cells you want to retrieve. example: A1:C2

    creds : Credentials
        Credentials of google api
    
    Returns
    -------
    list[list[str]]
        The value of the cells you retrived
    """
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                    range=range_name).execute()
        values = result.get('values', [])
        
        return values
    except HttpError as err:
        print(err)

def extractEmail_form(content):
    email = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', content)

    if email:
        return email.group(0)
    
def get_name_check_by_email(email):
    #get credentials
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())


    emails = get_values(SPREADSHEET_ID, RANGE_NAME, creds)
    row = 0
    for i in range(len(emails) - 1, -1, -1):
        if emails[i][0] == email:
            row = i + 1
            break
    # Throw err if didn't find mail
    if row == 0:
        raise Exception("Didn't find profile")
    
    name = get_values(SPREADSHEET_ID, 'B' + str(row), creds)[0][0]
    check = 'OPT (Optional Practical Training) Maintenance' in get_values(SPREADSHEET_ID, 'L' + str(row), creds)[0][0].split(',')
    return name, check

def getCandidateEmailnNameCheck(msg):
    if msg.is_multipart():
        content = ""
        for part in msg.get_payload():
            try:
                nn = part.get_payload(decode=True).decode()
                content += nn
            except:
                pass
    else:
        content = msg.get_payload(decode=True).decode()
    for i in range(len(content)-10):
        if(content[i:i+15] == "View Response: "):
            candidate_email = extractEmail_form(content[i+6:i+100])
            name, check = get_name_check_by_email(candidate_email)
            return (candidate_email, name, check)
        
    raise Exception("Didn't find email after View Response:")


    

def reply():
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = config.USER
    smtp_password = config.PASSWORD

    # IMAP settings
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(smtp_username, smtp_password)
    mail.select('inbox')

    # Search for emails with the desired subject line
    since_date = (datetime.now() - timedelta(days=2)).strftime('%d-%b-%Y')
    subject1 = 'Your form, Zenativity Volunteer Application Form, has new responses.'
    typ, data = mail.search(None, '(UNSEEN SUBJECT "{0}" SINCE {1})'.format(subject1, since_date))
    print(data)

    sending_list = []
    for num in data[0].split():
        typ, msg_data = mail.fetch(num, '(RFC822)')

        msg_str = msg_data[0][1].decode('utf-8')
        msg = email.message_from_string(msg_str)

        sending_list.append(getCandidateEmailnNameCheck(msg))
    print(sending_list)
       
    # check sent box
    mail.select('"[Gmail]/Sent Mail"')
    for to_email, to_name, to_check in sending_list:
        if not to_check:
            continue

        typ, data = mail.search(None, '(TO "{0}")'.format(to_email))
        if len(data[0].split()) > 0:
            print(data, len(data))
            print("have sent to " + to_email)
            continue
        
        sender_email = to_email


        # Compose the reply message
        reply_subject = 'Zenativity Volunteer Opportunity'


        # parallel reply email
        reply_name = 'Hi {0}, it was nice chatting with you today!'.format(to_name)

        # create http email content
        message = MIMEMultipart()
        message["From"] = smtp_username
        message["To"] = sender_email
        message["Subject"] = reply_subject

        # parallel reply email
        reply_body = 'Hi {0}, Thank you for your interest in this volunteer opportunity! We are glad to connect with you and get to know you better! Please make an appointment with us using the following link: {1}'.format(to_name, config.CALENDLY_LINK)
        html = """
        <html>
            <body>
                <p>{0}</p>
            </body>
        </html>
        """.format(reply_body)

        # Add HTML content to message body
        body = MIMEText(html, "html")
        message.attach(body)

        # Send the reply message
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(smtp_username, sender_email, message.as_string())
    
    mail.close()
    mail.logout()
