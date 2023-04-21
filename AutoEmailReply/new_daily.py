import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import config     # stores the email
from datetime import datetime, timedelta
import re

def extractEmail(content):
    email = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', content)

    if email:
        return email.group(0)

def extractName(content):
    return 'pass'

def getCandidateEmailnName(msg):
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
            candidate_email = extractEmail(content[i+6:i+100])
            break

    return (candidate_email)

def main():
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = config.USER
    smtp_password = config.PASSWORD

    # IMAP settings
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(smtp_username, smtp_password)
    mail.select('inbox')

    # Search for emails with the desired subject line
    since_date = (datetime.now() - timedelta(days=50)).strftime('%d-%b-%Y')
    subject1 = 'Your form, Zenativity Volunteer Application Form, has new responses.'
    typ, data = mail.search(None, '(SUBJECT "{0}" SINCE {1})'.format(subject1, since_date))
    print(data)

    for num in data[0].split():
        typ, msg_data = mail.fetch(num, '(RFC822)')

        msg_str = msg_data[0][1].decode('utf-8')
        msg = email.message_from_string(msg_str)

        toEmail = getCandidateEmailnName(msg)
        print(toEmail)

if __name__ == "__main__":
    main()