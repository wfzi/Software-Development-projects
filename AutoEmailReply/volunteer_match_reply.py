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
    name = re.search("(\w+)\s+(\w+)", content).group()
    return name

def getCandidateEmailnName(mgs):
    if mgs.is_multipart():
        content = ""
        for part in mgs.get_payload():
            try:
                nn = part.get_payload(decode=True).decode()
                content += nn
            except:
                pass
    else:
        content = mgs.get_payload(decode=True).decode()

    for i in range(len(content)-10):
        if(content[i:i+4] == "Name"):
            candidate_name = extractName(content[i+6: i+30])

        if(content[i:i+5] == "Email"):
            print(1)
            candidate_email = extractEmail(content[i+6:i+106])
            break
    return (candidate_name, candidate_email)

def obtain_header(msg):
    # decode the email subject
    subject, encoding = email.header.decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding)

    return subject

# SMTP settings
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
    subject1 = 'Someone wants to help: Volunteer with us and maintain your OPT status!'
    subject2 = 'Someone wants to help: Laid off? Losing OPT status? Volunteer with us and maintain your OPT status!'
    typ, data = mail.search(None, '(UNSEEN OR SUBJECT "{0}" SUBJECT "{1}" SINCE {2})'.format(subject1, subject2, since_date))
    print(data)

    # Loop through each email and send a reply
    for num in data[0].split():
        typ, msg_data = mail.fetch(num, '(RFC822)')

        msg_str = msg_data[0][1].decode('utf-8')
        msg = email.message_from_string(msg_str)

        msg_subject = obtain_header(msg)
        candidate_name, toEmail = getCandidateEmailnName(msg)
        print(toEmail)
        # Extract the sender's email address
        sender_email = toEmail


        # Compose the reply message
        reply_subject = 'RE: Zenativity Volunteer Opportunity'

        # parallel reply email
        reply_body = 'Hi {0}, Thank you for your interest in this volunteer opportunity! We are glad to connect with you and get to know you better! Please make an appointment with us using the following link: {1}'.format(candidate_name, config.CALENDLY_LINK)
        reply_info = "OPPORTUNITY INFORMATION: Title: {0} Organization: Zenativity, Inc.".format(msg_subject)

        # create http email content
        message = MIMEMultipart()
        message["From"] = smtp_username
        message["To"] = sender_email
        message["Subject"] = reply_subject

        html = """
        <html>
            <body>
                <p>{0}</p><br>
                <p>{1}</p>
            </body>
        </html>
        """.format(reply_body, reply_info)

        # Add HTML content to message body
        body = MIMEText(html, "html")
        message.attach(body)

        # Send the reply message
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            print(smtp_username, sender_email)
            server.sendmail(smtp_username, sender_email, message.as_string())

    # Close the connection to the mail server
    mail.close()
    mail.logout()
