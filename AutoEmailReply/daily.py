import imaplib
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import config     # stores the email
from datetime import datetime, timedelta
import re

import volunteer_match_reply
import google_form_reply

if __name__ == "__main__":
    volunteer_match_reply.reply()
    google_form_reply.reply()
