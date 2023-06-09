"""
Documentation conventions: numpy docstring conventions
https://numpydoc.readthedocs.io/en/latest/format.html#docstring-standard
"""

from __future__ import print_function

import os.path
# libraries for google API
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import datetime
import dateutil.relativedelta as dateDelta

#libraries for sending email
import imaplib
import smtplib
from email.mime.text import MIMEText

import email_data


weekDays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
dateFormat = '%m-%d-%y'


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


def get_sheet_id(name: 'str', creds: 'Credentials') -> 'str':
    """
    Get the weekly report sheets id by name and credentials

    Parameters
    ----------
    name : str
        Name of the people whose weekly report sheets id we want to get.

    creds : Credentials
        Credentials of google api
    
    Returns
    -------
    str
        The weekly report sheets id
    """
    try:
        service = build('drive', 'v3', credentials=creds)

        # Call the Drive v3 API
        results = service.files().list(
            q="'1w-41nhWBFJFjWGXTT4WOTfSbbN-GVbyx' in parents", fields="nextPageToken, files(id, name)").execute()
        items = results.get('files', [])
        
        for item in items:
            cur_name = item['name'].split('-')[0]
            if cur_name == name:
                return item['id']
            
        return 'No files found.'
    except HttpError as err:
        print(err)


def update_values(spreadsheet_id: 'str', range_name: 'str', value_input_option: 'str', values: 'list[list[str]]', creds: 'Credentials'):
    """
    Write the content into google sheets

    Parameters
    ----------
    spreadsheet_id : str
        The id of the spreadsheet.

    range_name : str
        Sheets name, colume and row name of the cells you want to retrieve. example: A1:C2

    value_input_option : str
        How we want to input our value

    values : list[list[str]]
        The values we want to write in

    creds : Credentials
        Credentials of google api
    """
    
    try:

        service = build('sheets', 'v4', credentials=creds)
        body = {
            'values': values
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption=value_input_option, body=body).execute()
        print(f"{result.get('updatedCells')} cells updated.")
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


def get_volunteer_info(mainTrackingForm_id: 'str', creds: 'Credentials') -> 'list[list[str]]':
    """
    Return the formated information retrieved from the main tracking google sheet for all active volunteers.

    Parameters
    ----------
    mainTrackingForm_id : str
        The Google Sheet URL ID for the main OPT tracking sheet

    creds : Credentials
        Credentials of google api

    Returns
    -------
    list[list[str]]
        Information for each active volunteers in format:
        [Name, Email, Start Week, Start Monday, Start Date, End Date, Sheet URL ID]
    """
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # Access the form and retrieve all data
    mainTrackingFormData = get_values(mainTrackingForm_id, "'OPT subscription tracking'!A2:G", creds)

    volunteerInfo = []
    for formData in mainTrackingFormData:

        renewed = 0
        if (formData[1] != ""):
            renewed = int(formData[1])

        if (renewed >= 0):
            name = formData[4]
            email = formData[5]

            endDate = formData[3].split('/')
            endDateFormated = datetime.date(int(endDate[2]), int(endDate[0]), int(endDate[1]))

            # [Year, Week Number, Weekday (1 - 7)]
            endDateWeekInfo = endDateFormated.isocalendar()

            # Calculate the start date based on end date and renewed times
            startDateFormated = (endDateFormated - 
                                    dateDelta.relativedelta(months = (renewed + 1) * 3) + 
                                    dateDelta.relativedelta(days = 1))
            startDateWeekInfo = startDateFormated.isocalendar()
            startDateAsWeekday = startDateWeekInfo[2]

            # Temperary Solution:   Some people will delay their start week to the week after their start date week
            #                       i.e., start at Wed/Fri, but start record by next Mon (week)
            deltaFromStartDateToClosestMonday = (1 if startDateAsWeekday < 3 else 8) - startDateAsWeekday
            mondayAtStartWeek = startDateFormated + dateDelta.relativedelta(days = deltaFromStartDateToClosestMonday)
            startWeekInfo = mondayAtStartWeek.isocalendar()

            sheetURL = formData[6]

            # Format the volunteer information
            info = [name,
                    email,
                    startWeekInfo,
                    mondayAtStartWeek.strftime(dateFormat),
                    startDateFormated.strftime(dateFormat),
                    endDateFormated.strftime(dateFormat),
                    sheetURL]

            volunteerInfo.append(info)

    return volunteerInfo


def verify_weekly_report(volunteerInfo: 'list[str]', spreadsheet_id: 'str', creds: 'Credentials') -> bool:
    """
    Verfiy the validity for a given weekly report sheet:

    1. Recorded activities for each week.
    2. Records should match the activity period (start/end date).
    3. Record for each week should be unique.

    TODO: Update the verification logic after confirming the sheet format

    Parameters
    ----------
    volunteerInfo : list[str]
        Volunteer information for a single volunteer
        [Name, Email, Start Week, Start Monday, Start Date, End Date, Sheet URL ID]

    spreadsheet_id : str
        The Google Sheet URL ID for a single volunteer weekly report sheet

    creds : Credentials
        Credentials of google api

    Returns
    -------
    bool
        Boolean indicates if the given report sheet is valid with given volunteer information
    """
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # Retrieve the basic information of the volunteer
        weeklyReport = sheet.get(spreadsheetId = spreadsheet_id).execute()

        # Verify if we are on the correct sheet
        volunteerName = weeklyReport['properties']['title'].split('-')[0]
        if volunteerName != volunteerInfo[0]:
            raise Exception("Given volunteer name doesn't match with name on given sheet!")

        currentDate = datetime.date.today()
        currentMonday = currentDate + dateDelta.relativedelta(days = 8 - currentDate.isocalendar()[2])

        startMonday = datetime.datetime.strptime(volunteerInfo[3], dateFormat).date()

        workingWeeksByStartEndDate = int((currentMonday - startMonday).days / 7)


        # Start checking the report content

        startWeekNumColumn = 'C'
        startWeekNumRow = 8
        endWeekNumRow = startWeekNumRow + workingWeeksByStartEndDate - 1

        startRecordColumn = 'D'
        endRecordColumn = 'F'
        startRecordRow = 8
        endRecordRow = startRecordRow + workingWeeksByStartEndDate - 1

        weekNumGrid = startWeekNumColumn + str(startWeekNumRow) + ':' + startWeekNumColumn + str(endWeekNumRow)
        recordGrid = startRecordColumn + str(startRecordRow) + ':' + endRecordColumn + str(endRecordRow)

        rawWeekInfo = get_values(spreadsheet_id, weekNumGrid, creds)
        rawRecordInfo = get_values(spreadsheet_id, recordGrid, creds)

        for week, record in zip(rawWeekInfo, rawRecordInfo):
            print(week, " ", record)

        return True
    
    except HttpError as err:
        print(err)
        
def send_email(to_email: 'str', to_name: 'str', duplicate: 'bool'):
    """
    send email to the email address assigned from email specified in email_data

    Parameters
    ----------
    to_email : str
        The email address you want to send email to
    to_name : str
        The name of the person you want to send email to
    duplicate: bool
        Whether the email is a reminder of duplicate content
    """
    # SMTP settings
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = email_data.USER
    smtp_password = email_data.PASSWORD

    # IMAP settings
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(smtp_username, smtp_password)
    mail.select('inbox')
    if duplicate:
        reply_body = ("Hi " + to_name + ",\n\n" + 
                            "Our weekly scanning system identifies that you have duplicate check-in contents in your weekly report, please DO NOT copy any weekly work content, please input new contents that summarizes your work.\n\n" +
                            "Zenativity Admin Team")
    else:
        reply_body = ("Hi " + to_name + ",\n\n" + 
                            "Our weekly scanning system identifies that you did not check in your weekly report in time, please check in asap. If you miss the check in for an accumulative of 3 times and no further action is taken, we will have to terminate your volunteer term with us IMMEDIATELY, as it does not meet with USCIS requirements for OPT student to have consistent record keeping of your work.\n\n" +
                            "Zenativity Admin Team")
    reply_msg = MIMEText(reply_body)
    reply_msg['From'] = smtp_username
    reply_msg['To'] = to_email
    reply_msg['Subject'] = 'Reminder for your Weekly Report'

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(smtp_username, to_email, reply_msg.as_string())
        
    mail.close()
    mail.logout()
