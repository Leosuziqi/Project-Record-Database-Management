import win32com.client, datetime
from datetime import date
from dateutil.parser import *
import calendar
import pandas as pd

def fetch_calendar():
    """Win32com.client and python"""
    "access outlook and get events from the calendar"
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    # turn this into a list to read more calendars
    recipient = ns.CreateRecipient("Support department shared calendar")  # cmd whoami to find this
    if recipient.Resolve():
        print("yes")
    else:
        print("no")
    resolved = recipient.Resolve()  # checks for username in address book

    appts = ns.GetSharedDefaultFolder(recipient, 9).Items
    "sort events by occurent and include recurring events"
    appts.Sort("[Start]")
    appts.IncludeRecurrences = "True"

    """Filter appointments to a date range"""
    "filter to the range: from = (today-10), to = (today)"
    """Future dates
    begin = date.today().strftime("%m/%d/%y")
    end = date.today()+datetime.timedelta(days=30)
    end = end.strftime("%m/%d/%y")
    appts = appts.Restrict("[Start] >= '"+begin+"' AND [END] <= '"+end+"'")
    """
    # History dates
    end = date.today().strftime("%m/%d/%y")
    begin = date.today()-datetime.timedelta(days=180)
    begin = begin.strftime("%m/%d/%y")
    appts = appts.Restrict("[Start] >= '"+begin+"' AND [END] <= '"+end+"'")
    """Extract specific attributes"""
    "create list of excluded meeting subjects"
    excluded_subjects=('<first excluded subject>', '<second excluded subject>', '<third excluded subject>', '<etc … >')
    "populate dictionary of meetings "
    apptDict = {}

    item = 0
    for indx, a in enumerate(appts):
        subject_temp = str(a.Subject)
        if subject_temp in (excluded_subjects):
            continue

        if subject_temp[0] == '(':
            subject_temp = subject_temp[1:]

        if ord(subject_temp[0]) > 48 and ord(subject_temp[0]) < 58:
            project_id = subject_temp[0:5]
            subject = subject_temp[6:]
        else:
            project_id = None;
            subject = subject_temp

        organizer = str(a.Organizer)
        meetingDate = str(a.Start)
        start_date = parse(meetingDate).date()
        #subject = str(a.Subject)
        #print(subject)
        number_day = a.duration/60/24
        duration = str(number_day)
        end_date = a.Start+datetime.timedelta(days=int(number_day))

        apptDict[item] = { "project_id": project_id, "Duration": duration,  "Subject": subject, "Start_Date":start_date.strftime("%m/%d/%y"),"Organizer": organizer, "End_Date":end_date.strftime("%m/%d/%y")
                          }
        item = item +1

    """Group the results by date"""
    "convert discretionary to dataframe and group_by Date"
    apt_df = pd.DataFrame.from_dict(apptDict, orient='index', columns=['Start_Date','project_id','Duration', 'End_Date', 'Organizer', 'Subject'])
    apt_df = apt_df.set_index('Organizer')

    """ Save to .excel file"""
    "add timestamp to filename and save"
    filename = date.today().strftime("%Y%m%d") + '_30days_meeting_list.xlsx'
    apt_df.to_excel(filename)

    return apt_df

