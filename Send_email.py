import xlrd
import schedule
import datetime as dt
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


def send_email(email_recipient,
               email_subject,
               email_message):

    email_sender = 'blueno_match@outlook.com'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login('blueno_match@outlook.com', 'Chen7nipun3Neel6')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent')
        server.quit()
    except:
        print("SMPT server connection error")
    return True

# read the file
path = '/Users/jimmyjianchen/Desktop/Brown/Blueno_Match/Final_Matches.xls'
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

# calculate row and column numbers
rowNumber = inputWorksheet.nrows
columnNumber = inputWorksheet.ncols

allReceivers = []
allMatchedGuys = []
allSendingTexts = []

for i in range(rowNumber):
    allReceivers.append(inputWorksheet.cell_value(i, 0))
    allMatchedGuys.append(inputWorksheet.cell_value(i, 1).split('\n')[0])

for i in range(rowNumber):
    curr = '''Dear participant,\n
     Hi,
     Thank you for participating in Blueno Match! We are glad to inform you that you have been matched with '''
    curr = curr + allMatchedGuys[i]
    curr = curr + ''' for this week. You can message them to meet up in person or on Zoom. Hope you guys have fun meeting and become good friends!
     Blueno Match is a weekly event so you can participate in next week’s matching as well. We will send out the link for that on Sunday. The earlier you fill out the survey, the more priority you will have in matching.
     Since the people you are matched with are selected from other people who filled out this survey, you will likely have better matches if more people participate in Blueno Match. Therefore, if this is something you end up benefitting from, try to get the word out so more people can join. 
     We are only responsible for matching people, not responsible for what happens during the meetups. Below are some reminders for you in case of extreme situations in in-person meetups. We don’t anticipate such situations to happen, but it’s important to have these resources ready beforehand.\n
Note:
1. It is safer to meet in public places.
2. Text each other before meeting.
3. Follow Brown's COVID rules.\n
     We also link here some resources you can use in case of emergency (they are also on the back of your Brown card):\n
Public Safety (Emergency): (401) 863-4111
Sexual Assault Response Line: 863-6000
Administrator on Call: 863-3322
Counselling & Psychological Services: 863-3476
Health Services: 863-3953
Report facility problems: 863-7800
Public Safety (Routine Calls): 863-3322
Request Safewalk: 863-1079
Brown University Shuttle: 863-2322\n
     Have a great weekend!\n
     Sincerely,
     Blueno Match'''
    allSendingTexts.append(curr)

def sendAllEmails():
    for i in range(rowNumber):
        send_email(allReceivers[i], 
        'Your Match Results for Week One', 
        allSendingTexts[i])

#schedule.every().day.at('08:00').do(sendAllEmails)
sendAllEmails()
#while 1:
#    schedule.run_pending()
#    time.sleep(1)
