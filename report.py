# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import smtplib
import datetime
import urllib
import calendar
from openpyxl import load_workbook
from openpyxl import utils
from oauth2client import client
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def GetMissingUrls(sheet):
    urls = []
    for line in sheet.rows:
        if (len(line) > 6 and line[0].value == 'Published' and line[6].value == None):
            urls.append(str(line[1].value))
    return urls

def GetAllSubmitted(sheet):
    count = 0
    for line in sheet.rows:
        if (len(line) > 1 and line[0].value == 'Submitted'):
            count += 1
    return count

def GetLastWeek(sheet,sheetName,days):
    last = {}
    for line in sheet.rows:
        if (len(line) > 8 and line[0].value == 'Published' and line[8].value != None):
            now = datetime.datetime.now()
            try:
                published =  utils.datetime.from_excel(line[8].value)
                diff  =  now - published
                if (diff.days <= days and line[1].value != None and line[6].value != None):
                    last[str(line[1].value)]=str(line[6].value)
            except:
                try:
                      published = datetime.datetime.strptime(str(line[8].value),'%Y-%m-%d %H:%M:%S' )
                      diff  =  now - published
                      if (diff.days <= days and line[1].value != None and line[6].value != None):
                          last[str(line[1].value)]=str(line[6].value)
                except:
                    print(sheetName + ', ' +line[8].value)
    return last
	
def GetArchived(sheet,sheetName, days):
    last = []
    for month in range(0,11):
        last.insert(month,0)
    for line in sheet.rows:
        if (len(line) > 8 and (line[0].value == 'Archived' or line[0].value == 'Published') and line[8].value != None):  
            try:       
                published =  utils.datetime.from_excel(line[8].value)			
            except:
                try:
                    published = datetime.datetime.strptime(str(line[8].value),'%Y-%m-%d %H:%M:%S' )
                except Exception as ex:
                    print(str(ex) + ' sheet name ' + sheetName)
            if (published.year == 2017):                
                last[published.month-1] =  last[published.month-1] + 1
               
            
    return last

def SendEmail(sheets,urls,last,archived,submitted, days,email,name,reportName,onlyTable=False,toAll=True): 
    msg = ''
    if (onlyTable):
        msg += '<p><b>Archived and Published articles summary for 2017 </b></p>'
        msg += '<table style="width:100%">'
        msg += '<tr>'
        msg += '<th>Region/Month</th>'    
        for month in range(1,12):
            msg += '<th>'+ calendar.month_name[month]+'</th>'
        msg += '</tr>'
        for h in archived.items():
            msg += '<tr>'
            msg += '<th>'+ h[0]+'</th>'
            for month in h[1]:
                if (month == 0):
                    msg += '<th style="background-color:powderblue;">'+ str(month)+'</th>'
                elif (month < 7):
                    msg += '<th style="background-color:rgb(255, 128, 128);">'+ str(month)+'</th>'
                elif (month < 15):
                    msg += '<th style="background-color:yellow;">'+ str(month)+'</th>'
                else:
                    msg += '<th style="background-color:LimeGreen;">'+ str(month)+'</th>'
            msg += '</tr>'
        msg += '</table>'
    msg += '<p><b>' + str(submitted) + ' Submitted articles are pending to be published. </b></p>'
    msg += '<p><b>{1} Published articles in the last {0} days : </b></p>'.format(days,len(last))
    for h in last.items():
        msg += '<div><a href="'+ h[1]+'">'+ h[0] + '</a></div>'
    msg +=  '<p><b>'+ str(len(urls))+' Published articles with missing Article URL : </b></p>'
    for h in urls:
        msg += '<div>\t' + h + '</div>'
    msg = "<html><head></head><body>" + msg + "</html></body>"

    with open('C:\SheetsHelper\msg.html', "wb") as wer:
        wer.write(msg.encode('utf-8'))

    fromAdd = 'bihshtein@hotmail.com'
    toAdd = ['bihshtein@hotmail.com']
    if (toAll):
        toAdd.append(email)
        #toAdd.append('Anthony.johnston@theculturetrip.com')
    emsg = MIMEMultipart('alternative')
    emsg['Subject'] = reportName + " for " + name
    part2 = MIMEText(msg, 'html')
    emsg.attach(part2         )
    s = smtplib.SMTP('smtp.live.com:587')#
    s.starttls()
    s.login(fromAdd, 'AlegAleg')
    s.sendmail(fromAdd, toAdd, emsg.as_string())
    s.quit()

def CreateReport(days,reportName,email,name):
    urls = []
    last = {}
    archived = {}
    submitted = 0
    wb = load_workbook('C:\SheetsHelper\calendar.xlsx',read_only=True)
    allSheets = wb.get_sheet_names()
    unusedSheets = ['Copy Editors & Writers']
    for sheet in unusedSheets:
        allSheets.remove(sheet)
    for sheet in allSheets:
        urls += GetMissingUrls(wb[sheet])
        for item in GetLastWeek(wb[sheet],sheet,days).items():
            last[item[0]] =item[1]
        archived[sheet] = []
        for item in GetArchived(wb[sheet],sheet,days):            
            archived[sheet].append(item)
        submitted += GetAllSubmitted(wb[sheet])
    SendEmail(allSheets,urls,last, archived, submitted,days,email,name, reportName,True,True)