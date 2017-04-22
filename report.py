# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import smtplib
import datetime
import urllib
from openpyxl import load_workbook
from openpyxl import utils
from oauth2client import client
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def removeNonAscii(s): return "".join(i for i in s if ord(i)<128)
def GetMissingIDS(sheet):
    ids = []
    for line in sheet.rows:
        if (line[0].value == 'Submitted' and line[5].value == None  and line[1].value != None):
            ids.append(str(removeNonAscii(line[1].value)))           
    return ids
                        
def GetMissingUrls(sheet):
    urls = []
    for line in sheet.rows:
        if (line[0].value == 'Published' and line[6].value == None):
            urls.append(str(line[1].value))
    return urls

def GetAllSubmitted(sheet):
    count = 0
    for line in sheet.rows:
        if (line[0].value == 'Submitted'):
            count += 1
    return count

def GetLastWeek(sheet,days):
    last = {}
    for line in sheet.rows:
        if (line[0].value == 'Published' and line[8].value != None):
            now = datetime.datetime.now()
          
            try:
                published =  utils.datetime.from_excel(line[8].value)
            except:
                try:
                      published = datetime.datetime.strptime(str(line[8].value),'%Y-%m-%d %H:%M:%S' )
                except:
                    print line[8].value
                    
                                
            diff  =  now - published
            if (diff.days <= days and line[1].value != None and line[6].value != None): 
                last[str(removeNonAscii(line[1].value))]=str(removeNonAscii(line[6].value))           
                
    return last

def SendEmail(sheets,urls,ids,last,submitted, days,email):  
    msg = '<p><b>Active sheets : </b></p>' 
    for s in sheets:
        if (sheets.index(s) % 10 == 0):
            msg += '<div></div>'
        msg += str(s) + ', '        
    msg += '<p><b>' + str(submitted) + ' Submitted articles are pending to be published. </b></p>'
    msg += '<p><b>{1} Published articles in the last {0} days : </b></p>'.format(days,len(last))   
    for h in last.iteritems():
        msg += '<div><a href="'+ h[1]+'">'+ h[0] + '</a></div>'
    msg += '<p><b>' + str(len(ids)) + ' Submitted articles with missing Post ID : </b></p>'
    for h in ids:
        msg += '<div>\t' + h + '</div>'                
    msg +=  '<p><b>'+ str(len(urls))+' Published articles with missing Article URL : </b></p>'
    for h in urls:
        msg += '<div>\t' + h + '</div>'        
    msg = "<html><head></head><body>" + msg + "</html></body>"
    
    with open('C:\SheetsHelper\msg.html', "wb") as wer:
        wer.write(msg)
            
    fromAdd = 'bihshtein@hotmail.com'
    toAdd = [email,'bihshtein@hotmail.com']
    emsg = MIMEMultipart('alternative')
    emsg['Subject'] = "Sheet Report"
    part2 = MIMEText(msg, 'html')
    emsg.attach(part2         )
    s = smtplib.SMTP('smtp.live.com:587')#
    s.starttls()
    s.login(fromAdd, 'AlegAleg')
    s.sendmail(fromAdd, toAdd, emsg.as_string())
    s.quit()
                
def CreateReport(days,email):
    ids = []
    urls = []
    last = {}
    submitted = 0   
    wb = load_workbook('C:\SheetsHelper\calendar.xlsx',read_only=True)
    allSheets = wb.get_sheet_names()
    unusedSheets = ['Copy Editors & Writers']
    for sheet in unusedSheets:
        allSheets.remove(sheet)
    for sheet in allSheets:   
        ids += GetMissingIDS(wb[sheet])
        urls += GetMissingUrls(wb[sheet])
        for item in GetLastWeek(wb[sheet],days).iteritems():
            last[item[0]] =item[1] 
        submitted += GetAllSubmitted(wb[sheet])
    SendEmail(allSheets,urls,ids, last, submitted,days,email)