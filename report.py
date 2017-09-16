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
                    print(sheetName + ', ' + str(line[8].value))
    return last
	
def GetArchived(sheet,sheetName, days):
    last = []   
    stackTrace = ''
    for month in range(0,14):
        last.append(0)
    for line in sheet.rows:
        published = None
        if (len(line) > 0 and (line[0].value == 'Archived' or line[0].value == 'Published')):  
            if (len(line) < 9 or line[8].value == None):
                last[13] =  last[13] + 1
            else:
                try:       
                    published =  utils.datetime.from_excel(line[8].value)			
                except:
                    try:
                        published = datetime.datetime.strptime(str(line[8].value),'%Y-%m-%d %H:%M:%S' )
                          
                    except Exception as ex:
                        msg = str(ex) + ' TAB Name ' + sheetName
                        print(msg)
                        stackTrace += '<div>'+msg + '</div>'
                        last[12] =  last[12] + 1
                if (published!= None  and published.year == 2017):                          
                    last[published.month-1] =  last[published.month-1] + 1 
                         
    t = last, stackTrace     
    return  t 	
    

def GetDailyReport(urls,last,submitted, days):
    msg = '<p><b>' + str(submitted) + ' Submitted articles are pending to be published. </b></p>'
    msg += '<p><b>{1} Published articles in the last {0} days : </b></p>'.format(days,len(last))
    for h in last.items():
    	msg += '<div><a href="'+ h[1]+'">'+ h[0] + '</a></div>'
    msg +=  '<p><b>'+ str(len(urls))+' Published articles with missing Article URL : </b></p>'
    for h in urls:
    	msg += '<div>\t' + h + '</div>'
    msg = "<html><head></head><body>" + msg + "</html></body>"
    return msg
def GetColor(num):
    if (num == 0):
        return 'powderblue'
    elif (num < 7):
        return'rgb(255, 128, 128)'
    elif (num < 15):
        return'yellow'
    else:
        return'LimeGreen'
def GetAnualReport(archived, year,name, stackTrace):
    msg = '<center><p style="font-size:40px"><b>{0} archived and published articles summary for 2017 </b></p></center>'.format(name)
    
    msg += '<table style="width:100%">'
    msg += '<tr>'
    msg += '<th>Region/Month</th>'    
    for month in range(1,13):
        msg += '<th>'+ calendar.month_name[month]+'</th>'
    msg += '<th>No Date</th>'
    msg += '<th>Invalid Date</th>'
    msg += '<th>Average</th>'
    msg += '</tr>'
    monthTotals = []
    writerTotals = 0
    writerActiveMonths = 0
    for month in range(0,14):
        monthTotals.append(0)   
    
    for h in archived.items():         
        msg += '<tr>'
        msg += '<th>'+ h[0]+'</th>'        
        count = 0
        for month in h[1]:
            monthTotals[count] += month
            msg += '<th style="background-color:{0};">'.format(GetColor(month))+ str(month)+'</th>'
            if (month > 0 and count < 13):
                writerActiveMonths+=1
                writerTotals+=month
            count += 1
        avg = round(writerTotals/writerActiveMonths,2)
        msg +='<th style="background-color:{0};">'.format(GetColor(avg))+str(avg)+'</th>'         
        
    msg += '</tr>'
    msg += '<tr>'
    msg += '<th>Total</th>'   
    for month in monthTotals:        
        msg += '<th style="background-color:LightGrey;">'+ str(month)+'</th>'
    msg += '</tr>'
    msg += '</table>'
    msg += '<p style="font-size:20px"><b> Invalid Dates Info</b></p>'
    msg += stackTrace
    return msg
def SendEmail(msg,email,name,reportName,toAll): 
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

def CreateReport(days,reportName,email,name,isAnualReport):
    stackTrace = ''
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
        if (not isAnualReport):
            urls += GetMissingUrls(wb[sheet])
            for item in GetLastWeek(wb[sheet],sheet,days).items():
                last[item[0]] =item[1]
            submitted += GetAllSubmitted(wb[sheet])
        else:
            archived[sheet] = []
            res = GetArchived(wb[sheet],sheet,days)
            stackTrace += res[1]
            for item in res[0]:            
                archived[sheet].append(item)                
    if (not isAnualReport):
        msg = GetDailyReport(urls,last, submitted,days)
    else:
        msg = GetAnualReport(archived,2017,name, stackTrace)
    SendEmail(msg,email,name, reportName,False)