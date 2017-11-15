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
from enum import Enum


class ReportType(Enum):
    Monthly = 1
    Weekly = 2
    Daily = 3   


invalidDateShift = 1
noDateShift = 2
averageShift = 4
submittedShift = 3	


    
def GetIndex(date, reportType):    
    if (reportType==ReportType.Monthly):
        return date.month -1
    elif (reportType==ReportType.Weekly):
        firstDay = datetime.datetime(date.year,1,1).weekday()
        daysToAdd = 0
        if (firstDay < 2):
            daysToAdd = 2-firstDay
        elif (firstDay > 2):
            daysToAdd = (6- firstDay+2+1)        
        return int((date - datetime.datetime(date.year,1,1 + daysToAdd)).days/7)
    else: 
        return (date - datetime.datetime(date.year,1,1)).days
    
def GetArchived(sheet,sheetName, year, reportType, max):    
    last = []   
    firstDay = datetime.datetime.now() - datetime.timedelta(days=1)
    stackTrace = ''
    publishedMsg = ''
    currIndex = GetIndex(firstDay,reportType)
    totalColumns =  max + averageShift    
    for month in range(0,totalColumns-1):
        liist = []       
        last.append(liist)    
    for line in sheet.rows:
        published = None
        if (len(line) > 0 and  line[0].value == 'Submitted'):            
            last[max+submittedShift-1].append(str(line[1].value))          
        if (len(line) > 1 and (line[0].value == 'Archived' or line[0].value == 'Published')):                          
            if (len(line) < 9 or line[8].value == None):                           			
                last[max + noDateShift-1].append(str(line[1].value))                
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
                        last[max+invalidDateShift-1].append(str(line[1].value))                        
                if (published != None  and published.year == year and ((currIndex - GetIndex(published,reportType))  < max)):                                                                
                   
                    index = currIndex-GetIndex(published,reportType)
                    last[index].append('aleg')                    
                    if (index == 0):
                        if (publishedMsg == ''):
                            publishedMsg = '<div>'+sheetName+'</div>'
                        publishedMsg += '<a href="'+ str(line[6].value)+'">'+ str(line[1].value) + ' | </a>'                                                 
    t = last, stackTrace , publishedMsg   
    return  t 	
    

def GetColor(num,reportType):
    if (reportType==ReportType.Monthly):
        if (num == 0):
            return 'LightGrey'
        elif (num < 7):
            return'rgb(255, 128, 128)'
        elif (num < 15):
            return'yellow'
        else:
            return'LimeGreen'
    elif (reportType==ReportType.Weekly):
        if (num == 0):
            return 'LightGrey'
        elif (num < 2):
            return'rgb(255, 128, 128)'
        elif (num < 4):
            return'yellow'
        else:
            return'LimeGreen'
    else:
        if (num == 0):
            return 'LightGrey'
        elif (num  <0.25):        		
            return 'rgb(255, 128, 128)'
        elif (num  < 0.5):
            return'yellow'
        else:
            return'LimeGreen'
def GetAnualReport(allSheets, archived, year,name, stackTrace, publishedMsg, reportType, max):
    firstDay = datetime.datetime.now() - datetime.timedelta(days=1)
    currMonth = GetIndex(firstDay, reportType)
    msg = '<center><p style="font-size:40px"><b>{0} {1} published articles summary</b></p></center>'.format(name,reportType)    
    msg += '<table style="width:100%">'
    msg += '<tr>'
    timeframe = 'Weekday'
    if (reportType==ReportType.Weekly):
        timeframe = 'Week number'
    if (reportType==ReportType.Monthly):
        timeframe = 'Month'
    msg += '<th>Region/{0}</th>'.format(timeframe)    
    published = ''
    for colNum in range(0,max):
        if (reportType==ReportType.Monthly):
            msg += '<th>'+ calendar.month_name[firstDay.month-colNum]+'</th>'
        elif (reportType==ReportType.Daily):
            msg += '<th>'+ calendar.day_name[(firstDay - datetime.timedelta(days=colNum)).weekday()]+'</th>'
        else:
            msg += '<th>'+ str(colNum)+'</th>'
    msg += '<th>Invalid Date</th>'
    msg += '<th>No Date</th>'    
    msg += '<th>Submitted</th>'
    msg += '<th>Average</th>'    
    msg += '</tr>'
    monthTotals = []
   
    for month in range(0,max+averageShift):
        monthTotals.append(0)   
    
    for sheet in allSheets:       
        msg += '<tr>'
        msg += '<th>'+ sheet+'</th>'        
        count = 0
        writerTotals = 0
        writerActiveMonths = 0
        for listPub in archived[sheet]:            
            sum = len(listPub)
            monthTotals[count] += sum
            if (count < max):
                msg += '<th style="background-color:{0};">'.format(GetColor(sum,reportType))+ str(sum)+'</th>'
            else:
                if (sum ==0):
                    msg += '<th style="background-color:{0};">'.format('LightGrey')+ str(sum)+'</th>'
                elif (sum <5):
                    msg += '<th style="background-color:{0};">'.format('yellow')+ str(sum)+'</th>'
                else:
                    msg += '<th style="background-color:{0};">'.format('rgb(255, 128, 128)')+ str(sum)+'</th>'
            if ((sum > 0 or reportType == ReportType.Daily) and  count < max):
                writerActiveMonths+=1
                writerTotals+=sum
            count += 1
        avg = 0
        if (writerActiveMonths > 0):
            avg = round(writerTotals/writerActiveMonths,2)
        msg +='<th style="background-color:{0};">'.format(GetColor(avg,reportType))+str(avg)+'</th>'         
        
    msg += '</tr>'
    msg += '<tr>'
    msg += '<th>Total</th>'   
    for month in monthTotals:        
        msg += '<th style="background-color:powderblue;">'+ str(month)+'</th>'
    msg += '</tr>'
    msg += '</table>'    
    msg += '<p style="font-size:20px"><b> Invalid Dates Info</b></p>'    
    msg += stackTrace
    msg += '<p style="font-size:20px"><b> Published links</b></p>'    
    msg += publishedMsg
    return msg
def SendEmail(msg,email,name,reportName,toAll): 
    with open('C:\SheetsHelper\msg.html', "wb") as wer:
        wer.write(msg.encode('utf-8'))

    fromAdd = 'bihshtein@hotmail.com'
    toAdd = ['bihshtein@hotmail.com']
    if (toAll):
        toAdd.append(email)
        toAdd.append('Anthony.johnston@theculturetrip.com')
    emsg = MIMEMultipart('alternative')
    emsg['Subject'] = reportName + " for " + name
    part2 = MIMEText(msg, 'html')
    emsg.attach(part2         )
    s = smtplib.SMTP('smtp.live.com:587')#
    s.starttls()
    s.login(fromAdd, 'AlegAleg')
    s.sendmail(fromAdd, toAdd, emsg.as_string())
    s.quit()

def CreateReport(reportName,email,name,reportType, max):
    stackTrace = ''   
    publishedMsg = ''   
    urls = []
    last = {}
    archived = {}
    submitted = 0
    wb = load_workbook('C:\SheetsHelper\calendar.xlsx',read_only=True)
    allSheets = wb.get_sheet_names()
    print(allSheets)
    unusedSheets = ['Copy Editors & Writers']
    for sheet in unusedSheets:
        allSheets.remove(sheet)
    for sheet in allSheets:       
            archived[sheet] = []
            res = GetArchived(wb[sheet],sheet, 2017, reportType, max)
            stackTrace += res[1]
            publishedMsg += res[2]
            for item in res[0]:            
                archived[sheet].append(item)                    
    
    msg = GetAnualReport(allSheets, archived,2017,name, stackTrace, publishedMsg, reportType, max)
    SendEmail(msg,email,name, reportName,True)