import sys
from oauth2client import client
import httplib2
import os
import io
from oauth2client.file import Storage
import json
from httplib2 import Http
from apiclient.discovery import build
from oauth2client import client
from oauth2client import tools
import apiclient
import datetime
import calendar
import report

def GetService():
    appName = 'sheetassist'
    file = 'keys.json'
    scopes = ['https://www.googleapis.com/auth/drive']
    
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'drive-python-quickstart.json')
    store = Storage(credential_path)
    credentials = store.get()
    
    if not credentials or credentials.invalid:			
        flow = client.flow_from_clientsecrets(file, scopes)
        flow.user_agent = appName    
        credentials = tools.run_flow(flow,store)
    http = credentials.authorize (httplib2.Http())
    return build('drive', 'v3', http=http)
	
def DownloadFile(file):
    drive_service = GetService()  
    request = drive_service.files().export_media(fileId=file, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response = request.execute()
    with open('C:\SheetsHelper\calendar.xlsx', "wb") as wer:
        wer.write(response)  

def DownloadAllAndResport(max, reportType):    
    if (report.ReportType[reportType] == report.ReportType.Weekly):
        if (datetime.datetime.today().weekday() != 2):
    	    return 
    if (report.ReportType[reportType] == report.ReportType.Monthly):
        if (datetime.datetime.today().day != 1):
    	    return 
    reportName = str(reportType)
    editors = {}
    editors['siukei@theculturetrip.com;lily.niu@culturetrip.com'] = ['1AqMxekEh_JerFfeEkY_omaJqsGHgjA5246QsZlfNJBY','Siukei']
    #editors['andrew.headspeath@culturetrip.com'] = ['1QMOp5ygwIGwlgZu7h80GdnFAO1jd8jWvkFnaTITP5sk','Andy']    
    #editors['grace@culturetrip.com'] = ['1STKWMSN2yi_Bk-LdS6MimFYMwXQ_z8fRwjPG-zxQFeE','Grace']
    #editors['tahiera@theculturetrip.com'] = ['1_ZEl2HqnKprC-hOIUD6ti79MbBkrdfHxxvNLhWzbBF8','Tahiera']
    #editors['lily.niu@culturetrip.com'] = ['10RNpzBXpFUcjIABr5eIUO8yxSoP7fBoJ-6oggBeeWug','Lily']
    #editors['mariam@theculturetrip.com'] = ['1COEqPSZ78R7gOvJbWk1jhMyKz7pI5hmn8cyvRZW4dgs', 'Mariam']
    #editors['charlotte.peet@theculturetrip.com'] = ['1pFC7mIMhFvN6_6MTWY9tbx-XV6s9Zy8gjpYwlFKXx4c','Charlotte']
    for editor in editors.items():
        DownloadFile(editor[1][0])
        report.CreateReport(reportName, editor[0],editor[1][1],report.ReportType[reportType], max)
	


max = int(sys.argv[1])
reportType  = str(sys.argv[2])
DownloadAllAndResport(max, reportType)


