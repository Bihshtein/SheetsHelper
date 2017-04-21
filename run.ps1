cd C:\SheetsHelper
$webClient = New-Object System.Net.WebClient
$url = "https://www.python.org/ftp/python/2.7.13/python-2.7.13.msi"
$file  = "C:\SheetsHelper\python-2.7.13.msi"
if(![System.IO.File]::Exists($file)){
    $webClient.DownloadFile($url,$file)
}
$url = "https://bootstrap.pypa.io/get-pip.py"
$file  = "C:\SheetsHelper\get-pip.py"
if(![System.IO.File]::Exists($file)){
    $webClient.DownloadFile($url,$file)
}
Start-Process "python-2.7.13.msi" /qn -Wait
C:\python27\python.exe get-pip.py
C:\python27\python.exe -m pip install -U pip openpyxl
C:\python27\python.exe -m pip install --upgrade google-api-python-client
C:\python27\python.exe download.py 7

