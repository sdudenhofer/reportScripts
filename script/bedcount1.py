import pandas as pd
import configparser
import pyodbc
import xlsxwriter
import time
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os, sys
import schedule
from ftplib import FTP
from time import sleep
import ftplib


#getting login/passwords from configuration file

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pword = config['MHD']['PASS']
fserver = config['FTP']['SERVER']
fuser = config['FTP']['USER']
fpass = config['FTP']['PASS']
fpath = config['FTP']['PATH']

# connection = fserver
ftp = FTP()
port = 21
username = fuser
password = fpass
ftp.connect(fserver, port)




try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=pword)
except pyodbc.Error as e:
    print(e)
    sleep(60)

query = "select nurst, room, bed, pat# from hospf0062.rmbed"
cursor = conn.cursor()
data = cursor.execute(query)
array = []
array1 = []

today = datetime.datetime.now()
today1 = today.isoformat()
    #today3 = today1.strftime("%Y-%m-%dT%H:%M:%SZ")
first = str(today1).split(".")[0] + "Z"
date = str(first).split(":")[0] + ":" + str(first).split(":")[1] + "Z"
print(date)
day = today.strftime("%Y%d%m%H%M%S")
writer = open("/home/itadmin/automation/files/MWMC." + str(day) + ".csv", "w+")
for row in data:
    if str(row[3]) == '0':
        row[3] = "FALSE"
    elif row[0] == "PCU":
        vent = "TRUE"
    else:
        block = "FALSE"
        vent = "FALSE"
        row[3] = "TRUE"
    array.append(str(date) + ", " + "McKenzie Willamette Medical Center,McKenzie Willamette Medical Center," + str(row[0]) + ", " + str(row[1]) + "," + str(row[2]) + "," + str(row[3]) + ", FALSE,FALSE")
    writer.write('Timestamp of record,System,Facility,Unit,Room,Bed,isOccupied,isBlocked,isVentCapable \n')
for r in array:
    bed = r.split(',')[5]
    dept = r.split(',')[3]
        #print(room)
    if bed == " P ":
        array1.append(r)
    elif dept == " NUR ":
        array1.append(r)
    else:
        writer.write(r + "\n")

    # original connection here
ftp_file = "/home/itadmin/automation/files/MWMC." + str(day) + ".csv"

fc = open(ftp_file, 'rb')
stor_file = str("STOR MWMC." + str(day) + ".csv")
    # ftp.connect(connection, port)

# trying to catch errors at any step in the sftp transfer. Errors occuring at ftp.cwd winerror 10060

for i in range(0, 10):
    while i <= 10:
        try:
            ftp.login(fuser, fpass)
        except ftplib.all_errors as f:
            # errorcode_string = str(f).split(None, 1)[0]
            print(str(f))
                #sleep(60)
            continue
        except ftplib.error_reply as h:
            print(str(h))
            sleep(60)
            continue
        break
try:
    ftp.cwd(fpath)
except ftplib.all_errors as g:
    print(str(g))
    
try:
    ftp.storbinary(stor_file, fc, 1024)
    fc.close()
    conn.close()
except ftplib.all_errors as i:
    print(str(i))
fc.close()

writer.close()
path = "/home/itadmin/automation/files/"
now = time.time()

for f in os.listdir(path):
    f = os.path.join(path, f)
    if os.stat(f).st_mtime <= now - 15  * 60:
        if os.path.isfile(f):
            os.remove(f)
    # conn.close()
    # ftp.quit()
