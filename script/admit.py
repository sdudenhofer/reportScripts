import configparser
import pyodbc
import pandas as pd
import openpyxl
import datetime
import schedule
import time
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os
import time
import logging

config = configparser.ConfigParser()
config.read('config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pwd = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

#logging.basicConfig(filename='/home/itadmin/logs/admit.log', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

count=0
while(count <= 2):
    try:
        conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)
        count = 4
    except pyodbc.Error as e:
        #logging.info("DB Error: " + str(e))
        print("DB ERROR SLEEP 2 minutes tried " + str(count) + " times")
        time.sleep(120)
        count = count + 1

today = datetime.datetime.now()
c = datetime.timedelta(days=1)
d = today - c
yesterday = d.strftime("%Y-%m-%d")
day = today.strftime('%Y-%m-%d')

writer = pd.ExcelWriter('admit-report-' + str(day) + '.xlsx', engine='openpyxl')
query = "SELECT t01.hssvc, t01.patno, t01.pname, t01.isadate, t02.room, t02.bed, t02.nurst \
FROM hospf0062.patients t01 LEFT OUTER JOIN hospf0062.rmbed t02 ON t01.patno = t02.pat# \
WHERE t01.isadate = '" + yesterday + "' and t02.nurst !='NULL' AND HSsvc !='OBS' ORDER BY t01.hssvc"
query2 = " SELECT t01.hssvc, t01.patno, t01.pname, t01.isadate, t02.room, t02.bed, t02.nurst \
FROM hospf0062.patients t01 LEFT OUTER JOIN hospf0062.rmbed t02 ON t01.patno = t02.pat# \
WHERE t01.isadate = '" + yesterday + "' and t02.nurst !='NULL' AND HSsvc = 'OBS' ORDER BY t01.hssvc"
    # add error handling to sql query
data = pd.read_sql(query, conn)
data2 = pd.read_sql(query2, conn)
df = pd.DataFrame(data)
df2 = pd.DataFrame(data2)
df.to_excel(writer, sheet_name='inpatient', index=False)
df2.to_excel(writer, sheet_name='OBS', index=False)
writer.save()

    # setup and send emails
msg = MIMEMultipart()
msg['Subject'] = "Admission Report For: " + str(day)
recipients = ['setdud@mckweb.com']
#, 'RBootes@MCKWeb.com', 'AneHum@MCKweb.com', 'EriMcc@mckweb.com']
#for to in recipients:
#    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = 'admit-report-' + str(day) + '.xlsx'

msg.attach(MIMEText("Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()

