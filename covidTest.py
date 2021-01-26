import pandas as pd
import configparser
import pyodbc
import xlsxwriter
import xlrd
import openpyxl
import time
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os
import logging

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

logging.basicConfig(filename="/home/itadmin/logs/covidTest.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

today = datetime.datetime.now()
c = datetime.timedelta(days=1)
b = today - c
yesterday = b.strftime("%Y-%m-%d")
day = b.strftime("%Y-%m-%d")
writer = pd.ExcelWriter('/home/itadmin/automation/files/community-test-' + str(day) + '.xlsx', engine='openpyxl')

query = "select t01.patno, t01.pname, t01.isadate, t02.phname  from hospf062.patients t01 \
left outer join hospf062.phymast t02 on t01.nwattphy = t02.nwdrnum \
where hssvc = 'RE2' and t01.isadate = '" + yesterday + "' order by t02.phname"

for i in range(0, 10):
    while i <= 10:
        try:
            conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
        except pyodbc.Error as f:
            logging.INFO(str(f))
            continue
        break

data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
df.to_excel(writer, sheet_name='covid', index=False)
writer.save()
#set up mail and email data to requested parties
msg = MIMEMultipart()
msg['Subject'] = 'Covid-19 Doctor Report'
recipients = ['setdud@mckweb.com', 'RBootes@MCKWeb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/community-test-' + str(day) + '.xlsx'

msg.attach(MIMEText("COVID19 Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()
print("Email Sent: " + str(today))

conn.close()
