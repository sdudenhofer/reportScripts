import pyodbc
import pandas as pd
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
import os
import configparser
import xlrd

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')

server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

conn = pyodbc.connect(DSN=server, UID=user, PWD=password)

today = datetime.datetime.now()
a = datetime.timedelta(days=1)
b = today -a
yesterday = b.strftime("%Y-%m-%d")
day = today.strftime("%Y-%m-%d")

query = "SELECT T01.patno, t01.adate, t01.iatme, t01.isqadte, t01.iqtime, t01.nwdocnum, t01.hssvc, t02.atype, t02.asrce, t01.age, t01.fincl, t01.pname, t02.arrvd, t01.diagn, t01.hstnum \
        FROM hospf062.patients t01 LEFT OUTER JOIN hospf062.admreg t02 on t01.patno = t02.patno and t01.hstnum = t02.hstnum \
        where t01.hssvc = 'OBS' and t01.isadate = '" + yesterday + "'"

data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
writer = pd.ExcelWriter('/home/itadmin/automation/files/observation-' + str(day) + '.xlsx', engine= 'openpyxl')
df.to_excel(writer, sheet_name='Observation Patients', index=False)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = "Daily Observation Report for: " + str(day)
recipients = ['setdud@mckweb.com', 'terdym@mckweb.com', 'jjones@mckweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application', 'octet-stream')
f = '/home/itadmin/automation/files/observation-' + str(day) + '.xlsx'
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
