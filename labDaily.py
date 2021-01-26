import pyodbc
import configparser
import pandas as pd
import openpyxl
import datetime
import time
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import os
import smtplib
import logging

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
pwd = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

today = datetime.datetime.now()
day = today.strftime('%y%m%d')

query = "select t02.opty, t02.oproc, t03.nurst, t02.oord#, t02.opat#, t02.osdate, \
        TIME(SUBSTR(DIGITS(t02.OSTIME),1,2) CONCAT ':' CONCAT SUBSTR(DIGITS(\
        t02.OSTIME),3,2)) AS STIME, t01.rhiscldt, TIME(SUBSTR(DIGITS(t01.rhcltm),1,2) CONCAT ':' CONCAT SUBSTR(DIGITS(\
        t01.rhcltm),3,2)) AS CTIME, t01.rhtc, \
        t01.rhisrndt, TIME(SUBSTR(DIGITS(t01.rhrntm),1,2) CONCAT ':' CONCAT SUBSTR(DIGITS(\
        t01.rhrntm),3,2)) AS RTIME, t01.rhisvfdt, TIME(SUBSTR(DIGITS(t01.rhvftm),1,2) CONCAT ':' CONCAT SUBSTR(DIGITS(\
        t01.rhvftm),3,2)) AS VTIME, t01.rhrcby \
    from orderf062.rh t01 LEFT OUTER JOIN \
    orderf062.oeorder t02 on t01.rhpt# = t02.opat# and t01.rhor# = t02.oord# LEFT OUTER JOIN \
    hospf062.rmbed t03 on t01.rhpt# = t03.pat# and t02.opat# = t03.pat# \
    WHERE t02.osdate = '" + day + "' \
	AND t02.ostime <= '400' \
    AND t02.OTODPT != 'RAD' \
    order by t03.nurst, STIME asc"

conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)



data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
# writer = pd.ExcelWriter('lab-report' + str(day) + '.xlsx')
df.to_excel('/home/itadmin/automation/files/lab_report.xlsx', index=False)

msg = MIMEMultipart()
msg['Subject'] = "Lab Tat Report For: " + str(day)
recipients = ['setdud@mckweb.com', 'TraHic@MCKweb.com', 'MelHub@MCKweb.com', 'KarKla@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/lab_report.xlsx'

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
print("Lab tat Report email Sent On " + day)
