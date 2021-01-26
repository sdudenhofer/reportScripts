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
import logging

logging.basicConfig(filename='/home/itadmin/logs/census.log', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
pwd = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']
try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)
except pyodbc.OperationalError as e:
    logging.INFO("Error number {0}: {1}.".format(e.args[0],e.args[1]))
    time.sleep(60)
    print("Trying again...")

today = datetime.datetime.now()
day = today.strftime('%Y-%m-%d')

query = "SELECT T01.NURST, T01.ROOM, T01.BED, T01.OBSERV, T02.PNAME, T02.AGE, T02.DIAGN, T03.PHNAME \
    FROM HOSPF062.RMBED T01 LEFT OUTER JOIN HOSPF062.PATIENTS T02 ON T01.PAT# = T02.PATNO \
    LEFT OUTER JOIN HOSPF062.PHYMAST T03 ON T02.NWDOCNUM = T03.NWDRNUM \
    WHERE NURST != 'SSU' AND NURST != 'EOP' AND PAT# > 0 ORDER BY NURST"

data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
df.to_excel('/home/itadmin/automation/files/census-report-' + str(day) + '.xlsx', index=False)

msg = MIMEMultipart()
msg['Subject'] = "Census Report For: " + str(day)
recipients = ['setdud@mckweb.com', 'DesShu@MCKweb.com', 'ValSim@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/census-report-' + str(day) + '.xlsx'

msg.attach(MIMEText("Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
