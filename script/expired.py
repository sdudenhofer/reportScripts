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

logging.basicConfig(filename="/home/itadmin/logs/expired.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pwd = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

for i in range(0, 10):
    while i <= 10:
        try:
            conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)
        except pyodbc.Error as f:
            #errorcode_string = str(f).split(None, 1)[0]
            logging.INFO(str(f))
            sleep(60)
            continue
        break
    break
today = datetime.datetime.now()
c = datetime.timedelta(days=1)
d = today - c
yesterday = d.strftime("%Y-%m-%d")
day = today.strftime('%Y-%m-%d')

query = "SELECT T01.HSSVC, T01.PATNO, T01.PNAME, T01.AGE, T01.SEX,\
            T01.ISADATE, T01.IATME, T01.DIAGN,\
            T01.ISDDATE, T01.DTIME, T03.DCSDS, T03.DCEXP, t04.room, t04.bed\
    FROM      HOSPF0062.PATIENTS T01 LEFT OUTER JOIN\
            HOSPF0062.DSSTAT T03\
    ON        T01.DCSTAT = T03.DCUBS\
    LEFT OUTER JOIN hospf0062.patrmbdp t04 ON t01.patno = t04.patn15\
    WHERE ISDDATE = '" + yesterday + "' AND T03.DCEXP = 'Y' AND RECID='Y' \
    ORDER BY  T01.PNAME ASC"

data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
df.to_excel('/home/itadmin/automation/files/expired-report-' + str(day) + '.xlsx', index=False)

    # setup and send emails
msg = MIMEMultipart()
msg['Subject'] = "Expired Report For: " + str(day)
recipients = ['setdud@mckweb.com', 'tracha@mckweb.com', 'AlaBee@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/expired-report-' + str(day) + '.xlsx'

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
print("Expired Report Email Sent On " + day)
