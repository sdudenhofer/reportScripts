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

config = configparser.ConfigParser()
config.read('config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pwd = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']


#connect to database
conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)


doctors = open('rapc-doctors.txt', 'r')
today = datetime.datetime.now()
c = datetime.timedelta(days=32)
d = today - c
yesterday = d.strftime("%Y-%m-%d")
day = today.strftime('%Y-%m-%d')

writer = pd.ExcelWriter('files/RAPC-phys-doc-' + str(day) + '.xlsx', engine='openpyxl')
   
for doc in doctors:
    d = doc.rstrip() + "%"
    query = "SELECT \
            T01.HSSVC, T01.HSTNUM, T02.ENCID, T01.PNAME, \
            T01.ISDOB, T02.TITL, T02.CREATEDT, T02.CRTDNAME \
            FROM HOSPF0062.PATIENTS T01 LEFT OUTER JOIN \
            HOSPF0062.CDNOTETB T02 \
            ON T02.ENCID = T01.PATNO LEFT OUTER JOIN \
            HOSPF0062.CDNTEATB T03 \
            ON T02.ENCID = T03.ENCTRID \
            AND T02.CREATEBY = T03.LSTMODBY \
            AND T03.ENCTRID = T01.PATNO \
            WHERE T02.CREATEDT BETWEEN '" + yesterday + "' \
            AND '" + day + "' \
            AND T02.CRTDNAME LIKE '" + d + "'"
    data = pd.read_sql(query, conn)
    df = pd.DataFrame(data)
    df.to_excel(writer, sheet_name=doc.rstrip(), index=False)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = "(secure) Monthly Report: " + str(day)
recipients = ['setdud@mckweb.com', 'dchaltraw@oregonimaging.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = 'files/RAPC-phys-doc-' + str(day) + '.xlsx'

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
print("Email Sent to: " + str(recipients))


