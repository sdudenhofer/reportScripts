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
import os
from time import sleep
import logging



config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    (e)
    sleep(60)




today = datetime.datetime.now()
day = today.strftime('%Y-%m-%d')

query = "SELECT T01.NURST, T01.ROOM,  T01.BED, T01.PAT#, \
            T02.PNAME, T03.TRRECVAL, T03.TRRECDT \
    FROM      HOSPF062.RMBED T01 INNER JOIN \
            HOSPF062.PATIENTS T02 \
    ON        T01.PAT# = T02.PATNO INNER JOIN \
            ORDERF062.NCTRN T03 \
    ON        T03.TRPAT# = T01.PAT# INNER JOIN \
            ORDERF062.NCPRM T04 \
    ON        T03.TRPRMID = T04.PRID \
    WHERE     T01.NURST IN ('CDU', 'CVIC', 'ICU', 'SCUJ', 'WHBC', 'PCU', 'CCU', \
            'MCU', 'NUR') \
    AND     T03.trprmid = 'Q0000005455' \
    AND     T04.PRSTS = 'A' \
    AND 	T03.TRRECDT >= '" + day + " 00:00:00' \
    AND 	T03.TRRECVAL not in ('Standard Isolation - Universal Precautions') \
    ORDER BY  T01.NURST ASC, t01.room, t01.bed ASC, t03.trrecdt asc"

query2 = """
select t01.nurst, t01.room, t01.bed, t02.dthpatno, t03.pname, t02.dthresp, t02.dthdttm
FROM hospf062.rmbed t01 LEFT OUTER JOIN hospf062.chpdtaph t02 on t01.pat# = t02.dthpatno
LEFT OUTER JOIN hospf062.patients t03 on t02.dthpatno = t03.patno
WHERE t02.dthresp = 'Isolation' and t01.nurst != '' order by t01.nurst, t01.room, t01.bed
"""
data = pd.read_sql(query, conn)
dataframe = pd.DataFrame(data)
writer = pd.ExcelWriter('/home/itadmin/automation/files/Nutrition-isolation-' + str(day) + '.xls', engine='openpyxl')
data2 = pd.read_sql(query2, conn)
dataframe2 = pd.DataFrame(data2)
dataframe.to_excel(writer, index=False, sheet_name='All Patients')
dataframe2.to_excel(writer, index=False, sheet_name='CHP')
writer.save()
conn.close()

msg = MIMEMultipart()
msg['Subject'] = 'Nutrition Isolation Report'

# recipients to send the email to
recipients = ['setdud@mckweb.com', 'rhocor@mckweb.com', 'krigre@mckweb.com', 'LolPut@MCKweb.com', 'SeaEgg@MCKweb.com', 'KarBoo@MCKweb.com', 'KatGat@MCKweb.com', 'CarCam@MCKweb.com', 'JolFeg@MCKweb.com', 'AmaSlo@MCKweb.com', 'ShaKna@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/Nutrition-isolation-' + str(day) + '.xls'

msg.attach(MIMEText("Isolation Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()


