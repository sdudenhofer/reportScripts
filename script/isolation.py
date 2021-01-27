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
import logging

logging.basicConfig(filename="/home/itadmin/logs/isolation.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
password = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
except pyodbc.Error as e:
    logging.INFO(e)
    sleep(60)

today = datetime.datetime.now()

query = """
SELECT
    ALL       T01.NURST, T01.ROOM,  T01.BED, T01.PAT#,
            T02.PNAME, T03.TRRECVAL, T03.TRRECDT,
            T03.TRRECDT
    FROM      HOSPF0062.RMBED T01 INNER JOIN
            HOSPF0062.PATIENTS T02
    ON        T01.PAT# = T02.PATNO INNER JOIN
            ORDERF0062.NCTRN T03
    ON        T03.TRPAT# = T01.PAT# INNER JOIN
            ORDERF0062.NCPRM T04
    ON        T03.TRPRMID = T04.PRID
    WHERE     T01.NURST IN ('CDU', 'CVIC', 'ICU', 'SCUJ', 'WHBC', 'PCU', 'CCU',
            'MCU', 'NUR')
    AND     T03.TRPRMID = 'Q0000005455'
    AND     T04.PRSTS = 'A'
    AND     T03.TRRECVAL <> 'Standard Isolation - Universal Precautions'
    AND     DATE(T03.TRRECDT) = CURRENT DATE-1 DAYS
    AND     TIME(T03.TRRECDT) > '16:00:00'
    ORDER BY  T01.NURST ASC, t01.room, t01.bed ASC
    """

query2 = """
select t01.nurst, t01.room, t01.bed, t02.dthpatno, t03.pname, t02.dthresp, t02.dthdttm
FROM hospf0062.rmbed t01 LEFT OUTER JOIN hospf0062.chpdtaph t02 on t01.pat# = t02.dthpatno
LEFT OUTER JOIN hospf0062.patients t03 on t02.dthpatno = t03.patno
WHERE t02.dthresp = 'Isolation' and t01.nurst != '' order by t01.nurst, t01.room, t01.bed
"""

writer = pd.ExcelWriter('/home/itadmin/automation/files/isolation.xlsx', engine='openpyxl')

data = pd.read_sql(query, conn)
data1 = pd.read_sql(query2, conn)
dataframe = pd.DataFrame(data)
dataframe1 = pd.DataFrame(data1)
dataframe.to_excel(writer, sheet_name='Nurse Orders', index=False)
dataframe1.to_excel(writer, sheet_name='CHP Locations', index=False)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = 'Isolation Daily Report'

# recipients to send the email to
recipients = ['setdud@mckweb.com', 'MWMC.DEPTMGRS@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/isolation.xlsx'

msg.attach(MIMEText("Isolation Daily Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
s.quit()
print("Isolation Email Sent: " + str(today))
