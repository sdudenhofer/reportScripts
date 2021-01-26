import configparser
import pyodbc
import pandas as pd
import datetime
from datetime import datetime
import openpyxl
from datetime import datetime
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
pwd = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']
conn = pyodbc.connect(DSN=server, UID=user, PWD=pwd)

today = datetime.now()
day = today.strftime('%Y-%m-%d')
writer = pd.ExcelWriter('files/Missing-RTOrders-' + str(day) + '.xlsx', engine='openpyxl')

query = """
SELECT
  ALL       T03.NURST, T03.ROOM CONCAT '-' CONCAT T03.BED AS RB, T01.PATNO,
            T01.PNAME, T02.OSDATE, T02.OSTIME, T02.OORD#, T02.OPROC, T04.POVDSC
  FROM      HOSPF062.PATIENTS T01 LEFT OUTER JOIN
            ORDERF062.OEORDER T02
  ON        T01.PATNO = T02.OPAT# LEFT OUTER JOIN
            HOSPF062.RMBED T03
  ON        T01.PATNO = T03.PAT# LEFT OUTER JOIN
            ORDERF062.OEPROC T04
  ON        T02.OPROC = T04.PPROC
  WHERE     T03.NURST IN ('CDU', 'CVIC', 'ICU', 'SCUJ', 'WHBC', 'PCU', 'CCU',
            'MCU')
    AND     T02.OPROC IN ('ABGCAPI', 'ABGCORDV', 'AP', 'ABGCORDA', 'ABGELGLU',
            'RSV', 'FLU/RVP', 'RESPDFA', 'BORDPER', 'BPRTAB', 'ABGERPNL',
            'ABGORPNL''ABGO2SAT', 'ABGVENBL', 'CULRSP', 'CARBONM', 'FLU/PCR',
            'FLUABAB', 'FLUABAG', 'FLUEABPCR', 'HINFBIGG', 'INFABDFA', 'FLUPCR'
,           'INFLUAB', 'INFLUWOP', 'RVPPCR', 'SPECCOLL', 'RVP PCR', 'FLUPCRAB',
            'INFLUPCR', 'RVPPCREX', 'CULRESP', 'CULLRES''CULAFBPL', 'RVP-PCR',
            'RVPPCREX', 'COVID-19', 'COVID-QL', 'COVID-LC', 'COVID-CM', 'COVID-RO')
    AND     T02.OSTAT = 3
    AND     T02.OINIT <> 'Y'
  ORDER BY  T03.NURST ASC, RB ASC
"""
data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
df.to_excel(writer, sheet_name='Missing RT Orders', index=False)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = "Missing RT Orders Report: " + str(day)
recipients = ['setdud@mckweb.com', 'MWMC.House.Coordinators@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = 'files/Missing-RTOrders-' + str(day) + '.xlsx'

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
