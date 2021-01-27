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
import schedule
import time
import logging

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')

logging.basicConfig(filename="/home/itadmin/logs/infectionPrevention.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

server = config['MHD']['ODBC']
user = config['MHD']['USER']
password = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

for i in range(0, 10):
    while i <= 10:
        try:
            conn = pyodbc.connect(DSN=server, UID=user, PWD=password)
        except pyodbc.Error as f:
            logging.INFO(str(f))
            sleep(60)
            continue
        break
    break

query_foley = "SELECT T01.NURST, T01.ROOM CONCAT '-' CONCAT T01.BED AS RB, T04.ISADATE, T01.PAT#, T04.PNAME, T02.TRRECDT, T02.TRRECVAL, T02.TRCOM \
FROM HOSPF0062.RMBED T01 LEFT OUTER JOIN ORDERF0062.NCTRN T02 ON T01.PAT# = T02.TRPAT# \
LEFT OUTER JOIN ORDERF0062.NCPRM T03 ON T02.TRPRMID = T03.PRID \
LEFT OUTER JOIN HOSPF0062.PATIENTS T04 ON T01.PAT# = T04.PATNO \
WHERE T01.NURST IN ('CDU', 'CVIC', 'ICU', 'SCUJ', 'WHBC', 'PCU', 'MCU', 'CCU') \
AND T02.TRPRMID IN ('Q0000000382', 'Q0000004576') AND T03.PRSTS = 'A' \
AND T02.TRRECVAL NOT IN ('Voiding', 'Incontinent', 'Other-See Comment', 'Straight Catheter') \
AND T02.TRRECDT BETWEEN CURRENT DATE-1 DAYS AND CURRENT DATE \
ORDER BY  T01.NURST ASC, T02.TRRECDT ASC"

query_cent = "SELECT T01.NURST, T01.ROOM CONCAT '-' CONCAT T01.BED AS RB, T04.ISADATE, T01.PAT#, T04.PNAME, T02.TRRECDT, T03.PRVAL, T02.TRRECVAL, T02.TRCOM \
FROM HOSPF0062.RMBED T01 LEFT OUTER JOIN ORDERF0062.NCTRN T02 ON T01.PAT# = T02.TRPAT# \
LEFT OUTER JOIN ORDERF0062.NCPRM T03 ON T02.TRPRMID = T03.PRID \
LEFT OUTER JOIN HOSPF0062.PATIENTS T04 ON T01.PAT# = T04.PATNO \
WHERE T01.NURST IN ('SCUJ', 'WHBC', 'PCU', 'MCU', 'CCU', 'CDU') \
AND T02.TRPRMID IN ('Q0000000163', 'Q0000004054', 'Q0000000168', 'Q0000000169', 'Q0000000667', 'Q0000000672', 'Q0000004013', 'Q0000004035', 'Q0000004027', 'Q0000003057') \
AND T04.ISADATE <> '0001-01-01' AND T03.PRSTS = 'A' \
AND T02.TRRECVAL NOT IN ('Peripheral;Saline Lock', 'Peripheral;Double Lumen', 'Peripheral', ' ', 'Peripheral;Saline Lock;Single Lumen', 'Peripheral;Single Lumen', 'Peripheral;Saline Lock;Double Lumen', 'Saline Lock') \
AND T04.ISDDATE = '0001-01-01' \
AND T02.TRRECDT BETWEEN CURRENT DATE-1 DAYS AND CURRENT DATE \
OR T01.NURST IN ('MCU', 'CCU', 'PCU', 'SCUJ', 'WHBC') \
AND T03.PRSTS = 'A' AND T04.ISADATE <> '0001-01-01' \
AND T02.TRPRMID IN ('Q0000004066', 'Q0000004049', 'Q0000003057', 'Q0000004030') \
AND T04.ISDDATE = '0001-01-01' AND T02.TRRECDT BETWEEN CURRENT DATE-1 DAYS AND CURRENT DATE \
ORDER BY T01.NURST ASC, T01.PAT# ASC, T02.TRRECDT ASC"

query_iso ="SELECT T01.NURST, T01.ROOM CONCAT '-' CONCAT T01.BED AS RMBED, T01.PAT#, T02.PNAME, T03.TRRECVAL, T03.TRRECDT \
FROM HOSPF0062.RMBED T01 INNER JOIN HOSPF0062.PATIENTS T02 ON T01.PAT# = T02.PATNO \
INNER JOIN ORDERF0062.NCTRN T03 ON T03.TRPAT# = T01.PAT# \
INNER JOIN ORDERF0062.NCPRM T04 ON T03.TRPRMID = T04.PRID \
WHERE T01.NURST IN ('CDU', 'CVIC', 'ICU', 'SCUJ', 'WHBC', 'PCU', 'CCU', 'MCU') \
AND T03.TRPRMID = 'Q0000005455' AND T04.PRSTS = 'A' \
AND T03.TRRECVAL <> 'Standard Isolation - Universal Precautions' \
AND DATE(T03.TRRECDT) = CURRENT DATE-1 DAYS \
AND TIME(T03.TRRECDT) > '16:00:00' ORDER BY T01.NURST ASC, RMBED ASC"

query_iso1 = """
select t01.nurst, t01.room, t01.bed, t02.dthpatno, t03.pname, t02.dthresp, t02.dthdttm
FROM hospf0062.rmbed t01 LEFT OUTER JOIN hospf0062.chpdtaph t02 on t01.pat# = t02.dthpatno
LEFT OUTER JOIN hospf0062.patients t03 on t02.dthpatno = t03.patno
WHERE t02.dthresp = 'Isolation' and t01.nurst != ''
"""
query_mdro = "SELECT T01.PATNO, T01.ISDDATE, T01.HSSVC, T02.CIINFCTN, T02.CIIDDTE, T02.CIRECDTE, T02.CIRESDTE, T02.CIRESBY, T02.CISTSCD, T02.CICMMT \
FROM HOSPF0062.PATIENTS T01 INNER JOIN HOSPF0062.CHPDRIP T02 ON T01.PATNO = T02.CIPAT# \
WHERE T01.ISDDATE = DATE(DAYS(CURRENT DATE)-1) AND T01.HSSVC IN ('MIP', 'SIP', 'ICU', 'CCU', 'OBS', 'BBN', ' NUR', 'OBI', 'GYN') "

query_vaccine = "SELECT T01.NURST, T01.ROOM, T01.BED, T02.PATNO, T02.PNAME, T02.AGE, T03.TRRECDT, T04.PRDS, T03.TRRECVAL, T03.TRUSR \
FROM HOSPF0062.RMBED T01 LEFT OUTER JOIN HOSPF0062.PATIENTS T02 ON T01.PAT# = T02.PATNO \
LEFT OUTER JOIN ORDERF0062.NCTRN T03 ON T01.PAT# = T03.TRPAT# \
LEFT OUTER JOIN ORDERF0062.NCPRM T04 ON T03.TRPRMID = T04.PRID \
WHERE T01.NURST IN ('CDU', 'CVIC', 'ICU', 'SCUJ', 'WHBC', 'PCU', 'CCU', 'MCU') \
AND T03.TRPRMID IN ('Q0000000124', 'Q0000000113') AND T04.PRSTS = 'A' \
AND SUBSTR(T03.TRRECVAL,1,3) IN ('No ', ' ', 'Una', 'Unr') \
AND DATE(T03.TRRECDT) = CURRENT DATE-1 DAYS \
ORDER BY  T01.NURST ASC, T01.ROOM ASC, T01.BED ASC, T04.PRDS ASC "

today = datetime.datetime.now()
a = datetime.timedelta(days=1)
b = today - a
census_date = b.strftime("%Y-%m-%d")
writer = pd.ExcelWriter('/home/itadmin/automation/files/infpre-reports-' + str(census_date) + '.xlsx', engine='openpyxl')

data0 = pd.read_sql(query_foley, conn)
data1 = pd.read_sql(query_cent, conn)
data2 = pd.read_sql(query_iso, conn)
data3 = pd.read_sql(query_mdro, conn)
data4 = pd.read_sql(query_vaccine, conn)
data5 = pd.read_sql(query_iso1, conn)
df0 = pd.DataFrame(data0)
df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)
df3 = pd.DataFrame(data3)
df4 = pd.DataFrame(data4)
df5 = pd.DataFrame(data5)
df0.to_excel(writer, sheet_name='Foley', index=False)
df1.to_excel(writer, sheet_name='Central Lines', index=False)
df2.to_excel(writer, sheet_name='Isolation', index=False)
df5.to_excel(writer, sheet_name='Isolation-CHP', index=False)
df3.to_excel(writer, sheet_name="MDRO", index=False)
df4.to_excel(writer, sheet_name="Vaccine", index=False)

writer.save()

msg = MIMEMultipart()
msg['Subject'] = 'Infection Prevention Daily Report'

# recipients to send the email to
recipients = ['setdud@mckweb.com', 'RacGei@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/infpre-reports-' + str(census_date) + '.xlsx'

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
print("Infection Prevention Email Sent: " + str(today))
