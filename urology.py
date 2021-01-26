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
config.read('/home/itadmin/automation/config.ini')
server = config['AS400']['ODBC']
user = config['AS400']['USER']
pword = config['AS400']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

conn = pyodbc.connect(DSN=server, UID=user, pwd=pword)

doctors = open('/home/itadmin/automation/udoctor.txt', 'r')
today = datetime.datetime.now()
c = datetime.timedelta(days=7)
d = today - c
yesterday = d.strftime("%Y-%m-%d")
day = today.strftime('%Y-%m-%d')

writer = pd.ExcelWriter('/home/itadmin/automation/files/urology-report-' + str(day) + '.xlsx', engine='openpyxl')
d_out = []
out_array = open('/home/itadmin/automation/files/urology.csv', 'w+')
out_array.write("MRN, Account Number, Patient Name, Physician, Document, Document Date \n")
for doc in doctors:
    d = doc.rstrip() + "%"
    query = "SELECT t01.patient_id, t01.encounter_id, t02.pname, \
        t01.createdby_name, t01.title, t01.created_date \
        from hospf062.cdnotetb t01 left outer join hospf062.patients t02 \
        on t01.encounter_id =t02.patno and t01.patient_id = t02.hstnum \
        where t01.createdt BETWEEN '" + str(yesterday) + "' \
        AND '" + str(day) + "' AND T01.CRTDNAME LIKE '" + d + "' order by t02.pname"
    cursor = conn.cursor()
    data = cursor.execute(query)
    for row in data:
        out = str(row[0]) + ", " + str(row[1]) + ", " + str(row[2]) + ", " + str(row[3]) \
            + ", " + str(row[4]) + ", " + str(row[5]) + "\n"
        d_out.append(out)


tran_query = "SELECT t01.dhhstno, t01.dhpatno, t03.pname, t02.phname, t01.dhfldr, t01.dhdate \
        FROM hospf062.trdochp t01 LEFT OUTER JOIN hospf062.phymast t02 on t01.dhdrno \
            = t02.nwdrnum LEFT OUTER JOIN hospf062.patients t03 on t01.dhhstno = t03.hstnum \
            and t01.dhpatno = t03.patno \
            WHERE t01.dhdate between '" + str(yesterday) + "' and '" + str(day) + "' and \
                dhdrno in (485, 1765, 481, 3349, 1927, 1002, 961, 325, 1590, 1650, 1523, \
                    850, 4090, 980, 346, 1436, 2935, 230, 330, 434, 2261, 701, 986, 1074)"
cur = conn.cursor()
data_tran = cur.execute(tran_query)
for r1 in data_tran:
    out1 = str(r1[0]) + "," + str(r1[1]) + "," + str(r1[2]) + ", " + str(r1[3]) \
        + ", " + str(r1[4]) + ", " + str(r1[5]) + "\n"
    d_out.append(out1)

for r in d_out:
    out_array.write(str(r))

out_array.close()
df1 = pd.read_csv('/home/itadmin/automation/files/urology.csv', sep=',')
df1.to_excel(writer, sheet_name="Urology List", index=False)

writer.save()
writer.close()    
    # setup and send emails
msg = MIMEMultipart()
msg['Subject'] = "Urology Report For: " + str(day)
recipients = ['setdud@mckweb.com', 'scosmi@mckweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/urology-report-' + str(day) + '.xlsx'

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
print("Urology Email Sent On " + day)
