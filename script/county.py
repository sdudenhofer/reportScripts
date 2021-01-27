import pandas as pd
import pyodbc
import configparser
import logging
import time
import datetime
from ftplib import FTP
import ftplib
import os
import schedule
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import smtplib
from time import sleep

logging.basicConfig(filename='/home/itadmin/logs/county.log', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pword = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

for i in range(0, 10):
        while i <= 10:
            try:
                conn = pyodbc.connect(DSN=server, UID=user, PWD=pword)
            except pyodbc.Error as f:
                logging.INFO(f)
                sleep(60)
                continue
            break
        break

today = datetime.datetime.now()
day = today.strftime('%Y-%m-%d')
writer = pd.ExcelWriter('/home/itadmin/automation/files/covid-weekly-report-' + str(day) + '.xlsx', engine='openpyxl')
tests = open('/home/itadmin/automation/test.txt', 'r')

for test in tests:
    t = test.rstrip()
    query = "select t02.pname, t02.dob, t06.rdisvfdt, t06.rdrs from \
        orderf0062.oeorder t01 left outer join hospf0062.patients t02 on t01.opat# = t02.patno \
        left outer join orderf0062.oeostat t04 on t01.ostat = t04.stat# \
        left outer join orderf0062.rd t06 on t02.patno = t06.rdpt# \
        where oproc in ('COVID-19', 'COVID-LC', 'COVID-QL', 'COVID-CM', 'COVID-RO') \
        and t02.patno not in ('4106640', '4106643', '4123971', '4123972', '4112283', '4112287',  '4112284')\
        and t06.rdrs != ''\
        and rdpf = '" + t + "'\
        and rdrs not in ('OR PUB.HLTH.LAB', 'LABCORP', 'CALL IF NEG', '        LABCORP', 'CALL IF POS', '*', '              *')\
        ORDER BY t01.osdate"


    data = pd.read_sql(query, conn)
    df = pd.DataFrame(data)
    df.to_excel(writer, sheet_name=test.rstrip(), index=False)
    df['RDRS'].value_counts(sort=True).to_excel(writer, sheet_name=test.rstrip() + " Totals")

writer.save()

msg = MIMEMultipart()
msg['Subject'] = "Covid Weekly Report"
recipients = ['setdud@mckweb.com', 'TraHic@MCKweb.com', 'LauKey@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/covid-weekly-report-' + str(day) + '.xlsx'

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
print("County Corona mail Sent On " + str(day))
