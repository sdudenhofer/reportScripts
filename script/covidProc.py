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

logging.basicConfig(filename='/home/itadmin/logs/covidProc.log', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
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
            logging.info(f)
            sleep(60)
            continue
        break
    break

today = datetime.datetime.now()
query = "select t01.oproc, t05.hssvc, t02.patno, t02.pname, t03.room, t03.bed, t03.nurst, t01.osdate, \
t04.stdsc, t01.oisrdate from \
orderf0)62.oeorder t01 left outer join hospf0062.patients t02 on t01.opat# = t02.patno \
left outer join hospf0062.rmbed t03 on t01.opat# = t03.pat# and t02.patno = t03.pat# \
left outer join orderf0062.oeostat t04 on t01.ostat = t04.stat# \
left outer join hospf0062.patients t05 on t01.opat# = t05.patno \
where oproc in ('COVID-19', 'COVID-LC', 'COVID-QL', 'COVID-CM', 'COVID-RO') \
and t02.patno != '4106640' and t02.patno != '4106643' ORDER BY t01.osdate"

data = pd.read_sql(query, conn)
dataframe = pd.DataFrame(data)
dataframe.to_excel('/home/itadmin/automation/files/cov-email-' + str(today) + '.xlsx', index=False)

msg = MIMEMultipart()
msg['Subject'] = 'Covid-19 Data Report'

# recipients to send the email to
recipients = ['setdud@mckweb.com', 'laukey@MCKweb.com', 'tanpar@mckweb.com', 'DesShu@MCKweb.com', 'TraHic@MCKweb.com', 'DavElg@MCKweb.com', 'MerNel@MCKweb.com', 'JulCav@MCKweb.com', 'JanWat@MCKweb.com', 'JamAbe@MCKweb.com', 'AdaLor@MCKweb.com', 'ChrGri@MCKweb.com', 'MWMC.House.Coordinators@MCKweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/cov-email-' + str(today) + '.xlsx'

msg.attach(MIMEText("COVID19 Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
for x in range(0,10):
    while x <= 10:
        try:
            s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
        except:
            logging.INFO("Unable to send email" + str(today))
            print("Unknown Error. Sleeping...")
            sleep(60)
            continue
        break
    break

s.quit()
print("Email Sent: " + str(today))
