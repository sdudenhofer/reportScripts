import pyodbc
import pandas as pd
import configparser
import datetime
import openpyxl
import time
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import os
import smtplib
import logging

logging.basicConfig(filename="/home/itadmin/logs/ccucensus.log", format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')
server = config['MHD']['ODBC']
user = config['MHD']['USER']
pword = config['MHD']['PASS']
eserv = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']

try:
    conn = pyodbc.connect(DSN=server, UID=user, PWD=pword)
except pyodbc.Error as e:
    logging.INFO(e)
    sleep(60)

query = "select t01.room, t01.bed, t01.pat#, t02.pname, t02.diagn  from hospf0062.rmbed t01 \
LEFT OUTER JOIN hospf0062.patients t02 on t01.pat# = t02.patno \
where t01.nurst = 'CCU' and t01.pat# > '0' order by t01.room, t01.bed"

cursor = conn.cursor()
data = cursor.execute(query)
output = open('/home/itadmin/automation/files/pt_ccu_census.csv', 'w+')
output.write("Room| Bed| Patient Number| Patient Name| Diagnosis\n")
for row in data:
    dout = str(row[0]) + "|" + str(row[1]) + "|" + str(row[2]) + "|" + str(row[3]) + "|" + str(row[4]) + "\n"
    empty = "NaN| NaN| NaN| NaN\n"
    output.write(dout)
    output.write(empty)
    output.write(empty)
output.close()

today = datetime.datetime.now()
day = today.strftime('%Y-%m-%d')

data = pd.read_csv('/home/itadmin/automation/files/pt_ccu_census.csv', sep="|")
d1 = data.replace("NaN", "", regex=True)
#df = pd.DataFrame(d1)
writer = pd.ExcelWriter('/home/itadmin/automation/files/pt-report-' + str(day) + '.xlsx', engine='openpyxl')
d1.to_excel(writer, sheet_name='CCU Census', index=False)
writer.save()

msg = MIMEMultipart()
msg['Subject'] = "CCU Census: " + str(day)
recipients = ['setdud@mckweb.com', 'JudBre@MCKweb.com', 'KevWhe@MCKweb.com', 'RenRue@MCKweb.com', 'CarWer@MCKweb.com', 'StaGre@MCKweb.com', 'LisBoy@MCKweb.com', 'AdaMar@mckweb.com', 'KenTau@MCKweb.com', 'AngLap@MCKweb.com', 'AnnRee@mckweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = '/home/itadmin/automation/files/pt-report-' + str(day) + '.xlsx'

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

