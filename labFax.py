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
import logging
import time

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
    print(e)
    sleep(60)

today = datetime.datetime.now()
day = today.strftime('%Y-%m-%d')
writer = pd.ExcelWriter('/home/itadmin/automation/files/lab-fax-' + str(day) + '.xlsx', engine='openpyxl')

query = """
SELECT t01.flpt#, t02.pname, t01.flor#, t04.povdsc, t01.flrcpt, t01.flsphne, t01.flstat, DATE(t01.flidt) as "Date", t01.flitm
FROM hospf062.rrfaxlog t01 
LEFT OUTER JOIN hospf062.patients t02 on t01.flpt# = t02.patno
LEFT OUTER JOIN orderf062.oeorder t03 on t01.flpt# = t03.opat# and t01.flor# = t03.oord#
LEFT OUTER JOIN orderf062.oeproc t04 on t03.oproc = t04.pproc
WHERE t01.flidt = CURRENT DATE-1 DAYS 
and t03.otodpt = 'LAB' order by t01.flidt
"""

data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
df['FLSTAT'].value_counts(sort=True).to_excel(writer, sheet_name="Total by Status")
df.loc[df['FLSTAT'] == 'S'].to_excel(writer, sheet_name='Sent Status')
df.loc[df['FLSTAT'] == 'U'].to_excel(writer, sheet_name='Unsent Status')
df.to_excel(writer, sheet_name="Patient Data")

writer.save()

filename = '/home/itadmin/automation/files/lab-fax-' + str(day) + '.xlsx'
msg = MIMEMultipart()
msg['Subject'] = 'Daily Lab Fax Report'
recipients = ['setdud@mckweb.com', 'TraHic@MCKweb.com', 'MelHub@MCKweb.com',
                'KarKla@MCKweb.com', 'LauKey@mckweb.com', 'SarCav@MCKweb.com']
for to in recipients:
        msg['To'] = to

attachment = MIMEBase('application','octet-stream')
    
msg.attach(MIMEText('Report Attached'))
attachment.set_payload(open(filename, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(filename))
msg.attach(attachment)
s = smtplib.SMTP(eserv)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', msg['To'], msg.as_string())
s.quit()
    #emailusers.emailusers(, 'Lab Fax Report', filename, 'webadmin@mckweb.com')
print('Emailed: ' + str(day))
