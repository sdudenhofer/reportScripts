import pandas as pd
import openpyxl
import pyodbc
import configparser
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from smtplib import SMTP, SMTPException, SMTPAuthenticationError
import smtplib
import os
import time
import logging
import schedule

# configure config parser
config = configparser.ConfigParser()
config.read('/home/itadmin/automation/config.ini')

server = config['AS400']['ODBC']
user = config['AS400']['USER']
password = config['AS400']['PASS']

email = config['OUTLOOK']['SERVER']
euser = config['OUTLOOK']['USER']
epass = config['OUTLOOK']['PASS']


# dates for the report
today = datetime.datetime.now()
a = datetime.timedelta(days=62)
b = today - a
c = datetime.timedelta(days=1)
d = today - c
yesterday = d.strftime("%Y-%m-%d")
month = b.strftime("%Y-%m-%d")

#configure logging information
logging.basicConfig(filename='/home/itadmin/logs/radrev.log', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)

# create first query
query = "SELECT\
  ALL       ACC.ISPDATEA, ACC.DEPT, GLK.GLDESC, ACC.PATNO, PAT.PNAME, PAT.HSSVC\
,           ACC.ISDATE, ACC.SVCCD, CHG.DESC, (ACC.QTY), (ACC.QTY),\
            (ACC.AMT1)\
  FROM      HOSPF062.ACCUMCHG ACC LEFT OUTER JOIN\
            HOSPF062.PATIENTS PAT\
  ON        ACC.PATNO = PAT.PATNO LEFT OUTER JOIN\
            HOSPF062.CHRGDESC CHG\
  ON        ACC.SVCCD = CHG.SVCCD LEFT OUTER JOIN\
            HOSPF062.GLKEYSM GLK\
  ON        ACC.DEPT = GLK.DEPTG\
  WHERE     ACC.ISPDATEA BETWEEN '" + month + "' AND\
            '" + yesterday + "'\
    AND     ACC.DEPT IN (161, 162, 163, 164, 281)\
    AND     PAT.RECTYP <> 'I'\
    AND     ACC.SVCCD NOT IN (2811277, 2811241, 2811255, 2811167, 2811256,\
            2811238, 2811243)\
  ORDER BY  006 ASC, 001 ASC, 005 ASC, 009 ASC"


conn = pyodbc.connect(DSN=server, UID=user, PWD=password)


#create second query
query2 = "SELECT\
  ALL       ACC.ISPDATEA, ACC.DEPT, GLK.GLDESC, ACC.PATNO, PAT.PNAME, PAT.HSSVC\
,           ACC.ISDATE, ACC.SVCCD, CHG.DESC, (ACC.QTY), (ACC.QTY),\
            (ACC.AMT1)\
  FROM      HOSPF062.ACCUMCHG ACC LEFT OUTER JOIN\
            HOSPF062.PATIENTS PAT\
  ON        ACC.PATNO = PAT.PATNO LEFT OUTER JOIN\
            HOSPF062.CHRGDESC CHG\
  ON        ACC.SVCCD = CHG.SVCCD LEFT OUTER JOIN\
            HOSPF062.GLKEYSM GLK\
  ON        ACC.DEPT = GLK.DEPTG\
  WHERE     ACC.ISPDATEA BETWEEN '" + month + "' AND\
            '" + yesterday + "'\
    AND     ACC.DEPT IN (161, 162, 163, 164, 281)\
    AND     PAT.RECTYP = 'I'\
    AND     ACC.SVCCD NOT IN (2811277, 2811241, 2811255, 2811167, 2811256,\
            2811238, 2811243)\
  ORDER BY  006 ASC, 001 ASC, 005 ASC, 009 ASC"

# Get data from Database
data1 = pd.read_sql(query, conn)
data2 = pd.read_sql(query2, conn)

#write to dataframe
df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)
writer = pd.ExcelWriter('Radiology-Revenue.xlsx', engine='openpyxl')

#convert data to excel
df1.to_excel(writer, sheet_name='Inpatient', index=False)
df2.to_excel(writer, sheet_name='Outpatient', index=False)
#save excel document
writer.save()
#create/send email
msg = MIMEMultipart()
msg['Subject'] = "IP/OP Rad Rev Report: " + str(yesterday)
recipients = ['setdud@mckweb.com', 'LisRam@mckweb.com']
for to in recipients:
    msg['To'] = to

attachment = MIMEBase('application','octet-stream')
f = 'Radiology-Revenue.xlsx'

msg.attach(MIMEText("Report Attached"))
attachment.set_payload(open(f, 'rb').read())
encoders.encode_base64(attachment)
attachment.add_header('Content-Disposition', 'attachment', filename = os.path.basename(f))
msg.attach(attachment)

# error catching for sending emails wait 2 minutes and try again
s = smtplib.SMTP(email)
#s.starttls()
#s.login(euser, epass)
s.sendmail('webadmin@mckweb.com', recipients, msg.as_string())
  #  c = 6
  #except SMTPAuthenticationError as f:
  #  logging.info('Email Authentication error: ' + f)
  #  time.sleep(120)
  #  c = c + 1
  #except:
  #  logging.info('Unknown email error')
  #  time.sleep(120)
  #  c = c + 1
s.quit()
logging.info("Email Sent to: " + str(recipients))

